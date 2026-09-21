"""Renderiza el estado a un .xlsx respetando la plantilla PRIS."""

from __future__ import annotations

import shutil
import zipfile
from datetime import date, time
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from .dates import (
    base_data_row_count,
    effective_entry,
    effective_exit,
    first_data_row,
    last_base_data_row,
    month_total_amount,
    month_total_hours,
    trackable_days,
)
from .models import Day, State
from .state import EXPORTS_DIR, TEMPLATE_FILE, load_state

DATA_START_ROW = first_data_row()              # 10
BASE_LAST_ROW = last_base_data_row()            # 31
BASE_TOTALS_ROW = 37
BASE_RATE_ROW = 38
BASE_AMOUNT_ROW = 39
BASE_SIG_ROW = 43
BASE_LAST_DAY_ROW = 43
BASE_SUP_NAME_ROW = 47
BASE_SUP_DATE_ROW = 47

HHMM_FMT = "HH:MM"
HOURS_FMT = "0.00"


def _set_cell(ws: Worksheet, row: int, col: int, value: Any) -> None:
    try:
        ws.cell(row=row, column=col).value = value
    except AttributeError:
        pass  # MergedCell: skip


def _set_cell_with_format(ws: Worksheet, row: int, col: int, value: Any, fmt: str) -> None:
    c = ws.cell(row=row, column=col)
    try:
        c.value = value
    except AttributeError:
        return  # MergedCell: skip
    c.number_format = fmt


def _has_media(path: Path) -> bool:
    try:
        with zipfile.ZipFile(path) as zf:
            return any(n.startswith("xl/media/") for n in zf.namelist())
    except (zipfile.BadZipFile, FileNotFoundError):
        return False


def _restore_media_via_zip(out_path: Path, template_path: Path) -> bool:
    """Plan B: si openpyxl perdio las imagenes, las re-inyecta copiando
    xl/media/ y xl/drawings/_rels/*.rels desde la plantilla original."""
    if not out_path.exists() or not template_path.exists():
        return False
    tmp = out_path.with_suffix(out_path.suffix + ".planb.tmp")
    with zipfile.ZipFile(template_path) as zt, zipfile.ZipFile(
        out_path, "r"
    ) as zo, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zw:
        out_names = set(zo.namelist())
        tpl_names = set(zt.namelist())
        for name in out_names:
            zw.writestr(name, zo.read(name))
        for name in tpl_names:
            if name.startswith("xl/media/") and name not in out_names:
                zw.writestr(name, zt.read(name))
            elif name.startswith("xl/drawings/_rels/") and name not in out_names:
                zw.writestr(name, zt.read(name))
    shutil.move(str(tmp), str(out_path))
    return _has_media(out_path)


def _shifted(inserted: int) -> dict[str, int]:
    return {
        "totals": BASE_TOTALS_ROW + inserted,
        "rate": BASE_RATE_ROW + inserted,
        "amount": BASE_AMOUNT_ROW + inserted,
        "sig": BASE_SIG_ROW + inserted,
        "last_day": BASE_LAST_DAY_ROW + inserted,
        "sup_name": BASE_SUP_NAME_ROW + inserted,
        "sup_date": BASE_SUP_DATE_ROW + inserted,
        "last_data": BASE_LAST_ROW + inserted,
    }


def _write_data_rows(
    ws: Worksheet, days_with_data: list[str], state: State, inserted: int
) -> None:
    """Escribe una fila por cada dia con entradas (en orden cronologico).

    Dias sin entradas se IGNORAN: se limpian las celdas A..F de la fila
    para que el Excel muestre solo los dias con data, aunque las filas
    del template siguen existiendo.
    """
    n = len(days_with_data)
    last_written_row = DATA_START_ROW + n - 1 + inserted
    first_blank_row = last_written_row + 1
    last_template_row = BASE_LAST_ROW + inserted

    for idx, iso in enumerate(days_with_data):
        row = DATA_START_ROW + idx
        day: Day | None = state.days.get(iso)
        d = date.fromisoformat(iso)
        _set_cell(ws, row, 1, d)  # A: fecha, formato heredado del template
        if day and day.items:
            joined = " ".join(it.description for it in day.items)
            _set_cell(ws, row, 2, joined)  # B (mergeada con C)
        else:
            _set_cell(ws, row, 2, None)
        entry = effective_entry(day) if day and day.items else (day.entry if day else None)
        exit_ = effective_exit(day) if day and day.items else (day.exit if day else None)
        _set_cell_with_format(ws, row, 4, entry, HHMM_FMT)
        _set_cell_with_format(ws, row, 5, exit_, HHMM_FMT)
        if day and day.items:
            total = round(sum(it.hours for it in day.items), 2)
        else:
            total = 0.0
        _set_cell_with_format(ws, row, 6, total, HOURS_FMT)

    # Limpiar las filas no usadas del template para que el Excel no muestre
    # las fechas/horas/formulas pre-existentes del mes anterior. Las celdas
    # que son parte de un merge (MergedCell) son read-only: solo la celda
    # top-left es escribible, el resto queda implicitamente vacia.
    for row in range(first_blank_row, last_template_row + 1):
        for col in range(1, 7):  # A..F
            cell = ws.cell(row=row, column=col)
            try:
                cell.value = None
            except AttributeError:
                # MergedCell: skip (la celda top-left del merge ya fue limpiada)
                pass


def _write_header_and_totals(
    ws: Worksheet,
    state: State,
    year: int,
    month: int,
    n: int,
    inserted: int,
    last_day_iso: str,
) -> None:
    """Header, totales, firma y supervisor.

    Layout unificado basado en `n` (dias con entradas):
    - Datos: filas 10 .. 10+n-1
    - 5 filas de separacion (del template)
    - Totales: fila n+15
    - Tarifa: fila n+16
    - Monto: fila n+17
    - Firma + ultimo dia: fila n+21
    - Supervisor: fila n+25

    Funciona igual cuando inserted>0 (la insercion corre 1..n-22 filas
    adicionales desde la fila 32, pero la formula n+15 sigue dando la
    fila correcta porque 37 + (n-22) = n+15).
    """
    totals_row = n + 15
    rate_row = n + 16
    amount_row = n + 17
    sig_row = n + 21
    sup_row = n + 25

    _set_cell(ws, 6, 1, state.professional.name)
    _set_cell(ws, 6, 3, state.professional.specialty)
    _set_cell(ws, 6, 5, date(year, month, 1))

    last_data_row = DATA_START_ROW + n - 1
    _set_cell(
        ws, totals_row, 6,
        f"=SUM(F{DATA_START_ROW}:F{last_data_row})",
    )

    _set_cell(ws, rate_row, 6, state.professional.hourly_rate)

    last_day = date.fromisoformat(last_day_iso)
    _set_cell(ws, sig_row, 5, last_day)

    _set_cell(ws, sup_row, 1, state.supervisor.name)
    sup_date = state.supervisor.sup_date
    _set_cell(ws, sup_row, 5, sup_date)


def _patch_sheet1(ws1: Worksheet, n: int) -> None:
    """Parchea las formulas cross-sheet en Hoja 1 segun el layout unificado.

    Las posiciones de referencia (totales, tarifa, firma, supervisor) se
    calculan desde `n`, asi el parche es valido para cualquier n con o
    sin filas insertadas.
    """
    totals_row = n + 15
    rate_row = n + 16
    sig_row = n + 21
    sup_row = n + 25
    ws1["D12"].value = f"='HOJA LABOR REALIZADA'!F{totals_row}"
    ws1["D14"].value = f"='HOJA LABOR REALIZADA'!F{rate_row}"
    ws1["A20"].value = f"='HOJA LABOR REALIZADA'!A{sig_row}:B{sig_row}"
    ws1["E20"].value = f"='HOJA LABOR REALIZADA'!E{sig_row}:F{sig_row}"
    ws1["A26"].value = f"='HOJA LABOR REALIZADA'!A{sup_row}:B{sup_row}"
    ws1["E26"].value = f"='HOJA LABOR REALIZADA'!E{sup_row}:F{sup_row}"


def render(
    state: State | None = None,
    year: int | None = None,
    month: int | None = None,
    template_path: Path | str | None = None,
    out_path: Path | str | None = None,
) -> dict:
    """Renderiza el mes (year, month) del estado a un .xlsx.

    - `state`: si es None, se carga del JSON (load_state se llama ANTES de
      load_workbook para evitar leer un .tmp huerfano; ver §11).
    - `year` y `month`: requeridos. Seleccionan que bloque de dias en state.days
      se vuelca al .xlsx (solo dias del mes correspondiente que tengan al
      menos un item; los dias sin entradas se ignoran y no aparecen en el Excel).
    - `template_path`: ruta a la plantilla .xlsx (default la del repo).
    - `out_path`: ruta de salida. Si es None, usa
      `exports/timesheet-<year>-<month>.xlsx`.
    """
    if year is None or month is None:
        raise ValueError("year y month son requeridos")
    if state is None:
        state = load_state()
    tpl = Path(template_path) if template_path else TEMPLATE_FILE
    if not tpl.exists():
        raise FileNotFoundError(f"plantilla no encontrada: {tpl}")

    if out_path is None:
        out_path = EXPORTS_DIR / f"timesheet-{year:04d}-{month:02d}.xlsx"
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    all_month_days = trackable_days(year, month)
    days_with_data: list[str] = sorted(
        iso for iso in state.days.keys()
        if iso.startswith(f"{year:04d}-{month:02d}-")
        and state.days[iso].items
    )
    n = len(days_with_data)
    if n == 0:
        return {
            "path": None,
            "year": year,
            "month": month,
            "days_with_data": 0,
            "days_in_month": len(all_month_days),
            "month_total_hours": 0.0,
            "month_total_amount": 0.0,
            "inserted_rows": 0,
            "warning": f"sin entradas con items para {year:04d}-{month:02d}",
            "isError": False,
        }

    inserted = max(0, n - base_data_row_count())

    wb = load_workbook(tpl, keep_links=True)
    ws = wb["HOJA LABOR REALIZADA"]
    ws1 = wb["FACT-SERV PROF"]

    if inserted > 0:
        ws.insert_rows(BASE_LAST_ROW + 1, amount=inserted)

    _write_data_rows(ws, days_with_data, state, inserted)
    _write_header_and_totals(
        ws, state, year, month, n, inserted, days_with_data[-1]
    )
    _patch_sheet1(ws1, n)

    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)

    warning: str | None = None
    if not _has_media(out_path):
        restored = _restore_media_via_zip(out_path, tpl)
        if restored:
            warning = "imagen restaurada desde la plantilla (plan B de zipfile)"
        else:
            warning = "imagen del logo no encontrada tras el render"

    page_msg = (
        f"mostrando {n} dias con entradas "
        f"(de {len(all_month_days)} posibles en {year:04d}-{month:02d})"
    )
    warning = page_msg if warning is None else f"{page_msg}; {warning}"

    return {
        "path": str(out_path.resolve()),
        "year": year,
        "month": month,
        "days_with_data": n,
        "days_in_month": len(all_month_days),
        "last_day": days_with_data[-1],
        "month_total_hours": month_total_hours(state, year, month),
        "month_total_amount": month_total_amount(state, year, month),
        "inserted_rows": inserted,
        "warning": warning,
    }


__all__ = ["render", "DATA_START_ROW", "BASE_LAST_ROW"]
