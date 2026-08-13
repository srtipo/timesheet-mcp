"""Renderiza el estado a un .xlsx respetando la plantilla PRIS."""

from __future__ import annotations

import shutil
import zipfile
from datetime import date, datetime, time
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .dates import (
    base_data_row_count,
    first_data_row,
    last_base_data_row,
    month_total_amount,
    month_total_hours,
    per_day_totals,
    working_days,
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
    ws.cell(row=row, column=col).value = value


def _set_cell_with_format(ws: Worksheet, row: int, col: int, value: Any, fmt: str) -> None:
    c = ws.cell(row=row, column=col)
    c.value = value
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
    ws: Worksheet, wd: list[date], state: State, inserted: int
) -> None:
    for idx, d in enumerate(wd[: BASE_LAST_ROW - DATA_START_ROW + 1 + inserted]):
        row = DATA_START_ROW + idx
        day: Day | None = state.days.get(d.isoformat())
        _set_cell(ws, row, 1, d)  # A: fecha, formato heredado del template
        if day and day.items:
            joined = " ".join(it.description for it in day.items)
            _set_cell(ws, row, 2, joined)  # B (mergeada con C)
        else:
            _set_cell(ws, row, 2, None)
        # Forzar HH:MM (24h) en D y E aunque el template traiga 12h.
        entry = day.entry if day else None
        exit_ = day.exit if day else None
        _set_cell_with_format(ws, row, 4, entry, HHMM_FMT)
        _set_cell_with_format(ws, row, 5, exit_, HHMM_FMT)
        if day and day.items:
            total = round(sum(it.hours for it in day.items), 2)
        else:
            total = 0.0
        _set_cell_with_format(ws, row, 6, total, HOURS_FMT)


def _write_header_and_totals(
    ws: Worksheet, wd: list[date], state: State, inserted: int
) -> None:
    sh = _shifted(inserted)
    _set_cell(ws, 6, 1, state.professional.name)
    _set_cell(ws, 6, 3, state.professional.specialty)
    _set_cell(ws, 6, 5, date(state.year, state.month, 1))

    last_data_row = sh["last_data"]
    if inserted > 0:
        _set_cell(ws, sh["totals"], 6, f"=SUM(F{DATA_START_ROW}:F{last_data_row})")
    else:
        _set_cell(ws, BASE_TOTALS_ROW, 6, f"=SUM(F{DATA_START_ROW}:F{BASE_LAST_ROW})")

    _set_cell(ws, sh["rate"], 6, state.professional.hourly_rate)

    last_day = wd[-1]
    _set_cell(ws, sh["last_day"], 5, last_day)

    _set_cell(ws, sh["sup_name"], 1, state.supervisor.name)
    sup_date = state.supervisor.sup_date
    _set_cell(ws, sh["sup_date"], 5, sup_date)


def _patch_sheet1(ws1: Worksheet, inserted: int) -> None:
    if inserted <= 0:
        return
    sh = _shifted(inserted)
    ws1["D12"].value = f"='HOJA LABOR REALIZADA'!F{sh['totals']}"
    ws1["D14"].value = f"='HOJA LABOR REALIZADA'!F{sh['rate']}"
    ws1["A20"].value = (
        f"='HOJA LABOR REALIZADA'!A{sh['sig']}:B{sh['sig']}"
    )
    ws1["E20"].value = (
        f"='HOJA LABOR REALIZADA'!E{sh['last_day']}:F{sh['last_day']}"
    )
    ws1["A26"].value = (
        f"='HOJA LABOR REALIZADA'!A{sh['sup_name']}:B{sh['sup_name']}"
    )
    ws1["E26"].value = (
        f"='HOJA LABOR REALIZADA'!E{sh['sup_date']}:F{sh['sup_date']}"
    )


def render(
    state: State | None = None,
    template_path: Path | str | None = None,
    out_path: Path | str | None = None,
) -> dict:
    """Renderiza el estado a un .xlsx.

    Si `state` es None, se carga del JSON (load_state se llama ANTES de
    load_workbook para evitar leer un .tmp huerfano; ver §11).
    """
    if state is None:
        state = load_state()
    tpl = Path(template_path) if template_path else TEMPLATE_FILE
    if not tpl.exists():
        raise FileNotFoundError(f"plantilla no encontrada: {tpl}")

    if out_path is None:
        out_path = EXPORTS_DIR / f"timesheet-{state.year:04d}-{state.month:02d}.xlsx"
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)

    wd = working_days(state.year, state.month)
    inserted = max(0, len(wd) - base_data_row_count())

    wb = load_workbook(tpl, keep_links=True)
    ws = wb["HOJA LABOR REALIZADA"]
    ws1 = wb["FACT-SERV PROF"]

    if inserted > 0:
        ws.insert_rows(BASE_LAST_ROW + 1, amount=inserted)

    _write_data_rows(ws, wd, state, inserted)
    _write_header_and_totals(ws, wd, state, inserted)
    _patch_sheet1(ws1, inserted)

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

    if inserted > 0:
        page_msg = (
            f"mes tiene {len(wd)} laborables; se insertaron {inserted} filas; "
            "el archivo puede abarcar 2 paginas"
        )
        warning = f"{page_msg}" if warning is None else f"{page_msg}; {warning}"

    return {
        "path": str(out_path.resolve()),
        "month_total_hours": month_total_hours(state),
        "month_total_amount": month_total_amount(state),
        "inserted_rows": inserted,
        "warning": warning,
    }


__all__ = ["render", "DATA_START_ROW", "BASE_LAST_ROW"]
