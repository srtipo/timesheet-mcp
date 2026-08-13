# Timesheet MCP (PRIS)

Servidor MCP (Model Context Protocol) en Python que mantiene el control
horario del profesional en un archivo JSON editable entre sesiones y, al
pedirlo, renderiza un `.xlsx` respetando la plantilla
`Timesheet Modelo Mayo 2025 PRIS.xlsx` (logo, merges, formulas).

## Stack

- Python 3.11+
- `openpyxl` 3.1+
- `mcp[cli]` 1.0+
- `pydantic` 2.6+

## Estructura

```
timesheet/
  models.py      # pydantic: Professional, Supervisor, DayItem, Day, State
  state.py       # load_state, save_state, mutadores puros
  dates.py       # working_days(year, month) y agregaciones
  renderer.py    # render(state, template, out) + plan B de imagen
template/        # copia inmutable de la plantilla PRIS
data/            # state.json (se crea al primer arranque)
exports/         # salida de export_to_excel
tests/           # smoke test (10 casos del spec + extras de borde)
server.py        # entrypoint MCP stdio, 11 tools
```

## Setup (Windows)

```powershell
# 1. Instalar Python 3.11+ desde python.org (marca "Add to PATH")
# 2. Desde la raiz del proyecto:
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -e .
# 3. Verificar que la plantilla este copiada en template/
# 4. Smoke test:
python tests/smoke_test.py
```

## Registro en cliente MCP

### OpenCode (recomendado — ya configurado en este proyecto)

El MCP `timesheet` ya esta registrado globalmente en
`%USERPROFILE%\.config\opencode\opencode.jsonc`:

```json
{
  "$schema": "https://opencode.ai/config.json",
  "mcp": {
    "timesheet": {
      "type": "local",
      "command": [
        "C:\\proyectos\\timesheet\\.venv\\Scripts\\python.exe",
        "C:\\proyectos\\timesheet\\server.py"
      ],
      "cwd": "C:\\proyectos\\timesheet",
      "enabled": true,
      "timeout": 15000
    }
  }
}
```

Disponible en cualquier sesion de opencode (en cualquier directorio de
trabajo) con solo mencionarlo en el prompt, p.ej.:

```
use timesheet para cargar las horas de hoy y exportar el excel del mes
```

Verificacion:

```powershell
opencode mcp list
# -> ✓ timesheet  connected
```

### Claude Desktop

`%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "timesheet": {
      "command": "C:\\proyectos\\timesheet\\.venv\\Scripts\\python.exe",
      "args": ["C:\\proyectos\\timesheet\\server.py"]
    }
  }
}
```

## Tools (11)

| #  | Tool                  | Proposito                                              |
| -- | --------------------- | ------------------------------------------------------ |
| 1  | `get_timesheet`       | Estado completo + bloque `computed` (totales por dia y mes) |
| 2  | `set_professional_info` | Editar nombre, especialidad, tarifa, mes/anio       |
| 3  | `set_supervisor`      | Editar nombre y fecha del supervisor                    |
| 4  | `set_day_metadata`    | Editar `entry` / `exit` de un dia                       |
| 5  | `set_day_item`        | Upsert de un item por descripcion (auto-default 8h)     |
| 6  | `set_day_items`       | Reemplazar la lista completa del dia                   |
| 7  | `delete_day_item`     | Borrar el primer item que matchee `description`         |
| 8  | `delete_day`          | Borrar el dia completo (items + entry/exit)             |
| 9  | `move_item`           | Mover un item por `item_id` entre dias                  |
| 10 | `calculate_hours`     | Totales calculados (sin escribir)                       |
| 11 | `export_to_excel`     | Renderizar el .xlsx segun la plantilla                  |

## Notas

- **Recalculo de formulas al abrir**: el servidor setea
  `wb.calculation.fullCalcOnLoad = True` para que Excel/LibreOffice
  recalcule automaticamente. Si no se ve el total actualizado al abrir
  el archivo, presione `F9` o `Ctrl+Alt+F9`.
- **Paginas multiples**: si el mes tiene mas de 22 laborables (caso
  conocido: 2027-07 y 2027-12), el archivo exportado puede ocupar 2
  paginas impresas. La impresion usa `fitToPage=true` del template.
- **Imagen del logo**: openpyxl generalmente la preserva. Como salvaguarda
  el renderer aplica un Plan B con `zipfile` tras el `save`, copiando
  `xl/media/` y `xl/drawings/_rels/*.rels` desde la plantilla original si
  hace falta. Si la imagen sigue sin aparecer, verifique la consola para
  el `warning` de respuesta.
- **Encoding**: el JSON se persiste en UTF-8 con `ensure_ascii=False`,
  preservando acentos y tildes (ej: "Raúl").
- **Persistencia atomica**: cada mutacion escribe a `state.json.tmp` y
  luego hace `os.replace`. No hay `state.json` huerfano a medio escribir.
