# Timesheet MCP (PRIS)

Servidor MCP (Model Context Protocol) en Python que mantiene el control
horario del profesional en un archivo JSON editable entre sesiones y, al
pedirlo, renderiza un `.xlsx` respetando la plantilla
`Timesheet Modelo Mayo 2025 PRIS.xlsx` (logo, merges, formulas).

**v0.3 — multi-mes**: cada entrada vive por su fecha ISO en `state.days`.
No hay operacion que borre data de otro mes: registrar una entrada nueva
siempre coexiste con lo anterior, sin importar el mes o anio.

## Stack

- Python 3.11+
- `openpyxl` 3.1+
- `mcp[cli]` 1.0+
- `pydantic` 2.6+

## Estructura

```
timesheet/
  models.py      # pydantic: Professional, Supervisor, DayItem, Day, State
  state.py       # load_state, save_state, mutadores puros (multi-mes)
  dates.py       # trackable_days, months_present, today_month
  renderer.py    # render(state, year, month, template, out) + plan B de imagen
template/        # copia inmutable de la plantilla PRIS
data/            # state.json (se crea al primer arranque)
exports/         # salida de export_to_excel
tests/           # smoke test (casos del spec + extras de borde, multi-mes)
server.py        # entrypoint MCP stdio, 12 tools
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

## Tools (12)

| #  | Tool                  | Proposito                                                          |
| -- | --------------------- | ------------------------------------------------------------------ |
| 1  | `get_timesheet`       | Estado completo + `computed` filtrado al mes activo                |
| 2  | `set_professional_info` | Editar nombre, especialidad o tarifa (nunca borra entradas)     |
| 3  | `set_supervisor`      | Editar nombre y fecha del supervisor                                |
| 4  | `set_day_metadata`    | Editar `entry` / `exit` de un dia                                  |
| 5  | `set_day_item`        | Upsert de un item por descripcion (auto-default 8h)                |
| 6  | `set_day_items`       | Reemplazar la lista completa del dia                               |
| 7  | `delete_day_item`     | Borrar el primer item que matchee `description`                     |
| 8  | `delete_day`          | Borrar el dia completo (items + entry/exit)                         |
| 9  | `move_item`           | Mover un item por `item_id` entre dias (cualquier mes)             |
| 10 | `calculate_hours`     | Totales calculados (sin escribir)                                  |
| 11 | `export_to_excel`     | Renderizar el .xlsx del mes pedido                                 |
| 12 | `list_months`         | Meses con datos + mes actual del sistema                           |

## Modelo multi-mes

Cada entrada se guarda con su fecha ISO. **No existe "mes activo"**:
agregar entradas en meses distintos no se pisa. Las herramientas que
necesitan saber "que mes miramos" aceptan `year` y `month` opcionales;
si no los pasas, usan el mes actual del sistema.

```python
# Agregar entrada en cualquier fecha (multiples meses conviven)
set_day_item(date="2026-08-15", description="RDM-X", hours=8)
set_day_item(date="2026-09-20", description="RDM-Y", hours=8)

# Ver que meses tienen datos
list_months()  # -> [{year:2026, month:8, day_count:1}, {year:2026, month:9, day_count:1}]

# Ver el mes actual del sistema (hoy)
get_timesheet()

# Ver un mes especifico
get_timesheet(year=2026, month=8)

# Exportar un mes especifico (requerido; warning si no hay data)
export_to_excel(year=2026, month=8, path="exports/agosto.xlsx")
```

**`set_professional_info` ya no acepta `year`/`month`** (no tendria
sentido: nada cambia entre meses). Cambiar nombre, especialidad o
tarifa NUNCA borra ni altera entradas existentes.

## Export: solo dias con entradas

El `.xlsx` muestra **unicamente los dias que tienen al menos un item**.
Si un mes tiene 21 dias con entradas de 31 posibles, el Excel muestra
solo esas 21 filas con fechas no consecutivas; las demas filas del
template quedan vacias (se limpian explicitamente para que no aparezca
contenido residual del mes anterior que la plantilla trae pre-cargado).

El layout se calcula desde `n` (cantidad de dias con entradas):

```
fila 10..10+n-1    -> dias con entradas (en orden cronologico)
fila n+15          -> totales (F)
fila n+16          -> tarifa (F)
fila n+17          -> monto (F)
fila n+21          -> firma + ultimo dia (A/E)
fila n+25          -> supervisor (A/E)
```

`last_day` en la respuesta y en el Excel es la fecha del **ultimo dia con
entradas**, no el ultimo dia del mes calendario.

## Notas

- **Recalculo de formulas al abrir**: el servidor setea
  `wb.calculation.fullCalcOnLoad = True` para que Excel/LibreOffice
  recalcule automaticamente. Si no se ve el total actualizado al abrir
  el archivo, presione `F9` o `Ctrl+Alt+F9`.
- **Paginas multiples**: el archivo exportado puede ocupar 1 o 2 paginas
  impresas dependiendo de cuantos dias con entradas haya. La impresion
  usa `fitToPage=true` del template.
- **Imagen del logo**: openpyxl generalmente la preserva. Como salvaguarda
  el renderer aplica un Plan B con `zipfile` tras el `save`, copiando
  `xl/media/` y `xl/drawings/_rels/*.rels` desde la plantilla original si
  hace falta. Si la imagen sigue sin aparecer, verifique la consola para
  el `warning` de respuesta.
- **Encoding**: el JSON se persiste en UTF-8 con `ensure_ascii=False`,
  preservando acentos y tildes (ej: "Raúl").
- **Persistencia atomica**: cada mutacion escribe a `state.json.tmp` y
  luego hace `os.replace`. No hay `state.json` huerfano a medio escribir.
- **Migracion automatica v0.2 -> v0.3**: si al cargar `data/state.json`
  se detectan los campos legacy `year`/`month`, se descartan
  silenciosamente y el archivo se reescribe sin ellos. Los dias en
  `state.days` se conservan intactos.
