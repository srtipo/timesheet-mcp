# Especificacion - MCP de Timesheet PRIS

Servidor MCP (Model Context Protocol) en Python que mantiene el control horario del profesional en un archivo JSON editable entre sesiones y, al pedirlo, renderiza un `.xlsx` respetando la plantilla `Timesheet Modelo Mayo 2025 PRIS.xlsx` (logo, merges, formulas).

- **Stack:** Python 3.11+, `openpyxl`, `mcp[cli]`, `pydantic`
- **Entorno objetivo:** Windows, sin Python preinstalado (se instala en el setup)

---

## 1. Objetivos y no-objetivos

**Objetivos**

- Llevar el registro de horas trabajadas de un profesional, con multiples tareas por dia.
- Editar la descripcion de cada dia, agregar/eliminar/mover tareas entre dias.
- Calcular totales por dia y por mes.
- Exportar a `.xlsx` con la misma apariencia, formulas e imagen de la plantilla.

**No-objetivos**

- No se modifica la plantilla base: vive en `template/` y queda intacta.
- No se mantienen multiples profesionales ni historiales cruzados: el estado es un unico mes activo.
- No se valida contra un sistema externo: es offline.

---

## 2. Arquitectura

```
+-----------------+        +----------------+        +---------------------+
|  Cliente MCP    | <----> |  server.py     | <----> |  state.json (data/) |
| (Claude/Cursor) |  stdio |  timesheet/    |        +---------------------+
+-----------------+        |   - state.py   |
                           |   - dates.py   |        +-----------------------+
                           |   - renderer.py| <----> | template/*.xlsx       |
                           +----------------+        +-----------------------+
                                                            |
                                                            v
                                                      exports/*.xlsx
```

- **Transporte:** `stdio` (estandar para MCP servers locales).
- **Persistencia:** JSON plano, lectura/escritura atomica (escribir a `.tmp` + rename).
- **Renderer:** `openpyxl.load_workbook(template, keep_links=True)`, edita celdas especificas, `wb.save(out)`.

---

## 3. Estructura de archivos

```
C:\proyectos\timesheet\
├── pyproject.toml
├── README.md
├── SPEC.md                      # este documento
├── server.py                    # entrypoint MCP, registra tools
├── timesheet\
│   ├── __init__.py
│   ├── models.py                # pydantic: Professional, Supervisor, DayItem, Day, State
│   ├── state.py                 # load_state, save_state, path helpers
│   ├── dates.py                 # working_days(year, month) -> list[date]
│   └── renderer.py              # render(state, template_path, out_path)
├── template\
│   └── Timesheet Modelo Mayo 2025 PRIS.xlsx
├── data\
│   └── state.json               # estado editable (default al primer arranque)
├── exports\                     # salida de export_to_excel
└── tests\
    └── smoke_test.py            # verificacion manual de los 11 tools
```

---

## 4. Esquema del estado (`data/state.json`)

```jsonc
{
  "schema_version": 1,
  "year": 2026,
  "month": 6,
  "professional": {
    "name": "Victor Acosta",
    "specialty": "Desarrollador Front-End",
    "hourly_rate": 3.5
  },
  "supervisor": {
    "name": "Raúl D. Olivero Carrucini",
    "date": "2026-06-30"
  },
  "days": {
    "2026-06-01": {
      "entry": "08:00",   // "HH:MM"  o null
      "exit":  "17:00",   // "HH:MM"  o null
      "items": [
        { "id": "a1b2c3d4-1111-2222-3333-444455556666", "description": "Reunion kick-off", "hours": 3 },
        { "id": "e5f6a7b8-7777-8888-9999-000011112222", "description": "Setup repo",       "hours": 5 }
      ]
    }
  }
}
```

**Decisiones del esquema**

- `days` es un dict indexado por fecha ISO `YYYY-MM-DD` para upsert O(1).
- `entry`/`exit` son metadata (no entran al calculo facturable).
- `items` es la fuente de verdad de las horas del dia.
- `items[].id` es un UUID4 inmutable, generado al crear el item. **No es clave unica dentro del dia**: dos items pueden compartir `description` (la regla de unicidad por descripcion se levanta; los tools que la usaban pasan a operar sobre la primera coincidencia o, en el caso de `move_item`, directamente por `id`).
- `items[].description` se persiste con `.strip()` aplicado; valores vacios o solo-espacios se rechazan.
- `items[].hours` es numero (`float`), redondeado a 2 decimales al persistir, `0` permitido, `>8` permitido (sumar a 11h es valido).
- **Migracion on-load:** si al cargar `state.json` algun item no trae `id`, se le asigna un UUID4 al vuelo y se reescribe el archivo. Asi, un JSON escrito antes de este cambio sigue funcionando sin intervencion manual.

---

## 5. Tools del MCP

Todos devuelven JSON. Errores de validacion devuelven `{"error": "mensaje", "isError": true}`. Warnings (operacion no-op o dato ausente) devuelven `{"warning": "mensaje", "data": ...}` con `isError=false` y el resultado concreto de la operacion.

**Validacion comun a todos los tools que reciben `date`:** la fecha debe caer en **Lun-Vie del mes/anio activo**. Sabados, domingos, o fechas fuera del mes activo se rechazan con error.

**Validacion comun a todos los tools que reciben `description`:** se aplica `.strip()`; si queda vacio, se rechaza con error.

### 5.1 `get_timesheet`

- **Input:** `{}`
- **Output:** estado completo + bloque `computed`:

```json
{
  "state": { "...": "..." },
  "computed": {
    "per_day": [
      { "date": "2026-06-01", "items_total": 8.0, "items_count": 2, "entry": "08:00", "exit": "17:00" }
    ],
    "month_total_hours": 176.0,
    "month_total_amount": 616.0
  }
}
```

### 5.2 `set_professional_info`

- **Input:** `name?`, `specialty?`, `year?`, `month?`, `hourly_rate?`.
- **Reglas:** `year` en [2000, 2100]; `month` en [1, 12]; `hourly_rate` >= 0.
- **Efecto colateral al cambiar `year` o `month`:** se **descartan TODOS los dias anteriores al nuevo mes activo** (no solo los no-laborables). Los dias descartados se reportan en el campo `discarded_days` de la respuesta con su fecha ISO, para que el cliente pueda confirmar al usuario.

### 5.3 `set_supervisor`

- **Input:** `name?`, `date?` (string `YYYY-MM-DD` o `null`).
- Sin restricciones adicionales sobre el mes.

### 5.4 `set_day_metadata`

- **Input:** `date` (requerido), `entry?` (`"HH:MM"` o `null`), `exit?`.
- **Valida:** `date` debe caer en Lun-Vie del mes/anio activo.
- `entry`/`exit` se normalizan a `HH:MM` con cero-padding (`"8:00"` se acepta y se guarda como `"08:00"`).

### 5.5 `set_day_item`

- **Input:** `date` (req), `description` (req, no vacio tras strip), `hours` (opcional, numero `>= 0`).
- **Comportamiento:**
  1. Se valida `date` ∈ laborables del mes activo.
  2. Se aplica `.strip()` a `description`; si queda vacio, error.
  3. Si `description` (ya strippeada) existe en el dia y `hours is None` → **no-op + warning** `"hours omitted, existing item preserved"`. El item existente no se modifica.
  4. Si `description` existe y `hours` no es `None` → **update**: se sobrescribe `hours` del **primer item** que matchee (si hay duplicados, solo el primero). `hours` se redondea a 2 decimales.
  5. Si `description` no existe y `hours is None` → `hours = max(0, 8 - sum(other_items.hours))`. Si la suma del resto ya es `>= 8`, `hours = 0`.
  6. Si `description` no existe y `hours` no es `None` → se crea un item nuevo con `id` UUID4 generado.
  7. `hours` explicito se acepta aunque supere 8 (puede dar 11h o mas). No hay techo impuesto.
- **Retorna:** el item actualizado/creado (con su `id`) + total del dia.

### 5.6 `set_day_items`

- **Input:** `date` (req, laborable), `items` (req, lista de `{id?, description, hours?}`).
- Reemplaza la lista completa del dia. Conserva `entry`/`exit` existentes.
- Cada item puede traer `id` propio o no. Si no trae, se genera uno al guardar.
- `items=[]` → equivalente a vaciar la lista (deja `entry`/`exit` intactos). Para borrar todo el dia, usar `delete_day`.
- `description` strip + rechazo de vacio por item.

### 5.7 `delete_day_item`

- **Input:** `date` (req, laborable), `description` (req, no vacio).
- Borra el **primer item** que matchee `description` (strippeada) en el dia.
- Si no existe ningun match, devuelve `{"warning": "item not found", "data": {...}}` con `isError=false` y el estado sin cambios.

### 5.8 `delete_day`

- **Input:** `date` (req, laborable).
- Borra la entrada completa del dia (items + entry/exit). Sin cambio respecto a la version anterior.

### 5.9 `move_item`

- **Input:** `from_date` (req, laborable), `to_date` (req, laborable), `item_id` (req, string UUID).
- Mueve el item identificado por `id` desde `from_date` a `to_date`, preservando `description` y `hours`.
- **No hay validacion de colision:** dos items con la misma `description` en `to_date` estan permitidos.
- `from_date` queda con `entry`/`exit` y los items restantes preservados; si `from_date` queda sin items, no se borra la metadata.
- Si `item_id` no existe en `from_date`, devuelve `{"warning": "item_id not found", "data": {...}}` con `isError=false`.

### 5.10 `calculate_hours`

- **Input:** `{}`
- **Output:** `month_total_hours`, `month_total_amount`, `per_day[]` con `items_total`, `items_count`, `entry`, `exit`, y `id` por item.
- No escribe nada.

### 5.11 `export_to_excel`

- **Input:** `path?` (opcional). Si no se pasa: `exports/timesheet-YYYY-MM.xlsx`.
- **Pasos:**
  1. Carga `template/Timesheet Modelo Mayo 2025 PRIS.xlsx` con `openpyxl`.
  2. Calcula `working_days(year, month)` y determina `inserted_rows = max(0, n - 22)`.
  3. Si `inserted_rows > 0`, llama a `ws.insert_rows(32, amount=inserted_rows)` y parchea las formulas afectadas (ver §6.1).
  4. Rellena celdas de la Hoja 2 segun seccion 6 (con el `number_format = "HH:MM"` forzado en D/E).
  5. Setea `wb.calculation.calcMode = "auto"` y `wb.calculation.fullCalcOnLoad = True` para forzar el recálculo en Excel al abrir.
  6. `wb.save(path)`.
  7. Verifica que el archivo guardado conserve la imagen (`xl/media/image1.jpeg`); si no, ejecuta el Plan B de post-procesado con `zipfile` (ver §11).
- **Output:**

```json
{
  "path": "<absoluta>",
  "month_total_hours": 176.0,
  "month_total_amount": 616.0,
  "inserted_rows": 0,
  "warning": null
}
```

`warning` se popula con un mensaje si `inserted_rows > 0` (ej: `"month has 23 working days; 1 row inserted; file may span 2 pages"`).

- **No evalua formulas.** Excel/LibreOffice las recalcula al abrir.

---

## 6. Reglas de renderizado (Hoja 2 — `HOJA LABOR REALIZADA`)

| Celda (caso base, 22 filas) | Origen | Notas |
| --- | --- | --- |
| `A6` | `professional.name` | Top-left del merge `A6:B6` |
| `C6` | `professional.specialty` | Top-left de `C6:D6` |
| `E6` | `datetime.date(year, month, 1)` (primer dia del mes activo) | Top-left de `E6:F6`. El `number_format` del template es `mmmm\ yyyy`, por lo que se ve como "June 2026" |
| `A10:A31` | `working_days(year, month)` (los primeros 22) | Formato fecha del template (`dd/mm/yyyy`). Si el mes tiene mas laborables, se insertan filas adicionales (ver §6.1) |
| `B10:B31` | `" ".join(item.description for item in day.items)` | Separador: un solo espacio. B esta mergeada con C (`B10:C10` ... `B31:C31`); escribir en B es correcto |
| `D10:D31` | `day.entry` | `datetime.time` con `number_format = "HH:MM"` (24h, forzado por el renderer aunque el template traiga 12h). Vacio si `entry is None` |
| `E10:E31` | `day.exit` | Idem |
| `F10:F31` | `round(sum(item.hours for item in day.items), 2)` | Numero con 2 decimales; **reemplaza la formula original** del template. `number_format = "0.00"` (heredado del template) |
| `F37` (caso base) | Formula `=SUM(F10:F31)` (existente en el template) | Se preserva. Si hubo `inserted_rows > 0`, se parchea a `=SUM(F10:F{31+inserted_rows})` |
| `F38` (caso base) | `professional.hourly_rate` | Reemplaza el `3.5` actual. `number_format` currency del template |
| `F39` (caso base) | Formula `=F37*F38` (existente) | Se preserva tal cual; las refs son por nombre, no se ven afectadas por `insert_rows` |
| `A43` (caso base) | **No se toca.** Es formula `=A6` que arrastra el valor del header. Escribir aqui machacaría la formula y romperia la referencia desde Hoja 1 | (Aclaracion: la tabla de la v0 decia "A43 = professional.name", pero la celda es formula, no valor) |
| `E43` (caso base) | `working_days[-1]` (ultimo dia del mes activo) | Formato `d/m/yyyy` del template |
| `A47` (caso base) | `supervisor.name` | Top-left de `A47:B47` |
| `E47` (caso base) | `supervisor.date` (parseado a `datetime.date`) | Formato `d/m/yyyy` |

**Hoja 1** (`FACT-SERV PROF`) no se toca directamente: las formulas `='HOJA LABOR REALIZADA'!A6:B6`, `!C6:D6`, `!E6:F6`, `!F37`, `!F38`, `!A43:B43`, `!E43:F43`, `!A47:B47`, `!E47:F47` ya estan en la plantilla. Cuando hubo `inserted_rows > 0`, las formulas que apuntan a filas shifted se parchean (ver §6.1).

**Formulas preservadas tras el save (caso base):**

- Hoja 2: `F37 = SUM(F10:F31)`, `F39 = F37*F38`.
- Hoja 1: `D12 = 'HOJA LABOR REALIZADA'!F37`, `D14 = 'HOJA LABOR REALIZADA'!F38`, `D16 = D12*D14`, `A7`, `C7`, `E7`, `A20`, `A26`, etc.

### 6.1 Manejo de meses con mas de 22 laborables

`openpyxl` corre `ws.insert_rows(32, amount=inserted_rows)` para abrir espacio justo despues de la ultima fila de datos (fila 31). Esta operacion:

- shiftea **celdas** (incluyendo las de totales en filas 37-39 y la firma en filas 43-49) hacia abajo, **conservando sus valores y formulas**;
- shiftea **merge ranges** que se solapan con las filas movidas;
- **NO actualiza referencias de formulas** (es un bug conocido de openpyxl).

Por eso, despues del `insert_rows` el renderer aplica estos parches explicitos:

```python
if inserted_rows > 0:
    new_totals = 37 + inserted_rows        # F37 -> F(new_totals)
    new_rate = 38 + inserted_rows          # F38 -> F(new_rate)
    new_amount = 39 + inserted_rows        # F39 -> F(new_amount)  (formula F37*F38, valida por nombre)
    new_signature = 43 + inserted_rows     # A43 -> A(new_signature)
    new_last_day = 43 + inserted_rows      # E43 -> E(new_last_day)
    new_sup_name = 47 + inserted_rows      # A47 -> A(new_sup_name)
    new_sup_date = 47 + inserted_rows      # E47 -> E(new_sup_date)
    new_last_data_row = 31 + inserted_rows

    # Hoja 2: extender la suma
    ws.cell(new_totals, 6).value = f"=SUM(F10:F{new_last_data_row})"
    # F39 = F37*F38 -> no necesita parche (refs por nombre)
    # A43 = A6       -> no necesita parche (refs por nombre)
    # A47, E47       -> se reescriben abajo en el flujo normal (fila new_sup_*)

    # Hoja 1: actualizar refs cross-sheet
    ws1["D12"].value = f"='HOJA LABOR REALIZADA'!F{new_totals}"
    ws1["D14"].value = f"='HOJA LABOR REALIZADA'!F{new_rate}"
    ws1["A20"].value = f"='HOJA LABOR REALIZADA'!A{new_signature}:B{new_signature}"
    ws1["E20"].value = f"='HOJA LABOR REALIZADA'!E{new_last_day}:F{new_last_day}"
    ws1["A26"].value = f"='HOJA LABOR REALIZADA'!A{new_sup_name}:B{new_sup_name}"
    ws1["E26"].value = f"='HOJA LABOR REALIZADA'!E{new_sup_date}:F{new_sup_date}"
```

**Casos en 2026** que disparan `inserted_rows > 0`: ninguno (todos los meses tienen ≤22 laborables). En 2027: julio (23) y diciembre (23). En general, cualquier mes de 31 dias donde los 3 dias "extra" (los que exceden 4 semanas) caigan todos en Lun-Vie.

**Consecuencia visible:** con `inserted_rows > 0` el archivo exportado puede ocupar **mas de una pagina impresa**. La pagina 1 contiene la cabecera de la factura y las primeras 22 filas de datos; la pagina 2 contiene las filas adicionales y el bloque de totales/firma. La impresion se ajusta automaticamente por la configuracion `fitToPage=true` del template, pero conviene verificar en Excel.

---

## 7. Reglas de negocio

1. **Dias laborables:** solo Lun-Vie. Sabados/domingos, rechazo.
2. **Filas de la plantilla:** 22 base (filas 10-31). Si el mes tiene **menos** laborables, las sobrantes quedan vacias. Si tiene **mas**, se insertan filas adicionales y se parchean las formulas (ver §6.1). **Ya no se trunca.**
3. **Cambio de mes/anio:** se descartan **TODOS** los dias del mes anterior, no solo los no-laborables, y se reportan en `discarded_days`.
4. **Default de horas (`set_day_item` sin `hours`):** `max(0, 8 - sum(other_items.hours))`.
5. **Exceder 8h:** permitido. `total_dia = 11h` es valido si el usuario lo pide. No hay techo.
6. **Persistencia:** cada tool que muta estado hace `save_state()` (escritura atomica via `.tmp` + `os.replace`).
7. **Validacion de entradas:** pydantic + validaciones custom en `state.py`.
8. **Encoding:** UTF-8 con `ensure_ascii=False` para preservar acentos/tildes (ej: "Raúl").
9. **Redondeo de horas:** `round(hours, 2)` al persistir y al renderizar.

---

## 8. Dependencias (`pyproject.toml`)

```toml
[project]
name = "timesheet-mcp"
version = "0.1.0"
requires-python = ">=3.11"
dependencies = [
  "openpyxl>=3.1",
  "mcp[cli]>=1.0",
  "pydantic>=2.6",
]

[build-system]
requires = ["setuptools>=68"]
build-backend = "setuptools.build_meta"

[project.scripts]
timesheet-mcp = "server:main"
```

---

## 9. Setup (orden de ejecucion)

1. Instalar Python 3.11+ desde python.org (marcar *Add to PATH*).
2. `cd C:\proyectos\timesheet`
3. `python -m venv .venv`
4. `.\.venv\Scripts\Activate.ps1`
5. `pip install -e .` (o `pip install openpyxl mcp[cli] pydantic`).
6. Copiar `Timesheet Modelo Mayo 2025 PRIS.xlsx` a `template/` (la copia, no se modifica).
7. El archivo `data/state.json` se crea automaticamente al primer arranque con valores por defecto.
8. `python server.py` para smoke test en consola.
9. Registrar en `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "timesheet": {
      "command": "C:\\proyectos\timesheet\\.venv\\Scripts\\python.exe",
      "args": ["C:\\proyectos\timesheet\\server.py"]
    }
  }
}
```

---

## 10. Smoke tests minimos

| #  | Accion                                                    | Esperado                                                                 |
| -- | --------------------------------------------------------- | ------------------------------------------------------------------------ |
| 1  | `export_to_excel()` con plantilla limpia                  | Archivo abre, logo visible, totales en 0                                 |
| 2  | `set_day_item("2026-06-01", "Tarea A")`                  | Item con `hours=8` y `id` UUID4                                          |
| 3  | `set_day_item("2026-06-01", "Tarea B", hours=3)`         | Item `B` con 3h, total dia = 11h                                         |
| 4  | `set_day_item("2026-06-01", "Tarea A", hours=2)`         | Update: `A=2h`, total dia = 5h                                           |
| 5  | `export_to_excel()` con 5 dias cargados                  | `F37` muestra 40 (al recalcular)                                         |
| 6  | `set_professional_info(year=2026, month=7)`              | Dias se regeneran para julio; junio descartado, listado en `discarded_days` |
| 7  | Reabrir `.xlsx` en Excel                                 | Logo, merges, formulas, totales correctos                                |
| 8  | `set_day_item` con desc existente y `hours=None`          | No-op + warning, item preservado                                         |
| 9  | `move_item` por `item_id`                                | Item se mueve, `from_date` conserva `entry`/`exit`                       |
| 10 | Render de un mes con 23 laborables (ej: 2027-07)         | Archivo se abre, F37 muestra el total correcto tras recálculo, sin errores |

---

## 11. Riesgos abiertos

- **Imagenes compartidas en openpyxl:** la plantilla tiene 1 imagen (`xl/media/image1.jpeg`) usada desde ambas hojas. openpyxl generalmente la preserva al `save`. **Plan B desde v1** (no como fallback opcional): el renderer, despues de `wb.save(out)`, abre el zip resultante y verifica que exista `xl/media/image1.jpeg`. Si falta, reescribe el archivo con `zipfile` copiando `xl/media/` y los `xl/drawings/_rels/*.rels` desde la plantilla original. Asi, en el peor caso el logo se restaura.
- **`insert_rows` y formulas:** openpyxl shiftea celdas y merges, pero **no actualiza referencias dentro de formulas**. Mitigación: parchear manualmente las 6 formulas cross-sheet en Hoja 1 + la `F37` en Hoja 2 cuando `inserted_rows > 0` (ver §6.1). Smoke test #10 es obligatorio: exportar 2027-07, abrir en Excel, verificar `F37` y `D12` en Hoja 1.
- **Formulas sin recalcular al abrir:** aceptable; igual forzamos `fullCalcOnLoad = True` al guardar. Se documenta en el README.
- **Zona horaria / formato de hora:** se usa `datetime.time` con `number_format = "HH:MM"` (24h, forzado por el renderer). Sin tz, ya que el template no la usa. El template trae formato 12h `h:mm:ss AM/PM`; el renderer lo sobreescribe a `HH:MM` al escribir las celdas de entry/exit.
- **Concurrencia:** el estado se guarda atomicamente (`os.replace`), pero no hay locking entre invocaciones del MCP. Aceptable para uso single-user. **Importante:** `export_to_excel` debe llamar a `load_state()` **antes** de `load_workbook(template)`, para evitar leer un `state.json.tmp` huerfano durante un `save_state` en curso.
- **Encoding del JSON:** UTF-8 con `ensure_ascii=False` tanto en `save_state` (dump) como en `load_state` (read), para preservar acentos/tildes en nombres (ej: "Raúl").

---

## 12. Out of scope (futuro)

- Multiples profesionales / multiples meses simultaneos.
- Validacion contra calendario del supervisor.
- Exportacion a PDF.
- Historial de versiones del JSON.
- UI web (no es el objetivo del MCP).

---

## 13. Orden de implementacion

1. `models.py` + `state.py` + `dates.py` (con `id` en DayItem + migración on-load).
2. `renderer.py` con tests contra la plantilla (validar imagen, formulas, formatos, **caso `inserted_rows > 0`**).
3. `server.py` registrando los 11 tools con las nuevas firmas.
4. Smoke tests de la seccion 10 + ejecucion manual de los tests extras (#8, #9, #10).
5. README con instrucciones de instalacion, configuracion del cliente MCP y notas sobre el recálculo automatico de Excel y las paginas multiples.
