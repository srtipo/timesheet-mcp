# Especificacion - MCP de Timesheet PRIS

Servidor MCP (Model Context Protocol) en Python que mantiene el control horario del profesional en un archivo JSON editable entre sesiones y, al pedirlo, renderiza un `.xlsx` respetando la plantilla `Timesheet Modelo Mayo 2025 PRIS.xlsx` (logo, merges, formulas).

**v0.3 — modelo multi-mes:** cada entrada vive por su fecha ISO en `state.days`. No existe el concepto de "mes activo" persistente: las operaciones que necesitan un mes (`get_timesheet`, `calculate_hours`, `export_to_excel`) lo reciben por parametro o, en su defecto, usan el mes actual del sistema. **Ninguna operacion puede borrar entradas de otro mes accidentalmente.**

- **Stack:** Python 3.11+, `openpyxl`, `mcp[cli]`, `pydantic`
- **Entorno objetivo:** Windows, sin Python preinstalado (se instala en el setup)

---

## 1. Objetivos y no-objetivos

**Objetivos**

- Llevar el registro de horas trabajadas de un profesional, con multiples tareas por dia.
- Editar la descripcion de cada dia, agregar/eliminar/mover tareas entre dias.
- Calcular totales por dia y por mes, sobre cualquier mes con datos.
- Exportar a `.xlsx` con la misma apariencia, formulas e imagen de la plantilla.

**No-objetivos**

- No se modifica la plantilla base: vive en `template/` y queda intacta.
- No se mantienen multiples profesionales. **Si se mantiene un estado multi-mes:** cada entrada persiste por su fecha ISO hasta que se borre explicitamente con `delete_day`.
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
- **Multi-mes:** `state.days` es un dict `str(YYYY-MM-DD) -> Day` sin asociacion a un mes. Los meses se derivan de las fechas en tiempo de consulta.

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
│   ├── state.py                 # load_state, save_state, mutadores (multi-mes)
│   ├── dates.py                 # trackable_days(year, month), months_present, today_month
│   └── renderer.py              # render(state, year, month, template_path, out_path)
├── template\
│   └── Timesheet Modelo Mayo 2025 PRIS.xlsx
├── data\
│   └── state.json               # estado editable (default al primer arranque)
├── exports\                     # salida de export_to_excel
└── tests\
    └── smoke_test.py            # verificacion manual de los 12 tools
```

---

## 4. Esquema del estado (`data/state.json`)

```jsonc
{
  "schema_version": 1,
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
    "2026-08-03": {
      "entry": "08:00",   // "HH:MM"  o null
      "exit":  "17:00",   // "HH:MM"  o null
      "items": [
        { "id": "a1b2c3d4-1111-2222-3333-444455556666", "description": "Reunion kick-off", "hours": 3 },
        { "id": "e5f6a7b8-7777-8888-9999-000011112222", "description": "Setup repo",       "hours": 5 }
      ]
    },
    "2026-08-04": {
      "entry": "08:00",
      "exit":  "17:00",
      "items": [
        { "id": "...", "description": "Otra tarea", "hours": 8 }
      ]
    }
  }
}
```

**Decisiones del esquema**

- **Sin `year`/`month`:** el estado es multi-mes. Cada entrada de `days` vive por su fecha ISO. Los meses se derivan al consultar (ver `months_present(state)` en `dates.py`).
- **`days`** es un dict indexado por fecha ISO `YYYY-MM-DD` para upsert O(1). Acepta claves de **cualquier mes o anio**.
- **`entry`/`exit`** son metadata (no entran al calculo facturable).
- **`items`** es la fuente de verdad de las horas del dia.
- **`items[].id`** es un UUID4 inmutable, generado al crear el item.
- **`items[].description`** se persiste con `.strip()` aplicado; valores vacios o solo-espacios se rechazan.
- **`items[].hours`** es numero (`float`), redondeado a 2 decimales al persistir, `0` permitido, `>8` permitido (sumar a 11h es valido).
- **Migracion on-load (v0.2 -> v0.3):** si al cargar `state.json` aparecen campos legacy `year`/`month`, se descartan silenciosamente y el archivo se reescribe sin ellos. Los dias en `days` se conservan.
- **Migracion on-load (items sin id):** si algun item no trae `id`, se le asigna un UUID4 al vuelo y se reescribe el archivo.
- **Defaults de `entry`/`exit`:** ver seccion 7 regla 9. `entry` por default es `08:00` y `exit` se calcula como `entry + sum(item.hours)`. Si el JSON trae `entry`/`exit` en `null`, el renderer y los `get_timesheet`/`calculate_hours` aplican los defaults solo en lectura.

---

## 5. Tools del MCP

Todos devuelven JSON. Errores de validacion devuelven `{"error": "mensaje", "isError": true}`. Warnings (operacion no-op o dato ausente) devuelven `{"warning": "mensaje", "data": ...}` con `isError=false` y el resultado concreto de la operacion.

**Validacion comun a todos los tools que reciben `date`:** la fecha debe ser un ISO valido `YYYY-MM-DD`. **No hay validacion contra un mes activo.** Cualquier fecha (pasada, presente o futura, en cualquier mes o anio dentro del rango razonable) se acepta. Esto es por diseno: el estado es multi-mes y la fecha la decide el usuario.

**Validacion comun a todos los tools que reciben `description`:** se aplica `.strip()`; si queda vacio, se rechaza con error.

**Parametros `year`/`month` opcionales:** los tools `get_timesheet`, `calculate_hours` y `export_to_excel` aceptan `year` y `month`. Si ambos faltan, usan el mes actual del sistema (`today.year`, `today.month`). Si solo uno viene, se rechaza con error (`year` y `month` deben ir juntos o ninguno).

### 5.1 `get_timesheet`

- **Input:** `year?`, `month?` (opcionales, juntos o ninguno).
- **Output:**

```json
{
  "active_month": { "year": 2026, "month": 8, "source": "explicit" },
  "state": { "...": "estado completo, todos los meses" },
  "computed": {
    "per_day": [
      { "date": "2026-08-01", "items_total": 8.0, "items_count": 1, "entry": "08:00", "exit": "17:00" }
    ],
    "month_total_hours": 168.0,
    "month_total_amount": 588.0,
    "months_present": [
      { "year": 2026, "month": 8, "day_count": 21 },
      { "year": 2026, "month": 9, "day_count": 1 }
    ]
  }
}
```

`active_month.source` es `"explicit"` si se pasaron `year`/`month`, `"today"` si se uso el mes actual del sistema.

`state.days` siempre contiene **todos** los dias de **todos** los meses (no se filtra). El filtrado se refleja en `computed.per_day`.

### 5.2 `set_professional_info`

- **Input:** `name?`, `specialty?`, `hourly_rate?` (todos opcionales).
- **Reglas:** `hourly_rate` >= 0. **`year`/`month` ya NO se aceptan** (el schema los rechaza como propiedades extra).
- **Efecto colateral:** **ninguno.** Cambiar nombre/especialidad/tarifa NUNCA borra ni altera entradas existentes, porque no hay operacion que pueda hacerlo.

### 5.3 `set_supervisor`

- **Input:** `name?`, `date?` (string `YYYY-MM-DD` o `null`).
- Sin restricciones adicionales.

### 5.4 `set_day_metadata`

- **Input:** `date` (requerido, ISO), `entry?` (`"HH:MM"` o `null`), `exit?`.
- `date` no se valida contra mes activo: cualquier ISO es valido.
- `entry`/`exit` se normalizan a `HH:MM` con cero-padding (`"8:00"` se acepta y se guarda como `"08:00"`).

### 5.5 `set_day_item`

- **Input:** `date` (req), `description` (req, no vacio tras strip), `hours` (opcional, numero `>= 0`).
- `date` no se valida contra mes activo.
- **Comportamiento:**
  1. Se aplica `.strip()` a `description`; si queda vacio, error.
  2. Si `description` (ya strippeada) existe en el dia y `hours is None` → **no-op + warning** `"hours omitted, existing item preserved"`. El item existente no se modifica.
  3. Si `description` existe y `hours` no es `None` → **update**: se sobrescribe `hours` del **primer item** que matchee.
  4. Si `description` no existe y `hours is None` → `hours = max(0, 8 - sum(other_items.hours))`. Si la suma del resto ya es `>= 8`, `hours = 0`.
  5. Si `description` no existe y `hours` no es `None` → se crea un item nuevo con `id` UUID4 generado.
  6. `hours` explicito se acepta aunque supere 8 (puede dar 11h o mas). No hay techo impuesto.
  7. **Aplicacion de defaults `entry`/`exit`** (ver seccion 7 regla 9): despues de crear/actualizar el item, si `day.entry is None` se setea `08:00`; si `day.exit is None` se calcula como `entry + sum(items.hours)`. Si el usuario habia fijado `entry`/`exit` explicitamente via `set_day_metadata`, esos valores se preservan.
- **Retorna:** el item actualizado/creado (con su `id`) + total del dia.

### 5.6 `set_day_items`

- **Input:** `date` (req, ISO), `items` (req, lista de `{id?, description, hours?}`).
- `date` no se valida contra mes activo.
- Reemplaza la lista completa del dia. Conserva `entry`/`exit` existentes.
- Cada item puede traer `id` propio o no. Si no trae, se genera uno al guardar.
- `items=[]` → equivalente a vaciar la lista (deja `entry`/`exit` intactos). Para borrar todo el dia, usar `delete_day`.
- **Aplicacion de defaults `entry`/`exit`:** ver seccion 7 regla 9.
- `description` strip + rechazo de vacio por item.

### 5.7 `delete_day_item`

- **Input:** `date` (req, ISO), `description` (req, no vacio).
- Borra el **primer item** que matchee `description` (strippeada) en el dia.
- Si no existe ningun match, devuelve `{"warning": "item not found", "data": {...}}` con `isError=false` y el estado sin cambios.

### 5.8 `delete_day`

- **Input:** `date` (req, ISO).
- Borra la entrada completa del dia (items + entry/exit).

### 5.9 `move_item`

- **Input:** `from_date` (req, ISO), `to_date` (req, ISO), `item_id` (req, string UUID).
- Mueve el item identificado por `id` desde `from_date` a `to_date`, preservando `description` y `hours`.
- `from_date` y `to_date` pueden ser de meses distintos (mover del 28/08 al 03/09 es valido).
- `from_date` queda con `entry`/`exit` y los items restantes preservados; si `from_date` queda sin items, no se borra la metadata.
- Si `item_id` no existe en `from_date`, devuelve `{"warning": "item_id not found", "data": {...}}` con `isError=false`.

### 5.10 `calculate_hours`

- **Input:** `year?`, `month?` (opcionales, juntos o ninguno).
- **Output:**

```json
{
  "active_month": { "year": 2026, "month": 8, "source": "explicit" },
  "per_day": [
    { "date": "2026-08-01", "items_total": 8.0, "items_count": 1, "entry": "08:00", "exit": "17:00",
      "items": [{ "id": "...", "description": "...", "hours": 8 }] }
  ],
  "month_total_hours": 168.0,
  "month_total_amount": 588.0,
  "months_present": [
    { "year": 2026, "month": 8, "day_count": 21 }
  ]
}
```

No escribe nada.

### 5.11 `export_to_excel`

- **Input:** `year?`, `month?` (opcionales; si ambos faltan, se usa el mes actual del sistema), `path?` (opcional; default `exports/timesheet-YYYY-MM.xlsx`).
- **Pre-condicion:** el mes a exportar debe tener al menos una entrada con items en `state.days`. Si no la tiene, devuelve `{"warning": "no hay entradas con items para YYYY-MM", "year": ..., "month": ..., "days_with_data": 0, "days_in_month": N, "path": null, "isError": false}` y no crea archivo.
- **Solo dias con entradas:** el Excel muestra **unicamente los dias que tienen al menos un item** (en orden cronologico). Los dias del mes sin entradas se IGNORAN completamente: no se escribe la fecha ni items ni horas en sus filas; ademas se limpian las celdas A..F de esas filas del template para que no quede contenido residual del mes anterior que la plantilla trae pre-cargado.
- **Pasos:**
  1. Carga `template/Timesheet Modelo Mayo 2025 PRIS.xlsx` con `openpyxl`.
  2. Calcula `days_with_data`: fechas ISO del mes que tienen al menos un item, ordenadas. `n = len(days_with_data)`.
  3. Calcula `inserted_rows = max(0, n - 22)`.
  4. Si `inserted_rows > 0`, llama a `ws.insert_rows(32, amount=inserted_rows)`.
  5. Escribe los `n` dias en filas `10..10+n-1`. Limpia las celdas A..F de las filas `10+n..31+inserted` (las celdas MergedCell se omiten silenciosamente).
  6. Header en fila 6 (nombre, especialidad, mes facturado).
  7. Totales/firma/supervisor en posiciones calculadas: totales en `n+15`, tarifa en `n+16`, monto en `n+17`, firma+ultimo dia en `n+21`, supervisor en `n+25`. Estas formulas son validas para cualquier `n` (con o sin inserts) porque `n+15 = 37 + (n-22)` cuando `inserted = n-22`.
  8. Formula totales: `=SUM(F10:F{10+n-1})`. Rango exacto de las filas con datos.
  9. Parchea formulas cross-sheet en Hoja 1 (`D12`, `D14`, `A20`, `E20`, `A26`, `E26`) usando las mismas posiciones `n+15`, `n+16`, `n+21`, `n+25`.
  10. Setea `wb.calculation.calcMode = "auto"` y `wb.calculation.fullCalcOnLoad = True` para forzar el recálculo en Excel al abrir.
  11. `wb.save(path)`.
  12. Verifica que el archivo guardado conserve la imagen (`xl/media/image1.jpeg`); si no, ejecuta el Plan B de post-procesado con `zipfile` (ver §11).
- **Output:**

```json
{
  "path": "<absoluta o null>",
  "year": 2026,
  "month": 8,
  "active_month_source": "explicit",
  "days_with_data": 21,
  "days_in_month": 31,
  "last_day": "2026-08-31",
  "month_total_hours": 168.0,
  "month_total_amount": 588.0,
  "inserted_rows": 0,
  "warning": "mostrando 21 dias con entradas (de 31 posibles en 2026-08)"
}
```

`days_with_data` es la cantidad de dias exportados. `days_in_month` es el total de dias del mes calendario. `last_day` es la fecha ISO del ultimo dia con entradas (no el ultimo dia del mes). `inserted_rows` refleja cuantas filas extra se insertaron despues de la 31 para acomodar los datos.

### 5.12 `list_months`

- **Input:** `{}`.
- **Output:**

```json
{
  "months_present": [
    { "year": 2026, "month": 8, "day_count": 21 },
    { "year": 2026, "month": 9, "day_count": 1 }
  ],
  "today_month": { "year": 2026, "month": 9 },
  "isError": false
}
```

`months_present` viene ordenado cronologicamente. Sirve para descubrir que meses tienen datos antes de pedir un `get_timesheet` o `export_to_excel` filtrado.

---

## 6. Reglas de renderizado (Hoja 2 — `HOJA LABOR REALIZADA`)

El layout es **unificado y basado en `n`** (cantidad de dias con entradas en el mes), no en la cantidad de dias del mes calendario. Esto se debe a que solo se exportan los dias con entradas.

**Layout para cualquier `n`:**

| Fila | Contenido |
| --- | --- |
| 6 | Header: nombre, especialidad, mes facturado |
| 10 a 10+n-1 | Datos: un dia por fila, ordenados por fecha |
| 10+n a 10+n+4 | Gap (5 filas vacias, vienen del template) |
| n+15 | Totales: `=SUM(F10:F{10+n-1})` |
| n+16 | Tarifa: `professional.hourly_rate` |
| n+17 | Monto: `=F{n+15}*F{n+16}` (formula por nombre, no necesita patch) |
| n+21 | Firma + ultimo dia: `professional.name` (en A) + `last_day` (en E) |
| n+25 | Supervisor: `supervisor.name` (en A) + `supervisor.date` (en E) |

**Reglas:**

- **Solo dias con entradas se exportan.** Si un mes tiene 5 dias con items y 25 sin items, el Excel muestra solo 5 filas de datos (fechas no consecutivas). Las demas filas del template quedan vacias (limpieza explicita).
- **Insert de filas:** si `n > 22`, se insertan `n-22` filas despues de la 31 antes de escribir los datos. Las posiciones del layout (`n+15`, `n+16`, etc.) ya contemplan este shift.
- **`last_day`** es la fecha ISO del ultimo dia con entradas, no el ultimo dia del mes.
- **Limpieza de filas no usadas:** para evitar que el template muestre contenido residual del mes anterior pre-cargado, las celdas A..F de las filas `10+n..31+inserted` se ponen en `None`. Las celdas `MergedCell` se omiten silenciosamente (read-only).

### 6.1 Manejo de meses con mas de 22 entradas

`openpyxl` corre `ws.insert_rows(32, amount=inserted_rows)` para abrir espacio justo despues de la ultima fila base de datos (fila 31). Esto solo se hace si `n > 22`. La operacion:

- shiftea **celdas** (incluyendo las de totales en filas 37-39 y la firma en filas 43-49) hacia abajo, **conservando sus valores y formulas**;
- shiftea **merge ranges** que se solapan con las filas movidas;
- **NO actualiza referencias de formulas** (es un bug conocido de openpyxl).

El renderer compensa esto escribiendo directamente en las posiciones calculadas desde `n`:
- totales en `n+15` (= `37 + (n-22)` cuando hay shift, = `37` cuando `n==22`, = `20` cuando `n==5`);
- parchea cross-sheet refs en Hoja 1 a esas mismas posiciones.

**Consecuencia visible:** con `inserted_rows > 0` el archivo exportado puede ocupar **mas de una pagina impresa**. La pagina 1 contiene la cabecera de la factura y las primeras filas de datos; la pagina 2 contiene el resto y el bloque de totales/firma.

---

## 7. Reglas de negocio

1. **Dias validos:** cualquier fecha ISO (cualquier mes, cualquier anio dentro del rango razonable, cualquier dia de la semana). No hay validacion contra un mes activo.
2. **Filas de la plantilla:** 22 base (filas 10-31). Si `n > 22` (dias con entradas), se insertan `n-22` filas. Si `n <= 22`, no se insertan. El layout se calcula desde `n` (ver §6), asi funciona uniforme con o sin inserts.
3. **Multi-mes:** `state.days` puede contener entradas de multiples meses simultaneamente. `months_present(state)` lista los meses con datos. No existe operacion que descarte entradas de un mes arbitrariamente.
4. **Export solo dias con entradas:** el Excel muestra unicamente los dias con al menos un item. Dias sin entradas del mes se IGNORAN: no se escribe la fecha ni items ni horas, y se limpian las celdas residuales del template. El resultado visual es como si esas filas no existieran.
5. **Default de horas (`set_day_item` sin `hours`):** `max(0, 8 - sum(other_items.hours))`.
6. **Exceder 8h:** permitido. `total_dia = 11h` es valido si el usuario lo pide. No hay techo.
7. **Persistencia:** cada tool que muta estado hace `save_state()` (escritura atomica via `.tmp` + `os.replace`).
8. **Validacion de entradas:** pydantic + validaciones custom en `state.py`.
9. **Encoding:** UTF-8 con `ensure_ascii=False` para preservar acentos/tildes (ej: "Raúl").
10. **Redondeo de horas:** `round(hours, 2)` al persistir y al renderizar.
11. **Defaults de `entry`/`exit`:** cuando se llama a `set_day_item` o `set_day_items`, se aplican defaults al dia mutado solo si los campos son `None`:
    - `entry` por default: `08:00`.
    - `exit` por default: `entry + sum(items.hours)` (wrap a 24h). Si la lista de items queda vacia, `exit` no se calcula y queda en `null`.
    - Si el usuario ya habia fijado `entry` o `exit` explicitamente via `set_day_metadata`, esos valores prevalecen en llamadas posteriores.
    - **Persistencia:** los defaults se escriben en `state.days[iso]` al mutar items. No hay migracion retroactiva: dias preexistentes en `data/state.json` sin `entry`/`exit` guardados mantienen `null` en el JSON hasta que se llame a un tool de mutacion. El renderer y los tools de lectura (`get_timesheet`, `calculate_hours`, `per_day_totals`) aplican los defaults **en lectura**.
    - **No aplica en `move_item` ni `delete_day_item`:** mover o borrar items no recalcula `exit`.

---

## 8. Dependencias (`pyproject.toml`)

```toml
[project]
name = "timesheet-mcp"
version = "0.3.0"
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
| 1  | `export_to_excel(year=2026, month=6)` con 1 entrada 0h    | Archivo abre, logo visible, totales en 0                                 |
| 2  | `set_day_item("2026-06-01", "Tarea A")`                  | Item con `hours=8` y `id` UUID4                                          |
| 3  | `set_day_item("2026-06-01", "Tarea B", hours=3)`         | Item `B` con 3h, total dia = 11h                                         |
| 4  | `set_day_item("2026-06-01", "Tarea A", hours=2)`         | Update: `A=2h`, total dia = 5h                                           |
| 5  | `export_to_excel(year=2026, month=6)` con 5 dias cargados | F20 (n+15) muestra 40, `=SUM(F10:F14)`                                   |
| 6  | `set_professional_info(name="X")` con entradas en ago+sep | Ambos meses intactos (multi-mes)                                         |
| 7  | Reabrir `.xlsx` en Excel                                 | Logo, merges, formulas, totales correctos                                |
| 8  | `set_day_item` con desc existente y `hours=None`          | No-op + warning, item preservado                                         |
| 9  | `move_item` por `item_id`                                | Item se mueve, `from_date` conserva `entry`/`exit`                       |
| 10 | Render con 23 dias en 2027-12                              | `inserted_rows=1`, totales en F38 (n+15), cross-sheet parcheado a F38     |
| 11 | `set_day_item("2026-08-15", "X")` y luego `set_day_item("2026-09-01", "Y")` | Ambos dias coexisten en `state.days` sin pisarse |
| 12 | `list_months` con datos en ago+sep                        | Devuelve `[{year:2026, month:8, day_count:N1}, {year:2026, month:9, day_count:N2}]` |
| 13 | Exportar junio con 5 entradas dispersas (no consecutivas) | Solo esas 5 filas con datos en el Excel, gap rows 15..19 vacias         |
| 14 | Exportar mes sin items                                    | Warning, `path=None`, sin archivo                                        |
| 15 | Exportar junio con 25 entradas                            | `inserted_rows=3`, totales en F40 (n+15)                                 |
| 16 | Exportar mes donde el ultimo dia con data es el 10         | `last_day = "2026-06-10"`, E25 (n+21) contiene esa fecha                |

---

## 11. Riesgos abiertos

- **Imagenes compartidas en openpyxl:** la plantilla tiene 1 imagen (`xl/media/image1.jpeg`) usada desde ambas hojas. openpyxl generalmente la preserva al `save`. **Plan B desde v1** (no como fallback opcional): el renderer, despues de `wb.save(out)`, abre el zip resultante y verifica que exista `xl/media/image1.jpeg`. Si falta, reescribe el archivo con `zipfile` copiando `xl/media/` y los `xl/drawings/_rels/*.rels` desde la plantilla original. Asi, en el peor caso el logo se restaura.
- **`insert_rows` y formulas:** openpyxl shiftea celdas y merges, pero **no actualiza referencias dentro de formulas**. Mitigación: parchear manualmente las 6 formulas cross-sheet en Hoja 1 + la `F37` en Hoja 2 cuando `inserted_rows > 0` (ver §6.1). Smoke test #10 es obligatorio: exportar 2027-12, abrir en Excel, verificar `F37` y `D12` en Hoja 1.
- **Formulas sin recalcular al abrir:** aceptable; igual forzamos `fullCalcOnLoad = True` al guardar. Se documenta en el README.
- **Zona horaria / formato de hora:** se usa `datetime.time` con `number_format = "HH:MM"` (24h, forzado por el renderer). Sin tz, ya que el template no la usa. El template trae formato 12h `h:mm:ss AM/PM`; el renderer lo sobreescribe a `HH:MM` al escribir las celdas de entry/exit.
- **Concurrencia:** el estado se guarda atomicamente (`os.replace`), pero no hay locking entre invocaciones del MCP. Aceptable para uso single-user. **Importante:** `export_to_excel` debe llamar a `load_state()` **antes** de `load_workbook(template)`, para evitar leer un `state.json.tmp` huerfano durante un `save_state` en curso.
- **Encoding del JSON:** UTF-8 con `ensure_ascii=False` tanto en `save_state` (dump) como en `load_state` (read), para preservar acentos/tildes en nombres (ej: "Raúl").
- **Multi-mes y export:** si se exporta un mes sin entradas, `export_to_excel` devuelve warning sin crear archivo. Esto es por diseno: un xlsx vacio no aporta valor.

---

## 12. Out of scope (futuro)

- Multiples profesionales.
- Validacion contra calendario del supervisor.
- Exportacion a PDF.
- Historial de versiones del JSON.
- UI web (no es el objetivo del MCP).
- Confirmacion interactiva antes de borrar entradas (ahora se borran con `delete_day` sin pedir confirmacion).

---

## 13. Orden de implementacion

1. `models.py` + `state.py` + `dates.py` (multi-mes, sin `year`/`month` en State; `today_month`, `months_present` en dates).
2. `renderer.py` con `render(state, year, month, ...)` (parametros requeridos).
3. `server.py` registrando los 12 tools con las nuevas firmas (incluido `list_months`).
4. Smoke tests de la seccion 10 + extras (multi-mes persistence, list_months, warning en export sin data).
5. README con instrucciones de instalacion, configuracion del cliente MCP y notas sobre el recálculo automatico de Excel y las paginas multiples.
