# Subir registros de Excel a Airtable

Script para cargar todos los días los registros de asistencia desde un archivo
Excel a la tabla **RegistroIngresos** de la base
[appEItiU345wNDzPP](https://airtable.com/appEItiU345wNDzPP/tblWocpsI4879rdSd).

## 1. Instalación (una sola vez)

```bash
pip install -r requirements.txt
```

## 2. Token de Airtable (una sola vez)

1. Entra a https://airtable.com/create/tokens y crea un *personal access token*.
2. Dale los permisos (*scopes*) `data.records:read` y `data.records:write`.
3. En **Access**, agrega la base de este proyecto.
4. Guarda el token en una variable de entorno:

```bash
export AIRTABLE_TOKEN="patXXXXXXXXXXXXXX.XXXXXXXX..."
```

En Windows (PowerShell): `$env:AIRTABLE_TOKEN = "patXXXX..."`

También puedes copiar `.env.example` a `.env` y cargarlo con tu herramienta
preferida, o pasar el token directo con `--token`.

## 3. Formato del Excel

El archivo debe tener una fila de encabezados y estas columnas:

| Columna | Obligatoria | Descripción |
|---|---|---|
| `Colaborador` | Sí | Nombre tal como aparece en la tabla **Colaboradores**. |
| `Fecha de Visita` | Sí | `2026-08-28` o `28/08/2026`, o una celda con formato fecha de Excel. |
| `Notas` | No | Texto libre. |

Puedes generar una plantilla lista para usar:

```bash
python subir_registros.py --crear-plantilla plantilla_registros.xlsx
```

Detalles útiles:

- Los nombres de columna no distinguen mayúsculas ni acentos, y se aceptan
  variantes comunes (`Nombre`, `Empleado`, `Fecha`, `Observaciones`, etc.).
  Si tus encabezados son distintos, agrégalos al diccionario `ALIAS_COLUMNAS`
  al inicio de `subir_registros.py`.
- El nombre del colaborador se busca ignorando mayúsculas y acentos, así que
  `maria fernanda lopez` encuentra a `María Fernanda López`.
- Si prefieres ser exacto, usa una columna `Colaborador ID` con el `recXXXX...`
  de Airtable; tiene prioridad sobre el nombre.
- Las filas vacías se ignoran.

## 4. Uso diario

Primero una simulación para revisar que todo esté bien (no escribe nada):

```bash
python subir_registros.py registros_de_hoy.xlsx --simulacion
```

Y cuando el resumen se vea correcto:

```bash
python subir_registros.py registros_de_hoy.xlsx
```

### Opciones

| Opción | Para qué sirve |
|---|---|
| `--simulacion` | Muestra lo que se subiría sin escribir en Airtable. |
| `--fecha hoy` | Rellena con la fecha de hoy las filas que no traen fecha (también acepta `--fecha 2026-08-28`). |
| `--hoja "Agosto"` | Usa una hoja específica del libro (por defecto, la primera). |
| `--crear-colaboradores` | Da de alta en **Colaboradores** los nombres que no existan todavía. |
| `--permitir-duplicados` | No omite filas ya cargadas (por defecto sí las omite). |
| `--token pat...` | Token de Airtable, si no quieres usar la variable de entorno. |
| `--crear-plantilla ruta.xlsx` | Genera un Excel vacío con las columnas correctas. |

## 5. Qué hace el script

1. Lee el `.xlsx`, `.xlsm` o `.csv` y detecta las columnas.
2. Descarga la tabla **Colaboradores** y traduce cada nombre a su `recId`
   (el campo `Colaborador` de RegistroIngresos es un enlace, no texto).
3. Valida las fechas y las convierte a `YYYY-MM-DD`.
4. **Evita duplicados**: omite las filas cuyo par colaborador + fecha ya existe
   en Airtable, y también los duplicados dentro del mismo Excel. Así puedes
   volver a correr el mismo archivo sin ensuciar la tabla.
5. Sube los registros en lotes de 10 (el máximo de la API), con reintentos
   automáticos si Airtable responde con límite de peticiones o error temporal.
6. Al final imprime cuántos registros se crearon y lista las filas con problemas
   (nombre no encontrado, fecha ilegible, colaborador vacío). Esas filas se
   omiten; el resto sí se sube.

El código de salida es `0` si todo salió bien y `1` si hubo filas con problemas,
por si quieres encadenarlo en una tarea programada.

## 6. Automatizarlo (opcional)

Ejemplo de cron para correr todos los días a las 8:00 con el archivo del día:

```cron
0 8 * * * cd /ruta/al/proyecto && AIRTABLE_TOKEN=pat... /usr/bin/python3 subir_registros.py "/ruta/a/registros/$(date +\%Y-\%m-\%d).xlsx" >> subida.log 2>&1
```
