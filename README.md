# Listas DIAN — actualización automática

Mantiene actualizadas, de forma automática, tres listas oficiales de la DIAN y las publica
como CSV estables que pueden leerse desde Excel/VBA por una URL fija:

1. **Proveedores ficticios**
2. **Contadores sancionados por la DIAN**
3. **Autorretenedores del impuesto sobre la renta**

- Fuente oficial: https://www.dian.gov.co/Paginas/Inicio.aspx
- El script ([`scraper.py`](scraper.py)) renderiza con Playwright la página de cada lista,
  localiza su enlace (un `<li data-content="...">` o un `<a href>`), resuelve la URL **actual**
  de su PDF (cambia cada vez que la DIAN actualiza), lo descarga y extrae la tabla con `pdfplumber`.
- Cada lista se procesa de forma **independiente**: si una falla, las demás se actualizan y
  publican igual (el workflow commitea lo que sí salió bien y aun así marca el run en rojo).
- Un GitHub Action programado ([`.github/workflows/actualizar.yml`](.github/workflows/actualizar.yml))
  lo ejecuta dos veces por semana (lunes y jueves, 6 a.m. Colombia) y manualmente cuando
  quieras, y commitea los archivos solo si cambian.

> ℹ️ El repositorio se llama `proveedores-ficticios-dian` por razones históricas (nació con
> esa única lista). No se renombra porque romperia las URLs raw que ya consumen las macros.

## Propósito y aviso

Este proyecto tiene **fines exclusivamente académicos y de apoyo a la comunidad contable y
tributaria**. Su único objetivo es facilitar el acceso a información **pública** que la DIAN ya
publica en su página oficial —las listas de **Proveedores ficticios**, **Contadores
sancionados por la DIAN** y **Autorretenedores del impuesto sobre la renta**—, de modo que
sirva como **ayuda para la toma de decisiones**.

- No persigue ningún fin comercial, ni distinto al de consultar y consolidar esa información pública.
- No modifica, interpreta ni certifica los datos: solo reproduce lo que la DIAN publica en sus PDF.
- La **fuente oficial y única válida** sigue siendo la DIAN
  (https://www.dian.gov.co/Paginas/Inicio.aspx). Ante cualquier diferencia, prevalece la
  publicación oficial.
- Los datos pueden contener errores de extracción o estar desactualizados respecto a la fuente;
  verifica siempre contra el PDF oficial antes de tomar decisiones con efectos legales.

## URLs raw de los CSV (las que consume Excel/VBA)

```
https://raw.githubusercontent.com/cantejuandavid/proveedores-ficticios-dian/main/proveedores_ficticios.csv
https://raw.githubusercontent.com/cantejuandavid/proveedores-ficticios-dian/main/contadores_sancionados.csv
https://raw.githubusercontent.com/cantejuandavid/proveedores-ficticios-dian/main/autorretenedores_renta.csv
```

> Son las URLs fijas que consumen las macros de Excel/VBA.

## Archivos publicados

| Archivo | Descripción |
|---|---|
| `proveedores_ficticios.csv` / `.json` | Lista de proveedores ficticios (automática). |
| `meta.json` | Meta de proveedores. Conserva su nombre histórico por compatibilidad con las macros ya desplegadas; es el meta de **una sola** fuente, no del repo. |
| `contadores_sancionados.csv` / `.json` | Lista de contadores sancionados (automática). |
| `contadores_sancionados.meta.json` | Meta de contadores. |
| `autorretenedores_renta.csv` / `.json` | Lista de autorretenedores de renta (automática). |
| `autorretenedores_renta.meta.json` | Meta de autorretenedores. |
| `entidades_no_sujetas.csv` | **Estático/manual**: entidades financieras no sujetas (bancos, corporaciones financieras, etc.). NO lo actualiza el robot; se sube a mano cuando cambia. Algunas razones sociales vienen truncadas (~80 caracteres) desde su origen. |
| `dias_inhabiles.csv` | **Estático/manual**: días inhábiles/festivos de Colombia 2024–2035 (fechas `YYYY-MM-DD`). NO lo actualiza el robot. |

Todos los CSV: **UTF-8 con BOM, separador `;`, con encabezados, estructura estable.**
(Los estáticos `dias_inhabiles.csv` y `entidades_no_sujetas.csv` pueden venir sin BOM; su
contenido es ASCII o ya viene normalizado.)

### Esquema de los `*.meta.json` (idéntico para las tres fuentes)

```json
{
  "fuente":              "nombre legible de la lista",
  "fecha_actualizacion": "última vez que los DATOS cambiaron (ISO-8601 UTC)",
  "url_pdf":             "URL del PDF de la DIAN usado en la última corrida",
  "num_registros":       0,
  "columnas":            ["..."],
  "pagina":              "página de la DIAN donde vive el enlace"
}
```

> `fecha_actualizacion` **no** es "última verificación": si el robot corre y no hay cambios,
> esa fecha no se mueve. La "última verificación" se consulta a la API de GitHub Actions
> (ver macro). Si la DIAN republica el mismo contenido con otro nombre de archivo, el meta
> se actualiza (`url_pdf`) pero la fecha se conserva.

### Estructura fija de los CSV

```
# proveedores_ficticios.csv
NIT;Razon_Social;Resolucion;Fecha;Estado

# contadores_sancionados.csv
No;Nombre;Cedula;Inscripcion_Profesional;Resolucion;Sancion;Fecha_Ejecutoria;Vencimiento;Autoridad

# autorretenedores_renta.csv
NIT;Razon_Social;Resolucion;Fecha

# entidades_no_sujetas.csv (estático)
NIT;Razon_Social;Detalle

# dias_inhabiles.csv (estático)
FECHA;MOTIVO
```

El mapeo de columnas del PDF a estos nombres canónicos está definido por fuente en la lista
`FUENTES` de [`scraper.py`](scraper.py). El parser detecta la fila de encabezado del PDF por
palabras clave (en tres pases: igualdad, prefijo, subcadena, robusto ante reordenamientos);
si no la detecta, asume el orden posicional. Si la DIAN cambia la estructura de algún PDF,
ajusta esa fuente en `FUENTES`.

## Robustez

- Si **no encuentra el enlace**, no puede descargar, o la tabla no supera las validaciones,
  el script **no sobrescribe** el CSV bueno anterior, registra el error y sale con código
  distinto de 0 (el run queda en rojo, pero las fuentes que sí funcionaron se commitean).
- Validaciones: mínimo de filas, % mínimo de identificadores numéricos plausibles y
  **validación con memoria**: si el número de registros cae más del 50% frente a la corrida
  anterior (PDF truncado, enlace equivocado), se rechaza. Si la reducción es legítima,
  borra el `*.meta.json` de esa fuente o ajusta `MAX_CAIDA_REGISTROS` y vuelve a correr.
- Escritura **atómica**: genera `*.tmp`, valida y solo entonces reemplaza los archivos finales.
- **Cron a prueba de inactividad**: GitHub desactiva los `schedule` tras ~60 días sin
  actividad en el repo, y este proyecto solo commitea cuando la DIAN cambia algo. Por eso el
  workflow incluye un paso *keep-alive* que se re-habilita a sí mismo en cada corrida.

Códigos de salida de `scraper.py`: `0` todas las fuentes OK · `2` alguna fuente con fallo
controlado (validación/estructura) · `1` fallo inesperado (excepción no controlada o no se
pudieron cargar las páginas).

## Uso local

```bash
python -m venv .venv
# Windows PowerShell:  .venv\Scripts\Activate.ps1
# Linux/Mac:           source .venv/bin/activate
pip install -r requirements.txt
python -m playwright install chromium
python scraper.py
```

## Consumo desde Excel/VBA

Los CSV mantienen encabezados fijos, separador `;` y UTF-8 para que una macro pueda hacer un
`GET` a la URL raw, partir por líneas y por `;`, y volcar a una hoja oculta.

> 📎 **Macro completa recomendada:** [`ejemplo_macro.bas`](ejemplo_macro.bas) carga el CSV
> (volcando como **texto** para que Excel no reinterprete fechas ni ejecute nada que parezca
> fórmula) y muestra tres fechas en `K1:L4`: *Lista DIAN actualizada al* (de `meta.json`),
> *Última verificación del robot* (API de GitHub Actions, **sin generar commits**) y
> *Consultado por mí el* (`Now()` local). Requiere al menos un run del workflow en Actions.

Ejemplo mínimo de referencia (usa `ServerXMLHTTP` —evita la caché de WinINet— y formatea las
celdas como texto antes de volcar):

```vba
Sub CargarProveedoresFicticios()
    Const URL As String = _
        "https://raw.githubusercontent.com/cantejuandavid/proveedores-ficticios-dian/main/proveedores_ficticios.csv"
    Dim http As Object, texto As String, lineas() As String, campos() As String
    Dim i As Long, j As Long, ws As Worksheet

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 5000, 10000, 10000, 30000
    http.Open "GET", URL & "?t=" & Format(Now, "yyyymmddhhnnss"), False ' evita caché CDN
    http.Send
    If http.Status <> 200 Then
        MsgBox "Error al descargar: " & http.Status
        Exit Sub
    End If

    texto = http.responseText
    If Len(texto) > 0 Then If AscW(Left(texto, 1)) = 65279 Then texto = Mid(texto, 2) ' BOM
    texto = Replace(texto, vbCrLf, vbLf)
    texto = Replace(texto, vbCr, vbLf)
    lineas = Split(texto, vbLf)

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ProveedoresFicticios")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "ProveedoresFicticios"
        ws.Visible = xlSheetVeryHidden
    End If
    ws.Cells.Clear
    ws.Columns("A:E").NumberFormat = "@"  ' texto: sin fechas reinterpretadas ni fórmulas

    For i = LBound(lineas) To UBound(lineas)
        If Len(Trim(lineas(i))) > 0 Then
            campos = Split(lineas(i), ";")
            For j = LBound(campos) To UBound(campos)
                ws.Cells(i + 1, j + 1).Value = campos(j)
            Next j
        End If
    Next i
End Sub
```
