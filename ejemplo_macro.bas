Attribute VB_Name = "ProveedoresFicticiosDIAN"
' =====================================================================================
'  Macro de ejemplo para consumir la lista de Proveedores Ficticios de la DIAN.
'
'  Llena una hoja oculta con el CSV publicado y, ADEMAS, muestra tres fechas:
'    1) "Lista DIAN actualizada al"  -> meta.json (fecha_actualizacion = ultimo CAMBIO real)
'    2) "Ultima verificacion robot"  -> API de GitHub Actions (run_started_at del ultimo
'                                       run exitoso = cuando el robot reviso, aunque NO
'                                       hubiera cambios). NO genera commits.
'    3) "Consultado por mi el"       -> Now() local (cuando TU corriste esta macro).
'
'  Pegar este modulo en el editor de VBA (Alt+F11) del libro consumidor.
'  Datos en columnas A:E (texto plano); panel de fechas en K1:L4.
'
'  NOTAS DE SEGURIDAD DE DATOS:
'   - Las celdas de datos se formatean como TEXTO ("@") ANTES de volcar: evita que
'     Excel reinterprete fechas ("3/09/2014" NO se convierte al formato regional) y
'     que un valor que empiece por '=', '+', '-' o '@' se ejecute como formula.
'   - Si adaptas esta macro a otro CSV del repo (contadores tiene 9 columnas A:I),
'     ajusta RANGO_DATOS y recuerda que el panel va en K:L para no pisar datos.
' =====================================================================================
Option Explicit

' --- URLs fijas -----------------------------------------------------------------------
Private Const URL_CSV As String = _
    "https://raw.githubusercontent.com/cantejuandavid/proveedores-ficticios-dian/main/proveedores_ficticios.csv"
Private Const URL_META As String = _
    "https://raw.githubusercontent.com/cantejuandavid/proveedores-ficticios-dian/main/meta.json"
Private Const URL_RUNS As String = _
    "https://api.github.com/repos/cantejuandavid/proveedores-ficticios-dian/actions/workflows/actualizar.yml/runs?status=success&per_page=1"

Private Const NOMBRE_HOJA As String = "ProveedoresFicticios"
Private Const RANGO_DATOS As String = "A:E"        ' columnas del CSV (proveedores = 5)
Private Const OFFSET_UTC_COLOMBIA As Double = -5   ' Colombia = UTC-5 (sin horario de verano)


Public Sub CargarProveedoresFicticios()
    Dim ws As Worksheet
    Set ws = ObtenerHojaOculta()
    ws.Cells.Clear

    ' Formatear como TEXTO antes de volcar (ver notas de seguridad arriba).
    ws.Columns(RANGO_DATOS).NumberFormat = "@"

    ' ---------------------------------------------------------------------------------
    ' 1) Descargar y volcar el CSV (fila 1 = encabezados)
    ' ---------------------------------------------------------------------------------
    Dim texto As String
    texto = HttpGet(URL_CSV, False)   ' False = no es API de GitHub
    If Len(texto) = 0 Then
        MsgBox "No se pudo descargar el CSV de proveedores ficticios.", vbExclamation
        Exit Sub
    End If

    ' Quitar BOM UTF-8 si viniera, y normalizar saltos de linea
    If Len(texto) > 0 Then If AscW(Left$(texto, 1)) = 65279 Then texto = Mid$(texto, 2)
    texto = Replace(texto, vbCrLf, vbLf)
    texto = Replace(texto, vbCr, vbLf)

    Dim lineas() As String, campos() As String
    Dim i As Long, j As Long, fila As Long
    lineas = Split(texto, vbLf)
    fila = 0
    For i = LBound(lineas) To UBound(lineas)
        If Len(Trim$(lineas(i))) > 0 Then
            fila = fila + 1
            campos = Split(lineas(i), ";")
            For j = LBound(campos) To UBound(campos)
                ws.Cells(fila, j + 1).Value = campos(j)
            Next j
        End If
    Next i

    ' ---------------------------------------------------------------------------------
    ' 2) Fecha de actualizacion del DATO (ultimo cambio) -> meta.json
    ' ---------------------------------------------------------------------------------
    Dim metaTxt As String, fechaActualizacion As Variant
    metaTxt = HttpGet(URL_META, False)
    fechaActualizacion = IsoUtcToLocal(ExtraerCampoJSON(metaTxt, "fecha_actualizacion"))

    ' ---------------------------------------------------------------------------------
    ' 3) Ultima verificacion del ROBOT -> API de GitHub Actions (sin commits)
    ' ---------------------------------------------------------------------------------
    Dim runsTxt As String, ultimaVerificacion As Variant
    runsTxt = HttpGet(URL_RUNS, True)    ' True = API de GitHub (manda cabeceras propias)
    ultimaVerificacion = IsoUtcToLocal(ExtraerCampoJSON(runsTxt, "run_started_at"))

    ' ---------------------------------------------------------------------------------
    ' 4) Panel de fechas en K1:L4 (lejos de los datos; contadores llega hasta la col. I)
    ' ---------------------------------------------------------------------------------
    ws.Range("K1").Value = "Lista DIAN actualizada al"
    ws.Range("K2").Value = "Ultima verificacion robot"
    ws.Range("K3").Value = "Consultado por mi el"
    ws.Range("K4").Value = "Registros"

    If IsNull(fechaActualizacion) Then
        ws.Range("L1").Value = "n/d"
    Else
        ws.Range("L1").Value = fechaActualizacion
        ws.Range("L1").NumberFormat = "dd/mm/yyyy hh:mm"
    End If

    If IsNull(ultimaVerificacion) Then
        ws.Range("L2").Value = "n/d"
    Else
        ws.Range("L2").Value = ultimaVerificacion
        ws.Range("L2").NumberFormat = "dd/mm/yyyy hh:mm"
    End If

    ws.Range("L3").Value = Now
    ws.Range("L3").NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Range("L4").Value = fila - 1   ' menos la fila de encabezados
End Sub


' =====================================================================================
'  Helpers
' =====================================================================================

' GET robusto con ServerXMLHTTP (evita la cache de IE y soporta HTTPS/proxy).
' esApiGitHub=True agrega las cabeceras que exige la API de GitHub.
Private Function HttpGet(ByVal url As String, ByVal esApiGitHub As Boolean) As String
    On Error GoTo fallo
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    ' Timeouts (ms): resolver DNS, conectar, enviar, recibir. Sin esto una llamada
    ' sincrona puede congelar Excel varios minutos si el host no responde.
    http.setTimeouts 5000, 10000, 10000, 30000

    Dim urlFinal As String
    If esApiGitHub Then
        urlFinal = url    ' la API no se cachea en el CDN
    Else
        ' cache-buster para la raw URL (el CDN de github cachea ~5 min)
        If InStr(url, "?") > 0 Then
            urlFinal = url & "&_=" & Format$(Now, "yyyymmddhhnnss")
        Else
            urlFinal = url & "?_=" & Format$(Now, "yyyymmddhhnnss")
        End If
    End If

    http.Open "GET", urlFinal, False
    http.setRequestHeader "Cache-Control", "no-cache"
    If esApiGitHub Then
        ' GitHub EXIGE User-Agent; Accept fija la version del API.
        http.setRequestHeader "User-Agent", "Excel-DIAN-Macro"
        http.setRequestHeader "Accept", "application/vnd.github+json"
    End If
    http.Send

    If http.Status = 200 Then
        HttpGet = http.responseText
    Else
        HttpGet = ""
    End If
    Exit Function
fallo:
    HttpGet = ""
End Function


' Extrae el valor del primer "campo": "valor" (o numerico) que aparezca en el JSON.
' Suficiente para meta.json (objeto plano) y para el primer run de la API.
Private Function ExtraerCampoJSON(ByVal json As String, ByVal campo As String) As String
    ExtraerCampoJSON = ""
    If Len(json) = 0 Then Exit Function

    Dim clave As String, p As Long, q As Long, c As String
    clave = """" & campo & """"
    p = InStr(1, json, clave, vbBinaryCompare)
    If p = 0 Then Exit Function

    p = InStr(p + Len(clave), json, ":")
    If p = 0 Then Exit Function
    p = p + 1

    ' saltar espacios en blanco
    Do While p <= Len(json)
        c = Mid$(json, p, 1)
        If c <> " " And c <> vbTab And c <> vbCr And c <> vbLf Then Exit Do
        p = p + 1
    Loop
    If p > Len(json) Then Exit Function

    If Mid$(json, p, 1) = """" Then
        ' valor entre comillas
        p = p + 1
        q = InStr(p, json, """")
        If q = 0 Then Exit Function
        ExtraerCampoJSON = Mid$(json, p, q - p)
    Else
        ' valor sin comillas (numero / bool) hasta , } ] o fin de linea
        q = p
        Do While q <= Len(json)
            c = Mid$(json, q, 1)
            If c = "," Or c = "}" Or c = "]" Or c = vbCr Or c = vbLf Then Exit Do
            q = q + 1
        Loop
        ExtraerCampoJSON = Trim$(Mid$(json, p, q - p))
    End If
End Function


' Convierte un timestamp ISO-8601 en UTC (ej "2026-06-15T11:03:21Z" o
' "2026-06-15T11:03:21.123+00:00") a fecha/hora local de Colombia (UTC-5).
' Devuelve Null si la cadena no es valida.
Private Function IsoUtcToLocal(ByVal iso As String) As Variant
    On Error GoTo fallo
    If Len(iso) < 19 Then GoTo fallo

    Dim y As Integer, mo As Integer, d As Integer
    Dim h As Integer, mi As Integer, s As Integer
    y = CInt(Mid$(iso, 1, 4))
    mo = CInt(Mid$(iso, 6, 2))
    d = CInt(Mid$(iso, 9, 2))
    h = CInt(Mid$(iso, 12, 2))
    mi = CInt(Mid$(iso, 15, 2))
    s = CInt(Mid$(iso, 18, 2))

    Dim dtUtc As Date
    dtUtc = DateSerial(y, mo, d) + TimeSerial(h, mi, s)
    IsoUtcToLocal = dtUtc + (OFFSET_UTC_COLOMBIA / 24#)
    Exit Function
fallo:
    IsoUtcToLocal = Null
End Function


' Devuelve (creandola si hace falta) la hoja oculta de datos.
Private Function ObtenerHojaOculta() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(NOMBRE_HOJA)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = NOMBRE_HOJA
    End If
    ws.Visible = xlSheetVeryHidden
    Set ObtenerHojaOculta = ws
End Function
