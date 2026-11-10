Attribute VB_Name = "SetupMemoriaYBuses"
Option Explicit

' ==============================================================================
' MODULO: SetupMemoriaYBuses.bas
' Genera los paneles faltantes que el docente solicito:
'   1. Hoja "MEMORIA": RAM Fisica (8 Marcos x 32 Bytes) + Area de SWAP (16 Paginas)
'   2. Visualizacion de Buses en la hoja CPU (Control, Datos, Direcciones, Interno)
'   3. Teclado Interactivo funcional con boton de escritura
' Ejecutar: ALT+F8 > GenerarMemoriaSwapYBuses > Ejecutar
' ==============================================================================

Public Sub GenerarMemoriaSwapYBuses()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' --- PARTE 1: Crear hoja MEMORIA con RAM y SWAP ---
    CrearHojaMemoria
    
    ' --- PARTE 2: Agregar buses visuales y teclado a la hoja CPU ---
    MejorarHojaCPU
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "Paneles de Memoria RAM, Swap, Buses visuales y Teclado interactivo generados." & vbCrLf & _
           "Revise las hojas 'CPU' y 'MEMORIA'.", vbInformation, "SIS-131"
End Sub

' =============================================================================
' PARTE 1: HOJA DE MEMORIA (RAM FISICA + SWAP)
' =============================================================================
Private Sub CrearHojaMemoria()
    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Set ws = wb.Worksheets("MEMORIA")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = "MEMORIA"
    Else
        On Error Resume Next
        Dim shp As Shape
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
        ws.Cells.Clear
        On Error GoTo 0
    End If
    
    ' Fondo oscuro
    ws.Cells.Interior.Color = RGB(24, 28, 36)
    ws.Cells.Font.Color = RGB(210, 215, 225)
    ws.Cells.Font.Name = "Consolas"
    ws.Cells.Font.Size = 9
    
    ' Columna A como separador
    ws.Columns("A").ColumnWidth = 2
    
    ' Configurar anchos para 16 columnas de datos + encabezado
    ws.Columns("B").ColumnWidth = 14
    Dim c As Long
    For c = 3 To 18  ' C to R = 16 columnas de bytes
        ws.Columns(c).ColumnWidth = 4.5
    Next c
    ws.Columns("S").ColumnWidth = 2  ' separador
    ws.Columns("T").ColumnWidth = 14
    Dim c2 As Long
    For c2 = 21 To 36  ' U to AJ = 16 columnas de bytes
        ws.Columns(c2).ColumnWidth = 4.5
    Next c2
    
    ' =================================================================
    ' BANNER PRINCIPAL
    ' =================================================================
    ws.Range("B1:R2").Merge
    ws.Range("B1").Value = "RAM FISICA (256 Bytes = 8 Marcos de 32 Bytes)"
    EstiloTituloMem ws.Range("B1:R2"), RGB(15, 76, 129)
    
    ws.Range("T1:AJ2").Merge
    ws.Range("T1").Value = "AREA DE SWAP / DISCO VIRTUAL (512 Bytes = 16 Paginas de 32 Bytes)"
    EstiloTituloMem ws.Range("T1:AJ2"), RGB(142, 68, 43)
    
    ' =================================================================
    ' RAM FISICA: 8 Marcos de 32 Bytes cada uno
    ' =================================================================
    Dim marco As Long
    Dim fila As Long
    fila = 3
    
    For marco = 0 To 7
        ' Encabezado del marco
        ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 18)).Merge
        ws.Cells(fila, 2).Value = "MARCO " & marco & "  |  Direccion Fisica: 0x" & Right$("00" & Hex$(marco * 32), 3) & " - 0x" & Right$("00" & Hex$(marco * 32 + 31), 3)
        
        Dim clrMarco As Long
        If marco Mod 2 = 0 Then
            clrMarco = RGB(20, 60, 90)
        Else
            clrMarco = RGB(30, 50, 70)
        End If
        EstiloTituloMem ws.Range(ws.Cells(fila, 2), ws.Cells(fila, 18)), clrMarco
        
        fila = fila + 1
        
        ' Encabezados de columna (+0 a +F)
        ws.Cells(fila, 2).Value = "Offset:"
        ws.Cells(fila, 2).Font.Size = 7
        ws.Cells(fila, 2).Font.Color = RGB(140, 150, 160)
        Dim col As Long
        For col = 0 To 15
            ws.Cells(fila, 3 + col).Value = "+" & Hex$(col)
            ws.Cells(fila, 3 + col).Font.Size = 7
            ws.Cells(fila, 3 + col).Font.Color = RGB(100, 120, 140)
            ws.Cells(fila, 3 + col).HorizontalAlignment = xlCenter
        Next col
        fila = fila + 1
        
        ' Fila 1 del marco (bytes 0-15)
        ws.Cells(fila, 2).Value = "0x" & Right$("00" & Hex$(marco * 32), 3) & ":"
        ws.Cells(fila, 2).Font.Bold = True
        ws.Cells(fila, 2).Font.Color = RGB(52, 152, 219)
        For col = 0 To 15
            ws.Cells(fila, 3 + col).Value = "00"
            ws.Cells(fila, 3 + col).HorizontalAlignment = xlCenter
            ws.Cells(fila, 3 + col).Font.Color = RGB(46, 204, 113)
            ws.Cells(fila, 3 + col).Interior.Color = RGB(30, 38, 48)
            ws.Cells(fila, 3 + col).Borders.LineStyle = xlContinuous
            ws.Cells(fila, 3 + col).Borders.Color = RGB(50, 60, 70)
        Next col
        fila = fila + 1
        
        ' Fila 2 del marco (bytes 16-31)
        ws.Cells(fila, 2).Value = "0x" & Right$("00" & Hex$(marco * 32 + 16), 3) & ":"
        ws.Cells(fila, 2).Font.Bold = True
        ws.Cells(fila, 2).Font.Color = RGB(52, 152, 219)
        For col = 0 To 15
            ws.Cells(fila, 3 + col).Value = "00"
            ws.Cells(fila, 3 + col).HorizontalAlignment = xlCenter
            ws.Cells(fila, 3 + col).Font.Color = RGB(46, 204, 113)
            ws.Cells(fila, 3 + col).Interior.Color = RGB(30, 38, 48)
            ws.Cells(fila, 3 + col).Borders.LineStyle = xlContinuous
            ws.Cells(fila, 3 + col).Borders.Color = RGB(50, 60, 70)
        Next col
        fila = fila + 1
        
        ' Espacio entre marcos
        fila = fila + 1
    Next marco
    
    ' =================================================================
    ' AREA DE SWAP: 16 Paginas de 32 Bytes cada una
    ' =================================================================
    Dim pagina As Long
    fila = 3
    
    For pagina = 0 To 15
        ' Encabezado de la pagina
        ws.Range(ws.Cells(fila, 20), ws.Cells(fila, 36)).Merge
        ws.Cells(fila, 20).Value = "PAGINA VIRTUAL " & pagina & "  |  Swap Offset: 0x" & Right$("000" & Hex$(pagina * 32), 3) & " - 0x" & Right$("000" & Hex$(pagina * 32 + 31), 3)
        
        Dim clrPag As Long
        If pagina < 8 Then
            clrPag = RGB(20, 70, 50)  ' Verde oscuro = en RAM (no en swap)
        Else
            clrPag = RGB(90, 40, 20)  ' Rojo oscuro = en swap
        End If
        EstiloTituloMem ws.Range(ws.Cells(fila, 20), ws.Cells(fila, 36)), clrPag
        
        fila = fila + 1
        
        ' Fila 1 de la pagina (bytes 0-15)
        ws.Cells(fila, 20).Value = "0x" & Right$("000" & Hex$(pagina * 32), 3) & ":"
        ws.Cells(fila, 20).Font.Bold = True
        ws.Cells(fila, 20).Font.Size = 8
        ws.Cells(fila, 20).Font.Color = RGB(231, 76, 60)
        For col = 0 To 15
            ws.Cells(fila, 21 + col).Value = "00"
            ws.Cells(fila, 21 + col).HorizontalAlignment = xlCenter
            ws.Cells(fila, 21 + col).Font.Size = 8
            ws.Cells(fila, 21 + col).Font.Color = RGB(200, 150, 100)
            ws.Cells(fila, 21 + col).Interior.Color = RGB(35, 30, 28)
            ws.Cells(fila, 21 + col).Borders.LineStyle = xlContinuous
            ws.Cells(fila, 21 + col).Borders.Color = RGB(60, 50, 40)
        Next col
        fila = fila + 1
        
        ' Fila 2 de la pagina (bytes 16-31)
        ws.Cells(fila, 20).Value = "0x" & Right$("000" & Hex$(pagina * 32 + 16), 3) & ":"
        ws.Cells(fila, 20).Font.Bold = True
        ws.Cells(fila, 20).Font.Size = 8
        ws.Cells(fila, 20).Font.Color = RGB(231, 76, 60)
        For col = 0 To 15
            ws.Cells(fila, 21 + col).Value = "00"
            ws.Cells(fila, 21 + col).HorizontalAlignment = xlCenter
            ws.Cells(fila, 21 + col).Font.Size = 8
            ws.Cells(fila, 21 + col).Font.Color = RGB(200, 150, 100)
            ws.Cells(fila, 21 + col).Interior.Color = RGB(35, 30, 28)
            ws.Cells(fila, 21 + col).Borders.LineStyle = xlContinuous
            ws.Cells(fila, 21 + col).Borders.Color = RGB(60, 50, 40)
        Next col
        fila = fila + 1
    Next pagina
    
    ws.Activate
    On Error Resume Next
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

' =============================================================================
' PARTE 2: MEJORAR HOJA CPU (Buses visuales + Teclado interactivo)
' =============================================================================
Private Sub MejorarHojaCPU()
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CPU")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    ' =================================================================
    ' A) DIAGRAMA DE BLOQUES CON BUSES (Filas 35+)
    ' =================================================================
    Dim startRow As Long
    startRow = 35
    
    ' Titulo del diagrama
    ws.Range(ws.Cells(startRow, 2), ws.Cells(startRow, 24)).Merge
    ws.Cells(startRow, 2).Value = "DIAGRAMA DE BLOQUES DEL SISTEMA Y BUSES DE INTERCONEXION"
    EstiloTituloMem ws.Range(ws.Cells(startRow, 2), ws.Cells(startRow, 24)), RGB(15, 76, 129)
    
    startRow = startRow + 2
    
    ' --- BLOQUE CPU/ALU (Izquierda arriba) ---
    DibujarBloque ws, startRow, 2, startRow + 4, 6, "CPU / ALU", _
        "AX BX CX DX" & vbCrLf & "PC IR MAR MDR" & vbCrLf & "Flags: ZF CF SF", _
        RGB(15, 76, 129), RGB(52, 152, 219)
    
    ' --- BLOQUE MEMORIA (Derecha arriba) ---
    DibujarBloque ws, startRow, 9, startRow + 4, 13, "MEMORIA RAM", _
        "256 Bytes Fisica" & vbCrLf & "8 Marcos x 32B" & vbCrLf & "Ver hoja MEMORIA", _
        RGB(39, 100, 60), RGB(46, 204, 113)
    
    ' --- BLOQUE CACHE (Centro arriba) ---
    DibujarBloque ws, startRow, 15, startRow + 4, 19, "CACHE L1/L2", _
        "L1I: 8B Instruc." & vbCrLf & "L1D: 8B Datos WB" & vbCrLf & "L2: 32B Unificada", _
        RGB(120, 40, 40), RGB(231, 76, 60)
    
    ' =================================================================
    ' BUSES HORIZONTALES (Las lineas que conectan todo)
    ' =================================================================
    Dim busRow As Long
    busRow = startRow + 5
    
    ' Bus de Control (Rojo)
    ws.Range(ws.Cells(busRow, 2), ws.Cells(busRow, 24)).Merge
    ws.Cells(busRow, 2).NumberFormat = "@"
    ws.Cells(busRow, 2).Value = "[ BUS DE CONTROL ]: [MEM_R | MEM_W | IO_R | IO_W | CLK]"
    ws.Cells(busRow, 2).Font.Name = "Consolas"
    ws.Cells(busRow, 2).Font.Size = 8
    ws.Cells(busRow, 2).Font.Bold = True
    ws.Cells(busRow, 2).Font.Color = RGB(255, 255, 255)
    ws.Cells(busRow, 2).Interior.Color = RGB(192, 57, 43)
    ws.Cells(busRow, 2).HorizontalAlignment = xlCenter
    
    busRow = busRow + 1
    
    ' Bus de Direcciones (Verde)
    ws.Range(ws.Cells(busRow, 2), ws.Cells(busRow, 24)).Merge
    ws.Cells(busRow, 2).NumberFormat = "@"
    ws.Cells(busRow, 2).Value = "[ BUS DE DIRECCIONES (16 bits) ]: [MAR -> Memoria / Perifericos]"
    ws.Cells(busRow, 2).Font.Name = "Consolas"
    ws.Cells(busRow, 2).Font.Size = 8
    ws.Cells(busRow, 2).Font.Bold = True
    ws.Cells(busRow, 2).Font.Color = RGB(255, 255, 255)
    ws.Cells(busRow, 2).Interior.Color = RGB(39, 174, 96)
    ws.Cells(busRow, 2).HorizontalAlignment = xlCenter
    
    busRow = busRow + 1
    
    ' Bus de Datos (Azul)
    ws.Range(ws.Cells(busRow, 2), ws.Cells(busRow, 24)).Merge
    ws.Cells(busRow, 2).NumberFormat = "@"
    ws.Cells(busRow, 2).Value = "[ BUS DE DATOS (8 bits, Bidireccional) ]: [MDR <-> Memoria / Perifericos]"
    ws.Cells(busRow, 2).Font.Name = "Consolas"
    ws.Cells(busRow, 2).Font.Size = 8
    ws.Cells(busRow, 2).Font.Bold = True
    ws.Cells(busRow, 2).Font.Color = RGB(255, 255, 255)
    ws.Cells(busRow, 2).Interior.Color = RGB(41, 128, 185)
    ws.Cells(busRow, 2).HorizontalAlignment = xlCenter
    
    busRow = busRow + 1
    
    ' Bus Interno (Morado)
    ws.Range(ws.Cells(busRow, 2), ws.Cells(busRow, 24)).Merge
    ws.Cells(busRow, 2).NumberFormat = "@"
    ws.Cells(busRow, 2).Value = "[ BUS INTERNO CPU (8 bits) ]: [ALU <-> Registros AX BX CX DX <-> MDR]"
    ws.Cells(busRow, 2).Font.Name = "Consolas"
    ws.Cells(busRow, 2).Font.Size = 8
    ws.Cells(busRow, 2).Font.Bold = True
    ws.Cells(busRow, 2).Font.Color = RGB(255, 255, 255)
    ws.Cells(busRow, 2).Interior.Color = RGB(142, 68, 173)
    ws.Cells(busRow, 2).HorizontalAlignment = xlCenter
    
    busRow = busRow + 2
    
    ' =================================================================
    ' BLOQUES INFERIORES: Monitor, Teclado, Disco/Swap
    ' =================================================================
    
    ' --- BLOQUE MONITOR / LCD ---
    DibujarBloque ws, busRow, 2, busRow + 3, 6, "MONITOR / LCD 8x1", _
        "Display: 8 caracteres" & vbCrLf & "Puerto: 0x03 (OUT)" & vbCrLf & "Buffer en RAM 0x78-0x7F", _
        RGB(22, 100, 83), RGB(26, 188, 156)
    
    ' --- BLOQUE TECLADO ---
    DibujarBloque ws, busRow, 9, busRow + 3, 13, "TECLADO (INPUT)", _
        "Cola FIFO: 8 bytes" & vbCrLf & "Puerto: 0x01/0x02 (IN)" & vbCrLf & "Buffer en RAM 0x70-0x77", _
        RGB(100, 80, 20), RGB(241, 196, 15)
    
    ' --- BLOQUE DISCO / SWAP ---
    DibujarBloque ws, busRow, 15, busRow + 3, 19, "DISCO / SWAP", _
        "512 Bytes virtuales" & vbCrLf & "16 Paginas x 32B" & vbCrLf & "Page Fault -> Swap I/O", _
        RGB(100, 40, 20), RGB(211, 84, 0)
    
    busRow = busRow + 5
    
    ' =================================================================
    ' B) TECLADO INTERACTIVO (Boton funcional)
    ' =================================================================
    ws.Range(ws.Cells(busRow, 2), ws.Cells(busRow, 19)).Merge
    ws.Cells(busRow, 2).Value = "TECLADO INTERACTIVO - PANEL DE ENTRADA"
    EstiloTituloMem ws.Range(ws.Cells(busRow, 2), ws.Cells(busRow, 19)), RGB(241, 196, 15)
    ws.Cells(busRow, 2).Font.Color = RGB(0, 0, 0)
    
    busRow = busRow + 1
    
    ' Instrucciones
    ws.Range(ws.Cells(busRow, 2), ws.Cells(busRow, 13)).Merge
    ws.Cells(busRow, 2).Value = "Haga clic en el boton [ESCRIBIR EN TECLADO] para enviar caracteres al buffer FIFO del teclado (Puerto 0x01/0x02). Luego ejecute el Programa 2 con RUN."
    ws.Cells(busRow, 2).Font.Size = 8
    ws.Cells(busRow, 2).Font.Color = RGB(180, 190, 200)
    ws.Cells(busRow, 2).WrapText = True
    ws.Rows(busRow).RowHeight = 30
    
    busRow = busRow + 1
    
    ' Cola FIFO visual (8 bytes)
    ws.Cells(busRow, 2).Value = "Cola FIFO:"
    ws.Cells(busRow, 2).Font.Bold = True
    ws.Cells(busRow, 2).Font.Color = RGB(241, 196, 15)
    
    Dim b As Long
    For b = 0 To 7
        ws.Cells(busRow, 3 + b).Value = "[ ]"
        ws.Cells(busRow, 3 + b).HorizontalAlignment = xlCenter
        ws.Cells(busRow, 3 + b).Font.Name = "Consolas"
        ws.Cells(busRow, 3 + b).Font.Size = 11
        ws.Cells(busRow, 3 + b).Font.Bold = True
        ws.Cells(busRow, 3 + b).Font.Color = RGB(241, 196, 15)
        ws.Cells(busRow, 3 + b).Interior.Color = RGB(40, 35, 20)
        ws.Cells(busRow, 3 + b).Borders.LineStyle = xlContinuous
        ws.Cells(busRow, 3 + b).Borders.Color = RGB(241, 196, 15)
    Next b
    ws.Cells(busRow, 11).Value = "Addr:"
    ws.Cells(busRow, 11).Font.Color = RGB(140, 150, 160)
    ws.Cells(busRow, 12).Value = "0x70-0x77"
    ws.Cells(busRow, 12).Font.Name = "Consolas"
    ws.Cells(busRow, 12).Font.Color = RGB(241, 196, 15)
    
    busRow = busRow + 1
    
    ' Boton de teclado
    Dim btnRng As Range
    Set btnRng = ws.Range(ws.Cells(busRow, 3), ws.Cells(busRow, 7))
    Dim btn As Button
    Set btn = ws.Buttons.Add(btnRng.Left + 2, btnRng.Top + 2, btnRng.Width - 4, btnRng.Height - 4)
    With btn
        .Caption = "ESCRIBIR EN TECLADO"
        .OnAction = "Boton_EscribirTeclado"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
    End With
    ws.Rows(busRow).RowHeight = 28
    
    ' =================================================================
    ' C) LEYENDA DE COLORES DE BUSES
    ' =================================================================
    busRow = busRow + 2
    ws.Range(ws.Cells(busRow, 2), ws.Cells(busRow, 13)).Merge
    ws.Cells(busRow, 2).Value = "LEYENDA DE BUSES DEL SISTEMA"
    EstiloTituloMem ws.Range(ws.Cells(busRow, 2), ws.Cells(busRow, 13)), RGB(52, 73, 94)
    
    busRow = busRow + 1
    ws.Cells(busRow, 2).Value = "  "
    ws.Cells(busRow, 2).Interior.Color = RGB(192, 57, 43)
    ws.Cells(busRow, 3).Value = "Bus de Control (MEM_R, MEM_W, IO_R, IO_W, CLK)"
    ws.Cells(busRow, 3).Font.Size = 8
    
    busRow = busRow + 1
    ws.Cells(busRow, 2).Value = "  "
    ws.Cells(busRow, 2).Interior.Color = RGB(39, 174, 96)
    ws.Cells(busRow, 3).Value = "Bus de Direcciones (16 bits, unidireccional: MAR -> Memoria)"
    ws.Cells(busRow, 3).Font.Size = 8
    
    busRow = busRow + 1
    ws.Cells(busRow, 2).Value = "  "
    ws.Cells(busRow, 2).Interior.Color = RGB(41, 128, 185)
    ws.Cells(busRow, 3).Value = "Bus de Datos Externo (8 bits, bidireccional: MDR <-> Memoria/Perifericos)"
    ws.Cells(busRow, 3).Font.Size = 8
    
    busRow = busRow + 1
    ws.Cells(busRow, 2).Value = "  "
    ws.Cells(busRow, 2).Interior.Color = RGB(142, 68, 173)
    ws.Cells(busRow, 3).Value = "Bus Interno CPU (8 bits: ALU <-> Registros <-> MDR, solo dentro del chip)"
    ws.Cells(busRow, 3).Font.Size = 8
End Sub

' =============================================================================
' DIBUJAR UN BLOQUE DE COMPONENTE (similar a la pizarra del docente)
' =============================================================================
Private Sub DibujarBloque(ByVal ws As Worksheet, _
    ByVal r1 As Long, ByVal c1 As Long, ByVal r2 As Long, ByVal c2 As Long, _
    ByVal titulo As String, ByVal contenido As String, _
    ByVal clrFondo As Long, ByVal clrBorde As Long)
    
    ' Titulo del bloque (primera fila)
    ws.Range(ws.Cells(r1, c1), ws.Cells(r1, c2)).Merge
    ws.Cells(r1, c1).Value = titulo
    ws.Cells(r1, c1).Font.Name = "Segoe UI"
    ws.Cells(r1, c1).Font.Size = 10
    ws.Cells(r1, c1).Font.Bold = True
    ws.Cells(r1, c1).Font.Color = RGB(255, 255, 255)
    ws.Cells(r1, c1).Interior.Color = clrFondo
    ws.Cells(r1, c1).HorizontalAlignment = xlCenter
    
    ' Contenido del bloque (filas siguientes)
    ws.Range(ws.Cells(r1 + 1, c1), ws.Cells(r2, c2)).Merge
    ws.Cells(r1 + 1, c1).Value = contenido
    ws.Cells(r1 + 1, c1).Font.Name = "Segoe UI"
    ws.Cells(r1 + 1, c1).Font.Size = 8
    ws.Cells(r1 + 1, c1).Font.Color = RGB(200, 210, 220)
    ws.Cells(r1 + 1, c1).Interior.Color = RGB(30, 38, 48)
    ws.Cells(r1 + 1, c1).HorizontalAlignment = xlCenter
    ws.Cells(r1 + 1, c1).VerticalAlignment = xlCenter
    ws.Cells(r1 + 1, c1).WrapText = True
    
    ' Borde del bloque completo
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(r1, c1), ws.Cells(r2, c2))
    With rng.Borders
        .LineStyle = xlContinuous
        .Color = clrBorde
        .Weight = xlMedium
    End With
End Sub

' =============================================================================
' ESTILO AUXILIAR
' =============================================================================
Private Sub EstiloTituloMem(ByVal rng As Range, ByVal clr As Long)
    With rng
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = clr
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub
