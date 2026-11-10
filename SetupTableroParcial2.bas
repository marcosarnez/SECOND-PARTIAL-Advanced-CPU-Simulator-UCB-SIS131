Attribute VB_Name = "SetupTableroParcial2"
Option Explicit

' ==============================================================================
' MODULO: SetupTableroParcial2.bas
' Generador del Tablero Visual - Segundo Parcial SIS-131
' Ejecutar: ALT+F8 > GenerarTableroSegundoParcial > Ejecutar
' ==============================================================================

Public Sub GenerarTableroSegundoParcial()
    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    ' --- Obtener o crear hoja CPU ---
    On Error Resume Next
    Set ws = wb.Worksheets("CPU")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(Before:=wb.Worksheets(1))
        ws.Name = "CPU"
    End If
    
    ws.Activate
    
    ' =====================================================================
    ' PASO PREVIO: DESPROTEGER Y DESBLOQUEAR TODAS LAS HOJAS DEL LIBRO
    ' =====================================================================
    Dim wsLoop As Worksheet
    For Each wsLoop In wb.Worksheets
        wsLoop.Unprotect ""
        wsLoop.Cells.Locked = False
    Next wsLoop
    
    ' =====================================================================
    ' PASO 0: LIMPIAR TODO (botones viejos, shapes, contenido, nombres)
    ' =====================================================================
    LimpiarHojaCompleta ws, wb
    
    ' =====================================================================
    ' PASO 1: CONFIGURAR COLUMNAS Y FILAS
    ' =====================================================================
    ConfigurarDimensiones ws
    
    ' =====================================================================
    ' PASO 2: PINTAR FONDO GENERAL
    ' =====================================================================
    ws.Range("A1:AQ60").Interior.Color = RGB(24, 28, 36)
    ws.Range("A1:AQ60").Font.Color = RGB(210, 215, 225)
    ws.Range("A1:AQ60").Font.Name = "Segoe UI"
    ws.Range("A1:AQ60").Font.Size = 9
    
    ' =====================================================================
    ' PASO 3: ENCABEZADO PRINCIPAL
    ' =====================================================================
    CrearEncabezado ws, wb
    
    ' =====================================================================
    ' PASO 4: REGISTROS CPU
    ' =====================================================================
    CrearPanelRegistros ws, wb
    
    ' =====================================================================
    ' PASO 5: BUSES TRI-ESTADO
    ' =====================================================================
    CrearPanelBuses ws, wb
    
    ' =====================================================================
    ' PASO 6: CACHE L1I, L1D, L2
    ' =====================================================================
    CrearPanelCache ws, wb
    
    ' =====================================================================
    ' PASO 7: MMU Y PAGE FAULT
    ' =====================================================================
    CrearPanelMMU ws, wb
    
    ' =====================================================================
    ' PASO 8: TABLA DE PAGINAS (16 VPNs)
    ' =====================================================================
    CrearTablaPaginas ws, wb
    
    ' =====================================================================
    ' PASO 9: PERIFERICOS (LCD + TECLADO)
    ' =====================================================================
    CrearPanelPerifericos ws, wb
    
    ' =====================================================================
    ' PASO 10: LOG DE EJECUCION
    ' =====================================================================
    CrearPanelLog ws, wb
    
    ' =====================================================================
    ' PASO 11: BOTONES INTERACTIVOS
    ' =====================================================================
    CrearBotones ws
    
    ' =====================================================================
    ' PASO 12: MATRIZ DE MEMORIA RAM 16x16 (256 BYTES)
    ' =====================================================================
    CrearPanelMemoriaRAM16x16 ws, wb
    
    ' =====================================================================
    ' PASO 13: ASEGURAR HOJA MEMORIA DETALLADA (RAM + SWAP) VISIBLE
    ' =====================================================================
    AsegurarHojaMemoriaVisible wb
    
    ' --- Desproteger y desbloquear todas las hojas del libro para edicion total ---
    For Each wsLoop In wb.Worksheets
        wsLoop.Unprotect ""
        wsLoop.Cells.Locked = False
    Next wsLoop
    
    ' --- Finalizar ---
    ws.Range("A1").Select
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    If Application.UserControl And Application.Visible Then
        MsgBox "Tablero del Segundo Parcial generado exitosamente!" & vbCrLf & vbCrLf & _
               "Haga clic en [Prog 1] para cargar un programa y luego [STEP] para ejecutar.", _
               vbInformation, "SIS-131 Simulador CPU"
    End If
End Sub

' =============================================================================
' LIMPIAR HOJA COMPLETA
' =============================================================================
Private Sub LimpiarHojaCompleta(ByVal ws As Worksheet, ByVal wb As Workbook)
    On Error Resume Next
    ws.Unprotect ""
    ws.Cells.Locked = False
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    
    ' Borrar contenido y formato del area usada
    ws.UsedRange.Clear
    ws.UsedRange.UnMerge
    On Error GoTo 0
End Sub

' =============================================================================
' CONFIGURAR DIMENSIONES DE COLUMNAS Y FILAS
' =============================================================================
Private Sub ConfigurarDimensiones(ByVal ws As Worksheet)
    ' Columna A = separador izquierdo
    ws.Columns("A").ColumnWidth = 2
    
    ' Bloque 1: REGISTROS CPU (B-F)
    ws.Columns("B").ColumnWidth = 7     ' Nombre reg
    ws.Columns("C").ColumnWidth = 8     ' Decimal
    ws.Columns("D").ColumnWidth = 8     ' Hex
    ws.Columns("E").ColumnWidth = 7     ' Binario
    ws.Columns("F").ColumnWidth = 16    ' Rol/Funcion
    
    ' Separador
    ws.Columns("G").ColumnWidth = 1.5
    
    ' Bloque 2: BUSES + PERIFERICOS (H-L)
    ws.Columns("H").ColumnWidth = 15    ' Nombre bus
    ws.Columns("I").ColumnWidth = 12    ' Valor
    ws.Columns("J").ColumnWidth = 12    ' Emisor
    ws.Columns("K").ColumnWidth = 10    ' Signal name
    ws.Columns("L").ColumnWidth = 6     ' Signal val
    
    ' Separador
    ws.Columns("M").ColumnWidth = 1.5
    
    ' Bloque 3: CACHE (N-R)
    ws.Columns("N").ColumnWidth = 16    ' Nivel
    ws.Columns("O").ColumnWidth = 8     ' Hits
    ws.Columns("P").ColumnWidth = 8     ' Misses
    ws.Columns("Q").ColumnWidth = 10    ' Hit Rate
    ws.Columns("R").ColumnWidth = 10    ' Status
    
    ' Separador
    ws.Columns("S").ColumnWidth = 1.5
    
    ' Bloque 4: PAGE TABLE (T-X)
    ws.Columns("T").ColumnWidth = 8     ' VPN
    ws.Columns("U").ColumnWidth = 5     ' P
    ws.Columns("V").ColumnWidth = 10    ' Ubicacion
    ws.Columns("W").ColumnWidth = 5     ' D
    ws.Columns("X").ColumnWidth = 5     ' R
    
    ' Separador
    ws.Columns("Y").ColumnWidth = 1.5
    
    ' Bloque 5: MEMORIA RAM FISICA 16x16 (Z-AP) - 256 Bytes
    ws.Columns("Z").ColumnWidth = 5.5   ' Direccion base (0, 10, 20... F0)
    Dim colMemIdx As Long
    For colMemIdx = 27 To 42            ' Col AA a AP (16 bytes: 0 a F)
        ws.Columns(colMemIdx).ColumnWidth = 4.2
    Next colMemIdx
    ws.Columns("AQ").ColumnWidth = 2    ' Margen derecho
    
    ' Filas
    ws.Rows("1").RowHeight = 6
    ws.Rows("2:3").RowHeight = 22
    ws.Rows("4").RowHeight = 6
    ws.Rows("5").RowHeight = 16
    ws.Rows("6:7").RowHeight = 26
    ws.Rows("8").RowHeight = 6
    
    Dim i As Long
    For i = 9 To 50
        ws.Rows(i).RowHeight = 17
    Next i
End Sub

' =============================================================================
' ENCABEZADO PRINCIPAL
' =============================================================================
Private Sub CrearEncabezado(ByVal ws As Worksheet, ByVal wb As Workbook)
    ' Banner superior 1: Identidad Corporativa y Logotipo
    Dim rBanner1 As Range
    Set rBanner1 = ws.Range("B2:AP2")
    rBanner1.Merge
    rBanner1.Value = "GERIBUS COMPUTING CORP.  |  DIGITAL COMPUTING ARCHITECTURE LAB — UCB SANTA CRUZ"
    With rBanner1
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(241, 196, 15) ' Dorado corporativo
        .Interior.Color = RGB(20, 30, 45)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Banner superior 2: Titulo del Simulador
    Dim rBanner2 As Range
    Set rBanner2 = ws.Range("B3:AP3")
    rBanner2.Merge
    rBanner2.Value = "SIMULADOR AVANZADO DE CPU 8-BITS  |  BUSES TRI-ESTADO  |  CACHE L1/L2  |  MEMORIA VIRTUAL  |  RAM 16x16"
    With rBanner2
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(15, 76, 129)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Sub-banner: Info academica y evaluacion
    Dim rSub As Range
    Set rSub = ws.Range("B4:AP4")
    rSub.Merge
    rSub.Value = "SIS-131 Arquitectura de Computadoras  |  Docente: Ing. Paulo Cesar Loayza Carrasco  |  Evaluacion Segundo Parcial"
    With rSub
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Color = RGB(160, 175, 190)
        .Interior.Color = RGB(28, 34, 46)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows("4").RowHeight = 15
    
    ' Titulos de secciones de botones
    Dim rCtrl As Range
    Set rCtrl = ws.Range("B5:F5")
    rCtrl.Merge
    rCtrl.Value = "CONTROLES DE EJECUCION"
    EstiloEtiqueta rCtrl, RGB(41, 128, 185)
    
    Dim rProg As Range
    Set rProg = ws.Range("H5:L5")
    rProg.Merge
    rProg.Value = "PROGRAMAS DE DEMOSTRACION"
    EstiloEtiqueta rProg, RGB(39, 174, 96)
    
    Dim rCiclo As Range
    Set rCiclo = ws.Range("N5:R5")
    rCiclo.Merge
    rCiclo.Value = "ESTADO DEL CICLO DE INSTRUCCION"
    EstiloEtiqueta rCiclo, RGB(142, 68, 173)
    
    ' Estado del Ciclo (valores)
    ws.Range("N6").Value = "Micro-paso:"
    ws.Range("N6").Font.Color = RGB(160, 165, 175)
    ws.Range("O6").Value = "T0"
    NombreRango wb, "MicroPaso_T", ws.Range("O6")
    ws.Range("O6").Font.Bold = True
    ws.Range("O6").Font.Color = RGB(46, 204, 113)
    
    ws.Range("P6").Value = "Fase:"
    ws.Range("P6").Font.Color = RGB(160, 165, 175)
    ws.Range("Q6:R6").Merge
    ws.Range("Q6").Value = "FETCH"
    NombreRango wb, "Fase_Nombre", ws.Range("Q6")
    ws.Range("Q6").Font.Bold = True
    ws.Range("Q6").Font.Color = RGB(52, 152, 219)
    
    ws.Range("N7").Value = "Instruccion:"
    ws.Range("N7").Font.Color = RGB(160, 165, 175)
    ws.Range("O7:P7").Merge
    ws.Range("O7").Value = "NOP"
    NombreRango wb, "Sel_Mnem", ws.Range("O7")
    ws.Range("O7").Font.Bold = True
    ws.Range("O7").Font.Color = RGB(241, 196, 15)
    
    ws.Range("Q7").Value = "Total:"
    ws.Range("Q7").Font.Color = RGB(160, 165, 175)
    ws.Range("R7").Value = 0
    NombreRango wb, "Total_Instrucciones", ws.Range("R7")
    ws.Range("R7").Font.Bold = True
    
    ' Info Page Fault grande (derecha del encabezado)
    Dim rPF As Range
    Set rPF = ws.Range("T5:X7")
    rPF.Merge
    rPF.Value = "NORMAL"
    NombreRango wb, "Alerta_PageFault", ws.Range("T5")
    With rPF
        .Font.Name = "Segoe UI"
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = RGB(46, 204, 113)
        .Interior.Color = RGB(30, 39, 46)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(46, 204, 113)
        .Borders.Weight = xlMedium
    End With
    
    ' Encabezado resumen para la Matriz de Memoria (Columnas Z a AP, Filas 5 a 7)
    Dim rMemTop As Range
    Set rMemTop = ws.Range("Z5:AP5")
    rMemTop.Merge
    rMemTop.Value = "MAPA FISICO DE MEMORIA RAM (256 BYTES)"
    EstiloEtiqueta rMemTop, RGB(41, 128, 185)
    
    ws.Range("Z6:AP6").Merge
    ws.Range("Z6").Value = "Capacidad: 256 Bytes  |  8 Marcos de 32 Bytes (M0..M7)  |  0x00-0x7F Codigo / 0x80-0xFF Datos"
    ws.Range("Z6").Font.Size = 8
    ws.Range("Z6").Font.Color = RGB(160, 175, 190)
    ws.Range("Z6").HorizontalAlignment = xlCenter
    ws.Range("Z6").VerticalAlignment = xlCenter
    
    ws.Range("Z7:AP7").Merge
    ws.Range("Z7").Value = "Direccion Activa apuntada por MAR/PC resaltada en AMARILLO dinamicamente"
    ws.Range("Z7").Font.Size = 8
    ws.Range("Z7").Font.Bold = True
    ws.Range("Z7").Font.Color = RGB(241, 196, 15)
    ws.Range("Z7").HorizontalAlignment = xlCenter
    ws.Range("Z7").VerticalAlignment = xlCenter
    
    ws.Range("Z6:AP6").Borders.LineStyle = xlContinuous
    ws.Range("Z6:AP6").Borders.Color = RGB(50, 60, 70)
    ws.Range("Z7:AP7").Borders.LineStyle = xlContinuous
    ws.Range("Z7:AP7").Borders.Color = RGB(50, 60, 70)
End Sub

' =============================================================================
' PANEL DE REGISTROS CPU (Columnas B-F, Filas 9+)
' =============================================================================
Private Sub CrearPanelRegistros(ByVal ws As Worksheet, ByVal wb As Workbook)
    ' Titulo
    Dim rTit As Range
    Set rTit = ws.Range("B9:F9")
    rTit.Merge
    rTit.Value = "BANCO DE REGISTROS (CPU 8-BIT)"
    EstiloTitulo rTit, RGB(15, 76, 129)
    
    ' Encabezados
    ws.Range("B10").Value = "REG"
    ws.Range("C10").Value = "DEC"
    ws.Range("D10").Value = "HEX"
    ws.Range("E10").Value = "BIN"
    ws.Range("F10").Value = "FUNCION"
    EstiloEncabezado ws.Range("B10:F10")
    
    ' Datos de registros
    Dim regData As Variant
    regData = Array( _
        Array("AX", "AX_Dec", "AX_Hex", "AX_Bin", "Acumulador", 11), _
        Array("BX", "BX_Dec", "BX_Hex", "BX_Bin", "Base / Indice", 12), _
        Array("CX", "CX_Dec", "CX_Hex", "CX_Bin", "Contador Loop", 13), _
        Array("DX", "DX_Dec", "DX_Hex", "DX_Bin", "Datos / E-S", 14), _
        Array("PC", "PC_Dec", "PC_Hex", "PC_Bin", "Program Counter", 16), _
        Array("IR", "IR_Dec", "IR_Hex", "IR_Bin", "Instr. Register", 17), _
        Array("MAR", "MAR_Dec", "MAR_Hex", "MAR_Bin", "Addr Register", 18), _
        Array("MDR", "MDR_Dec", "MDR_Hex", "MDR_Bin", "Data Register", 19) _
    )
    
    Dim k As Long
    For k = LBound(regData) To UBound(regData)
        Dim r As Long
        r = regData(k)(5)
        
        ' Separador visual entre generales y de control
        If r = 16 Then
            ws.Range("B15:F15").Merge
            ws.Range("B15").Value = "REGISTROS DE CONTROL"
            EstiloEtiqueta ws.Range("B15:F15"), RGB(52, 73, 94)
        End If
        
        ws.Cells(r, 2).Value = regData(k)(0)
        ws.Cells(r, 2).Font.Bold = True
        ws.Cells(r, 2).Font.Color = RGB(52, 152, 219)
        
        ws.Cells(r, 3).Value = 0
        NombreRango wb, CStr(regData(k)(1)), ws.Cells(r, 3)
        
        ws.Cells(r, 4).Value = "0x00"
        NombreRango wb, CStr(regData(k)(2)), ws.Cells(r, 4)
        ws.Cells(r, 4).Font.Name = "Consolas"
        
        ws.Cells(r, 5).Value = "00000000"
        NombreRango wb, CStr(regData(k)(3)), ws.Cells(r, 5)
        ws.Cells(r, 5).Font.Name = "Consolas"
        ws.Cells(r, 5).Font.Size = 8
        
        ws.Cells(r, 6).Value = regData(k)(4)
        ws.Cells(r, 6).Font.Size = 8
        ws.Cells(r, 6).Font.Italic = True
        ws.Cells(r, 6).Font.Color = RGB(140, 150, 160)
        
        EstiloFila ws.Range(ws.Cells(r, 2), ws.Cells(r, 6))
    Next k
    
    ' FLAGS
    ws.Range("B21:F21").Merge
    ws.Range("B21").Value = "BANDERAS ALU (FLAGS)"
    EstiloEtiqueta ws.Range("B21:F21"), RGB(52, 73, 94)
    
    ws.Range("B22").Value = "ZF"
    ws.Range("B22").Font.Bold = True
    ws.Range("C22").Value = 0
    NombreRango wb, "Flag_ZF", ws.Range("C22")
    
    ws.Range("D22").Value = "CF"
    ws.Range("D22").Font.Bold = True
    ws.Range("E22").Value = 0
    NombreRango wb, "Flag_CF", ws.Range("E22")
    
    ws.Range("B23").Value = "SF"
    ws.Range("B23").Font.Bold = True
    ws.Range("C23").Value = 0
    NombreRango wb, "Flag_SF", ws.Range("C23")
    
    ws.Range("D23").Value = "Status:"
    ws.Range("D23").Font.Color = RGB(140, 150, 160)
    ws.Range("E23:F23").Merge
    ws.Range("E23").Value = "Sistema listo"
    NombreRango wb, "Status_Msg", ws.Range("E23")
    ws.Range("E23").Font.Size = 8
    
    EstiloFila ws.Range("B22:F23")
End Sub

' =============================================================================
' PANEL DE BUSES TRI-ESTADO (Columnas H-L, Filas 9+)
' =============================================================================
Private Sub CrearPanelBuses(ByVal ws As Worksheet, ByVal wb As Workbook)
    ' Titulo
    Dim rTit As Range
    Set rTit = ws.Range("H9:L9")
    rTit.Merge
    rTit.Value = "BUSES DEL SISTEMA (TRI-STATE / HIGH-Z)"
    EstiloTitulo rTit, RGB(211, 84, 0)
    
    ' Encabezados
    ws.Range("H10").Value = "BUS"
    ws.Range("I10:J10").Merge
    ws.Range("I10").Value = "VALOR / EMISOR"
    ws.Range("K10:L10").Merge
    ws.Range("K10").Value = "ESTADO"
    EstiloEncabezado ws.Range("H10:L10")
    
    ' Bus Interno
    ws.Range("H11").Value = "Bus Interno"
    ws.Range("H11").Font.Bold = True
    ws.Range("I11:J11").Merge
    ws.Range("I11").Value = "High-Z"
    ws.Range("I11").Font.Name = "Consolas"
    NombreRango wb, "Bus_Interno_Val", ws.Range("I11")
    ws.Range("K11:L11").Merge
    ws.Range("K11").Value = "IDLE"
    ws.Range("K11").Font.Color = RGB(149, 165, 166)
    
    ' Bus Direcciones
    ws.Range("H12").Value = "Bus Direcciones"
    ws.Range("H12").Font.Bold = True
    ws.Range("I12:J12").Merge
    ws.Range("I12").Value = "High-Z"
    ws.Range("I12").Font.Name = "Consolas"
    NombreRango wb, "Bus_Dir_Val", ws.Range("I12")
    ws.Range("K12:L12").Merge
    ws.Range("K12").Value = "IDLE"
    ws.Range("K12").Font.Color = RGB(149, 165, 166)
    
    ' Bus Datos Ext
    ws.Range("H13").Value = "Bus Datos Ext"
    ws.Range("H13").Font.Bold = True
    ws.Range("I13:J13").Merge
    ws.Range("I13").Value = "High-Z"
    ws.Range("I13").Font.Name = "Consolas"
    NombreRango wb, "Bus_Datos_Val", ws.Range("I13")
    ws.Range("K13:L13").Merge
    ws.Range("K13").Value = "IDLE"
    ws.Range("K13").Font.Color = RGB(149, 165, 166)
    
    EstiloFila ws.Range("H11:L13")
    
    ' SENALES DE CONTROL
    ws.Range("H15:L15").Merge
    ws.Range("H15").Value = "SENALES DEL BUS DE CONTROL"
    EstiloEtiqueta ws.Range("H15:L15"), RGB(52, 73, 94)
    
    ws.Range("H16").Value = "MEM_R"
    ws.Range("I16").Value = 0
    NombreRango wb, "Sig_MemRead", ws.Range("I16")
    ws.Range("J16").Value = "MEM_W"
    ws.Range("K16").Value = 0
    NombreRango wb, "Sig_MemWrite", ws.Range("K16")
    
    ws.Range("H17").Value = "IO_R"
    ws.Range("I17").Value = 0
    NombreRango wb, "Sig_IoRead", ws.Range("I17")
    ws.Range("J17").Value = "IO_W"
    ws.Range("K17").Value = 0
    NombreRango wb, "Sig_IoWrite", ws.Range("K17")
    EstiloFila ws.Range("H16:L17")
    
    ' MICRO-PALABRA
    ws.Range("H19:L19").Merge
    ws.Range("H19").Value = "MICRO-PALABRA DE CONTROL (CONTROL WORD)"
    EstiloEtiqueta ws.Range("H19:L19"), RGB(52, 73, 94)
    
    ws.Range("H20:L21").Merge
    ws.Range("H20").Value = "[En_PC=0 Ld_MAR=0 Mem_R=0 Ld_MDR=0 En_MDR=0 Ld_IR=0 ...]"
    NombreRango wb, "MicroPalabra_Bin", ws.Range("H20")
    ws.Range("H20").Font.Name = "Consolas"
    ws.Range("H20").Font.Size = 8
    ws.Range("H20").Font.Color = RGB(241, 196, 15)
    ws.Range("H20").WrapText = True
    ws.Range("H20").HorizontalAlignment = xlCenter
    ws.Range("H20").VerticalAlignment = xlCenter
    EstiloBorde ws.Range("H20:L21"), RGB(60, 70, 80)
End Sub

' =============================================================================
' PANEL DE CACHE (Columnas N-R, Filas 9+)
' =============================================================================
Private Sub CrearPanelCache(ByVal ws As Worksheet, ByVal wb As Workbook)
    ' Titulo
    Dim rTit As Range
    Set rTit = ws.Range("N9:R9")
    rTit.Merge
    rTit.Value = "JERARQUIA DE CACHE (L1I / L1D / L2)"
    EstiloTitulo rTit, RGB(192, 57, 43)
    
    ' Encabezados
    ws.Range("N10").Value = "NIVEL"
    ws.Range("O10").Value = "HITS"
    ws.Range("P10").Value = "MISSES"
    ws.Range("Q10").Value = "HIT RATE"
    ws.Range("R10").Value = "STATUS"
    EstiloEncabezado ws.Range("N10:R10")
    
    ' L1I
    ws.Range("N11").Value = "L1I (Instr. 8B)"
    ws.Range("N11").Font.Bold = True
    ws.Range("O11").Value = 0
    NombreRango wb, "L1I_Hits", ws.Range("O11")
    ws.Range("P11").Value = 0
    NombreRango wb, "L1I_Misses", ws.Range("P11")
    ws.Range("Q11").Value = "0.0%"
    NombreRango wb, "L1I_HitRate", ws.Range("Q11")
    ws.Range("R11").Value = "IDLE"
    NombreRango wb, "L1I_Status", ws.Range("R11")
    ws.Range("R11").Font.Color = RGB(149, 165, 166)
    
    ' L1D
    ws.Range("N12").Value = "L1D (Data 8B)"
    ws.Range("N12").Font.Bold = True
    ws.Range("O12").Value = 0
    NombreRango wb, "L1D_Hits", ws.Range("O12")
    ws.Range("P12").Value = 0
    NombreRango wb, "L1D_Misses", ws.Range("P12")
    ws.Range("Q12").Value = "0.0%"
    NombreRango wb, "L1D_HitRate", ws.Range("Q12")
    ws.Range("R12").Value = "IDLE"
    NombreRango wb, "L1D_Status", ws.Range("R12")
    ws.Range("R12").Font.Color = RGB(149, 165, 166)
    
    ' L2
    ws.Range("N13").Value = "L2 (Unif. 32B)"
    ws.Range("N13").Font.Bold = True
    ws.Range("O13").Value = 0
    NombreRango wb, "L2_Hits", ws.Range("O13")
    ws.Range("P13").Value = 0
    NombreRango wb, "L2_Misses", ws.Range("P13")
    ws.Range("Q13").Value = "0.0%"
    NombreRango wb, "L2_HitRate", ws.Range("Q13")
    ws.Range("R13").Value = "IDLE"
    NombreRango wb, "L2_Status", ws.Range("R13")
    ws.Range("R13").Font.Color = RGB(149, 165, 166)
    
    EstiloFila ws.Range("N11:R13")
    
    ' Resumen global
    ws.Range("N15:R15").Merge
    ws.Range("N15").Value = "RESUMEN GLOBAL DE CACHE"
    EstiloEtiqueta ws.Range("N15:R15"), RGB(52, 73, 94)
    
    ws.Range("N16").Value = "Global Hit Rate:"
    ws.Range("N16").Font.Color = RGB(160, 165, 175)
    ws.Range("O16:P16").Merge
    ws.Range("O16").Value = "0.0%"
    NombreRango wb, "Global_HitRate", ws.Range("O16")
    ws.Range("O16").Font.Bold = True
    ws.Range("O16").Font.Size = 11
    ws.Range("O16").Font.Color = RGB(46, 204, 113)
    
    ws.Range("Q16").Value = "Evento:"
    ws.Range("Q16").Font.Color = RGB(160, 165, 175)
    ws.Range("R16").Value = "INIT"
    NombreRango wb, "Cache_Evento", ws.Range("R16")
    ws.Range("R16").Font.Bold = True
    EstiloFila ws.Range("N16:R16")
    
    ' MMU Telemetria debajo de cache
    ws.Range("N18:R18").Merge
    ws.Range("N18").Value = "TELEMETRIA MMU (MEMORIA VIRTUAL)"
    EstiloEtiqueta ws.Range("N18:R18"), RGB(52, 73, 94)
    
    ws.Range("N19").Value = "Accesos Mem:"
    ws.Range("N19").Font.Color = RGB(160, 165, 175)
    ws.Range("O19").Value = 0
    NombreRango wb, "Total_Accesos_Mem", ws.Range("O19")
    ws.Range("O19").Font.Bold = True
    
    ws.Range("P19").Value = "Page Faults:"
    ws.Range("P19").Font.Color = RGB(160, 165, 175)
    ws.Range("Q19").Value = 0
    NombreRango wb, "Total_Page_Faults", ws.Range("Q19")
    ws.Range("Q19").Font.Bold = True
    ws.Range("Q19").Font.Color = RGB(231, 76, 60)
    EstiloFila ws.Range("N19:R19")
    
    ' Mensaje Page Fault
    ws.Range("N20:R21").Merge
    ws.Range("N20").Value = "Sin fallos de pagina detectados."
    NombreRango wb, "Msg_Page_Fault", ws.Range("N20")
    ws.Range("N20").Font.Size = 8
    ws.Range("N20").Font.Color = RGB(140, 150, 160)
    ws.Range("N20").WrapText = True
    EstiloBorde ws.Range("N20:R21"), RGB(60, 70, 80)
End Sub

' =============================================================================
' PANEL MMU (no necesario, integrado en cache)
' =============================================================================
Private Sub CrearPanelMMU(ByVal ws As Worksheet, ByVal wb As Workbook)
    ' Ya integrado arriba
End Sub

' =============================================================================
' TABLA DE PAGINAS (Columnas T-X, Filas 9+)
' =============================================================================
Private Sub CrearTablaPaginas(ByVal ws As Worksheet, ByVal wb As Workbook)
    ' Titulo
    Dim rTit As Range
    Set rTit = ws.Range("T9:X9")
    rTit.Merge
    rTit.Value = "TABLA DE PAGINAS (16 VPNs -> 8 MARCOS)"
    EstiloTitulo rTit, RGB(44, 62, 80)
    
    ' Encabezados
    ws.Range("T10").Value = "VPN"
    ws.Range("U10").Value = "P"
    ws.Range("V10").Value = "UBICACION"
    ws.Range("W10").Value = "D"
    ws.Range("X10").Value = "R"
    EstiloEncabezado ws.Range("T10:X10")
    
    Dim p As Long
    For p = 0 To 15
        Dim rowP As Long
        rowP = 11 + p
        
        ws.Cells(rowP, 20).Value = "VPN " & p
        ws.Cells(rowP, 20).Font.Bold = True
        ws.Cells(rowP, 20).Font.Size = 8
        
        If p < 8 Then
            ws.Cells(rowP, 21).Value = 1
            ws.Cells(rowP, 21).Font.Color = RGB(46, 204, 113)
            ws.Cells(rowP, 22).Value = "Marco " & p
            ws.Cells(rowP, 22).Font.Color = RGB(46, 204, 113)
        Else
            ws.Cells(rowP, 21).Value = 0
            ws.Cells(rowP, 21).Font.Color = RGB(231, 76, 60)
            ws.Cells(rowP, 22).Value = "Swap"
            ws.Cells(rowP, 22).Font.Color = RGB(231, 76, 60)
        End If
        
        ws.Cells(rowP, 23).Value = 0
        ws.Cells(rowP, 24).Value = 0
        
        ws.Cells(rowP, 22).Font.Size = 8
        
        EstiloFila ws.Range(ws.Cells(rowP, 20), ws.Cells(rowP, 24))
    Next p
End Sub

' =============================================================================
' PANEL DE PERIFERICOS (Columnas H-L, Filas 23+)
' =============================================================================
Private Sub CrearPanelPerifericos(ByVal ws As Worksheet, ByVal wb As Workbook)
    ' Titulo
    ws.Range("H23:L23").Merge
    ws.Range("H23").Value = "PERIFERICOS MAPEADOS EN MEMORIA (E/S)"
    EstiloTitulo ws.Range("H23:L23"), RGB(22, 160, 133)
    
    ' LCD Display
    ws.Range("H24").Value = "LCD 8x1:"
    ws.Range("H24").Font.Bold = True
    ws.Range("H24").Font.Color = RGB(160, 165, 175)
    
    ws.Range("I24:L24").Merge
    ws.Range("I24").NumberFormat = "@"
    ws.Range("I24").Value = "[ READY ]"
    NombreRango wb, "LCD_Display", ws.Range("I24")
    With ws.Range("I24")
        .Font.Name = "Consolas"
        .Font.Size = 14
        .Font.Bold = True
        .Font.Color = RGB(0, 40, 0)
        .Interior.Color = RGB(46, 204, 113)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows("24").RowHeight = 28
    
    ' Teclado info
    ws.Range("H25").Value = "KBD Status:"
    ws.Range("H25").Font.Color = RGB(160, 165, 175)
    ws.Range("I25").Value = "0"
    NombreRango wb, "KBD_Status", ws.Range("I25")
    
    ws.Range("J25").Value = "Dato:"
    ws.Range("J25").Font.Color = RGB(160, 165, 175)
    ws.Range("K25:L25").Merge
    ws.Range("K25").Value = "0x00 (' ')"
    ws.Range("K25").Font.Name = "Consolas"
    NombreRango wb, "KBD_Data", ws.Range("K25")
    
    ws.Range("H26").Value = "Cola Head:"
    ws.Range("H26").Font.Color = RGB(160, 165, 175)
    ws.Range("I26").Value = 0
    NombreRango wb, "Cola_Head", ws.Range("I26")
    
    ws.Range("J26").Value = "Tail/Count:"
    ws.Range("J26").Font.Color = RGB(160, 165, 175)
    ws.Range("K26:L26").Merge
    ws.Range("K26").Value = "Tail: 0 | Cnt: 0"
    NombreRango wb, "Cola_Tail", ws.Range("K26")
    
    EstiloFila ws.Range("H25:L26")
End Sub

' =============================================================================
' PANEL DE LOG (Columnas B-R, Filas 25+)
' =============================================================================
Private Sub CrearPanelLog(ByVal ws As Worksheet, ByVal wb As Workbook)
    ' Titulo
    ws.Range("B25:F25").Merge
    ws.Range("B25").Value = "LOG DE EJECUCION Y TRAZA"
    EstiloTitulo ws.Range("B25:F25"), RGB(52, 73, 94)
    
    Dim i As Long
    For i = 1 To 8
        Dim rowL As Long
        rowL = 25 + i
        ws.Range(ws.Cells(rowL, 2), ws.Cells(rowL, 6)).Merge
        ws.Cells(rowL, 2).Value = "[INIT] Sistema inicializado. Esperando programa..."
        ws.Cells(rowL, 2).Font.Name = "Consolas"
        ws.Cells(rowL, 2).Font.Size = 8
        ws.Cells(rowL, 2).Font.Color = RGB(120, 130, 140)
        NombreRango wb, "Log_" & i, ws.Cells(rowL, 2)
    Next i
    EstiloBorde ws.Range(ws.Cells(26, 2), ws.Cells(33, 6)), RGB(50, 60, 70)
End Sub

' =============================================================================
' BOTONES INTERACTIVOS (FORM CONTROLS)
' =============================================================================
Private Sub CrearBotones(ByVal ws As Worksheet)
    On Error Resume Next
    
    ' --- Botones de Ejecucion (Fila 6-7, Columnas B-F) ---
    Dim colores As Variant
    
    ' STEP (azul)
    CrearUnBoton ws, ws.Range("B6:C6"), "STEP", "Boton_STEP", RGB(52, 152, 219)
    ' RUN (verde)
    CrearUnBoton ws, ws.Range("D6:E6"), "RUN", "Boton_RUN", RGB(46, 204, 113)
    ' PAUSE (amarillo)
    CrearUnBoton ws, ws.Range("B7:C7"), "PAUSE", "Boton_PAUSE", RGB(241, 196, 15)
    ' RESET (rojo)
    CrearUnBoton ws, ws.Range("D7:E7"), "RESET", "Boton_RESET", RGB(231, 76, 60)
    
    ' --- Parametros Editables de Ejecucion (Columna F, Filas 6 y 7) ---
    ' F6: Cantidad de Pasos a dar con cada pulsacion de [STEP] (o lote de pasos)
    With ws.Range("F6")
        .Value = 1
        .NumberFormat = """Pasos: ""0"
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(241, 196, 15) ' Dorado llamativo
        .Interior.Color = RGB(35, 45, 60)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(70, 90, 115)
        .Borders.Weight = xlThin
        .Locked = False
    End With
    NombreRango ws.Parent, "Config_Pasos", ws.Range("F6")
    
    ' F7: Retardo / Delay en ms para [RUN] (0 = Turbo / Maxima velocidad, 50, 100, 300)
    With ws.Range("F7")
        .Value = 50
        .NumberFormat = """Delay: ""0"" ms"""
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(100, 220, 255) ' Cyan llamativo
        .Interior.Color = RGB(35, 45, 60)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(70, 90, 115)
        .Borders.Weight = xlThin
        .Locked = False
    End With
    NombreRango ws.Parent, "Config_Delay", ws.Range("F7")
    
    ' --- Botones de Programas (Fila 6-7, Columnas H-L) ---
    CrearUnBoton ws, ws.Range("H6:I6"), "Prog 1: Aritmetica", "Cargar_Programa_1", RGB(39, 174, 96)
    CrearUnBoton ws, ws.Range("J6:L6"), "Prog 2: Teclado / LCD", "Cargar_Programa_2", RGB(39, 174, 96)
    CrearUnBoton ws, ws.Range("H7:L7"), "Prog 3: Page Fault + Cache Miss", "Cargar_Programa_3", RGB(39, 174, 96)
    
    On Error GoTo 0
End Sub

Private Sub CrearUnBoton(ByVal ws As Worksheet, ByVal rng As Range, ByVal texto As String, ByVal macroName As String, ByVal clr As Long)
    Dim btn As Button
    Set btn = ws.Buttons.Add(rng.Left + 2, rng.Top + 2, rng.Width - 4, rng.Height - 4)
    With btn
        .Caption = texto
        .OnAction = macroName
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
    End With
End Sub

' =============================================================================
' ESTILOS REUTILIZABLES
' =============================================================================
Private Sub EstiloTitulo(ByVal rng As Range, ByVal clr As Long)
    With rng
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = clr
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub EstiloEtiqueta(ByVal rng As Range, ByVal clr As Long)
    With rng
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = RGB(200, 210, 220)
        .Interior.Color = clr
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub EstiloEncabezado(ByVal rng As Range)
    With rng
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = RGB(200, 210, 220)
        .Interior.Color = RGB(40, 50, 60)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub EstiloFila(ByVal rng As Range)
    With rng
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(50, 60, 70)
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Locked = False
    End With
End Sub

Private Sub EstiloBorde(ByVal rng As Range, ByVal clr As Long)
    With rng
        .Borders.LineStyle = xlContinuous
        .Borders.Color = clr
        .Borders.Weight = xlThin
    End With
End Sub

Private Sub NombreRango(ByVal wb As Workbook, ByVal nombre As String, ByVal rng As Range)
    On Error Resume Next
    wb.Names(nombre).Delete
    On Error GoTo 0
    wb.Names.Add Name:=nombre, RefersTo:=rng
End Sub

' =============================================================================
' PANEL DE MEMORIA RAM FISICA 16x16 (Columnas Z-AP, Filas 9-27)
' Representacion visual exacta del 1er Parcial, adaptada al 2do Parcial
' =============================================================================
Private Sub CrearPanelMemoriaRAM16x16(ByVal ws As Worksheet, ByVal wb As Workbook)
    On Error Resume Next
    ' Titulo Principal del Panel
    Dim rTit As Range
    Set rTit = ws.Range("Z9:AP9")
    rTit.Merge
    rTit.Value = "MATRIZ DE MEMORIA RAM FISICA (16x16 = 256 Bytes)"
    EstiloTitulo rTit, RGB(15, 76, 129)
    
    ' Encabezados de Columna (Addr, 0 a F)
    ws.Cells(10, 26).Value = "Addr"
    
    Dim c As Long
    For c = 0 To 15
        ws.Cells(10, 27 + c).Value = Hex$(c)
    Next c
    
    With ws.Range(ws.Cells(10, 26), ws.Cells(10, 42))
        .Font.Name = "Consolas"
        .Font.Bold = True
        .Font.Size = 8
        .Font.Color = RGB(200, 210, 220)
        .Interior.Color = RGB(40, 50, 60)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Filas 0 a F0 (Columna Z, Filas Excel 11 a 26)
    Dim r As Long, rowExcel As Long
    For r = 0 To 15
        rowExcel = 11 + r
        If r = 0 Then
            ws.Cells(rowExcel, 26).Value = "0"
        Else
            ws.Cells(rowExcel, 26).Value = Hex$(r) & "0"
        End If
    Next r
    
    With ws.Range(ws.Cells(11, 26), ws.Cells(26, 26))
        .Font.Name = "Consolas"
        .Font.Bold = True
        .Font.Size = 8
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(26, 54, 93)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' Llenado y formateo en bloque de la matriz de datos (Filas 11-26, Columnas 27-42)
    Dim arr00(1 To 16, 1 To 16) As Variant
    For r = 1 To 16
        For c = 1 To 16
            arr00(r, c) = "00"
        Next c
    Next r
    
    With ws.Range(ws.Cells(11, 27), ws.Cells(26, 42))
        .Value = arr00
        .Font.Name = "Consolas"
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(50, 65, 80)
        .Borders.Weight = xlThin
        .Locked = False
    End With
    
    ' Colorear en bloque: Filas 11-18 Azul (Codigo 0x00-0x7F, Marcos 0-3)
    With ws.Range(ws.Cells(11, 27), ws.Cells(18, 42))
        .Interior.Color = RGB(22, 36, 52)
        .Font.Color = RGB(147, 197, 253)
    End With
    
    ' Colorear en bloque: Filas 19-26 Verde (Datos 0x80-0xFF, Marcos 4-7)
    With ws.Range(ws.Cells(19, 27), ws.Cells(26, 42))
        .Interior.Color = RGB(18, 42, 30)
        .Font.Color = RGB(134, 239, 172)
    End With
    
    ' Bordes de columna Z de direcciones
    With ws.Range(ws.Cells(10, 26), ws.Cells(26, 26))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(50, 65, 80)
        .Borders.Weight = xlThin
    End With
    
    ' Pie / Leyenda Informativa en Fila 27
    ws.Range("Z27:AP27").Merge
    ws.Range("Z27").Value = "Azul: Marcos 0-3 (Codigo/Instr. 0x00-0x7F)  |  Verde: Marcos 4-7 (Datos/E-S 0x80-0xFF)  |  Amarillo: Cursor Activo (MAR/PC)"
    With ws.Range("Z27")
        .Font.Name = "Segoe UI"
        .Font.Size = 7
        .Font.Italic = True
        .Font.Color = RGB(150, 160, 175)
        .Interior.Color = RGB(24, 28, 36)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    On Error GoTo 0
End Sub

Private Sub AsegurarHojaMemoriaVisible(ByVal wb As Workbook)
    On Error Resume Next
    Dim wsMem As Worksheet
    Set wsMem = wb.Worksheets("MEMORIA")
    If wsMem Is Nothing Then Set wsMem = wb.Worksheets("Memoria")
    If Not wsMem Is Nothing Then
        wsMem.Visible = -1 ' xlSheetVisible
    End If
    On Error GoTo 0
End Sub

' ==============================================================================
' MACRO MAESTRA: ARMAR TODO EL SIMULADOR DE UN SOLO CLIC
' ==============================================================================
Public Sub IniciarTodoElSimulador()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' 1. Generar tablero principal oscuro en hoja CPU
    GenerarTableroSegundoParcial
    
    ' 2. Generar hoja MEMORIA (RAM Fisica 256B y Area de Swap 512B)
    GenerarMemoriaSwapYBuses
    
    ' 3. Generar hoja ARQUITECTURA (Diagrama de la pizarra del profesor)
    GenerarArquitecturaPizarraCompleta
    
    ' 4. Cargar Programa 1 de Aritmetica por defecto
    CargarPrograma1_Aritmetica
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    If Application.UserControl And Application.Visible Then
        MsgBox "SIMULADOR COMPLETAMENTE ARMADO Y LISTO:" & vbCrLf & vbCrLf & _
               "1. Hoja 'CPU': Tablero interactivo con STEP, RUN, PAUSE, RESET y 3 programas." & vbCrLf & _
               "2. Hoja 'MEMORIA': RAM Fisica (256 bytes) y Swap (512 bytes) que se actualizan en vivo." & vbCrLf & _
               "3. Hoja 'ARQUITECTURA': Diagrama de la pizarra con Buses, Teclado manejable y Editor de programas.", _
               vbInformation, "SIS-131 Arquitectura de Computadoras"
    End If
End Sub
