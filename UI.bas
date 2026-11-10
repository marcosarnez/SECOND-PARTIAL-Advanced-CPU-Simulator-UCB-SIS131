Attribute VB_Name = "UI"
Option Explicit

' ==============================================================================
' MODULO: UI.bas
' Simulador Avanzado de Arquitectura de Computadoras - Segundo Parcial (SIS-131)
' Control Grafico y Visualizacion en Microsoft Excel:
'   1. Registros Generales: AX, BX, CX, DX y Registros de Control (PC, IR, MAR, MDR).
'   2. Visualizacion de Buses y Buffers Tri-Estado (High-Z dinamico).
'   3. Micro-palabra de Control y Senales (MEM_R, MEM_W, IO_R, IO_W).
'   4. Panel de Jerarquia de Cache L1I, L1D, L2 con contadores y Hit Rates.
'   5. Tabla de Paginas (16 VPNs) y alerta visual de PAGE FAULT en Rojo.
'   6. Display Pantalla LCD 8x1 y Panel interactivo de Teclado.
'   7. Log de eventos y traza de ejecucion.
'   8. Actualizacion en vivo de la hoja MEMORIA (RAM Fisica + Swap).
'   9. Senales del Bus de Control visuales.
' ==============================================================================

Private Function NomSeguro(ByVal nombre As String) As Range
    On Error Resume Next
    Set NomSeguro = ThisWorkbook.Names(nombre).RefersToRange
    On Error GoTo 0
End Function

Private Function HojaCPU() As Worksheet
    On Error Resume Next
    Set HojaCPU = ThisWorkbook.Worksheets("CPU")
    On Error GoTo 0
End Function

Private Function HojaMemoria() As Worksheet
    On Error Resume Next
    Set HojaMemoria = ThisWorkbook.Worksheets("Memoria")
    If HojaMemoria Is Nothing Then Set HojaMemoria = ThisWorkbook.Worksheets("RAM")
    On Error GoTo 0
End Function

Private Function HojaMEMORIA_Visual() As Worksheet
    On Error Resume Next
    Set HojaMEMORIA_Visual = ThisWorkbook.Worksheets("MEMORIA")
    On Error GoTo 0
End Function

Private Function HojaArquitectura() As Worksheet
    On Error Resume Next
    Set HojaArquitectura = ThisWorkbook.Worksheets("ARQUITECTURA")
    On Error GoTo 0
End Function

' ------------------------------------------------------------------------------
' ACTUALIZACION INTEGRAL DE LA INTERFAZ (Invocada en cada T-State)
' ------------------------------------------------------------------------------
Public Sub ActualizarTodoUI()
    On Error Resume Next
    Dim prevScreenUpdating As Boolean
    Dim prevEvents As Boolean
    prevScreenUpdating = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    ' 1. Actualizar Registros de Datos y Control
    ActualizarRegistrosVisuales
    
    ' 2. Actualizar Banderas ZF, CF, SF
    ActualizarBanderasVisuales
    
    ' 3. Actualizar Buses y Micro-palabra
    ActualizarBusesVisuales
    
    ' 4. Actualizar Jerarquia de Cache
    ActualizarCacheVisual
    
    ' 5. Actualizar Memoria Virtual y Tabla de Paginas
    ActualizarMemoriaVirtualVisual
    
    ' 6. Actualizar Perifericos (Teclado y LCD 8x1)
    ActualizarPerifericosVisual
    
    ' 7. Actualizar Estado del Reloj y Log
    ActualizarEstadoRelojVisual
    
    ' 8. Actualizar hoja MEMORIA (RAM + Swap) en vivo
    ActualizarHojaMemoriaVisual
    
    ' 9. Actualizar cola FIFO visual del teclado en la hoja CPU
    ActualizarColaFIFOVisual
    
    ' 10. Actualizar Matriz de Memoria RAM 16x16 en la hoja CPU
    ActualizarMatrizRAM16x16Visual
    
    ' 11. Actualizar diagrama interactivo en hoja ARQUITECTURA (Pizarra)
    ActualizarArquitecturaPizarraVisual
    
    Application.ScreenUpdating = prevScreenUpdating
    Application.EnableEvents = prevEvents
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' 1. REGISTROS DE LA CPU (AX, BX, CX, DX, PC, IR, MAR, MDR)
' ------------------------------------------------------------------------------
Private Sub ActualizarRegistrosVisuales()
    EscribirCelda "AX_Dec", RegAX
    EscribirCelda "BX_Dec", RegBX
    EscribirCelda "CX_Dec", RegCX
    EscribirCelda "DX_Dec", RegDX
    EscribirCelda "PC_Dec", RegPC
    EscribirCelda "IR_Dec", RegIR
    EscribirCelda "MAR_Dec", RegMAR
    EscribirCelda "MDR_Dec", RegMDR
    
    ' Representacion Hexadecimal
    EscribirCelda "AX_Hex", "0x" & Right$("0" & Hex$(RegAX), 2)
    EscribirCelda "BX_Hex", "0x" & Right$("0" & Hex$(RegBX), 2)
    EscribirCelda "CX_Hex", "0x" & Right$("0" & Hex$(RegCX), 2)
    EscribirCelda "DX_Hex", "0x" & Right$("0" & Hex$(RegDX), 2)
    EscribirCelda "PC_Hex", "0x" & Right$("00" & Hex$(RegPC), 3)
    EscribirCelda "IR_Hex", "0x" & Right$("0" & Hex$(RegIR), 2)
    EscribirCelda "MAR_Hex", "0x" & Right$("00" & Hex$(RegMAR), 3)
    EscribirCelda "MDR_Hex", "0x" & Right$("0" & Hex$(RegMDR), 2)
End Sub

' ------------------------------------------------------------------------------
' 2. BANDERAS ZF, CF, SF
' ------------------------------------------------------------------------------
Private Sub ActualizarBanderasVisuales()
    PintarFlagVisual "Flag_ZF", FlagZF
    PintarFlagVisual "Flag_CF", FlagCF
    PintarFlagVisual "Flag_SF", FlagSF
End Sub

Private Sub PintarFlagVisual(ByVal nombreNombrado As String, ByVal valor As Long)
    Dim r As Range
    Set r = NomSeguro(nombreNombrado)
    If Not r Is Nothing Then
        r.value = valor
        r.Font.Bold = True
        r.HorizontalAlignment = -4108
        If valor = 1 Then
            r.Interior.Color = RGB(120, 100, 30)  ' Amarillo oscuro tema dark
            r.Font.Color = RGB(241, 196, 15)
        Else
            r.Interior.Color = RGB(35, 40, 50)
            r.Font.Color = RGB(120, 130, 140)
        End If
    End If
End Sub

' ------------------------------------------------------------------------------
' 3. BUSES DEL SISTEMA Y MICRO-PALABRA DE CONTROL
' ------------------------------------------------------------------------------
Private Sub ActualizarBusesVisuales()
    ' Bus Interno
    If BusInterno = HIGH_Z Then
        EscribirCelda "Bus_Interno_Val", "High-Z"
        PintarBusCelda "Bus_Interno_Val", False
    Else
        EscribirCelda "Bus_Interno_Val", "0x" & Right$("0" & Hex$(BusInterno), 2) & " (" & EmisorBusInterno & ")"
        PintarBusCelda "Bus_Interno_Val", True
    End If
    
    ' Bus Externo de Direcciones
    If BusDireccionesExt = HIGH_Z Then
        EscribirCelda "Bus_Dir_Val", "High-Z"
        PintarBusCelda "Bus_Dir_Val", False
    Else
        EscribirCelda "Bus_Dir_Val", "0x" & Right$("00" & Hex$(BusDireccionesExt), 3)
        PintarBusCelda "Bus_Dir_Val", True
    End If
    
    ' Bus Externo de Datos
    If BusDatosExt = HIGH_Z Then
        EscribirCelda "Bus_Datos_Val", "High-Z"
        PintarBusCelda "Bus_Datos_Val", False
    Else
        EscribirCelda "Bus_Datos_Val", "0x" & Right$("0" & Hex$(BusDatosExt), 2) & " (" & EmisorBusDatosExt & ")"
        PintarBusCelda "Bus_Datos_Val", True
    End If
    
    ' Micro-palabra de Control
    EscribirCelda "MicroPalabra_Bin", ControlWord_Str
    
    ' Senales individuales
    PintarSenalControl "Sig_MemRead", MemRead
    PintarSenalControl "Sig_MemWrite", MemWrite
    PintarSenalControl "Sig_IoRead", IoRead
    PintarSenalControl "Sig_IoWrite", IoWrite
End Sub

Private Sub PintarBusCelda(ByVal nombreNombrado As String, ByVal activo As Boolean)
    Dim r As Range
    Set r = NomSeguro(nombreNombrado)
    If Not r Is Nothing Then
        If activo Then
            r.Interior.Color = RGB(20, 70, 40)   ' Verde oscuro tema dark
            r.Font.Color = RGB(46, 204, 113)
            r.Font.Bold = True
        Else
            r.Interior.Color = RGB(35, 40, 50)    ' Gris oscuro High-Z
            r.Font.Color = RGB(100, 110, 120)
            r.Font.Bold = False
        End If
    End If
End Sub

Private Sub PintarSenalControl(ByVal nombreNombrado As String, ByVal valor As Long)
    Dim r As Range
    Set r = NomSeguro(nombreNombrado)
    If Not r Is Nothing Then
        r.value = valor
        If valor = 1 Then
            r.Interior.Color = RGB(120, 100, 30)  ' Amarillo oscuro activo
            r.Font.Color = RGB(241, 196, 15)
            r.Font.Bold = True
        Else
            r.Interior.Color = RGB(35, 40, 50)
            r.Font.Color = RGB(100, 110, 120)
            r.Font.Bold = False
        End If
    End If
End Sub

' ------------------------------------------------------------------------------
' 4. JERARQUIA DE MEMORIA CACHE (L1I, L1D, L2)
' ------------------------------------------------------------------------------
Private Sub ActualizarCacheVisual()
    EscribirCelda "L1I_Status", UltimoEstadoL1I
    EscribirCelda "L1D_Status", UltimoEstadoL1D
    EscribirCelda "L2_Status", UltimoEstadoL2
    EscribirCelda "Cache_Evento", UltimoEventoCache
    
    EscribirCelda "L1I_Hits", Hits_L1I
    EscribirCelda "L1I_Misses", Misses_L1I
    EscribirCelda "L1I_HitRate", HitRate_L1I & " %"
    
    EscribirCelda "L1D_Hits", Hits_L1D
    EscribirCelda "L1D_Misses", Misses_L1D
    EscribirCelda "L1D_HitRate", HitRate_L1D & " %"
    
    EscribirCelda "L2_Hits", Hits_L2
    EscribirCelda "L2_Misses", Misses_L2
    EscribirCelda "L2_HitRate", HitRate_L2 & " %"
    
    EscribirCelda "Global_HitRate", HitRate_Global & " %"
End Sub

' ------------------------------------------------------------------------------
' 5. MEMORIA VIRTUAL, TABLA DE PAGINAS Y PAGE FAULT
' ------------------------------------------------------------------------------
Private Sub ActualizarMemoriaVirtualVisual()
    EscribirCelda "Total_Accesos_Mem", TotalAccesosMemoria
    EscribirCelda "Total_Page_Faults", TotalPageFaults
    EscribirCelda "Msg_Page_Fault", MensajePageFault
    
    Dim rAlerta As Range
    Set rAlerta = NomSeguro("Alerta_PageFault")
    If Not rAlerta Is Nothing Then
        If UltimoPageFaultOcurrido Then
            rAlerta.value = "PAGE FAULT!"
            rAlerta.Interior.Color = RGB(180, 40, 30)  ' ROJO ALERTA oscuro
            rAlerta.Font.Color = RGB(255, 255, 255)
            rAlerta.Font.Bold = True
        Else
            rAlerta.value = "NORMAL"
            rAlerta.Interior.Color = RGB(20, 70, 40)   ' VERDE oscuro
            rAlerta.Font.Color = RGB(46, 204, 113)
            rAlerta.Font.Bold = False
        End If
    End If
    
    ' Actualizar entradas de la Tabla de Paginas en la hoja CPU
    ' La tabla empieza en fila 11, columnas T-X (20-24)
    Dim ws As Worksheet
    Set ws = HojaCPU
    If Not ws Is Nothing Then
        Dim i As Long
        Dim baseRow As Long
        baseRow = 11  ' Fila donde empieza VPN 0 en el tablero
        For i = 0 To 15
            Dim rowPT As Long
            rowPT = baseRow + i
            
            ws.Cells(rowPT, 20).value = "VPN " & i
            
            ' Columna P (Presente)
            ws.Cells(rowPT, 21).value = TablaPaginas(i).Presente
            
            ' Columna UBICACION
            If TablaPaginas(i).Presente = 1 Then
                ws.Cells(rowPT, 22).value = "Marco " & TablaPaginas(i).pfn
            Else
                ws.Cells(rowPT, 22).value = "Swap"
            End If
            
            ' Columna D (Dirty)
            ws.Cells(rowPT, 23).value = TablaPaginas(i).Dirty
            
            ' Columna R (Referenciado)
            ws.Cells(rowPT, 24).value = TablaPaginas(i).Referenciado
            
            ' Colores segun estado (tema oscuro)
            If UltimoPageFaultOcurrido And i = UltimaVPN_Accedida Then
                ' VPN que causo PAGE FAULT = rojo intenso
                ws.Cells(rowPT, 20).Font.Color = RGB(255, 100, 80)
                ws.Cells(rowPT, 21).Font.Color = RGB(255, 100, 80)
                ws.Cells(rowPT, 22).Font.Color = RGB(255, 100, 80)
                ws.Cells(rowPT, 23).Font.Color = RGB(255, 100, 80)
                ws.Cells(rowPT, 24).Font.Color = RGB(255, 100, 80)
                ws.Range(ws.Cells(rowPT, 20), ws.Cells(rowPT, 24)).Interior.Color = RGB(80, 20, 15)
            ElseIf TablaPaginas(i).Presente = 1 Then
                ' En RAM = verde
                ws.Cells(rowPT, 20).Font.Color = RGB(46, 204, 113)
                ws.Cells(rowPT, 21).Font.Color = RGB(46, 204, 113)
                ws.Cells(rowPT, 22).Font.Color = RGB(46, 204, 113)
                ws.Cells(rowPT, 23).Font.Color = RGB(46, 204, 113)
                ws.Cells(rowPT, 24).Font.Color = RGB(46, 204, 113)
                ws.Range(ws.Cells(rowPT, 20), ws.Cells(rowPT, 24)).Interior.Color = RGB(20, 50, 30)
            Else
                ' En Swap = rojo tenue
                ws.Cells(rowPT, 20).Font.Color = RGB(231, 76, 60)
                ws.Cells(rowPT, 21).Font.Color = RGB(231, 76, 60)
                ws.Cells(rowPT, 22).Font.Color = RGB(231, 76, 60)
                ws.Cells(rowPT, 23).Font.Color = RGB(231, 76, 60)
                ws.Cells(rowPT, 24).Font.Color = RGB(231, 76, 60)
                ws.Range(ws.Cells(rowPT, 20), ws.Cells(rowPT, 24)).Interior.Color = RGB(50, 25, 20)
            End If
        Next i
    End If
End Sub

' ------------------------------------------------------------------------------
' 6. PERIFERICOS (TECLADO Y PANTALLA LCD 8x1)
' ------------------------------------------------------------------------------
Private Sub ActualizarPerifericosVisual()
    ' LCD 8x1: actualizar los 8 caracteres en display
    Dim textoLCD As String
    textoLCD = LCD_ObtenerTexto
    EscribirCelda "LCD_Display", textoLCD
    
    ' Teclado Hardware
    EscribirCelda "KBD_Status", RegKbdStatus
    EscribirCelda "KBD_Data", "0x" & Right$("0" & Hex$(RegKbdData), 2) & " ('" & Chr$(IIf(RegKbdData >= 32 And RegKbdData <= 126, RegKbdData, 32)) & "')"
    
    ' Cola FIFO en RAM
    Dim headVal As Long, tailVal As Long, countVal As Long
    headVal = MemoriaVirtual_LeerByte(DIR_KB_HEAD)
    tailVal = MemoriaVirtual_LeerByte(DIR_KB_TAIL)
    countVal = MemoriaVirtual_LeerByte(DIR_KB_COUNT)
    EscribirCelda "Cola_Head", headVal
    EscribirCelda "Cola_Tail", "Tail: " & tailVal & " | Cnt: " & countVal
End Sub

' ------------------------------------------------------------------------------
' 7. RELOJ, MNEMONICO, MICROPASO Y LOG
' ------------------------------------------------------------------------------
Private Sub ActualizarEstadoRelojVisual()
    EscribirCelda "MicroPaso_T", "T" & MicroPasoT
    EscribirCelda "Fase_Nombre", FaseNombre
    EscribirCelda "Sel_Mnem", MnemonicoActual
    EscribirCelda "Status_Msg", DetalleMicroPaso
    EscribirCelda "Total_Instrucciones", TotalInstruccionesEjecutadas
    
    ' Registrar en Log superior
    If MicroPasoT = 0 And TotalInstruccionesEjecutadas > 0 Then
        AgregarLog "[Paso " & TotalInstruccionesEjecutadas & "] " & MnemonicoActual & " | " & DetalleMicroPaso
    End If
End Sub

Private Sub AgregarLog(ByVal mensaje As String)
    On Error Resume Next
    Dim i As Long
    For i = 10 To 2 Step -1
        EscribirCelda "Log_" & i, LeerCeldaTexto("Log_" & (i - 1))
    Next i
    EscribirCelda "Log_1", mensaje
    On Error GoTo 0
End Sub

Private Function LeerCeldaTexto(ByVal nombre As String) As String
    Dim r As Range
    Set r = NomSeguro(nombre)
    If Not r Is Nothing Then LeerCeldaTexto = CStr(r.value) Else LeerCeldaTexto = ""
End Function

Private Sub EscribirCelda(ByVal nombre As String, ByVal valor As Variant)
    Dim r As Range
    Set r = NomSeguro(nombre)
    If Not r Is Nothing Then r.value = valor
End Sub

' ==============================================================================
' 8. ACTUALIZACION EN VIVO DE LA HOJA MEMORIA (RAM FISICA + SWAP)
' ==============================================================================
' Esta funcion escribe los valores actuales de RAM_Fisica() y Disco_Swap()
' en las celdas de la hoja "MEMORIA" creada por SetupMemoriaYBuses.bas
'
' Layout de la hoja MEMORIA:
'   RAM: 8 marcos, cada marco ocupa 5 filas (encabezado + offset + 2 filas datos + espacio)
'         Marco 0 empieza en fila 3, datos en filas 5-6, columnas C-R (3-18)
'         Marco N empieza en fila 3 + N*5
'   SWAP: 16 paginas, cada pagina ocupa 3 filas (encabezado + 2 filas datos)
'         Pagina 0 empieza en fila 3, datos en filas 4-5, columnas U-AJ (21-36)
'         Pagina N empieza en fila 3 + N*3
' ==============================================================================
Private Sub ActualizarHojaMemoriaVisual()
    Dim wsMem As Worksheet
    Set wsMem = HojaMEMORIA_Visual
    If wsMem Is Nothing Then Exit Sub
    
    ' Si esta corriendo a maxima velocidad (RetardoMs = 0), solo actualizar al completar instruccion
    If EnEjecucion And RetardoMs = 0 And MicroPasoT <> 0 Then Exit Sub
    
    Dim marco As Long, col As Long, addr As Long
    Dim arr16(1 To 1, 1 To 16) As Variant
    
    ' ---- RAM FISICA: 8 marcos x 32 bytes ----
    For marco = 0 To 7
        Dim filaBaseMrc As Long
        filaBaseMrc = 3 + (marco * 5)  ' fila del encabezado del marco
        
        ' Fila de datos 1 (bytes 0-15): filaBaseMrc + 2
        Dim filaD1 As Long
        filaD1 = filaBaseMrc + 2
        For col = 0 To 15
            addr = marco * 32 + col
            arr16(1, col + 1) = Right$("0" & Hex$(RAM_Fisica(addr)), 2)
        Next col
        wsMem.Range(wsMem.Cells(filaD1, 3), wsMem.Cells(filaD1, 18)).value = arr16
        
        ' Fila de datos 2 (bytes 16-31): filaBaseMrc + 3
        Dim filaD2 As Long
        filaD2 = filaBaseMrc + 3
        For col = 0 To 15
            addr = marco * 32 + 16 + col
            arr16(1, col + 1) = Right$("0" & Hex$(RAM_Fisica(addr)), 2)
        Next col
        wsMem.Range(wsMem.Cells(filaD2, 3), wsMem.Cells(filaD2, 18)).value = arr16
        
        ' Actualizar encabezado del marco con la VPN asignada
        Dim vpnAsig As Long
        vpnAsig = MarcoAsignadoA_VPN(marco)
        Dim encabezado As String
        encabezado = "MARCO " & marco & "  |  Dir. Fisica: 0x" & Right$("00" & Hex$(marco * 32), 3) & " - 0x" & Right$("00" & Hex$(marco * 32 + 31), 3)
        If vpnAsig >= 0 Then
            encabezado = encabezado & "  [VPN " & vpnAsig & " CARGADA]"
        Else
            encabezado = encabezado & "  [LIBRE]"
        End If
        wsMem.Cells(filaBaseMrc, 2).value = encabezado
        
        ' Color del encabezado: azul si tiene VPN, gris si libre
        If vpnAsig >= 0 Then
            wsMem.Cells(filaBaseMrc, 2).Interior.Color = RGB(20, 60, 100)
        Else
            wsMem.Cells(filaBaseMrc, 2).Interior.Color = RGB(40, 45, 55)
        End If
    Next marco
    
    ' ---- AREA DE SWAP: 16 paginas x 32 bytes ----
    Dim pagina As Long
    For pagina = 0 To 15
        Dim filaBasePag As Long
        filaBasePag = 3 + (pagina * 3)  ' fila del encabezado de la pagina
        
        ' Fila de datos 1 (bytes 0-15): filaBasePag + 1
        Dim filaPD1 As Long
        filaPD1 = filaBasePag + 1
        For col = 0 To 15
            addr = pagina * 32 + col
            arr16(1, col + 1) = Right$("0" & Hex$(Disco_Swap(addr)), 2)
        Next col
        wsMem.Range(wsMem.Cells(filaPD1, 21), wsMem.Cells(filaPD1, 36)).value = arr16
        
        ' Fila de datos 2 (bytes 16-31): filaBasePag + 2
        Dim filaPD2 As Long
        filaPD2 = filaBasePag + 2
        For col = 0 To 15
            addr = pagina * 32 + 16 + col
            arr16(1, col + 1) = Right$("0" & Hex$(Disco_Swap(addr)), 2)
        Next col
        wsMem.Range(wsMem.Cells(filaPD2, 21), wsMem.Cells(filaPD2, 36)).value = arr16
        
        ' Color del encabezado de swap: verde si esta en RAM, rojo si en swap
        If TablaPaginas(pagina).Presente = 1 Then
            wsMem.Cells(filaBasePag, 20).Interior.Color = RGB(20, 70, 50)
        Else
            wsMem.Cells(filaBasePag, 20).Interior.Color = RGB(90, 40, 20)
        End If
    Next pagina
End Sub

' ==============================================================================
' 9. COLA FIFO VISUAL DEL TECLADO (En la hoja CPU, seccion inferior)
' ==============================================================================
' Actualiza las 8 celdas visuales de la cola FIFO del teclado
' que SetupMemoriaYBuses.bas crea en la hoja CPU
' ==============================================================================
Private Sub ActualizarColaFIFOVisual()
    Dim ws As Worksheet
    Set ws = HojaCPU
    If ws Is Nothing Then Exit Sub
    
    Static FilaFIFOCached As Long
    
    If FilaFIFOCached = 0 Then
        Dim celBuscar As Range
        On Error Resume Next
        Set celBuscar = ws.Range("B1:B70").Find(What:="Cola FIFO:", LookIn:=xlValues, LookAt:=xlWhole)
        On Error GoTo 0
        If Not celBuscar Is Nothing Then
            FilaFIFOCached = celBuscar.Row
        Else
            Exit Sub
        End If
    End If
    
    Dim filaFIFO As Long
    filaFIFO = FilaFIFOCached
    
    Dim b As Long
    Dim arrCola(1 To 1, 1 To 8) As Variant
    For b = 0 To 7
        Dim byteEnCola As Long
        byteEnCola = MemoriaVirtual_LeerByte(DIR_COLA_KBD + b)
        
        If byteEnCola >= 32 And byteEnCola <= 126 Then
            arrCola(1, b + 1) = "[" & Chr$(byteEnCola) & "]"
        ElseIf byteEnCola > 0 Then
            arrCola(1, b + 1) = "[" & Right$("0" & Hex$(byteEnCola), 2) & "]"
        Else
            arrCola(1, b + 1) = "[ ]"
        End If
    Next b
    
    ws.Range(ws.Cells(filaFIFO, 3), ws.Cells(filaFIFO, 10)).value = arrCola
End Sub

' ==============================================================================
' 10. MATRIZ DE MEMORIA RAM FISICA 16x16 (256 BYTES) EN HOJA CPU
' Actualiza en vivo los 256 bytes de memoria fisica en formato hexadecimal
' y resalta dinamicamente en AMARILLO la casilla activa (apuntada por MAR/PC)
' ==============================================================================
Private Sub ActualizarMatrizRAM16x16Visual()
    Dim wsCPU As Worksheet
    Set wsCPU = HojaCPU
    If wsCPU Is Nothing Then Exit Sub
    
    ' Si esta corriendo a maxima velocidad (RetardoMs = 0), solo actualizar al completar instruccion
    If EnEjecucion And RetardoMs = 0 And MicroPasoT <> 0 Then Exit Sub
    
    Static UltimaFilaResaltada As Long
    Static UltimaColResaltada As Long
    Static UltimoColorPrevio As Long
    Static UltimoFontPrevio As Long
    
    Dim r As Long, c As Long, addr As Long
    Dim arr16x16(1 To 16, 1 To 16) As Variant
    
    ' Llenar matriz 16x16 desde RAM_Fisica (0 a 255)
    For r = 1 To 16
        For c = 1 To 16
            addr = ((r - 1) * 16) + (c - 1)
            arr16x16(r, c) = Right$("0" & Hex$(RAM_Fisica(addr)), 2)
        Next c
    Next r
    
    ' Escritura ultra-rapida en bloque en el rango Z11:AP26 (Col 27 a 42, Filas 11 a 26)
    wsCPU.Range(wsCPU.Cells(11, 27), wsCPU.Cells(26, 42)).value = arr16x16
    
    ' Determinar direccion fisica activa a resaltar en Amarillo (como en el 1er Parcial)
    Dim addrActiva As Long
    addrActiva = -1
    
    If RegMAR >= 0 And RegMAR <= 511 Then
        Dim vpn As Long, pfn As Long
        vpn = (RegMAR And 511) \ TAM_PAGINA
        If TablaPaginas(vpn).Presente = 1 Then
            pfn = TablaPaginas(vpn).pfn
            addrActiva = (pfn * TAM_PAGINA) + (RegMAR Mod TAM_PAGINA)
        ElseIf RegMAR < 256 Then
            addrActiva = RegMAR
        End If
    ElseIf RegPC >= 0 And RegPC <= 511 Then
        Dim vpnPC As Long, pfnPC As Long
        vpnPC = (RegPC And 511) \ TAM_PAGINA
        If TablaPaginas(vpnPC).Presente = 1 Then
            pfnPC = TablaPaginas(vpnPC).pfn
            addrActiva = (pfnPC * TAM_PAGINA) + (RegPC Mod TAM_PAGINA)
        ElseIf RegPC < 256 Then
            addrActiva = RegPC
        End If
    End If
    
    Dim filaNueva As Long, colNueva As Long
    If addrActiva >= 0 And addrActiva <= 255 Then
        filaNueva = (addrActiva \ 16) + 11
        colNueva = (addrActiva Mod 16) + 27
    Else
        filaNueva = 0
        colNueva = 0
    End If
    
    ' Si cambio la celda activa, restaurar el color anterior
    If UltimaFilaResaltada >= 11 And UltimaFilaResaltada <= 26 And UltimaColResaltada >= 27 And UltimaColResaltada <= 42 Then
        If UltimaFilaResaltada <> filaNueva Or UltimaColResaltada <> colNueva Then
            With wsCPU.Cells(UltimaFilaResaltada, UltimaColResaltada)
                .Interior.Color = UltimoColorPrevio
                .Font.Color = UltimoFontPrevio
                .Font.Bold = False
            End With
            UltimaFilaResaltada = 0
            UltimaColResaltada = 0
        End If
    End If
    
    ' Resaltar la nueva celda activa en Amarillo brillante (como en el 1er Parcial)
    If filaNueva >= 11 And filaNueva <= 26 And colNueva >= 27 And colNueva <= 42 Then
        If filaNueva <> UltimaFilaResaltada Or colNueva <> UltimaColResaltada Then
            If filaNueva <= 17 Then
                ' Azul oscuro codigo (0x00-0x6F)
                UltimoColorPrevio = RGB(22, 36, 52)
                UltimoFontPrevio = RGB(147, 197, 253)
            ElseIf filaNueva = 18 Then
                ' Fila 70: Cola FIFO Teclado (0x70-0x77) y Punteros (0x78-0x7A)
                If colNueva >= 27 And colNueva <= 34 Then
                    UltimoColorPrevio = RGB(20, 48, 70)
                    UltimoFontPrevio = RGB(125, 211, 252)
                ElseIf colNueva >= 35 And colNueva <= 37 Then
                    UltimoColorPrevio = RGB(28, 58, 85)
                    UltimoFontPrevio = RGB(186, 230, 253)
                Else
                    UltimoColorPrevio = RGB(22, 36, 52)
                    UltimoFontPrevio = RGB(147, 197, 253)
                End If
            ElseIf filaNueva = 26 Then
                ' Fila F0: RESERVA DE VRAM / BUFFER DE PANTALLA (0xF0-0xFF)
                If colNueva >= 35 Then
                    ' 0xF8 - 0xFF Buffer Activo Pantalla LCD 8x1
                    UltimoColorPrevio = RGB(85, 40, 12)
                    UltimoFontPrevio = RGB(254, 215, 170)
                Else
                    ' 0xF0 - 0xF7 Reserva E/S Mapeada
                    UltimoColorPrevio = RGB(55, 30, 10)
                    UltimoFontPrevio = RGB(251, 146, 60)
                End If
            Else
                ' Verde oscuro datos (0x80-0xEF, Filas 19-25)
                UltimoColorPrevio = RGB(18, 42, 30)
                UltimoFontPrevio = RGB(134, 239, 172)
            End If
            
            With wsCPU.Cells(filaNueva, colNueva)
                .Interior.Color = RGB(241, 196, 15)  ' Amarillo intenso
                .Font.Color = RGB(0, 0, 0)           ' Texto negro
                .Font.Bold = True
            End With
            UltimaFilaResaltada = filaNueva
            UltimaColResaltada = colNueva
        End If
    End If
End Sub

' ==============================================================================
' 11. ACTUALIZACION EN VIVO DE LA HOJA ARQUITECTURA (PIZARRA DEL DOCENTE)
' Sincroniza registros, marcos, buses, monitor LCD, teclado y cola FIFO 0x70-0x77
' ==============================================================================
Private Sub ActualizarArquitecturaPizarraVisual()
    Dim wsArq As Worksheet
    Set wsArq = HojaArquitectura
    If wsArq Is Nothing Then Exit Sub
    
    If EnEjecucion And RetardoMs = 0 And MicroPasoT <> 0 Then Exit Sub
    
    On Error Resume Next
    
    ' --- 1. Registros CPU ---
    wsArq.Cells(8, 6).Value = RegAX
    wsArq.Cells(9, 6).Value = RegBX
    wsArq.Cells(10, 6).Value = RegCX
    wsArq.Cells(11, 6).Value = RegDX
    wsArq.Cells(13, 3).Value = "0x" & Right$("00" & Hex$(RegPC), 3)
    wsArq.Cells(13, 5).Value = "0x" & Right$("0" & Hex$(RegIR), 2)
    wsArq.Cells(14, 3).Value = "0x" & Right$("00" & Hex$(RegMAR), 3)
    wsArq.Cells(14, 5).Value = "0x" & Right$("0" & Hex$(RegMDR), 2)
    
    If BusInterno = HIGH_Z Then
        wsArq.Range("B15").Value = "[ BUS INTERNO CPU (8 bits) ]: High-Z"
    Else
        wsArq.Range("B15").Value = "[ BUS INTERNO CPU (8 bits) ]: 0x" & Right$("0" & Hex$(BusInterno), 2) & " <- " & EmisorBusInterno
    End If
    
    ' --- 2. Marcos de Memoria RAM Fisica (Marcos 0..7 en Filas 7..14) ---
    Dim m As Long, v As Long, vpnAsig As String
    For m = 0 To 7
        vpnAsig = "Libre"
        For v = 0 To 15
            If TablaPaginas(v).Presente = 1 And TablaPaginas(v).PFN = m Then
                vpnAsig = "VPN " & v
                Exit For
            End If
        Next v
        wsArq.Cells(7 + m, 10).Value = vpnAsig
        If vpnAsig = "Libre" Then
            wsArq.Cells(7 + m, 11).Value = "[Libre]"
            wsArq.Cells(7 + m, 11).Font.Color = RGB(149, 165, 166)
        Else
            wsArq.Cells(7 + m, 11).Value = "[Presente]"
            wsArq.Cells(7 + m, 11).Font.Color = RGB(39, 174, 96)
        End If
    Next m
    
    ' --- 3. Disco / Area de Swap ---
    Dim cntRam As Long, cntSwap As Long
    cntRam = 0: cntSwap = 0
    For v = 0 To 15
        If TablaPaginas(v).Presente = 1 Then cntRam = cntRam + 1 Else cntSwap = cntSwap + 1
    Next v
    wsArq.Cells(9, 16).Value = cntRam & " Marcos"
    wsArq.Cells(10, 16).Value = cntSwap & " Paginas"
    wsArq.Cells(12, 16).Value = IIf(TotalPageFaults > 0, "PF #" & TotalPageFaults, "Ninguno")
    wsArq.Cells(13, 16).Value = IIf(UltimoPageFaultOcurrido, "VPN " & UltimaVPN_Accedida, "Ninguno")
    
    ' --- 4. Los 3 Buses Horizontales (Control, Direcciones, Datos) ---
    wsArq.Range("B18").Value = "[ BUS DE CONTROL ]:  MEM_R=" & MemRead & "  |  MEM_W=" & MemWrite & "  |  IO_R=" & IoRead & "  |  IO_W=" & IoWrite & "  |  RELOJ (CLK)=" & (TotalInstruccionesEjecutadas Mod 2)
    If BusDireccionesExt = HIGH_Z Then
        wsArq.Range("B20").Value = "[ BUS DE DIRECCIONES (16 bits) ]: High-Z (Inactivo)"
    Else
        wsArq.Range("B20").Value = "[ BUS DE DIRECCIONES (16 bits) ]: MAR -> 0x" & Right$("00" & Hex$(BusDireccionesExt), 3)
    End If
    If BusDatosExt = HIGH_Z Then
        wsArq.Range("B22").Value = "[ BUS DE DATOS (8 bits, Bidireccional) ]: MDR <-> High-Z"
    Else
        wsArq.Range("B22").Value = "[ BUS DE DATOS (8 bits, Bidireccional) ]: MDR <-> 0x" & Right$("0" & Hex$(BusDatosExt), 2) & " (" & EmisorBusDatosExt & ")"
    End If
    
    ' --- 5. Monitor / LCD 8x1 ---
    Dim txtLCD As String
    txtLCD = LCD_ObtenerTexto()
    If Len(Trim$(txtLCD)) = 0 Then
        wsArq.Range("C26").Value = "[ READY ]"
    Else
        wsArq.Range("C26").Value = txtLCD
    End If
    
    ' --- 6. Teclado Interactivo ---
    wsArq.Cells(26, 13).Value = RegKbdStatus
    Dim chKbd As String
    If RegKbdData >= 32 And RegKbdData <= 126 Then chKbd = Chr$(RegKbdData) Else chKbd = " "
    wsArq.Cells(27, 13).Value = "0x" & Right$("0" & Hex$(RegKbdData), 2) & " ('" & chKbd & "')"
    
    ' --- 7. Cola Circular FIFO del Teclado en RAM (0x70 - 0x77) ---
    Dim headK As Long, tailK As Long, cntK As Long
    headK = MemoriaVirtual_LeerByte(DIR_KB_HEAD)
    tailK = MemoriaVirtual_LeerByte(DIR_KB_TAIL)
    cntK = MemoriaVirtual_LeerByte(DIR_KB_COUNT)
    
    Dim slotIdx As Long, offsetIdx As Long, estaActivo(0 To 7) As Boolean
    For slotIdx = 0 To 7
        estaActivo(slotIdx) = False
    Next slotIdx
    For offsetIdx = 0 To cntK - 1
        estaActivo((headK + offsetIdx) Mod 8) = True
    Next offsetIdx
    
    Dim bFifo As Long
    For slotIdx = 0 To 7
        With wsArq.Cells(27, 17 + slotIdx)
            If estaActivo(slotIdx) Then
                bFifo = MemoriaVirtual_LeerByte(DIR_COLA_KBD + slotIdx)
                If bFifo >= 32 And bFifo <= 126 Then
                    .Value = "[ " & Chr$(bFifo) & " ]"
                Else
                    .Value = "0x" & Right$("0" & Hex$(bFifo), 2)
                End If
                .Interior.Color = RGB(255, 243, 205)
                .Font.Color = RGB(0, 0, 0)
            Else
                .Value = "[ ]"
                .Interior.Color = RGB(255, 255, 255)
                .Font.Color = RGB(127, 140, 141)
            End If
        End With
    Next slotIdx
    
    wsArq.Cells(29, 19).Value = headK
    wsArq.Cells(29, 23).Value = tailK
    wsArq.Cells(29, 25).Value = cntK
    On Error GoTo 0
End Sub
