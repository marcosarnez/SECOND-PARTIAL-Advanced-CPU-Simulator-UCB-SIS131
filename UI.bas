Option Explicit

' ==============================================================================
' MODULO: UI.bas
' Simulador Avanzado de Arquitectura de Computadoras - Segundo Parcial (SIS-131)
' Control Gráfico y Visualización en Microsoft Excel:
'   1. Registros Generales: AX, BX, CX, DX y Registros de Control (PC, IR, MAR, MDR).
'   2. Visualización de Buses y Buffers Tri-Estado (High-Z dinámico).
'   3. Micro-palabra de Control y Señales (MEM_R, MEM_W, IO_R, IO_W).
'   4. Panel de Jerarquía de Caché L1I, L1D, L2 con contadores y Hit Rates.
'   5. Tabla de Páginas (16 VPNs) y alerta visual de PAGE FAULT en Rojo.
'   6. Display Pantalla LCD 8x1 y Panel interactivo de Teclado.
'   7. Log de eventos y traza de ejecución.
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

' ------------------------------------------------------------------------------
' ACTUALIZACIÓN INTEGRAL DE LA INTERFAZ (Invocada en cada T-State)
' ------------------------------------------------------------------------------
Public Sub ActualizarTodoUI()
    On Error Resume Next
    
    ' 1. Actualizar Registros de Datos y Control
    ActualizarRegistrosVisuales
    
    ' 2. Actualizar Banderas ZF, CF, SF
    ActualizarBanderasVisuales
    
    ' 3. Actualizar Buses y Micro-palabra
    ActualizarBusesVisuales
    
    ' 4. Actualizar Jerarquía de Caché
    ActualizarCacheVisual
    
    ' 5. Actualizar Memoria Virtual y Tabla de Páginas
    ActualizarMemoriaVirtualVisual
    
    ' 6. Actualizar Periféricos (Teclado y LCD 8x1)
    ActualizarPerifericosVisual
    
    ' 7. Actualizar Estado del Reloj y Log
    ActualizarEstadoRelojVisual
    
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
    
    ' Representación Hexadecimal
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
        r.Value = valor
        r.Font.Bold = True
        r.HorizontalAlignment = -4108
        If valor = 1 Then
            r.Interior.Color = RGB(249, 231, 159) ' Amarillo iluminado
            r.Font.Color = RGB(31, 78, 121)
        Else
            r.Interior.Color = RGB(245, 245, 245)
            r.Font.Color = RGB(120, 120, 120)
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
    
    ' Señales individuales
    PintarSeñalControl "Sig_MemRead", MemRead
    PintarSeñalControl "Sig_MemWrite", MemWrite
    PintarSeñalControl "Sig_IoRead", IoRead
    PintarSeñalControl "Sig_IoWrite", IoWrite
End Sub

Private Sub PintarBusCelda(ByVal nombreNombrado As String, ByVal activo As Boolean)
    Dim r As Range
    Set r = NomSeguro(nombreNombrado)
    If Not r Is Nothing Then
        If activo Then
            r.Interior.Color = RGB(213, 245, 227) ' Verde claro activo
            r.Font.Color = RGB(20, 90, 50)
            r.Font.Bold = True
        Else
            r.Interior.Color = RGB(242, 244, 244) ' Gris High-Z
            r.Font.Color = RGB(127, 140, 141)
            r.Font.Bold = False
        End If
    End If
End Sub

Private Sub PintarSeñalControl(ByVal nombreNombrado As String, ByVal valor As Long)
    Dim r As Range
    Set r = NomSeguro(nombreNombrado)
    If Not r Is Nothing Then
        r.Value = valor
        If valor = 1 Then
            r.Interior.Color = RGB(249, 231, 159)
            r.Font.Color = RGB(31, 78, 121)
            r.Font.Bold = True
        Else
            r.Interior.Color = RGB(245, 245, 245)
            r.Font.Color = RGB(150, 150, 150)
            r.Font.Bold = False
        End If
    End If
End Sub

' ------------------------------------------------------------------------------
' 4. JERARQUÍA DE MEMORIA CACHÉ (L1I, L1D, L2)
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
' 5. MEMORIA VIRTUAL, TABLA DE PÁGINAS Y PAGE FAULT
' ------------------------------------------------------------------------------
Private Sub ActualizarMemoriaVirtualVisual()
    EscribirCelda "Total_Accesos_Mem", TotalAccesosMemoria
    EscribirCelda "Total_Page_Faults", TotalPageFaults
    EscribirCelda "Msg_Page_Fault", MensajePageFault
    
    Dim rAlerta As Range
    Set rAlerta = NomSeguro("Alerta_PageFault")
    If Not rAlerta Is Nothing Then
        If UltimoPageFaultOcurrido Then
            rAlerta.Value = "¡PAGE FAULT!"
            rAlerta.Interior.Color = RGB(231, 76, 60) ' ROJO ALERTA
            rAlerta.Font.Color = RGB(255, 255, 255)
            rAlerta.Font.Bold = True
        Else
            rAlerta.Value = "NORMAL"
            rAlerta.Interior.Color = RGB(213, 245, 227) ' VERDE
            rAlerta.Font.Color = RGB(39, 174, 96)
            rAlerta.Font.Bold = False
        End If
    End If
    
    ' Actualizar entradas de la Tabla de Páginas (si existe el rango en hoja)
    Dim ws As Worksheet
    Set ws = HojaCPU
    If Not ws Is Nothing Then
        Dim i As Long
        Dim baseRow As Long
        baseRow = 38 ' Fila típica para la Page Table en la hoja CPU
        For i = 0 To 15
            Dim cVPN As Range, cP As Range, cPFN As Range, cD As Range, cR As Range
            Set cVPN = ws.Cells(baseRow + i, 20)
            Set cP = ws.Cells(baseRow + i, 21)
            Set cPFN = ws.Cells(baseRow + i, 22)
            Set cD = ws.Cells(baseRow + i, 23)
            Set cR = ws.Cells(baseRow + i, 24)
            
            If Not cVPN Is Nothing And cVPN.Row > 0 Then
                cVPN.Value = "VPN " & i
                cP.Value = TablaPaginas(i).Presente
                cPFN.Value = IIf(TablaPaginas(i).Presente = 1, "Marco " & TablaPaginas(i).PFN, "Swap")
                cD.Value = TablaPaginas(i).Dirty
                cR.Value = TablaPaginas(i).Referenciado
                
                ' Resaltar en rojo si fue la VPN causante del Page Fault
                If UltimoPageFaultOcurrido And i = UltimaVPN_Accedida Then
                    ws.Range(cVPN, cR).Interior.Color = RGB(242, 215, 213)
                ElseIf TablaPaginas(i).Presente = 1 Then
                    ws.Range(cVPN, cR).Interior.Color = RGB(234, 250, 234)
                Else
                    ws.Range(cVPN, cR).Interior.Color = RGB(255, 255, 255)
                End If
            End If
        Next i
    End If
End Sub

' ------------------------------------------------------------------------------
' 6. PERIFÉRICOS (TECLADO Y PANTALLA LCD 8x1)
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
    EscribirCelda "Cola_Head", MemoriaVirtual_LeerByte(DIR_KB_HEAD)
    EscribirCelda "Cola_Tail", MemoriaVirtual_LeerByte(DIR_KB_TAIL)
    EscribirCelda "Cola_Count", MemoriaVirtual_LeerByte(DIR_KB_COUNT)
End Sub

' ------------------------------------------------------------------------------
' 7. RELOJ, MNEMÓNICO, MICROPASO Y LOG
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
    If Not r Is Nothing Then LeerCeldaTexto = CStr(r.Value) Else LeerCeldaTexto = ""
End Function

Private Sub EscribirCelda(ByVal nombre As String, ByVal valor As Variant)
    Dim r As Range
    Set r = NomSeguro(nombre)
    If Not r Is Nothing Then r.Value = valor
End Sub
