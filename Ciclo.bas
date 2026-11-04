Option Explicit

#If VBA7 Then
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ==============================================================================
' MODULO: Ciclo.bas
' Simulador Avanzado de Arquitectura de Computadoras - Segundo Parcial (SIS-131)
' Motor del Reloj del Sistema:
'   1. Contador de Micro-Paso (T-States): T0, T1, T2 (Fetch), T3 (Decode), T4..T6 (Execute/Store).
'   2. Fases Temporales Estrictas en cada T-State:
'      - Fase 1: Setup / Preparación (Nivel Alto): UC activa señales y el emisor expone al bus.
'      - Fase 2: Latch / Disparo (Flanco de Reloj): Receptores con Load=1 capturan el dato.
'   3. Control de Ejecución: STEP, RUN, PAUSE, RESET y ajuste de Delay (ms).
'   4. Validación de Exclusión Mutua en cada pulso de reloj.
' ==============================================================================

Public MicroPasoT As Long          ' 0 a 6 (T0..T6)
Public FaseNombre As String        ' "FETCH", "DECODE", "EXECUTE", "STORE"
Public DetalleMicroPaso As String  ' Explicación técnica de la micro-operación activa
Public EnEjecucion As Boolean      ' True si RUN está corriendo
Public EsPausado As Boolean        ' True si el usuario pulsó PAUSE
Public TotalInstruccionesEjecutadas As Long
Public RetardoMs As Long           ' Delay entre micro-pasos en milisegundos

' Operandos decodificados de la instrucción en curso
Public OpcodeActual As Long
Public Operando1 As Long
Public Operando2 As Long
Public TamanoInstruccionActual As Long
Public MnemonicoActual As String

' Resetea el motor de ejecución del ciclo
Public Sub ResetCiclo()
    EnEjecucion = False
    EsPausado = False
    MicroPasoT = 0
    FaseNombre = "REPOSO"
    DetalleMicroPaso = "Simulador en reposo. Presione STEP o RUN para iniciar."
    TotalInstruccionesEjecutadas = 0
    RetardoMs = 300
    OpcodeActual = 0
    Operando1 = 0
    Operando2 = 0
    TamanoInstruccionActual = 0
    MnemonicoActual = "---"
    ResetBuses
    ResetSeñalesControl
End Sub

' ------------------------------------------------------------------------------
' AVANCE DE UN SOLO MICRO-PASO (Pulsar botón STEP)
' ------------------------------------------------------------------------------
Public Sub EjecutarUnMicroPaso()
    If RegIR = &H0 And MicroPasoT = 3 Then
        ' CPU detenida por HLT
        DetalleMicroPaso = "CPU Detenida por instrucción HLT. Presione RESET para reiniciar."
        EnEjecucion = False
        ActualizarTodoUI
        Exit Sub
    End If
    
    ' Limpiar señales de control previas y buses
    ResetBuses
    ResetSeñalesControl
    
    Select Case MicroPasoT
        Case 0 ' ================= T0: FETCH (PC -> MAR) =================
            FaseNombre = "FETCH"
            DetalleMicroPaso = "T0 [FETCH]: PC expone dirección al Bus Interno -> MAR captura dirección."
            
            ' Fase 1: Setup
            EmitirMicroPalabra 7, 5, 0, 0, 0, 0 ' Emisor PC(7) -> Destino MAR(5)
            CPU_EmitirABusInterno "PC"
            Ld_MAR = True
            
            ' Fase 2: Latch
            CPU_CapturarDeBusInterno "MAR"
            LiberarBusInterno
            MicroPasoT = 1
            
        Case 1 ' ================= T1: FETCH (Solicitar OpCode, PC++) =================
            FaseNombre = "FETCH"
            DetalleMicroPaso = "T1 [FETCH]: MAR expone Dir al Bus Ext, MEM_READ=1. Caché L1I/L2/RAM entrega OpCode. PC incrementa."
            
            ' Fase 1: Setup
            EmitirMicroPalabra 0, 6, 1, 0, 0, 0 ' MEM_READ = 1, Destino MDR
            CPU_MAR_ExponerDireccion
            MemRead = 1
            
            ' Obtener OpCode a través de la Jerarquía de Memoria
            Dim dirFisicaInstr As Long
            dirFisicaInstr = MMU_Traducir(RegMAR, False)
            Dim byteOpcode As Long
            byteOpcode = Cache_LeerInstruccion(dirFisicaInstr)
            
            ' Exponer en Bus de Datos Externo
            EmitirAlBusDatosExt "Caché/RAM", byteOpcode
            Ld_MDR = True
            
            ' Fase 2: Latch
            CPU_MDR_CapturarDeBusExterno
            LiberarBusDatosExt
            LiberarBusDireccionesExt
            MemRead = 0
            
            ' PC se incrementa en 1 (apunta al siguiente byte)
            RegPC = (RegPC + 1) And 511
            MicroPasoT = 2
            
        Case 2 ' ================= T2: FETCH (MDR -> IR) =================
            FaseNombre = "FETCH"
            DetalleMicroPaso = "T2 [FETCH]: MDR vuelca OpCode al Bus Interno -> IR captura la instrucción."
            
            ' Fase 1: Setup
            EmitirMicroPalabra 6, 7, 0, 0, 0, 0 ' Emisor MDR(6) -> Destino IR
            CPU_EmitirABusInterno "MDR"
            Ld_IR = True
            
            ' Fase 2: Latch
            CPU_CapturarDeBusInterno "IR"
            LiberarBusInterno
            OpcodeActual = RegIR
            MicroPasoT = 3
            
        Case 3 ' ================= T3: DECODE =================
            FaseNombre = "DECODE"
            DecodificarInstruccionCompleta
            DetalleMicroPaso = "T3 [DECODE]: UC decodifica OpCode 0x" & Right$("0" & Hex$(OpcodeActual), 2) & _
                               " (" & MnemonicoActual & ")."
            
            If OpcodeActual = &H0 Then ' HLT
                DetalleMicroPaso = "T3 [DECODE]: Instrucción HLT reconocida. El reloj se detiene."
                EnEjecucion = False
            Else
                MicroPasoT = 4
            End If
            
        Case 4, 5, 6 ' ================= T4..T6: EXECUTE / STORE =================
            FaseNombre = IIf(MicroPasoT = 6, "STORE", "EXECUTE")
            EjecutarMicroPasoExecuteStore
    End Select
    
    ' Refrescar interfaz de usuario
    ActualizarTodoUI
End Sub

' ------------------------------------------------------------------------------
' DECODIFICACIÓN COMPLETA DE INSTRUCCIÓN (Fase T3)
' ------------------------------------------------------------------------------
Private Sub DecodificarInstruccionCompleta()
    Dim b1 As Long, b2 As Long
    
    Select Case OpcodeActual
        Case &H0 ' HLT
            TamanoInstruccionActual = 1
            MnemonicoActual = "HLT"
            
        Case &H1 ' MOV r, imm
            TamanoInstruccionActual = 3
            Operando1 = MemoriaVirtual_LeerByte(RegPC)
            Operando2 = MemoriaVirtual_LeerByte(RegPC + 1)
            RegPC = (RegPC + 2) And 511
            MnemonicoActual = "MOV " & NombreRegPorId(Operando1) & ", 0x" & Hex$(Operando2)
            
        Case &H2 ' MOV r, r
            TamanoInstruccionActual = 3
            Operando1 = MemoriaVirtual_LeerByte(RegPC)
            Operando2 = MemoriaVirtual_LeerByte(RegPC + 1)
            RegPC = (RegPC + 2) And 511
            MnemonicoActual = "MOV " & NombreRegPorId(Operando1) & ", " & NombreRegPorId(Operando2)
            
        Case &H3 ' LOAD r, [dir]
            TamanoInstruccionActual = 3
            b1 = MemoriaVirtual_LeerByte(RegPC)
            b2 = MemoriaVirtual_LeerByte(RegPC + 1)
            RegPC = (RegPC + 2) And 511
            Operando1 = b1 And 3 ' Reg ID
            ' Dirección de 9 bits (0..511)
            Operando2 = ((b1 \ 4) * 256) + b2
            MnemonicoActual = "LOAD " & NombreRegPorId(Operando1) & ", [0x" & Hex$(Operando2) & "]"
            
        Case &H4 ' STORE [dir], r
            TamanoInstruccionActual = 3
            b1 = MemoriaVirtual_LeerByte(RegPC)
            b2 = MemoriaVirtual_LeerByte(RegPC + 1)
            RegPC = (RegPC + 2) And 511
            Operando1 = b1 And 3 ' Reg origen ID
            Operando2 = ((b1 \ 4) * 256) + b2
            MnemonicoActual = "STORE [0x" & Hex$(Operando2) & "], " & NombreRegPorId(Operando1)
            
        Case &H10 ' ADD r, imm
            TamanoInstruccionActual = 3
            Operando1 = MemoriaVirtual_LeerByte(RegPC)
            Operando2 = MemoriaVirtual_LeerByte(RegPC + 1)
            RegPC = (RegPC + 2) And 511
            MnemonicoActual = "ADD " & NombreRegPorId(Operando1) & ", 0x" & Hex$(Operando2)
            
        Case &H11 ' ADD r, r
            TamanoInstruccionActual = 3
            Operando1 = MemoriaVirtual_LeerByte(RegPC)
            Operando2 = MemoriaVirtual_LeerByte(RegPC + 1)
            RegPC = (RegPC + 2) And 511
            MnemonicoActual = "ADD " & NombreRegPorId(Operando1) & ", " & NombreRegPorId(Operando2)
            
        Case &H12 ' SUB r, imm
            TamanoInstruccionActual = 3
            Operando1 = MemoriaVirtual_LeerByte(RegPC)
            Operando2 = MemoriaVirtual_LeerByte(RegPC + 1)
            RegPC = (RegPC + 2) And 511
            MnemonicoActual = "SUB " & NombreRegPorId(Operando1) & ", 0x" & Hex$(Operando2)
            
        Case &H13 ' SUB r, r
            TamanoInstruccionActual = 3
            Operando1 = MemoriaVirtual_LeerByte(RegPC)
            Operando2 = MemoriaVirtual_LeerByte(RegPC + 1)
            RegPC = (RegPC + 2) And 511
            MnemonicoActual = "SUB " & NombreRegPorId(Operando1) & ", " & NombreRegPorId(Operando2)
            
        Case &H14 ' INC r
            TamanoInstruccionActual = 2
            Operando1 = MemoriaVirtual_LeerByte(RegPC)
            RegPC = (RegPC + 1) And 511
            MnemonicoActual = "INC " & NombreRegPorId(Operando1)
            
        Case &H15 ' DEC r
            TamanoInstruccionActual = 2
            Operando1 = MemoriaVirtual_LeerByte(RegPC)
            RegPC = (RegPC + 1) And 511
            MnemonicoActual = "DEC " & NombreRegPorId(Operando1)
            
        Case &H16 ' CMP r, imm
            TamanoInstruccionActual = 3
            Operando1 = MemoriaVirtual_LeerByte(RegPC)
            Operando2 = MemoriaVirtual_LeerByte(RegPC + 1)
            RegPC = (RegPC + 2) And 511
            MnemonicoActual = "CMP " & NombreRegPorId(Operando1) & ", 0x" & Hex$(Operando2)
            
        Case &H17 ' CMP r, r
            TamanoInstruccionActual = 3
            Operando1 = MemoriaVirtual_LeerByte(RegPC)
            Operando2 = MemoriaVirtual_LeerByte(RegPC + 1)
            RegPC = (RegPC + 2) And 511
            MnemonicoActual = "CMP " & NombreRegPorId(Operando1) & ", " & NombreRegPorId(Operando2)
            
        Case &H18 To &H1A ' AND, OR, XOR
            TamanoInstruccionActual = 3
            Operando1 = MemoriaVirtual_LeerByte(RegPC)
            Operando2 = MemoriaVirtual_LeerByte(RegPC + 1)
            RegPC = (RegPC + 2) And 511
            Dim opLogica As String
            opLogica = IIf(OpcodeActual = &H18, "AND", IIf(OpcodeActual = &H19, "OR", "XOR"))
            MnemonicoActual = opLogica & " " & NombreRegPorId(Operando1) & ", " & NombreRegPorId(Operando2)
            
        Case &H1B ' NOT r
            TamanoInstruccionActual = 2
            Operando1 = MemoriaVirtual_LeerByte(RegPC)
            RegPC = (RegPC + 1) And 511
            MnemonicoActual = "NOT " & NombreRegPorId(Operando1)
            
        Case &H20 To &H23 ' JMP, JZ, JNZ, JC
            TamanoInstruccionActual = 3
            b1 = MemoriaVirtual_LeerByte(RegPC)
            b2 = MemoriaVirtual_LeerByte(RegPC + 1)
            RegPC = (RegPC + 2) And 511
            Operando2 = (b1 * 256) + b2
            Dim mnSalto As String
            Select Case OpcodeActual
                Case &H20: mnSalto = "JMP"
                Case &H21: mnSalto = "JZ"
                Case &H22: mnSalto = "JNZ"
                Case &H23: mnSalto = "JC"
            End Select
            MnemonicoActual = mnSalto & " 0x" & Hex$(Operando2)
            
        Case &H30 ' IN r, [puerto]
            TamanoInstruccionActual = 3
            Operando1 = MemoriaVirtual_LeerByte(RegPC)
            Operando2 = MemoriaVirtual_LeerByte(RegPC + 1)
            RegPC = (RegPC + 2) And 511
            MnemonicoActual = "IN " & NombreRegPorId(Operando1) & ", [0x" & Hex$(Operando2) & "]"
            
        Case &H31 ' OUT [puerto], r
            TamanoInstruccionActual = 3
            Operando1 = MemoriaVirtual_LeerByte(RegPC) ' Reg origen
            Operando2 = MemoriaVirtual_LeerByte(RegPC + 1) ' Puerto destino
            RegPC = (RegPC + 2) And 511
            MnemonicoActual = "OUT [0x" & Hex$(Operando2) & "], " & NombreRegPorId(Operando1)
            
        Case Else
            TamanoInstruccionActual = 1
            MnemonicoActual = "OP_DESCONOCIDO (0x" & Hex$(OpcodeActual) & ")"
    End Select
End Sub

' ------------------------------------------------------------------------------
' EJECUCIÓN DETALLADA DE MICROPASOS T4, T5, T6 (EXECUTE Y STORE)
' ------------------------------------------------------------------------------
Private Sub EjecutarMicroPasoExecuteStore()
    Select Case OpcodeActual
        ' ----------------------------------------------------------------------
        ' LOAD r, [dir]
        ' ----------------------------------------------------------------------
        Case &H3
            Select Case MicroPasoT
                Case 4
                    DetalleMicroPaso = "T4 [LOAD]: Dirección 0x" & Hex$(Operando2) & " viaja al MAR."
                    EmitirMicroPalabra 0, 5, 0, 0, 0, 0
                    RegMAR = Operando2
                    ExponerDireccionExterna RegMAR
                    MicroPasoT = 5
                Case 5
                    DetalleMicroPaso = "T5 [LOAD]: MEM_READ=1. Caché L1D/L2/RAM entrega dato al Bus de Datos Ext -> MDR absorbe."
                    EmitirMicroPalabra 0, 6, 1, 0, 0, 0
                    MemRead = 1
                    Dim dirFisica As Long
                    dirFisica = MMU_Traducir(RegMAR, False)
                    Dim datoLeido As Long
                    datoLeido = Cache_LeerDato(dirFisica)
                    EmitirAlBusDatosExt "Caché L1D/RAM", datoLeido
                    Ld_MDR = True
                    CPU_MDR_CapturarDeBusExterno
                    LiberarBusDatosExt
                    LiberarBusDireccionesExt
                    MemRead = 0
                    MicroPasoT = 6
                Case 6
                    DetalleMicroPaso = "T6 [STORE]: MDR vuelca dato a Bus Interno -> " & NombreRegPorId(Operando1) & " captura el dato. Ciclo completado."
                    EmitirMicroPalabra 6, Operando1 + 1, 0, 0, 0, 0
                    CPU_EmitirABusInterno "MDR"
                    SetRegPorId Operando1, RegMDR
                    LiberarBusInterno
                    TerminarCicloInstruccion
            End Select
            
        ' ----------------------------------------------------------------------
        ' STORE [dir], r
        ' ----------------------------------------------------------------------
        Case &H4
            Select Case MicroPasoT
                Case 4
                    DetalleMicroPaso = "T4 [STORE]: Dirección destino 0x" & Hex$(Operando2) & " viaja al MAR."
                    EmitirMicroPalabra 0, 5, 0, 0, 0, 0
                    RegMAR = Operando2
                    ExponerDireccionExterna RegMAR
                    MicroPasoT = 5
                Case 5
                    DetalleMicroPaso = "T5 [STORE]: Registro " & NombreRegPorId(Operando1) & " vuelca dato a MDR vía Bus Interno."
                    EmitirMicroPalabra Operando1 + 1, 6, 0, 0, 0, 0
                    RegMDR = GetRegPorId(Operando1)
                    MicroPasoT = 6
                Case 6
                    DetalleMicroPaso = "T6 [STORE]: MAR expone Dir, MDR expone dato, MEM_WRITE=1. Grabación en Caché L1D (Write-Back) / RAM. Ciclo completado."
                    EmitirMicroPalabra 6, 0, 0, 1, 0, 0
                    MemWrite = 1
                    EmitirAlBusDatosExt "MDR", RegMDR
                    Dim dirFisicaW As Long
                    dirFisicaW = MMU_Traducir(RegMAR, True)
                    Cache_EscribirDato dirFisicaW, RegMDR
                    LiberarBusDatosExt
                    LiberarBusDireccionesExt
                    MemWrite = 0
                    TerminarCicloInstruccion
            End Select
            
        ' ----------------------------------------------------------------------
        ' ENTRADA POR PUERTO: IN r, [puerto]
        ' ----------------------------------------------------------------------
        Case &H30
            Select Case MicroPasoT
                Case 4
                    DetalleMicroPaso = "T4 [IN]: Número de puerto 0x" & Hex$(Operando2) & " se envía al MAR."
                    EmitirMicroPalabra 0, 5, 0, 0, 0, 0
                    RegMAR = Operando2
                    ExponerDireccionExterna RegMAR
                    MicroPasoT = 5
                Case 5
                    DetalleMicroPaso = "T5 [IN]: MAR expone puerto, IO_READ=1 (RAM inhibida). Periférico Teclado vuelca dato a Bus Ext -> MDR captura."
                    EmitirMicroPalabra 0, 6, 0, 0, 1, 0
                    IoRead = 1
                    Dim datoPuerto As Long
                    datoPuerto = Periferico_LeerPuerto(RegMAR)
                    EmitirAlBusDatosExt "Teclado (Puerto 0x" & Hex$(RegMAR) & ")", datoPuerto
                    Ld_MDR = True
                    CPU_MDR_CapturarDeBusExterno
                    LiberarBusDatosExt
                    LiberarBusDireccionesExt
                    IoRead = 0
                    MicroPasoT = 6
                Case 6
                    DetalleMicroPaso = "T6 [IN]: MDR vuelca dato a Bus Interno -> Registro " & NombreRegPorId(Operando1) & " captura el valor. Ciclo completado."
                    EmitirMicroPalabra 6, Operando1 + 1, 0, 0, 0, 0
                    SetRegPorId Operando1, RegMDR
                    TerminarCicloInstruccion
            End Select
            
        ' ----------------------------------------------------------------------
        ' SALIDA POR PUERTO: OUT [puerto], r
        ' ----------------------------------------------------------------------
        Case &H31
            Select Case MicroPasoT
                Case 4
                    DetalleMicroPaso = "T4 [OUT]: Número de puerto 0x" & Hex$(Operando2) & " al MAR y dato de " & NombreRegPorId(Operando1) & " al MDR."
                    RegMAR = Operando2
                    RegMDR = GetRegPorId(Operando1)
                    MicroPasoT = 5
                Case 5
                    DetalleMicroPaso = "T5 [OUT]: MAR expone puerto, MDR expone dato en Bus Ext. Preparando pulso de IO_WRITE."
                    ExponerDireccionExterna RegMAR
                    EmitirAlBusDatosExt "MDR", RegMDR
                    MicroPasoT = 6
                Case 6
                    DetalleMicroPaso = "T6 [OUT]: IO_WRITE=1. Cerrojo (Latch) de Pantalla LCD congela el carácter visualmente. Ciclo completado."
                    EmitirMicroPalabra 6, 0, 0, 0, 0, 1
                    IoWrite = 1
                    Periferico_EscribirPuerto RegMAR, RegMDR
                    LiberarBusDatosExt
                    LiberarBusDireccionesExt
                    IoWrite = 0
                    TerminarCicloInstruccion
            End Select
            
        ' ----------------------------------------------------------------------
        ' MOV r, imm y MOV r, r
        ' ----------------------------------------------------------------------
        Case &H1, &H2
            Select Case MicroPasoT
                Case 4, 5
                    DetalleMicroPaso = "T4-T5 [EXECUTE]: Transferencia interna directa al registro destino " & NombreRegPorId(Operando1) & "."
                    MicroPasoT = 6
                Case 6
                    If OpcodeActual = &H1 Then
                        SetRegPorId Operando1, Operando2
                    Else
                        SetRegPorId Operando1, GetRegPorId(Operando2)
                    End If
                    DetalleMicroPaso = "T6 [STORE]: " & MnemonicoActual & " completado. Ciclo finalizado."
                    TerminarCicloInstruccion
            End Select
            
        ' ----------------------------------------------------------------------
        ' OPERACIONES ARITMÉTICO-LÓGICAS (ALU)
        ' ----------------------------------------------------------------------
        Case &H10 To &H1B
            Dim nombreOpALU As String
            Select Case OpcodeActual
                Case &H10, &H11: nombreOpALU = "ADD"
                Case &H12, &H13: nombreOpALU = "SUB"
                Case &H14: nombreOpALU = "INC"
                Case &H15: nombreOpALU = "DEC"
                Case &H16, &H17: nombreOpALU = "CMP"
                Case &H18: nombreOpALU = "AND"
                Case &H19: nombreOpALU = "OR"
                Case &H1A: nombreOpALU = "XOR"
                Case &H1B: nombreOpALU = "NOT"
            End Select
            
            Select Case MicroPasoT
                Case 4
                    DetalleMicroPaso = "T4 [ALU]: " & NombreRegPorId(Operando1) & " se vuelca a TempA. Buses externos en silencio."
                    RegTempA = GetRegPorId(Operando1)
                    MicroPasoT = 5
                Case 5
                    DetalleMicroPaso = "T5 [ALU]: Segundo operando se vuelca a TempB."
                    If OpcodeActual = &H11 Or OpcodeActual = &H13 Or OpcodeActual = &H17 Or (OpcodeActual >= &H18 And OpcodeActual <= &H1A) Then
                        RegTempB = GetRegPorId(Operando2)
                    ElseIf OpcodeActual = &H10 Or OpcodeActual = &H12 Or OpcodeActual = &H16 Then
                        RegTempB = Operando2
                    Else
                        RegTempB = 1 ' Para INC / DEC
                    End If
                    MicroPasoT = 6
                Case 6
                    Dim resALU As Long
                    resALU = EjecutarALU(nombreOpALU, RegTempA, RegTempB)
                    If nombreOpALU <> "CMP" Then
                        SetRegPorId Operando1, resALU
                    End If
                    DetalleMicroPaso = "T6 [ALU]: " & nombreOpALU & " ejecutado. Banderas: ZF=" & FlagZF & ", CF=" & FlagCF & ", SF=" & FlagSF & ". Ciclo finalizado."
                    TerminarCicloInstruccion
            End Select
            
        ' ----------------------------------------------------------------------
        ' SALTOS: JMP, JZ, JNZ, JC
        ' ----------------------------------------------------------------------
        Case &H20 To &H23
            Select Case MicroPasoT
                Case 4, 5
                    DetalleMicroPaso = "T4-T5 [BRANCH]: Evaluación de condiciones de banderas en la UC."
                    MicroPasoT = 6
                Case 6
                    Dim salta As Boolean
                    salta = False
                    Select Case OpcodeActual
                        Case &H20: salta = True
                        Case &H21: If FlagZF = 1 Then salta = True
                        Case &H22: If FlagZF = 0 Then salta = True
                        Case &H23: If FlagCF = 1 Then salta = True
                    End Select
                    
                    If salta Then
                        RegPC = Operando2 And 511
                        DetalleMicroPaso = "T6 [BRANCH]: Salto tomado -> PC actualizado a 0x" & Hex$(RegPC) & "."
                    Else
                        DetalleMicroPaso = "T6 [BRANCH]: Condición no cumplida. No se toma el salto."
                    End If
                    TerminarCicloInstruccion
            End Select
            
        Case Else
            TerminarCicloInstruccion
    End Select
End Sub

Private Sub TerminarCicloInstruccion()
    MicroPasoT = 0
    TotalInstruccionesEjecutadas = TotalInstruccionesEjecutadas + 1
End Sub

' ------------------------------------------------------------------------------
' MODOS DE CONTROL: RUN, PAUSE, STEP COMPLETO, RESET
' ------------------------------------------------------------------------------

' Ejecución continua automática (RUN)
Public Sub EjecutarModoRUN()
    If EnEjecucion Then Exit Sub
    EnEjecucion = True
    EsPausado = False
    
    ' Leer retardo de la celda F7 de la hoja CPU si existe
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("CPU")
    If Not ws Is Nothing Then
        If IsNumeric(ws.Range("F7").Value) Then
            RetardoMs = CLng(ws.Range("F7").Value)
            If RetardoMs < 0 Then RetardoMs = 0
        End If
    End If
    On Error GoTo 0
    
    Do While EnEjecucion And Not EsPausado
        EjecutarUnMicroPaso
        DoEvents
        If RetardoMs > 0 Then Sleep RetardoMs
        If RegIR = &H0 And MicroPasoT = 0 Then
            EnEjecucion = False
            Exit Do
        End If
    Loop
    EnEjecucion = False
End Sub

' Pausa la ejecución (PAUSE)
Public Sub EjecutarModoPAUSE()
    EsPausado = True
    EnEjecucion = False
    DetalleMicroPaso = "Ejecución pausada por el usuario. Presione STEP o RUN para continuar."
    ActualizarTodoUI
End Sub

' Reinicia todos los módulos del computador (RESET)
Public Sub EjecutarModoRESET()
    ResetCPU
    ResetMemoriaVirtual
    ResetCache
    ResetPerifericos
    ResetBuses
    ResetCiclo
    ActualizarTodoUI
End Sub

' ==============================================================================
' ENLACES EXACTOS PARA LOS BOTONES GRÁFICOS DE LA HOJA EXCEL
' ==============================================================================

Public Sub BotonSTEP()
    EjecutarUnMicroPaso
End Sub

Public Sub BotonRUN()
    EjecutarModoRUN
End Sub

Public Sub BotonPAUSE()
    EjecutarModoPAUSE
End Sub

Public Sub BotonResetCiclo()
    EjecutarModoRESET
End Sub

Public Sub BotonRESET()
    EjecutarModoRESET
End Sub

Public Sub BotonCargarPrograma()
    CargarProgramaCatalogo 1
End Sub

Public Sub BotonCargarFibo()
    CargarProgramaCatalogo 2
End Sub

