Option Explicit

' ==============================================================================
' MODULO: CPU.bas
' Simulador Avanzado de Arquitectura de Computadoras - Segundo Parcial (SIS-131)
' Modela:
'   1. Registros de Propósito General (8 bits):
'      - AX (0): Acumulador principal
'      - BX (1): Registro Base
'      - CX (2): Contador de bucles
'      - DX (3): Datos y periféricos de E/S
'   2. Registros de Control del Ciclo:
'      - PC (Contador de Programa, 0..511 para Memoria Virtual)
'      - IR (Registro de Instrucción, 8 bits)
'      - MAR (Memory Address Register - Aduana de Direcciones)
'      - MDR (Memory Data Register - Aduana Transceptora de Datos)
'   3. Registros de la ALU:
'      - TempA (Operando 1)
'      - TempB (Operando 2)
'      - ALU_OUT (Resultado retenido)
'   4. Banderas (Flags): ZF (Zero), CF (Carry), SF (Sign)
'   5. Conexiones con Buses mediante Buffers Tri-Estado
' ==============================================================================

' Registros de Propósito General (8 bits)
Public RegAX As Long
Public RegBX As Long
Public RegCX As Long
Public RegDX As Long

' Registros de Control
Public RegPC As Long   ' 0..511 (Espacio Virtual)
Public RegIR As Long   ' 0..255
Public RegMAR As Long  ' 0..511
Public RegMDR As Long  ' 0..255

' Registros Temporales de la ALU
Public RegTempA As Long
Public RegTempB As Long
Public RegALUOut As Long

' Banderas de la CPU (0 o 1)
Public FlagZF As Long
Public FlagCF As Long
Public FlagSF As Long

' Estados de activación de señales Load (receptores)
Public Ld_AX As Boolean
Public Ld_BX As Boolean
Public Ld_CX As Boolean
Public Ld_DX As Boolean
Public Ld_PC As Boolean
Public Ld_IR As Boolean
Public Ld_MAR As Boolean
Public Ld_MDR As Boolean
Public Ld_TempA As Boolean
Public Ld_TempB As Boolean

' Estados de activación de señales Enable (emisores a bus interno)
Public En_AX As Boolean
Public En_BX As Boolean
Public En_CX As Boolean
Public En_DX As Boolean
Public En_PC As Boolean
Public En_MDR As Boolean
Public En_ALU As Boolean

' Limita cualquier número a un byte (0..255)
Public Function Byte8(ByVal valor As Long) As Long
    Byte8 = valor And 255
End Function

' Resetea todos los registros y banderas de la CPU
Public Sub ResetCPU()
    RegAX = 0
    RegBX = 0
    RegCX = 0
    RegDX = 0
    RegPC = 0
    RegIR = 0
    RegMAR = 0
    RegMDR = 0
    RegTempA = 0
    RegTempB = 0
    RegALUOut = 0
    FlagZF = 0
    FlagCF = 0
    FlagSF = 0
    ResetSeñalesControl
End Sub

' Desactiva todas las señales de Load y Enable (reposo)
Public Sub ResetSeñalesControl()
    Ld_AX = False: Ld_BX = False: Ld_CX = False: Ld_DX = False
    Ld_PC = False: Ld_IR = False: Ld_MAR = False: Ld_MDR = False
    Ld_TempA = False: Ld_TempB = False
    En_AX = False: En_BX = False: En_CX = False: En_DX = False
    En_PC = False: En_MDR = False: En_ALU = False
End Sub

' Obtener valor de un registro general por su código (0=AX, 1=BX, 2=CX, 3=DX)
Public Function GetRegPorId(ByVal id As Long) As Long
    Select Case id And 3
        Case 0: GetRegPorId = RegAX
        Case 1: GetRegPorId = RegBX
        Case 2: GetRegPorId = RegCX
        Case 3: GetRegPorId = RegDX
    End Select
End Function

' Escribir valor en un registro general por su código
Public Sub SetRegPorId(ByVal id As Long, ByVal valor As Long)
    Select Case id And 3
        Case 0: RegAX = Byte8(valor)
        Case 1: RegBX = Byte8(valor)
        Case 2: RegCX = Byte8(valor)
        Case 3: RegDX = Byte8(valor)
    End Select
End Sub

' Nombre del registro según ID
Public Function NombreRegPorId(ByVal id As Long) As String
    Select Case id And 3
        Case 0: NombreRegPorId = "AX"
        Case 1: NombreRegPorId = "BX"
        Case 2: NombreRegPorId = "CX"
        Case 3: NombreRegPorId = "DX"
    End Select
End Function

' Obtener ID según nombre
Public Function IdRegPorNombre(ByVal nombre As String) As Long
    Select Case UCase$(Trim$(nombre))
        Case "AX": IdRegPorNombre = 0
        Case "BX": IdRegPorNombre = 1
        Case "CX": IdRegPorNombre = 2
        Case "DX": IdRegPorNombre = 3
        Case Else: IdRegPorNombre = 0
    End Select
End Function

' ------------------------------------------------------------------------------
' ALU de 8 bits con cálculo de banderas ZF, CF, SF
' ------------------------------------------------------------------------------
Public Function EjecutarALU(ByVal op As String, ByVal a As Long, ByVal b As Long) As Long
    Dim r As Long
    Dim c As Long
    Dim logica As Boolean
    
    a = Byte8(a)
    b = Byte8(b)
    c = 0
    logica = False
    
    Select Case UCase$(Trim$(op))
        Case "ADD"
            r = a + b
            If r > 255 Then c = 1
            r = Byte8(r)
            
        Case "SUB", "CMP"
            If a < b Then
                c = 1
                r = Byte8(a - b + 256)
            Else
                r = a - b
            End If
            
        Case "INC"
            r = a + 1
            If r > 255 Then c = 1
            r = Byte8(r)
            
        Case "DEC"
            If a = 0 Then
                c = 1
                r = 255
            Else
                r = a - 1
            End If
            
        Case "AND"
            r = a And b
            logica = True
            
        Case "OR"
            r = a Or b
            logica = True
            
        Case "XOR"
            r = a Xor b
            logica = True
            
        Case "NOT"
            r = (Not a) And 255
            logica = True
            
        Case Else
            r = a
    End Select
    
    RegALUOut = Byte8(r)
    
    ' Actualización de Banderas
    FlagZF = IIf(RegALUOut = 0, 1, 0)
    FlagSF = IIf((RegALUOut And 128) <> 0, 1, 0)
    If logica Then
        FlagCF = 0
    Else
        FlagCF = c
    End If
    
    EjecutarALU = RegALUOut
End Function

' ------------------------------------------------------------------------------
' INTERFAZ CON BUSES Y ADUANAS MAR / MDR
' ------------------------------------------------------------------------------

' Emite el contenido de un registro al Bus Interno vía su Buffer Tri-Estado
Public Sub CPU_EmitirABusInterno(ByVal componente As String)
    Select Case UCase$(Trim$(componente))
        Case "AX"
            En_AX = True
            EmitirAlBusInterno "AX", RegAX
        Case "BX"
            En_BX = True
            EmitirAlBusInterno "BX", RegBX
        Case "CX"
            En_CX = True
            EmitirAlBusInterno "CX", RegCX
        Case "DX"
            En_DX = True
            EmitirAlBusInterno "DX", RegDX
        Case "PC"
            En_PC = True
            EmitirAlBusInterno "PC", RegPC
        Case "MDR"
            En_MDR = True
            EmitirAlBusInterno "MDR", RegMDR
        Case "ALU"
            En_ALU = True
            EmitirAlBusInterno "ALU_OUT", RegALUOut
    End Select
End Sub

' Captura el dato presente en el Bus Interno (Flanco de Reloj / Latch)
Public Sub CPU_CapturarDeBusInterno(ByVal destino As String)
    If BusInterno = HIGH_Z Then
        Err.Raise vbObjectError + 1004, "CPU::CPU_CapturarDeBusInterno", _
            "No se puede capturar en '" & destino & "': El Bus Interno esta en Alta Impedancia (High-Z)."
    End If
    
    Select Case UCase$(Trim$(destino))
        Case "AX"
            RegAX = Byte8(BusInterno)
        Case "BX"
            RegBX = Byte8(BusInterno)
        Case "CX"
            RegCX = Byte8(BusInterno)
        Case "DX"
            RegDX = Byte8(BusInterno)
        Case "MAR"
            RegMAR = BusInterno And &H1FF ' Admite hasta 512B
            ExponerDireccionExterna RegMAR
        Case "MDR"
            RegMDR = Byte8(BusInterno)
        Case "IR"
            RegIR = Byte8(BusInterno)
        Case "TEMPA"
            RegTempA = Byte8(BusInterno)
        Case "TEMPB"
            RegTempB = Byte8(BusInterno)
        Case "PC"
            RegPC = BusInterno And &H1FF
    End Select
End Sub

' La Aduana MAR expone su dirección al Bus Externo
Public Sub CPU_MAR_ExponerDireccion()
    ExponerDireccionExterna RegMAR
End Sub

' La Aduana MDR emite su dato al Bus de Datos Externo (Escritura hacia RAM/Periféricos)
Public Sub CPU_MDR_EmitirABusExterno()
    EmitirAlBusDatosExt "MDR (CPU)", RegMDR
End Sub

' La Aduana MDR absorbe el dato proveniente del Bus de Datos Externo (Lectura desde RAM/Periféricos)
Public Sub CPU_MDR_CapturarDeBusExterno()
    If BusDatosExt = HIGH_Z Then
        Err.Raise vbObjectError + 1005, "CPU::CPU_MDR_CapturarDeBusExterno", _
            "MDR no puede capturar: El Bus de Datos Externo esta en Alta Impedancia (High-Z)."
    End If
    RegMDR = Byte8(BusDatosExt)
End Sub
