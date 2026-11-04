Option Explicit

' ==============================================================================
' MODULO: Buses.bas
' Simulador Avanzado de Arquitectura de Computadoras - Segundo Parcial (SIS-131)
' Modela:
'   1. Bus Interno de Datos (Datapath de la CPU)
'   2. Bus Externo de Direcciones (Address Bus)
'   3. Bus Externo de Datos bidireccional (Data Bus)
'   4. Bus de Control (MEM_READ, MEM_WRITE, IO_READ, IO_WRITE, CLK)
'   5. Buffers Tri-Estado (0, 1, Alta Impedancia 'Z') y Regla de Exclusión Mutua
'   6. Micro-palabra de Control (Control Word)
' ==============================================================================

' Estado de alta impedancia representado formalmente
Public Const HIGH_Z As Long = -1

' Variables que representan los buses fisicos en cada instante de reloj
Public BusInterno As Long          ' Valor en el bus interno (0..255) o HIGH_Z (-1)
Public BusDireccionesExt As Long   ' Direccion emitida por MAR (0..255) o HIGH_Z (-1)
Public BusDatosExt As Long         ' Valor en el bus de datos externo (0..255) o HIGH_Z (-1)

' Señales de control del Bus Externo (1 = Activo, 0 = Inactivo)
Public MemRead As Long
Public MemWrite As Long
Public IoRead As Long
Public IoWrite As Long

' Nombre del modulo emisor activo en cada bus (para inspeccion visual)
Public EmisorBusInterno As String
Public EmisorBusDatosExt As String

' Vector de Micro-Palabra de Control:
' Estructura: [EMISOR_ID : 3b] | [DESTINO_ID : 3b] | MEM_R | MEM_W | IO_R | IO_W
' IDs Internos:
'   000 (0) = NINGUNO
'   001 (1) = AX
'   010 (2) = BX
'   011 (3) = CX
'   100 (4) = DX
'   101 (5) = MAR
'   110 (6) = MDR
'   111 (7) = ALU_OUT / PC / TEMP
Public ControlWord_EmisorId As Long
Public ControlWord_DestinoId As Long
Public ControlWord_Str As String

' Buffer Tri-Estado: estructura para cada registro/periferico
Public Type TriStateBuffer
    Nombre As String
    Habilitado As Boolean   ' Output Enable (OE = 1)
    ValorEntrada As Long    ' Dato interno a exponer (0..255)
    EstadoVisual As String  ' "0", "1", o "Z" (High-Z)
End Type

' Inicializa los buses en estado de Alta Impedancia (Z)
Public Sub ResetBuses()
    BusInterno = HIGH_Z
    BusDireccionesExt = HIGH_Z
    BusDatosExt = HIGH_Z
    MemRead = 0
    MemWrite = 0
    IoRead = 0
    IoWrite = 0
    EmisorBusInterno = "NINGUNO (High-Z)"
    EmisorBusDatosExt = "NINGUNO (High-Z)"
    ControlWord_EmisorId = 0
    ControlWord_DestinoId = 0
    ControlWord_Str = "000|000|0|0|0|0"
End Sub

' ------------------------------------------------------------------------------
' Regla de Oro: Ecuacion de Exclusion Mutua Externa
' MEM_READ + MEM_WRITE + IO_READ + IO_WRITE <= 1
' Suma = 1: Operacion externa valida.
' Suma = 0: Operacion interna pura (ALU o registros; buses externos en reposo).
' Suma > 1: ERROR DE COLISION DE BUS (Cortocircuito logico).
' ------------------------------------------------------------------------------
Public Function ValidarExclusionMutuaExterna() As Boolean
    Dim suma As Long
    suma = MemRead + MemWrite + IoRead + IoWrite
    If suma > 1 Then
        Err.Raise vbObjectError + 1001, "Buses::ValidarExclusionMutuaExterna", _
            "¡ERROR DE COLISION DE BUS! Cortocircuito logico detectado: " & _
            "MEM_R=" & MemRead & ", MEM_W=" & MemWrite & ", IO_R=" & IoRead & ", IO_W=" & IoWrite
        ValidarExclusionMutuaExterna = False
    Else
        ValidarExclusionMutuaExterna = True
    End If
End Function

' Coloca un dato en el Bus Interno respetando exclusion mutua
Public Sub EmitirAlBusInterno(ByVal nombreEmisor As String, ByVal valor As Long)
    If BusInterno <> HIGH_Z And EmisorBusInterno <> "NINGUNO (High-Z)" Then
        Err.Raise vbObjectError + 1002, "Buses::EmitirAlBusInterno", _
            "¡COLISION EN BUS INTERNO! Emisor '" & nombreEmisor & "' intenta hablar pero '" & _
            EmisorBusInterno & "' ya tiene el control del bus."
    End If
    BusInterno = valor And 255
    EmisorBusInterno = nombreEmisor
End Sub

' Libera el Bus Interno colocandolo en Alta Impedancia (High-Z)
Public Sub LiberarBusInterno()
    BusInterno = HIGH_Z
    EmisorBusInterno = "NINGUNO (High-Z)"
End Sub

' Coloca un dato en el Bus Externo de Datos
Public Sub EmitirAlBusDatosExt(ByVal nombreEmisor As String, ByVal valor As Long)
    If BusDatosExt <> HIGH_Z And EmisorBusDatosExt <> "NINGUNO (High-Z)" Then
        Err.Raise vbObjectError + 1003, "Buses::EmitirAlBusDatosExt", _
            "¡COLISION EN BUS DE DATOS EXTERNO! Emisor '" & nombreEmisor & "' compite con '" & _
            EmisorBusDatosExt & "'."
    End If
    BusDatosExt = valor And 255
    EmisorBusDatosExt = nombreEmisor
End Sub

' Libera el Bus Externo de Datos (High-Z)
Public Sub LiberarBusDatosExt()
    BusDatosExt = HIGH_Z
    EmisorBusDatosExt = "NINGUNO (High-Z)"
End Sub

' El MAR expone la direccion en el Bus de Direcciones Externo
Public Sub ExponerDireccionExterna(ByVal direccion As Long)
    BusDireccionesExt = direccion And &H1FF ' Soporta hasta 512B virtuales / 256B fisicos
End Sub

' Libera el Bus de Direcciones Externo
Public Sub LiberarBusDireccionesExt()
    BusDireccionesExt = HIGH_Z
End Sub

' Emite la micro-palabra de control en la Unidad de Control
Public Sub EmitirMicroPalabra(ByVal emisorId As Long, ByVal destinoId As Long, _
                              ByVal mr As Long, ByVal mw As Long, _
                              ByVal ior As Long, ByVal iow As Long)
    ControlWord_EmisorId = emisorId
    ControlWord_DestinoId = destinoId
    MemRead = mr
    MemWrite = mw
    IoRead = ior
    IoWrite = iow
    
    ' Validar inmediatamente la exclusion mutua
    ValidarExclusionMutuaExterna
    
    ' Generar representacion en cadena binaria
    ControlWord_Str = DecToBin3(emisorId) & "|" & DecToBin3(destinoId) & "|" & _
                      mr & "|" & mw & "|" & ior & "|" & iow
End Sub

Private Function DecToBin3(ByVal n As Long) As String
    Dim b As String
    b = ""
    b = IIf((n And 4) <> 0, "1", "0") & _
        IIf((n And 2) <> 0, "1", "0") & _
        IIf((n And 1) <> 0, "1", "0")
    DecToBin3 = b
End Function

' Retorna el estado visual formateado de un buffer ("Z" o "0xNN")
Public Function FormatoBufferVisual(ByVal enable As Boolean, ByVal valor As Long) As String
    If enable Then
        FormatoBufferVisual = "0x" & Right$("0" & Hex$(valor And 255), 2)
    Else
        FormatoBufferVisual = "High-Z"
    End If
End Function
