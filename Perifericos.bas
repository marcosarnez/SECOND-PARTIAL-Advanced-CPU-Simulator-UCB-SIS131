Option Explicit

' ==============================================================================
' MODULO: Perifericos.bas
' Simulador Avanzado de Arquitectura de Computadoras - Segundo Parcial (SIS-131)
' Modela:
'   1. Periférico de Entrada: Teclado interactivo
'      - Puerto 0x01: Registro de Estado (Bit 0 = 1 si hay tecla disponible)
'      - Puerto 0x02: Registro de Datos (Código ASCII; autolimpia el estado al leer)
'      - Buffer Tri-Estado: En_KBD
'   2. Cola Circular FIFO en Memoria RAM (8 Bytes: 0x70 a 0x77):
'      - 0x78: KB_HEAD (Puntero de lectura 0..7)
'      - 0x79: KB_TAIL (Puntero de escritura 0..7)
'      - 0x7A: KB_COUNT (Cantidad de teclas en cola 0..8)
'   3. Periférico de Salida: Pantalla LCD 8x1 con Cerrojo (Video Latch)
'      - 8 Celdas de caracteres (0xF8 a 0xFF en RAM)
'      - Latch de retención para evitar parpadeos dinámicos del bus
'      - Señal de captura: Ld_LCD
' ==============================================================================

' Direcciones de recursos mapeados en RAM
Public Const DIR_COLA_KBD As Long = &H70      ' 0x70 a 0x77 (8 bytes)
Public Const DIR_KB_HEAD As Long = &H78       ' Puntero Head
Public Const DIR_KB_TAIL As Long = &H79       ' Puntero Tail
Public Const DIR_KB_COUNT As Long = &H7A      ' Contador de elementos
Public Const DIR_VRAM_LCD As Long = &HF8      ' 0xF8 a 0xFF (8 bytes)

' Puertos de E/S Aislada
Public Const PUERTO_KBD_STATUS As Long = &H1  ' 0x01
Public Const PUERTO_KBD_DATA As Long = &H2    ' 0x02
Public Const PUERTO_LCD_DATA As Long = &H3    ' 0x03

' Registros locales del hardware del Teclado
Public RegKbdStatus As Long  ' Bit 0 = 1 si tecla lista
Public RegKbdData As Long    ' Código ASCII retenido en el teclado físico
Public En_KBD As Boolean     ' Señal Output Enable del buffer Tri-Estado del teclado

' Registros del Cerrojo (Video Latch) de la Pantalla LCD 8x1
Public LCD_Latch(0 To 7) As Long
Public Ld_LCD As Boolean     ' Señal Load para congelar el dato en la pantalla
Public PosicionCursorLCD As Long

' Resetea el hardware de Teclado y Pantalla LCD
Public Sub ResetPerifericos()
    Dim i As Long
    
    RegKbdStatus = 0
    RegKbdData = 0
    En_KBD = False
    Ld_LCD = False
    PosicionCursorLCD = 0
    
    ' Inicializar el Latch de pantalla con espacios en blanco (ASCII 32)
    For i = 0 To 7
        LCD_Latch(i) = 32
    Next i
    
    ' Inicializar la cola FIFO en la memoria RAM
    InicializarColaRAM
End Sub

' Inicializa los punteros y celdas de la cola de teclado en RAM
Public Sub InicializarColaRAM()
    Dim i As Long
    For i = 0 To 7
        MemoriaVirtual_CargarByte DIR_COLA_KBD + i, 0
    Next i
    MemoriaVirtual_CargarByte DIR_KB_HEAD, 0
    MemoriaVirtual_CargarByte DIR_KB_TAIL, 0
    MemoriaVirtual_CargarByte DIR_KB_COUNT, 0
    
    ' Inicializar celdas de la Reserva VRAM en RAM (0xF0 - 0xFF) con 0
    For i = 0 To 15
        MemoriaVirtual_CargarByte &HF0 + i, 0
    Next i
End Sub

' ------------------------------------------------------------------------------
' INTERACCIÓN DE ENTRADA: El usuario presiona una tecla en el panel de Excel
' ------------------------------------------------------------------------------
Public Sub Teclado_PulsarTecla(ByVal teclaAscii As Long)
    teclaAscii = teclaAscii And 255
    RegKbdData = teclaAscii
    RegKbdStatus = 1 ' Bit 0 activo: tecla lista
End Sub

' Simula escribir una cadena de texto en el teclado ingresando los caracteres a la cola FIFO
Public Sub Teclado_EncolarTexto(ByVal texto As String)
    Dim i As Long, codigo As Long
    For i = 1 To Len(texto)
        codigo = Asc(Mid$(texto, i, 1))
        Teclado_InsertarEnColaFIFO codigo
    Next i
    ' Si el registro de datos del teclado está libre, transferir automáticamente la primera tecla
    If RegKbdStatus = 0 And MemoriaVirtual_LeerByte(DIR_KB_COUNT) > 0 Then
        codigo = Teclado_ExtraerDeColaFIFO()
        Teclado_PulsarTecla codigo
    End If
End Sub

' Inserción de una tecla en la Cola Circular de RAM (Productor)
Public Function Teclado_InsertarEnColaFIFO(ByVal codigoAscii As Long) As Boolean
    Dim count As Long, tail As Long
    
    count = MemoriaVirtual_LeerByte(DIR_KB_COUNT)
    
    ' Detección de Desbordamiento (Overflow): si count = 8, descartar tecla
    If count >= 8 Then
        Teclado_InsertarEnColaFIFO = False
        Exit Function
    End If
    
    tail = MemoriaVirtual_LeerByte(DIR_KB_TAIL)
    
    ' Escribir el carácter en 0x70 + KB_TAIL
    MemoriaVirtual_CargarByte DIR_COLA_KBD + tail, codigoAscii
    
    ' KB_TAIL = (KB_TAIL + 1) mod 8
    tail = (tail + 1) Mod 8
    MemoriaVirtual_CargarByte DIR_KB_TAIL, tail
    
    ' KB_COUNT++
    count = count + 1
    MemoriaVirtual_CargarByte DIR_KB_COUNT, count
    
    Teclado_InsertarEnColaFIFO = True
End Function

' Extracción de una tecla de la Cola Circular de RAM (Consumidor)
Public Function Teclado_ExtraerDeColaFIFO() As Long
    Dim count As Long, head As Long, codigo As Long
    
    count = MemoriaVirtual_LeerByte(DIR_KB_COUNT)
    If count = 0 Then
        Teclado_ExtraerDeColaFIFO = 0 ' Cola vacía
        Exit Function
    End If
    
    head = MemoriaVirtual_LeerByte(DIR_KB_HEAD)
    codigo = MemoriaVirtual_LeerByte(DIR_COLA_KBD + head)
    
    ' KB_HEAD = (KB_HEAD + 1) mod 8
    head = (head + 1) Mod 8
    MemoriaVirtual_CargarByte DIR_KB_HEAD, head
    
    ' KB_COUNT--
    count = count - 1
    MemoriaVirtual_CargarByte DIR_KB_COUNT, count
    
    Teclado_ExtraerDeColaFIFO = codigo
End Function

' ------------------------------------------------------------------------------
' LECTURA POR PUERTO (Instrucción IN reg, [puerto]) - E/S Aislada
' ------------------------------------------------------------------------------
Public Function Periferico_LeerPuerto(ByVal numPuerto As Long) As Long
    Select Case numPuerto
        Case PUERTO_KBD_STATUS ' Puerto 0x01: Estado
            En_KBD = True
            Periferico_LeerPuerto = RegKbdStatus
            
        Case PUERTO_KBD_DATA   ' Puerto 0x02: Dato ASCII
            En_KBD = True
            Periferico_LeerPuerto = RegKbdData
            ' Al leer el dato del puerto 0x02:
            ' Si aún quedan teclas en la cola FIFO, transferir automáticamente la siguiente al registro
            If MemoriaVirtual_LeerByte(DIR_KB_COUNT) > 0 Then
                Dim proximaTecla As Long
                proximaTecla = Teclado_ExtraerDeColaFIFO()
                RegKbdData = proximaTecla
                RegKbdStatus = 1
            Else
                RegKbdStatus = 0
            End If
            
        Case Else
            Periferico_LeerPuerto = 0
    End Select
End Function

' ------------------------------------------------------------------------------
' ESCRITURA EN PUERTO (Instrucción OUT [puerto], reg) - Latch de Pantalla LCD
' ------------------------------------------------------------------------------
Public Sub Periferico_EscribirPuerto(ByVal numPuerto As Long, ByVal valor As Long)
    valor = valor And 255
    Select Case numPuerto
        Case PUERTO_LCD_DATA
            Ld_LCD = True
            LCD_EscribirCaracter valor
        Case Else
            ' Puerto de depuración o futuro uso
    End Select
End Sub

' ------------------------------------------------------------------------------
' PANTALLA LCD 8x1: Operación con Cerrojo (Video Latch) y Soporte de Backspace
' ------------------------------------------------------------------------------
Public Sub LCD_EscribirCaracter(ByVal asciiVal As Long)
    ' Soporte de Backspace (ASCII 0x08)
    If asciiVal = 8 Then
        If PosicionCursorLCD > 0 Then
            PosicionCursorLCD = PosicionCursorLCD - 1
            LCD_Latch(PosicionCursorLCD) = 32 ' Espacio en blanco
            MemoriaVirtual_CargarByte DIR_VRAM_LCD + PosicionCursorLCD, 0
        End If
        Exit Sub
    End If
    
    ' Si el cursor llegó al final de la pantalla (8 caracteres), hacer Scroll hacia la izquierda
    If PosicionCursorLCD >= 8 Then
        LCD_HacerScrollIzquierda
        PosicionCursorLCD = 7
    End If
    
    ' Congelar el dato en el Latch y en la memoria de video RAM
    LCD_Latch(PosicionCursorLCD) = asciiVal
    MemoriaVirtual_CargarByte DIR_VRAM_LCD + PosicionCursorLCD, asciiVal
    
    PosicionCursorLCD = PosicionCursorLCD + 1
End Sub

' Desplazamiento horizontal (Scroll) cuando se llena la pantalla
Public Sub LCD_HacerScrollIzquierda()
    Dim i As Long
    For i = 0 To 6
        LCD_Latch(i) = LCD_Latch(i + 1)
        MemoriaVirtual_CargarByte DIR_VRAM_LCD + i, LCD_Latch(i)
    Next i
    LCD_Latch(7) = 32 ' Espacio al final
    MemoriaVirtual_CargarByte DIR_VRAM_LCD + 7, 0
End Sub

' Devuelve el texto actualmente proyectado en el display LCD (8 caracteres)
Public Function LCD_ObtenerTexto() As String
    Dim res As String, i As Long, ch As Long
    res = ""
    For i = 0 To 7
        ch = LCD_Latch(i)
        If ch >= 32 And ch <= 126 Then
            res = res & Chr$(ch)
        Else
            res = res & " "
        End If
    Next i
    LCD_ObtenerTexto = res
End Function
