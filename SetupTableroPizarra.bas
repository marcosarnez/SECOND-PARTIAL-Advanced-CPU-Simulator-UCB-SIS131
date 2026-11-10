Attribute VB_Name = "SetupTableroPizarra"
Option Explicit

' ==============================================================================
' MODULO: SetupTableroPizarra.bas
' Integra en la hoja principal de Excel la arquitectura exacta de la pizarra:
'   - Bloques: CPU/ALU, MEMORIA RAM, DISCO/SWAP, MONITOR LCD, TECLADO
'   - 4 Buses visuales: Control, Direcciones, Datos y Bus Interno
'   - Editor de codigo ASM interactivo (como la foto del docente)
'   - Teclado 100% manejable por el usuario
'   - Visualizador de celdas de RAM fisica y Disco Swap
' ==============================================================================

Public Sub GenerarArquitecturaPizarraCompleta()
    Dim ws As Worksheet
    Dim wb As Workbook
    Set wb = ThisWorkbook
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Asegurar hoja "ARQUITECTURA" o integrarlo en la hoja principal
    On Error Resume Next
    Set ws = wb.Worksheets("ARQUITECTURA")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = "ARQUITECTURA"
    Else
        ' Limpiar shapes y contenido anterior
        ws.Unprotect ""
        ws.Cells.Locked = False
        Dim shp As Shape
        For Each shp In ws.Shapes
            shp.Delete
        Next shp
        ws.Cells.Clear
    End If
    
    ws.Unprotect ""
    ws.Cells.Locked = False
    
    ws.Activate
    ActiveWindow.DisplayGridlines = True
    
    ' Estilo general
    ws.Cells.Interior.Color = RGB(245, 247, 250)
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 9
    
    ' Ajustar columnas principales
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B:F").ColumnWidth = 9   ' CPU
    ws.Columns("G").ColumnWidth = 4     ' Espacio
    ws.Columns("H:L").ColumnWidth = 9   ' Memoria RAM
    ws.Columns("M").ColumnWidth = 4     ' Espacio
    ws.Columns("N:R").ColumnWidth = 9   ' Disco / Swap
    ws.Columns("S").ColumnWidth = 4     ' Espacio
    ws.Columns("T:Y").ColumnWidth = 9   ' Editor Programas
    
    ' =========================================================================
    ' 1. BANNER SUPERIOR
    ' =========================================================================
    With ws.Range("B2:Y3")
        .Merge
        .value = "SIMULADOR DE ARQUITECTURA DE COMPUTADORAS - SISTEMA DE BUSES, MEMORIA Y E/S"
        .Font.Size = 13
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(26, 37, 48)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    ' =========================================================================
    ' 2. BLOQUE CPU / ALU (Naranja / Azul) - Filas 5 a 16, Col B a F
    ' =========================================================================
    CrearBloqueVisual ws, 5, 2, 16, 6, "CPU (UNIDAD CENTRAL DE PROCESAMIENTO)", RGB(230, 126, 34), RGB(253, 242, 233)
    
    ' Sub-bloque ALU
    ws.Range("B8:D11").Merge
    ws.Range("B8").value = "ALU" & vbCrLf & "(Unidad Aritmetico-Logica)"
    ws.Range("B8").HorizontalAlignment = xlCenter
    ws.Range("B8").VerticalAlignment = xlCenter
    ws.Range("B8").Font.Bold = True
    ws.Range("B8").Interior.Color = RGB(243, 156, 18)
    ws.Range("B8").Font.Color = RGB(255, 255, 255)
    
    ' Registros AX, BX, CX, DX
    ws.Cells(8, 5).value = "AX:"
    ws.Cells(8, 6).value = 0
    ws.Cells(9, 5).value = "BX:"
    ws.Cells(9, 6).value = 0
    ws.Cells(10, 5).value = "CX:"
    ws.Cells(10, 6).value = 0
    ws.Cells(11, 5).value = "DX:"
    ws.Cells(11, 6).value = 0
    
    ' Registros de Control PC, IR, MAR, MDR
    ws.Cells(13, 2).value = "PC:"
    ws.Cells(13, 3).value = "0x000"
    ws.Cells(13, 4).value = "IR:"
    ws.Cells(13, 5).value = "0x00"
    ws.Cells(14, 2).value = "MAR:"
    ws.Cells(14, 3).value = "0x000"
    ws.Cells(14, 4).value = "MDR:"
    ws.Cells(14, 5).value = "0x00"
    
    ' Bus Interno dentro de la CPU
    ws.Range("B15:F15").Merge
    ws.Range("B15").value = "[ BUS INTERNO CPU (8 bits) ]: High-Z"
    ws.Range("B15").Font.Size = 8
    ws.Range("B15").Font.Bold = True
    ws.Range("B15").Interior.Color = RGB(155, 89, 182)
    ws.Range("B15").Font.Color = RGB(255, 255, 255)
    ws.Range("B15").HorizontalAlignment = xlCenter
    
    ' =========================================================================
    ' 3. BLOQUE MEMORIA RAM (Azul) - Filas 5 a 16, Col H a L
    ' =========================================================================
    CrearBloqueVisual ws, 5, 8, 16, 12, "MEMORIA RAM FISICA (256 Bytes)", RGB(41, 128, 185), RGB(235, 245, 251)
    
    ' Contenido resumido de Marcos
    Dim m As Long
    For m = 0 To 3
        ws.Cells(7 + m, 8).value = "Marco " & m & ":"
        ws.Cells(7 + m, 9).value = "0x" & Right$("00" & Hex$(m * 32), 2)
        ws.Cells(7 + m, 10).value = "VPN " & m
        ws.Cells(7 + m, 11).value = "[Presente]"
        ws.Cells(7 + m, 11).Font.Color = RGB(39, 174, 96)
    Next m
    For m = 4 To 7
        ws.Cells(7 + m, 8).value = "Marco " & m & ":"
        ws.Cells(7 + m, 9).value = "0x" & Right$("00" & Hex$(m * 32), 2)
        ws.Cells(7 + m, 10).value = "VPN " & m
        ws.Cells(7 + m, 11).value = "[Presente]"
        ws.Cells(7 + m, 11).Font.Color = RGB(39, 174, 96)
    Next m
    
    ws.Range("H15:L15").Merge
    ws.Range("H15").value = "MMU: Trad. Dir. Virtual -> Fisica"
    ws.Range("H15").Font.Size = 8
    ws.Range("H15").Font.Italic = True
    ws.Range("H15").HorizontalAlignment = xlCenter
    
    ' =========================================================================
    ' 4. BLOQUE DISCO / SWAP (Naranja oscuro / Marron) - Filas 5 a 16, Col N a R
    ' =========================================================================
    CrearBloqueVisual ws, 5, 14, 16, 18, "DISCO / AREA DE SWAP (512 Bytes)", RGB(211, 84, 0), RGB(251, 238, 230)
    
    ws.Cells(7, 14).value = "Total Paginas:"
    ws.Cells(7, 16).value = "16 (VPN 0-15)"
    ws.Cells(8, 14).value = "Tamano Pagina:"
    ws.Cells(8, 16).value = "32 Bytes"
    ws.Cells(9, 14).value = "Paginas en RAM:"
    ws.Cells(9, 16).value = "8 Marcos"
    ws.Cells(10, 14).value = "Paginas en Swap:"
    ws.Cells(10, 16).value = "8 Paginas"
    ws.Cells(12, 14).value = "Ultimo Swap-Out:"
    ws.Cells(12, 16).value = "Ninguno"
    ws.Cells(13, 14).value = "Ultimo Swap-In:"
    ws.Cells(13, 16).value = "Ninguno"
    
    ws.Range("N15:R15").Merge
    ws.Range("N15").value = "Fallo de Pagina -> Atencion por Swap"
    ws.Range("N15").Font.Size = 8
    ws.Range("N15").Font.Italic = True
    ws.Range("N15").HorizontalAlignment = xlCenter
    
    ' =========================================================================
    ' 5. EDITOR INTERACTIVO DE PROGRAMAS (Como en la foto 2 del docente)
    ' Col T a Y, Filas 5 a 16
    ' =========================================================================
    CrearBloqueVisual ws, 5, 20, 16, 25, "PROGRAMAS / EDITOR DE INSTRUCCIONES", RGB(52, 73, 94), RGB(242, 244, 244)
    
    ' Selector o titulo del programa
    ws.Range("T7:Y7").Merge
    ws.Range("T7").value = "Programa Actual: [ Modificable ]"
    ws.Range("T7").Font.Bold = True
    ws.Range("T7").Interior.Color = RGB(255, 235, 59)
    ws.Range("T7").HorizontalAlignment = xlCenter
    
    ' Area de texto para escribir mnem??nicos
    ws.Range("T8:Y13").Merge
    ws.Range("T8").NumberFormat = "@"
    ws.Range("T8").value = "POLL:" & vbCrLf & _
                           "IN AX, [1]" & vbCrLf & _
                           "CMP AX, 0" & vbCrLf & _
                           "JZ FIN" & vbCrLf & _
                           "IN DX, [2]" & vbCrLf & _
                           "OUT [3], DX" & vbCrLf & _
                           "JMP POLL" & vbCrLf & _
                           "FIN:" & vbCrLf & _
                           "HLT"
    ws.Range("T8").Font.Name = "Consolas"
    ws.Range("T8").Font.Size = 8
    ws.Range("T8").Interior.Color = RGB(255, 255, 255)
    ws.Range("T8").Borders.LineStyle = xlContinuous
    ws.Range("T8").VerticalAlignment = xlTop
    
    ' Boton 1: CARGAR / ENSAMBLAR
    Dim rBtn As Range
    Set rBtn = ws.Range("T14:U15")
    Dim btnCargar As Button
    Set btnCargar = ws.Buttons.Add(rBtn.Left, rBtn.Top, rBtn.Width, rBtn.Height)
    With btnCargar
        .Caption = "CARGAR ASM"
        .OnAction = "Boton_CargarProgramaCustom"
        .Font.Bold = True
        .Font.Size = 8
    End With
    
    ' Boton 2: EJECUTAR RUN
    Set rBtn = ws.Range("V14:W15")
    Dim btnRun As Button
    Set btnRun = ws.Buttons.Add(rBtn.Left, rBtn.Top, rBtn.Width, rBtn.Height)
    With btnRun
        .Caption = "EJECUTAR"
        .OnAction = "Boton_EjecutarProgramaCustom"
        .Font.Bold = True
        .Font.Size = 8
    End With
    
    ' Boton 3: PASO A PASO (STEP)
    Set rBtn = ws.Range("X14:Y15")
    Dim btnStep As Button
    Set btnStep = ws.Buttons.Add(rBtn.Left, rBtn.Top, rBtn.Width, rBtn.Height)
    With btnStep
        .Caption = "PASO A PASO"
        .OnAction = "Boton_STEP"
        .Font.Bold = True
        .Font.Size = 8
    End With
    
    ' =========================================================================
    ' 6. LOS 3 BUSES PRINCIPALES (Lineas horizontales que unen todo)
    ' Filas 18 a 22
    ' =========================================================================
    ' Bus de Control (Rojo)
    ws.Range("B18:Y18").Merge
    ws.Range("B18").value = "[ BUS DE CONTROL ]:  MEM_R=0  |  MEM_W=0  |  IO_R=0  |  IO_W=0  |  RELOJ (CLK)=0"
    ws.Range("B18").Interior.Color = RGB(192, 57, 43)
    ws.Range("B18").Font.Color = RGB(255, 255, 255)
    ws.Range("B18").Font.Bold = True
    ws.Range("B18").Font.Name = "Consolas"
    ws.Range("B18").Font.Size = 9
    ws.Range("B18").HorizontalAlignment = xlCenter
    
    ' Bus de Direcciones (Verde)
    ws.Range("B20:Y20").Merge
    ws.Range("B20").value = "[ BUS DE DIRECCIONES (16 bits) ]: MAR -> 0x000 (Direccion actual activa)"
    ws.Range("B20").Interior.Color = RGB(39, 174, 96)
    ws.Range("B20").Font.Color = RGB(255, 255, 255)
    ws.Range("B20").Font.Bold = True
    ws.Range("B20").Font.Name = "Consolas"
    ws.Range("B20").Font.Size = 9
    ws.Range("B20").HorizontalAlignment = xlCenter
    
    ' Bus de Datos (Azul)
    ws.Range("B22:Y22").Merge
    ws.Range("B22").value = "[ BUS DE DATOS (8 bits, Bidireccional) ]: MDR <-> High-Z"
    ws.Range("B22").Interior.Color = RGB(41, 128, 185)
    ws.Range("B22").Font.Color = RGB(255, 255, 255)
    ws.Range("B22").Font.Bold = True
    ws.Range("B22").Font.Name = "Consolas"
    ws.Range("B22").Font.Size = 9
    ws.Range("B22").HorizontalAlignment = xlCenter
    
    ' =========================================================================
    ' 7. BLOQUES INFERIORES: MONITOR LCD Y TECLADO (Verde / Amarillo)
    ' Filas 24 a 32
    ' =========================================================================
    ' MONITOR / PANTALLA LCD 8x1 (Verde) - Col B a G
    CrearBloqueVisual ws, 24, 2, 32, 7, "MONITOR / LCD 8x1 (SALIDA)", RGB(22, 160, 133), RGB(232, 248, 245)
    ws.Range("C26:F27").Merge
    ws.Range("C26").NumberFormat = "@"
    ws.Range("C26").value = "[ READY ]"
    ws.Range("C26").Font.Name = "Consolas"
    ws.Range("C26").Font.Size = 14
    ws.Range("C26").Font.Bold = True
    ws.Range("C26").Font.Color = RGB(0, 40, 0)
    ws.Range("C26").Interior.Color = RGB(46, 204, 113)
    ws.Range("C26").HorizontalAlignment = xlCenter
    ws.Range("C26").VerticalAlignment = xlCenter
    
    ws.Cells(29, 3).value = "Puerto E/S:"
    ws.Cells(29, 5).value = "0x03 (OUT)"
    ws.Cells(30, 3).value = "Buffer VRAM:"
    ws.Cells(30, 5).value = "0xF8 - 0xFF (Fila F0 RAM)"
    
    ' TECLADO INTERACTIVO Y MANEJABLE (Amarillo) - Col I a N
    CrearBloqueVisual ws, 24, 9, 32, 14, "TECLADO INTERACTIVO (ENTRADA)", RGB(241, 196, 15), RGB(254, 249, 231)
    
    ws.Cells(26, 10).value = "Registro Estado:"
    ws.Cells(26, 13).value = 0
    ws.Cells(27, 10).value = "Ultimo Caracter:"
    ws.Cells(27, 13).value = "0x00 (' ')"
    
    ' Boton para que el usuario escriba directamente en el teclado
    Set rBtn = ws.Range("J29:M31")
    Dim btnKbd As Button
    Set btnKbd = ws.Buttons.Add(rBtn.Left, rBtn.Top, rBtn.Width, rBtn.Height)
    With btnKbd
        .Caption = "ESCRIBIR EN EL TECLADO"
        .OnAction = "Boton_EscribirTeclado"
        .Font.Bold = True
        .Font.Size = 9
    End With
    
    ' COLA CIRCULAR FIFO DEL TECLADO - Col P a Y
    CrearBloqueVisual ws, 24, 16, 32, 25, "COLA FIFO DEL TECLADO EN RAM (0x70 - 0x77)", RGB(127, 140, 141), RGB(248, 249, 249)
    
    Dim c As Long
    For c = 0 To 7
        ws.Cells(26, 17 + c).value = "0x7" & c
        ws.Cells(26, 17 + c).Font.Size = 7
        ws.Cells(26, 17 + c).HorizontalAlignment = xlCenter
        
        ws.Cells(27, 17 + c).value = "[ ]"
        ws.Cells(27, 17 + c).Font.Name = "Consolas"
        ws.Cells(27, 17 + c).Font.Bold = True
        ws.Cells(27, 17 + c).HorizontalAlignment = xlCenter
        ws.Cells(27, 17 + c).Interior.Color = RGB(255, 255, 255)
        ws.Cells(27, 17 + c).Borders.LineStyle = xlContinuous
    Next c
    
    ws.Cells(29, 17).value = "Head:"
    ws.Cells(29, 19).value = 0
    ws.Cells(29, 21).value = "Tail:"
    ws.Cells(29, 23).value = 0
    ws.Cells(29, 24).value = "Count:"
    ws.Cells(29, 25).value = 0
    
    ws.Range("Q31:Y31").Merge
    ws.Range("Q31").value = "Escribe texto usando el boton y observa como se llena la cola y luego la CPU la lee."
    ws.Range("Q31").Font.Size = 8
    ws.Range("Q31").Font.Italic = True
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    If Application.UserControl Then
        MsgBox "Tablero con la Arquitectura de la Pizarra generado exitosamente en la hoja 'ARQUITECTURA'." & vbCrLf & _
               "- Bloques CPU, Memoria, Swap, Monitor y Teclado conectados por buses." & vbCrLf & _
               "- Editor de programas interactivo listo." & vbCrLf & _
               "- Teclado interactivo funcional.", vbInformation, "SIS-131 Arquitectura"
    End If
End Sub

Private Sub CrearBloqueVisual(ByVal ws As Worksheet, _
                              ByVal r1 As Long, ByVal c1 As Long, _
                              ByVal r2 As Long, ByVal c2 As Long, _
                              ByVal titulo As String, _
                              ByVal clrEncabezado As Long, _
                              ByVal clrCuerpo As Long)
    ' Encabezado
    With ws.Range(ws.Cells(r1, c1), ws.Cells(r1 + 1, c2))
        .Merge
        .value = titulo
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = clrEncabezado
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    ' Cuerpo
    With ws.Range(ws.Cells(r1 + 2, c1), ws.Cells(r2, c2))
        .Interior.Color = clrCuerpo
        .Borders.LineStyle = xlContinuous
        .Borders.Color = clrEncabezado
    End With
End Sub

' ==============================================================================
' CARGAR / ENSAMBLAR PROGRAMA EDITADO (Sin ejecutar de inmediato)
' ==============================================================================
Public Sub Boton_CargarProgramaCustom()
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ARQUITECTURA")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "No se encontro la hoja 'ARQUITECTURA'.", vbExclamation
        Exit Sub
    End If
    
    Dim codigoTexto As String
    codigoTexto = ws.Range("T8").value
    
    If Len(Trim$(codigoTexto)) = 0 Then
        MsgBox "Escriba al menos una instruccion en la caja de codigo.", vbExclamation
        Exit Sub
    End If
    
    ' Resetear sistema
    EjecutarModoRESET
    
    ' Ensamblar y cargar todo el programa en la memoria virtual
    Dim ok As Boolean
    ok = EnsamblarYCargar(codigoTexto, 0)
    
    If Not ok Then
        MsgBox "Error de sintaxis: " & UltimoErrorEnsamblador, vbCritical, "Error Ensamblador"
        Exit Sub
    End If
    
    ActualizarTodoUI
    If Application.UserControl Then
        MsgBox "Codigo ensamblado y cargado con exito en Memoria Virtual!" & vbCrLf & _
               "Lineas: " & TotalLineasEnsambladas & " | Bytes: " & TotalBytesGenerados & vbCrLf & _
               "Puede avanzar con PASO A PASO (STEP) o presionar EJECUTAR (RUN).", vbInformation, "Ensamblador SIS-131"
    End If
End Sub

' ==============================================================================
' BOTON PARA EJECUTAR EL PROGRAMA QUE EL USUARIO ESCRIBA EN LA CAJA DEL EXCEL
' ==============================================================================
Public Sub Boton_EjecutarProgramaCustom()
    Dim ws As Worksheet
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ARQUITECTURA")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "No se encontro la hoja 'ARQUITECTURA'.", vbExclamation
        Exit Sub
    End If
    
    Dim codigoTexto As String
    codigoTexto = ws.Range("T8").value
    
    If Len(Trim$(codigoTexto)) = 0 Then
        MsgBox "Escriba al menos una instruccion en la caja de codigo.", vbExclamation
        Exit Sub
    End If
    
    ' Preservar cola FIFO del teclado si el usuario ya ingreso texto antes de ejecutar
    Dim kbdActivo As Boolean, savedStatus As Long, savedData As Long
    Dim savedHead As Long, savedTail As Long, savedCount As Long
    Dim savedCola(0 To 7) As Long, idxK As Long
    kbdActivo = (RegKbdStatus = 1)
    If kbdActivo Then
        savedStatus = RegKbdStatus
        savedData = RegKbdData
        savedHead = MemoriaVirtual_LeerByte(DIR_KB_HEAD)
        savedTail = MemoriaVirtual_LeerByte(DIR_KB_TAIL)
        savedCount = MemoriaVirtual_LeerByte(DIR_KB_COUNT)
        For idxK = 0 To 7
            savedCola(idxK) = MemoriaVirtual_LeerByte(DIR_COLA_KBD + idxK)
        Next idxK
    End If
    
    ' Resetear sistema
    EjecutarModoRESET
    
    ' Restaurar teclado si estaba cargado
    If kbdActivo Then
        For idxK = 0 To 7
            MemoriaVirtual_CargarByte DIR_COLA_KBD + idxK, savedCola(idxK)
        Next idxK
        MemoriaVirtual_CargarByte DIR_KB_HEAD, savedHead
        MemoriaVirtual_CargarByte DIR_KB_TAIL, savedTail
        MemoriaVirtual_CargarByte DIR_KB_COUNT, savedCount
        RegKbdStatus = savedStatus
        RegKbdData = savedData
    End If
    
    ' Ensamblar y cargar todo el programa directamente en la memoria virtual
    Dim ok As Boolean
    ok = EnsamblarYCargar(codigoTexto, 0)
    
    If Not ok Then
        MsgBox "Error al ensamblar el codigo.", vbExclamation
        Exit Sub
    End If
    
    ' Iniciar ejecucion
    Boton_RUN
End Sub

' ==============================================================================
' INTERACCION CON EL TECLADO: El usuario escribe texto y va al buffer
' ==============================================================================
Public Sub Boton_EscribirTeclado()
    Dim textoUsuario As String
    textoUsuario = InputBox("Ingrese el texto que desea enviar al Teclado del simulador:" & vbCrLf & _
                            "(Maximo 8 caracteres para el buffer circular)", "Teclado Interactivo SIS-131", "HOLA")
    
    If Len(textoUsuario) = 0 Then Exit Sub
    
    ' Truncar a 8 caracteres maximo
    If Len(textoUsuario) > 8 Then
        textoUsuario = Left$(textoUsuario, 8)
    End If
    
    ' Encolar en el hardware simulado (ya expone el primer caracter en RegKbdData y Status=1)
    Teclado_EncolarTexto textoUsuario
    
    ' Actualizar interfaz en todas las hojas (CPU y ARQUITECTURA)
    ActualizarTodoUI
    
    If Application.UserControl And Application.Visible Then
        MsgBox "Texto '" & textoUsuario & "' cargado en la cola FIFO del teclado." & vbCrLf & _
               "Ahora presione [EJECUTAR] o [PASO A PASO] para verlo pasar al Monitor LCD.", vbInformation, "Teclado"
    End If
End Sub
