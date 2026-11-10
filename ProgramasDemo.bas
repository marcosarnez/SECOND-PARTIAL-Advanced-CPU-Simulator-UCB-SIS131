Option Explicit

' ==============================================================================
' MODULO: ProgramasDemo.bas
' Simulador Avanzado de Arquitectura de Computadoras - Segundo Parcial (SIS-131)
' Catálogo de Programas Oficiales requeridos por la Rúbrica:
'   1. Prog 1: Cálculo aritmético con bucles usando los 4 registros (AX, BX, CX, DX).
'   2. Prog 2: Entrada de texto desde Teclado con eco interactivo a la Pantalla LCD 8x1 (con Backspace).
'   3. Prog 3: Accesos dispersos a memoria que demuestran Page Faults y fallos en Caché (L1/L2).
' ==============================================================================

Public Function ObtenerCodigoPrograma1() As String
    Dim s As String
    s = "; ==========================================" & vbCrLf
    s = s & "; PROGRAMA 1: CALCULO ARITMETICO CON BUCLES" & vbCrLf
    s = s & "; Uso obligatorio de AX, BX, CX y DX" & vbCrLf
    s = s & "; ==========================================" & vbCrLf
    s = s & "MOV AX, 0x00      ; AX = Acumulador total = 0" & vbCrLf
    s = s & "MOV BX, 0x05      ; BX = Valor a sumar en cada paso (5)" & vbCrLf
    s = s & "MOV CX, 0x04      ; CX = Contador de iteraciones del bucle (4)" & vbCrLf
    s = s & "MOV DX, 0x01      ; DX = Indicador de estado de ejecucion" & vbCrLf
    s = s & "BUCLE:" & vbCrLf
    s = s & "ADD AX, BX        ; Acumular: AX = AX + BX" & vbCrLf
    s = s & "INC DX            ; DX cuenta los pasos internos" & vbCrLf
    s = s & "DEC CX            ; Decrementar contador de bucle" & vbCrLf
    s = s & "JNZ BUCLE         ; Repetir mientras CX != 0" & vbCrLf
    s = s & "STORE [0x60], AX  ; Guardar resultado final (20 = 0x14) en RAM de Datos" & vbCrLf
    s = s & "STORE [0x61], DX  ; Guardar contador de pasos en RAM de Datos" & vbCrLf
    s = s & "HLT               ; Detener el procesador"
    ObtenerCodigoPrograma1 = s
End Function

Public Function ObtenerCodigoPrograma2() As String
    Dim s As String
    s = "; ==========================================" & vbCrLf
    s = s & "; PROGRAMA 2: ENTRADA/SALIDA INTERACTIVA" & vbCrLf
    s = s & "; Teclado (Puertos 1,2) -> LCD 8x1 (Puerto 3)" & vbCrLf
    s = s & "; ==========================================" & vbCrLf
    s = s & "MOV CX, 0x08      ; CX = Limite maximo de caracteres (8)" & vbCrLf
    s = s & "POLL_TECLA:" & vbCrLf
    s = s & "IN AX, [0x01]     ; Leer Puerto 0x01 (Status del Teclado)" & vbCrLf
    s = s & "CMP AX, 0x00     ; ¿Hay tecla disponible? (Bit 0 == 1)" & vbCrLf
    s = s & "JZ POLL_TECLA     ; Si no hay tecla, seguir sondeando" & vbCrLf
    s = s & "IN DX, [0x02]     ; Leer codigo ASCII del Puerto 0x02 (y resetea Status)" & vbCrLf
    s = s & "OUT [0x03], DX    ; Enviar caracter al Puerto 0x03 (Pantalla LCD con Latch)" & vbCrLf
    s = s & "DEC CX            ; Decrementar cuenta" & vbCrLf
    s = s & "JNZ POLL_TECLA    ; Continuar hasta completar la pantalla" & vbCrLf
    s = s & "HLT               ; Detener al llenar el display"
    ObtenerCodigoPrograma2 = s
End Function

Public Function ObtenerCodigoPrograma3() As String
    Dim s As String
    s = "; ==========================================" & vbCrLf
    s = s & "; PROGRAMA 3: FALLOS DE PAGINA (PAGE FAULTS)" & vbCrLf
    s = s & "; Y ACCESOS DISPERSOS EN CACHE MULTINIVEL" & vbCrLf
    s = s & "; ==========================================" & vbCrLf
    s = s & "MOV AX, 0xAA      ; Patron de datos de prueba A" & vbCrLf
    s = s & "MOV BX, 0xBB      ; Patron de datos de prueba B" & vbCrLf
    s = s & "; 1. Acceso a Pagina Virtual 0 (Dir 0x010)" & vbCrLf
    s = s & "STORE [0x10], AX" & vbCrLf
    s = s & "; 2. Acceso a Pagina Virtual 2 (Dir 0x040)" & vbCrLf
    s = s & "STORE [0x40], BX" & vbCrLf
    s = s & "; 3. Acceso a Pagina Virtual 4 (Dir 0x080)" & vbCrLf
    s = s & "STORE [0x80], AX" & vbCrLf
    s = s & "; 4. Acceso a Pagina Virtual 6 (Dir 0x0C0)" & vbCrLf
    s = s & "STORE [0xC0], BX" & vbCrLf
    s = s & "; 5. Acceso a Pagina Virtual 8 (Dir 0x100 -> Espacio Virtual Alto)" & vbCrLf
    s = s & "STORE [0x100], AX" & vbCrLf
    s = s & "; 6. Acceso a Pagina Virtual 10 (Dir 0x140)" & vbCrLf
    s = s & "STORE [0x140], BX" & vbCrLf
    s = s & "; 7. Acceso a Pagina Virtual 12 (Dir 0x180)" & vbCrLf
    s = s & "STORE [0x180], AX" & vbCrLf
    s = s & "; 8. Acceso a Pagina Virtual 14 (Dir 0x1C0) -> RAM Llena (8 marcos)" & vbCrLf
    s = s & "STORE [0x1C0], BX" & vbCrLf
    s = s & "; 9. Acceso a Pagina Virtual 15 (Dir 0x1E0) -> ¡FORZAR PAGE FAULT Y DESALOJO!" & vbCrLf
    s = s & "STORE [0x1E0], AX" & vbCrLf
    s = s & "; 10. Re-acceso a Pagina Virtual 0 (desalojada a Swap) -> ¡SEGUNDO PAGE FAULT!" & vbCrLf
    s = s & "LOAD CX, [0x10]" & vbCrLf
    s = s & "HLT"
    ObtenerCodigoPrograma3 = s
End Function

Public Sub CargarProgramaCatalogo(ByVal numeroPrograma As Long)
    Dim src As String
    Select Case numeroPrograma
        Case 1: src = ObtenerCodigoPrograma1
        Case 2: src = ObtenerCodigoPrograma2
        Case 3: src = ObtenerCodigoPrograma3
        Case Else: src = ObtenerCodigoPrograma1
    End Select
    
    ResetCPU
    ResetMemoriaVirtual
    ResetCache
    ResetPerifericos
    ResetBuses
    
    If Not EnsamblarYCargar(src, 0) Then
        MsgBox "Error al ensamblar programa " & numeroPrograma & ": " & UltimoErrorEnsamblador, vbCritical, "Error Ensamblador"
    Else
        If numeroPrograma = 2 Then
            Teclado_EncolarTexto "HOLA"
        End If
        ActualizarTodoUI
    End If
End Sub

' ==============================================================================
' MACROS SIN PARÁMETROS PARA ASIGNAR DIRECTAMENTE A BOTONES EN EXCEL (ALT + F8)
' ==============================================================================

Public Sub Cargar_Programa_1()
    CargarProgramaCatalogo 1
End Sub

Public Sub Cargar_Programa_2()
    CargarProgramaCatalogo 2
End Sub

Public Sub Cargar_Programa_3()
    CargarProgramaCatalogo 3
End Sub

Public Sub Boton_STEP()
    EjecutarUnMicroPaso
End Sub

Public Sub Boton_RUN()
    EjecutarModoRUN
End Sub

Public Sub Boton_PAUSE()
    EjecutarModoPAUSE
End Sub

Public Sub Boton_RESET()
    EjecutarModoRESET
End Sub

