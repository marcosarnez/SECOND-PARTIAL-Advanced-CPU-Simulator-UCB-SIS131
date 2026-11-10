Attribute VB_Name = "Ensamblador"
Option Explicit

' ==============================================================================
' MODULO: Ensamblador.bas
' Simulador Avanzado de Arquitectura de Computadoras - Segundo Parcial (SIS-131)
' Funcionalidades:
'   1. Parser y Ensamblador de Código Mnemónico a Código Máquina (8 bits).
'   2. Soporte completo para los 4 Registros: AX (0), BX (1), CX (2), DX (3).
'   3. Soporte para instrucciones de E/S: IN reg, [puerto] y OUT [puerto], reg.
'   4. Soporte para etiquetas (Labels) y saltos (JMP, JZ, JNZ, JC).
'   5. Direccionamiento para todo el espacio de Memoria Virtual (0x000 a 0x1FF, 512B).
'   6. Carga automática de binarios a las páginas virtuales de memoria.
' ==============================================================================

Public UltimoErrorEnsamblador As String
Public TotalLineasEnsambladas As Long
Public TotalBytesGenerados As Long

' Estructura para resolución de etiquetas en dos pasadas
Private Type Etiqueta
    Nombre As String
    Direccion As Long
End Type

Private Etiquetas(0 To 63) As Etiqueta
Private NumEtiquetas As Long

' Resetea el estado del ensamblador
Public Sub ResetEnsamblador()
    UltimoErrorEnsamblador = ""
    TotalLineasEnsambladas = 0
    TotalBytesGenerados = 0
    NumEtiquetas = 0
End Sub

' ------------------------------------------------------------------------------
' ENSAMBLADOR PRINCIPAL (Dos Pasadas)
' Entrada: Array de strings o texto multilínea con código ASM
' Salida: True si compila con éxito, False con mensaje en UltimoErrorEnsamblador
' ------------------------------------------------------------------------------
Public Function EnsamblarYCargar(ByVal codigoFuente As String, ByVal dirInicio As Long) As Boolean
    Dim lineas() As String
    Dim i As Long
    Dim lineaLimpia As String
    Dim pcVirtual As Long
    Dim bytesInstr() As Long
    Dim numBytes As Long
    Dim k As Long
    
    ResetEnsamblador
    dirInicio = dirInicio And 511
    
    ' Separar líneas por saltos de línea (CRLF o LF)
    codigoFuente = Replace$(codigoFuente, vbCrLf, vbLf)
    codigoFuente = Replace$(codigoFuente, vbCr, vbLf)
    lineas = Split(codigoFuente, vbLf)
    
    ' ==========================================================================
    ' PASADA 1: Identificar etiquetas y calcular direcciones de memoria
    ' ==========================================================================
    pcVirtual = dirInicio
    For i = 0 To UBound(lineas)
        lineaLimpia = LimpiarLinea(lineas(i))
        If Len(lineaLimpia) > 0 Then
            ' Verificar si la línea contiene una etiqueta (termina en ':')
            If InStr(lineaLimpia, ":") > 0 Then
                Dim partesLabel() As String
                partesLabel = Split(lineaLimpia, ":")
                RegistrarEtiqueta Trim$(partesLabel(0)), pcVirtual
                lineaLimpia = Trim$(partesLabel(1))
            End If
            
            If Len(lineaLimpia) > 0 Then
                numBytes = CalcularTamanoInstruccion(lineaLimpia)
                If numBytes = 0 Then
                    UltimoErrorEnsamblador = "Error en línea " & (i + 1) & ": Instrucción desconocida '" & lineaLimpia & "'"
                    EnsamblarYCargar = False
                    Exit Function
                End If
                pcVirtual = pcVirtual + numBytes
            End If
        End If
    Next i
    
    ' ==========================================================================
    ' PASADA 2: Generar código máquina y cargar en Memoria Virtual
    ' ==========================================================================
    pcVirtual = dirInicio
    For i = 0 To UBound(lineas)
        lineaLimpia = LimpiarLinea(lineas(i))
        If Len(lineaLimpia) > 0 Then
            If InStr(lineaLimpia, ":") > 0 Then
                Dim p() As String
                p = Split(lineaLimpia, ":")
                lineaLimpia = Trim$(p(1))
            End If
            
            If Len(lineaLimpia) > 0 Then
                If Not CodificarInstruccion(lineaLimpia, pcVirtual, bytesInstr) Then
                    UltimoErrorEnsamblador = "Error en línea " & (i + 1) & ": " & UltimoErrorEnsamblador
                    EnsamblarYCargar = False
                    Exit Function
                End If
                
                ' Grabar bytes en la Memoria Virtual
                For k = 0 To UBound(bytesInstr)
                    MemoriaVirtual_CargarByte pcVirtual + k, bytesInstr(k)
                Next k
                
                pcVirtual = pcVirtual + UBound(bytesInstr) + 1
                TotalLineasEnsambladas = TotalLineasEnsambladas + 1
            End If
        End If
    Next i
    
    TotalBytesGenerados = pcVirtual - dirInicio
    EnsamblarYCargar = True
End Function

' ------------------------------------------------------------------------------
' RUTINAS DE PARSEO Y CODIFICACIÓN
' ------------------------------------------------------------------------------

Private Function LimpiarLinea(ByVal s As String) As String
    Dim posComentario As Long
    posComentario = InStr(s, ";")
    If posComentario > 0 Then s = Left$(s, posComentario - 1)
    LimpiarLinea = Trim$(s)
End Function

Private Sub RegistrarEtiqueta(ByVal nombre As String, ByVal direccion As Long)
    If NumEtiquetas < 64 Then
        Etiquetas(NumEtiquetas).Nombre = UCase$(Trim$(nombre))
        Etiquetas(NumEtiquetas).Direccion = direccion
        NumEtiquetas = NumEtiquetas + 1
    End If
End Sub

Private Function BuscarEtiqueta(ByVal nombre As String) As Long
    Dim i As Long
    nombre = UCase$(Trim$(nombre))
    For i = 0 To NumEtiquetas - 1
        If Etiquetas(i).Nombre = nombre Then
            BuscarEtiqueta = Etiquetas(i).Direccion
            Exit Function
        End If
    Next i
    BuscarEtiqueta = -1
End Function

Private Function ParseValorODir(ByVal s As String) As Long
    s = Trim$(s)
    ' Si está entre corchetes [0xNN]
    If Left$(s, 1) = "[" And Right$(s, 1) = "]" Then
        s = Trim$(Mid$(s, 2, Len(s) - 2))
    End If
    
    ' Si es etiqueta
    Dim dirEtiqueta As Long
    dirEtiqueta = BuscarEtiqueta(s)
    If dirEtiqueta <> -1 Then
        ParseValorODir = dirEtiqueta
        Exit Function
    End If
    
    ' Si es hexadecimal 0xNN o NNh
    If LCase$(Left$(s, 2)) = "0x" Then
        ParseValorODir = Val("&H" & Mid$(s, 3))
    ElseIf LCase$(Right$(s, 1)) = "h" Then
        ParseValorODir = Val("&H" & Left$(s, Len(s) - 1))
    Else
        ParseValorODir = Val(s)
    End If
End Function

Private Function CalcularTamanoInstruccion(ByVal linea As String) As Long
    Dim partes() As String
    Dim mnem As String
    partes = Split(linea, " ")
    mnem = UCase$(Trim$(partes(0)))
    
    Select Case mnem
        Case "HLT": CalcularTamanoInstruccion = 1
        Case "INC", "DEC", "NOT": CalcularTamanoInstruccion = 2
        Case "JMP", "JZ", "JNZ", "JC": CalcularTamanoInstruccion = 3
        Case "MOV", "LOAD", "STORE", "ADD", "SUB", "CMP", "AND", "OR", "XOR", "IN", "OUT": CalcularTamanoInstruccion = 3
        Case Else: CalcularTamanoInstruccion = 0
    End Select
End Function

Private Function CodificarInstruccion(ByVal linea As String, ByVal pcActual As Long, ByRef salida() As Long) As Boolean
    Dim partes() As String
    Dim mnem As String
    Dim resto As String
    Dim ops() As String
    Dim op1 As String, op2 As String
    Dim posEspacio As Long
    
    linea = Trim$(linea)
    posEspacio = InStr(linea, " ")
    If posEspacio > 0 Then
        mnem = UCase$(Trim$(Left$(linea, posEspacio - 1)))
        resto = Trim$(Mid$(linea, posEspacio + 1))
    Else
        mnem = UCase$(linea)
        resto = ""
    End If
    
    If InStr(resto, ",") > 0 Then
        ops = Split(resto, ",")
        op1 = Trim$(ops(0))
        op2 = Trim$(ops(1))
    Else
        op1 = resto
        op2 = ""
    End If
    
    Select Case mnem
        Case "HLT"
            ReDim salida(0 To 0)
            salida(0) = &H0
            CodificarInstruccion = True
            
        Case "MOV"
            ReDim salida(0 To 2)
            salida(1) = IdRegPorNombre(op1)
            If EsNombreRegistro(op2) Then
                salida(0) = &H2 ' MOV r, r
                salida(2) = IdRegPorNombre(op2)
            Else
                salida(0) = &H1 ' MOV r, imm
                salida(2) = ParseValorODir(op2) And 255
            End If
            CodificarInstruccion = True
            
        Case "LOAD"
            ' LOAD r, [dir]
            ReDim salida(0 To 2)
            salida(0) = &H3
            Dim dirL As Long
            dirL = ParseValorODir(op2)
            ' Codificación de 9 bits: bit 8 se guarda en bit 2 del byte de registro
            salida(1) = (IdRegPorNombre(op1) And 3) Or ((dirL \ 256) * 4)
            salida(2) = dirL Mod 256
            CodificarInstruccion = True
            
        Case "STORE"
            ' STORE [dir], r
            ReDim salida(0 To 2)
            salida(0) = &H4
            Dim dirS As Long
            dirS = ParseValorODir(op1)
            salida(1) = (IdRegPorNombre(op2) And 3) Or ((dirS \ 256) * 4)
            salida(2) = dirS Mod 256
            CodificarInstruccion = True
            
        Case "ADD"
            ReDim salida(0 To 2)
            salida(1) = IdRegPorNombre(op1)
            If EsNombreRegistro(op2) Then
                salida(0) = &H11: salida(2) = IdRegPorNombre(op2)
            Else
                salida(0) = &H10: salida(2) = ParseValorODir(op2) And 255
            End If
            CodificarInstruccion = True
            
        Case "SUB"
            ReDim salida(0 To 2)
            salida(1) = IdRegPorNombre(op1)
            If EsNombreRegistro(op2) Then
                salida(0) = &H13: salida(2) = IdRegPorNombre(op2)
            Else
                salida(0) = &H12: salida(2) = ParseValorODir(op2) And 255
            End If
            CodificarInstruccion = True
            
        Case "INC"
            ReDim salida(0 To 1)
            salida(0) = &H14
            salida(1) = IdRegPorNombre(op1)
            CodificarInstruccion = True
            
        Case "DEC"
            ReDim salida(0 To 1)
            salida(0) = &H15
            salida(1) = IdRegPorNombre(op1)
            CodificarInstruccion = True
            
        Case "CMP"
            ReDim salida(0 To 2)
            salida(1) = IdRegPorNombre(op1)
            If EsNombreRegistro(op2) Then
                salida(0) = &H17: salida(2) = IdRegPorNombre(op2)
            Else
                salida(0) = &H16: salida(2) = ParseValorODir(op2) And 255
            End If
            CodificarInstruccion = True
            
        Case "AND"
            ReDim salida(0 To 2): salida(0) = &H18: salida(1) = IdRegPorNombre(op1): salida(2) = IdRegPorNombre(op2)
            CodificarInstruccion = True
        Case "OR"
            ReDim salida(0 To 2): salida(0) = &H19: salida(1) = IdRegPorNombre(op1): salida(2) = IdRegPorNombre(op2)
            CodificarInstruccion = True
        Case "XOR"
            ReDim salida(0 To 2): salida(0) = &H1A: salida(1) = IdRegPorNombre(op1): salida(2) = IdRegPorNombre(op2)
            CodificarInstruccion = True
        Case "NOT"
            ReDim salida(0 To 1): salida(0) = &H1B: salida(1) = IdRegPorNombre(op1)
            CodificarInstruccion = True
            
        Case "JMP", "JZ", "JNZ", "JC"
            ReDim salida(0 To 2)
            Select Case mnem
                Case "JMP": salida(0) = &H20
                Case "JZ":  salida(0) = &H21
                Case "JNZ": salida(0) = &H22
                Case "JC":  salida(0) = &H23
            End Select
            Dim dirJ As Long
            dirJ = ParseValorODir(op1)
            salida(1) = dirJ \ 256
            salida(2) = dirJ Mod 256
            CodificarInstruccion = True
            
        Case "IN"
            ' IN reg, [puerto]
            ReDim salida(0 To 2)
            salida(0) = &H30
            salida(1) = IdRegPorNombre(op1)
            salida(2) = ParseValorODir(op2) And 255
            CodificarInstruccion = True
            
        Case "OUT"
            ' OUT [puerto], reg
            ReDim salida(0 To 2)
            salida(0) = &H31
            salida(1) = IdRegPorNombre(op2)
            salida(2) = ParseValorODir(op1) And 255
            CodificarInstruccion = True
            
        Case Else
            UltimoErrorEnsamblador = "Mnemónico no reconocido: " & mnem
            CodificarInstruccion = False
    End Select
End Function

Private Function EsNombreRegistro(ByVal s As String) As Boolean
    s = UCase$(Trim$(s))
    EsNombreRegistro = (s = "AX" Or s = "BX" Or s = "CX" Or s = "DX")
End Function
