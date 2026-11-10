Attribute VB_Name = "MemoriaRAM"
Option Explicit

' ==============================================================================
' MODULO: MemoriaRAM.bas
' Simulador Avanzado de Arquitectura de Computadoras - Segundo Parcial (SIS-131)
' Puente de compatibilidad con las hojas de cálculo de memoria:
'   - Sincroniza la hoja "RAM" (256 bytes) con la memoria física del simulador.
'   - Proporciona rutinas para botones interactivos de lectura/escritura manual.
' ==============================================================================

Private Function HojaRAM() As Worksheet
    On Error Resume Next
    Set HojaRAM = ThisWorkbook.Worksheets("RAM")
    On Error GoTo 0
End Function

Private Function HojaMem() As Worksheet
    On Error Resume Next
    Set HojaMem = ThisWorkbook.Worksheets("Memoria")
    On Error GoTo 0
End Function

' Lee directamente de la memoria física y sincroniza MAR/MDR
Public Function ReadRAM(ByVal address As Long) As Long
    address = address And 255
    Dim v As Long
    v = RAM_Fisica_LeerByteDirecto(address)
    
    If Not HojaMem Is Nothing Then
        HojaMem.Range("F4").Value = address
        HojaMem.Range("F5").Value = v
    End If
    
    RegMAR = address
    RegMDR = v
    ReadRAM = v
End Function

' Escribe directamente en la memoria física y actualiza la celda en la hoja
Public Sub WriteRAM(ByVal address As Long, ByVal value As Long)
    address = address And 255
    value = value And 255
    
    RAM_Fisica_EscribirByteDirecto address, value
    
    If Not HojaRAM Is Nothing Then
        HojaRAM.Cells(address + 1, 1).Value = value
    End If
    
    If Not HojaMem Is Nothing Then
        HojaMem.Range("F4").Value = address
        HojaMem.Range("F5").Value = value
    End If
    
    RegMAR = address
    RegMDR = value
End Sub

' Botón Leer de la interfaz manual de Excel
Public Sub BotonLeer()
    If HojaMem Is Nothing Then Exit Sub
    Dim addr As Long
    addr = NumeroEn(HojaMem.Range("B4"), 0, 511, "Casilla")
    If addr < 0 Then Exit Sub
    
    Dim v As Long
    v = MemoriaVirtual_LeerByte(addr)
    HojaMem.Range("B5").Value = v
    HojaMem.Range("A7").Value = "Leída dirección virtual 0x" & Hex$(addr) & " (" & addr & "). Valor: " & v
End Sub

' Botón Escribir de la interfaz manual de Excel
Public Sub BotonEscribir()
    If HojaMem Is Nothing Then Exit Sub
    Dim addr As Long, v As Long
    addr = NumeroEn(HojaMem.Range("B4"), 0, 511, "Casilla")
    If addr < 0 Then Exit Sub
    v = NumeroEn(HojaMem.Range("B5"), 0, 255, "Valor")
    If v < 0 Then Exit Sub
    
    MemoriaVirtual_CargarByte addr, v
    HojaMem.Range("A7").Value = "Guardado valor " & v & " en dirección virtual 0x" & Hex$(addr)
End Sub

' Botón Limpiar Memoria
Public Sub BotonLimpiar()
    ResetMemoriaVirtual
    ResetCache
    
    If Not HojaRAM Is Nothing Then
        Dim i As Long
        For i = 1 To 256
            HojaRAM.Cells(i, 1).Value = 0
        Next i
    End If
    
    If Not HojaMem Is Nothing Then
        HojaMem.Range("B4").Value = 0
        HojaMem.Range("B5").Value = 0
        HojaMem.Range("F4").Value = 0
        HojaMem.Range("F5").Value = 0
        HojaMem.Range("A7").Value = "Memoria Virtual y Física restablecidas a cero."
    End If
    
    ActualizarTodoUI
End Sub

Private Function NumeroEn(ByVal celda As Range, ByVal minimo As Long, ByVal maximo As Long, ByVal nombre As String) As Long
    If Len(Trim$(CStr(celda.Value))) = 0 Or Not IsNumeric(celda.Value) Then
        MsgBox nombre & " debe ser un número entero.", vbExclamation
        NumeroEn = -1
        Exit Function
    End If
    Dim valNum As Long
    valNum = CLng(celda.Value)
    If valNum < minimo Or valNum > maximo Then
        MsgBox nombre & " debe estar entre " & minimo & " y " & maximo & ".", vbExclamation
        NumeroEn = -1
        Exit Function
    End If
    NumeroEn = valNum
End Function
