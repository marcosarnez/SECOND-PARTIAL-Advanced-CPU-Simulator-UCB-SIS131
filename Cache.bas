Option Explicit

' ==============================================================================
' MODULO: Cache.bas
' Simulador Avanzado de Arquitectura de Computadoras - Segundo Parcial (SIS-131)
' Jerarquía de Memoria Caché Multinivel:
'   1. Caché L1I (Instrucciones): 8 Bytes (1 línea de 8B, flujo de PC).
'   2. Caché L1D (Datos): 8 Bytes (1 línea de 8B, lectura/escritura de operandos).
'   3. Caché L2 (Unificada): 32 Bytes (4 líneas de 8B o 1 bloque de 32B).
' Métricas:
'   - Contadores acumulativos de Hit y Miss (L1I, L1D, L2).
'   - Cálculo dinámico de Hit Rate (%).
'   - Notificaciones y estados de animación para la interfaz de usuario.
' Política:
'   - Write-Back con Dirty Bit (escritura en RAM solo al desalojar).
' ==============================================================================

' Estructura de una línea de caché de 8 bytes
Public Type CacheLine8
    Valida As Boolean
    Dirty As Boolean
    Tag As Long          ' Dirección base del bloque (múltiplo de 8)
    Datos(0 To 7) As Long
End Type

' Instancias de la Jerarquía
Public L1I As CacheLine8            ' 1 línea de 8 bytes para instrucciones
Public L1D As CacheLine8            ' 1 línea de 8 bytes para datos
Public L2Lines(0 To 3) As CacheLine8 ' 4 líneas de 8 bytes = 32 bytes unificados
Public L2_FIFO_Index As Long        ' Puntero circular para reemplazo en L2

' Contadores de Rendimiento
Public Hits_L1I As Long
Public Misses_L1I As Long

Public Hits_L1D As Long
Public Misses_L1D As Long

Public Hits_L2 As Long
Public Misses_L2 As Long

' Mensajes de estado del último acceso para visualización en Excel
Public UltimoEstadoL1I As String
Public UltimoEstadoL1D As String
Public UltimoEstadoL2 As String
Public UltimoEventoCache As String

' Resetea completamente toda la jerarquía de caché y métricas
Public Sub ResetCache()
    Dim i As Long, j As Long
    
    ' Reset L1I
    L1I.Valida = False
    L1I.Dirty = False
    L1I.Tag = -1
    For j = 0 To 7: L1I.Datos(j) = 0: Next j
    
    ' Reset L1D
    L1D.Valida = False
    L1D.Dirty = False
    L1D.Tag = -1
    For j = 0 To 7: L1D.Datos(j) = 0: Next j
    
    ' Reset L2
    For i = 0 To 3
        L2Lines(i).Valida = False
        L2Lines(i).Dirty = False
        L2Lines(i).Tag = -1
        For j = 0 To 7: L2Lines(i).Datos(j) = 0: Next j
    Next i
    L2_FIFO_Index = 0
    
    ' Reset Métricas
    Hits_L1I = 0: Misses_L1I = 0
    Hits_L1D = 0: Misses_L1D = 0
    Hits_L2 = 0: Misses_L2 = 0
    
    UltimoEstadoL1I = "INACTIVO"
    UltimoEstadoL1D = "INACTIVO"
    UltimoEstadoL2 = "INACTIVO"
    UltimoEventoCache = "Caché inicializada y vacía"
End Sub

' Calcula la tasa de aciertos (Hit Rate) en porcentaje
Public Function HitRate_L1I() As Double
    Dim total As Long
    total = Hits_L1I + Misses_L1I
    If total > 0 Then HitRate_L1I = Round((Hits_L1I / CDbl(total)) * 100, 1) Else HitRate_L1I = 0
End Function

Public Function HitRate_L1D() As Double
    Dim total As Long
    total = Hits_L1D + Misses_L1D
    If total > 0 Then HitRate_L1D = Round((Hits_L1D / CDbl(total)) * 100, 1) Else HitRate_L1D = 0
End Function

Public Function HitRate_L2() As Double
    Dim total As Long
    total = Hits_L2 + Misses_L2
    If total > 0 Then HitRate_L2 = Round((Hits_L2 / CDbl(total)) * 100, 1) Else HitRate_L2 = 0
End Function

Public Function HitRate_Global() As Double
    Dim hits As Long, total As Long
    hits = Hits_L1I + Hits_L1D + Hits_L2
    total = hits + Misses_L2 ' Todo fallo final de L2 va a RAM
    If total > 0 Then HitRate_Global = Round((hits / CDbl(total)) * 100, 1) Else HitRate_Global = 0
End Function

' ------------------------------------------------------------------------------
' LECTURA DE INSTRUCCIÓN (Flujo del PC -> L1I -> L2 -> RAM)
' ------------------------------------------------------------------------------
Public Function Cache_LeerInstruccion(ByVal dirFisica As Long) As Long
    Dim base8 As Long
    Dim offset As Long
    Dim idxL2 As Long
    Dim j As Long
    Dim valor As Long
    
    base8 = (dirFisica \ 8) * 8
    offset = dirFisica Mod 8
    
    ' 1. Consultar L1I
    If L1I.Valida And (L1I.Tag = base8) Then
        Hits_L1I = Hits_L1I + 1
        UltimoEstadoL1I = "HIT (0x" & Hex$(dirFisica) & ")"
        UltimoEventoCache = "L1I HIT en 0x" & Hex$(dirFisica)
        Cache_LeerInstruccion = L1I.Datos(offset)
        Exit Function
    End If
    
    ' Miss en L1I
    Misses_L1I = Misses_L1I + 1
    UltimoEstadoL1I = "MISS"
    
    ' 2. Consultar L2
    idxL2 = BuscarLineaEnL2(base8)
    If idxL2 <> -1 Then
        ' Hit en L2
        Hits_L2 = Hits_L2 + 1
        UltimoEstadoL2 = "HIT (Línea " & idxL2 & ")"
        UltimoEventoCache = "L1I MISS -> L2 HIT -> Carga L1I"
        
        ' Cargar línea de L2 a L1I
        L1I.Valida = True
        L1I.Dirty = False
        L1I.Tag = base8
        For j = 0 To 7
            L1I.Datos(j) = L2Lines(idxL2).Datos(j)
        Next j
        
        Cache_LeerInstruccion = L1I.Datos(offset)
        Exit Function
    End If
    
    ' Miss en L2 -> Traer bloque desde RAM física
    Misses_L2 = Misses_L2 + 1
    UltimoEstadoL2 = "MISS -> Lectura de RAM"
    UltimoEventoCache = "MISS L1I + MISS L2 -> Acceso a RAM -> Recarga L2 y L1I"
    
    ' Traer 8 bytes de RAM física e insertar en L2
    InsertarEnL2 base8
    
    ' Cargar también a L1I
    L1I.Valida = True
    L1I.Dirty = False
    L1I.Tag = base8
    For j = 0 To 7
        L1I.Datos(j) = RAM_Fisica_LeerByteDirecto(base8 + j)
    Next j
    
    Cache_LeerInstruccion = L1I.Datos(offset)
End Function

' ------------------------------------------------------------------------------
' LECTURA DE DATOS (Operandos / Variables -> L1D -> L2 -> RAM)
' ------------------------------------------------------------------------------
Public Function Cache_LeerDato(ByVal dirFisica As Long) As Long
    Dim base8 As Long
    Dim offset As Long
    Dim idxL2 As Long
    Dim j As Long
    
    base8 = (dirFisica \ 8) * 8
    offset = dirFisica Mod 8
    
    ' 1. Consultar L1D
    If L1D.Valida And (L1D.Tag = base8) Then
        Hits_L1D = Hits_L1D + 1
        UltimoEstadoL1D = "HIT (0x" & Hex$(dirFisica) & ")"
        UltimoEventoCache = "L1D HIT en 0x" & Hex$(dirFisica)
        Cache_LeerDato = L1D.Datos(offset)
        Exit Function
    End If
    
    ' Miss en L1D
    Misses_L1D = Misses_L1D + 1
    UltimoEstadoL1D = "MISS"
    
    ' Si la línea saliente de L1D está sucia (Dirty), desalojarla a L2 antes de sobreescribir
    DesalojarL1DSiSucia
    
    ' 2. Consultar L2
    idxL2 = BuscarLineaEnL2(base8)
    If idxL2 <> -1 Then
        Hits_L2 = Hits_L2 + 1
        UltimoEstadoL2 = "HIT (Línea " & idxL2 & ")"
        UltimoEventoCache = "L1D MISS -> L2 HIT -> Carga L1D"
        
        L1D.Valida = True
        L1D.Dirty = False
        L1D.Tag = base8
        For j = 0 To 7
            L1D.Datos(j) = L2Lines(idxL2).Datos(j)
        Next j
        
        Cache_LeerDato = L1D.Datos(offset)
        Exit Function
    End If
    
    ' Miss en L2 -> Traer desde RAM
    Misses_L2 = Misses_L2 + 1
    UltimoEstadoL2 = "MISS -> Lectura de RAM"
    UltimoEventoCache = "MISS L1D + MISS L2 -> Acceso a RAM -> Recarga L2 y L1D"
    
    InsertarEnL2 base8
    
    L1D.Valida = True
    L1D.Dirty = False
    L1D.Tag = base8
    For j = 0 To 7
        L1D.Datos(j) = RAM_Fisica_LeerByteDirecto(base8 + j)
    Next j
    
    Cache_LeerDato = L1D.Datos(offset)
End Function

' ------------------------------------------------------------------------------
' ESCRITURA DE DATOS (Write-Back en L1D con Dirty Bit)
' ------------------------------------------------------------------------------
Public Sub Cache_EscribirDato(ByVal dirFisica As Long, ByVal valor As Long)
    Dim base8 As Long
    Dim offset As Long
    Dim idxL2 As Long
    Dim j As Long
    
    base8 = (dirFisica \ 8) * 8
    offset = dirFisica Mod 8
    valor = valor And 255
    
    ' Si ya está presente en L1D (Hit de escritura)
    If L1D.Valida And (L1D.Tag = base8) Then
        Hits_L1D = Hits_L1D + 1
        L1D.Datos(offset) = valor
        L1D.Dirty = True
        UltimoEstadoL1D = "WRITE HIT (Dirty=1)"
        UltimoEventoCache = "L1D WRITE HIT en 0x" & Hex$(dirFisica)
        Exit Sub
    End If
    
    ' Write Miss en L1D: desalojar línea anterior si estaba sucia
    Misses_L1D = Misses_L1D + 1
    DesalojarL1DSiSucia
    
    ' Buscar en L2 o traer de RAM para tener el bloque completo
    idxL2 = BuscarLineaEnL2(base8)
    If idxL2 = -1 Then
        Misses_L2 = Misses_L2 + 1
        InsertarEnL2 base8
    Else
        Hits_L2 = Hits_L2 + 1
    End If
    
    ' Cargar a L1D, modificar el byte y marcar Dirty=1
    L1D.Valida = True
    L1D.Tag = base8
    For j = 0 To 7
        L1D.Datos(j) = RAM_Fisica_LeerByteDirecto(base8 + j)
    Next j
    L1D.Datos(offset) = valor
    L1D.Dirty = True
    
    UltimoEstadoL1D = "WRITE ALLOCATE (Dirty=1)"
    UltimoEventoCache = "L1D WRITE MISS -> Asignado y modificado con Dirty=1"
End Sub

' ------------------------------------------------------------------------------
' RUTINAS AUXILIARES DE L2 Y DESALOJO WRITE-BACK
' ------------------------------------------------------------------------------

Private Function BuscarLineaEnL2(ByVal base8 As Long) As Long
    Dim i As Long
    For i = 0 To 3
        If L2Lines(i).Valida And (L2Lines(i).Tag = base8) Then
            BuscarLineaEnL2 = i
            Exit Function
        End If
    Next i
    BuscarLineaEnL2 = -1
End Function

Private Sub DesalojarL1DSiSucia()
    Dim idxL2 As Long, j As Long
    If L1D.Valida And L1D.Dirty Then
        ' Escribir línea sucia en L2
        idxL2 = BuscarLineaEnL2(L1D.Tag)
        If idxL2 = -1 Then
            idxL2 = AsignarSlotL2(L1D.Tag)
        End If
        L2Lines(idxL2).Valida = True
        L2Lines(idxL2).Tag = L1D.Tag
        L2Lines(idxL2).Dirty = True
        For j = 0 To 7
            L2Lines(idxL2).Datos(j) = L1D.Datos(j)
        Next j
        L1D.Dirty = False
    End If
End Sub

Private Sub InsertarEnL2(ByVal base8 As Long)
    Dim slot As Long, j As Long
    slot = AsignarSlotL2(base8)
    L2Lines(slot).Valida = True
    L2Lines(slot).Tag = base8
    L2Lines(slot).Dirty = False
    For j = 0 To 7
        L2Lines(slot).Datos(j) = RAM_Fisica_LeerByteDirecto(base8 + j)
    Next j
End Sub

Private Function AsignarSlotL2(ByVal nuevoTag As Long) As Long
    Dim i As Long
    ' 1. Buscar un slot libre
    For i = 0 To 3
        If Not L2Lines(i).Valida Then
            AsignarSlotL2 = i
            Exit Function
        End If
    Next i
    
    ' 2. Si están todos llenos, política FIFO de reemplazo
    Dim victima As Long
    victima = L2_FIFO_Index
    L2_FIFO_Index = (L2_FIFO_Index + 1) Mod 4
    
    ' Si la línea víctima está sucia, escribir a RAM física (Write-Back)
    If L2Lines(victima).Dirty Then
        Dim j As Long
        For j = 0 To 7
            RAM_Fisica_EscribirByteDirecto L2Lines(victima).Tag + j, L2Lines(victima).Datos(j)
        Next j
    End If
    
    AsignarSlotL2 = victima
End Function

' Fuerza el vaciado de todas las líneas sucias de L1D y L2 hacia la RAM física (Flush)
Public Sub Cache_Flush()
    DesalojarL1DSiSucia
    Dim i As Long, j As Long
    For i = 0 To 3
        If L2Lines(i).Valida And L2Lines(i).Dirty Then
            For j = 0 To 7
                RAM_Fisica_EscribirByteDirecto L2Lines(i).Tag + j, L2Lines(i).Datos(j)
            Next j
            L2Lines(i).Dirty = False
        End If
    Next i
    UltimoEventoCache = "Caché vaciada (Flush completado hacia RAM)"
End Sub
