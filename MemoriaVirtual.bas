Option Explicit

' ==============================================================================
' MODULO: MemoriaVirtual.bas
' Simulador Avanzado de Arquitectura de Computadoras - Segundo Parcial (SIS-131)
' Modela:
'   1. Memoria Virtual (512 Bytes): 16 Páginas de 32 Bytes (VPN 0 a 15).
'   2. Memoria Principal Física RAM (256 Bytes): 8 Marcos de 32 Bytes (PFN 0 a 7).
'   3. Disco / Área de Swap (512 Bytes): Respaldo persistente de las 16 páginas.
'   4. Tabla de Páginas (Page Table): VPN, P (Presente), PFN, D (Dirty), R (Ref).
'   5. Unidad de Manejo de Memoria (MMU) y Atención de Fallo de Página (Page Fault).
'   6. Algoritmo de Reemplazo de Páginas (FIFO / LRU).
' ==============================================================================

Public Const TAM_PAGINA As Long = 32
Public Const NUM_PAGINAS_VIRTUALES As Long = 16 ' 16 * 32 = 512 Bytes
Public Const NUM_MARCOS_FISICOS As Long = 8    ' 8 * 32 = 256 Bytes

' Entrada de la Tabla de Páginas
Public Type EntradaTablaPaginas
    VPN As Long         ' 0 a 15
    Presente As Long    ' 1 = En RAM física, 0 = En Disco / Swap
    PFN As Long         ' Marco Físico asignado (0 a 7), o -1 si P=0
    Dirty As Long       ' 1 = Modificado en RAM
    Referenciado As Long' 1 = Accedido recientemente
    TiempoCarga As Long ' Para algoritmo FIFO / LRU
End Type

' Estructuras de Memoria
Public TablaPaginas(0 To NUM_PAGINAS_VIRTUALES - 1) As EntradaTablaPaginas
Public RAM_Fisica(0 To (NUM_MARCOS_FISICOS * TAM_PAGINA) - 1) As Long  ' 256 Bytes
Public Disco_Swap(0 To (NUM_PAGINAS_VIRTUALES * TAM_PAGINA) - 1) As Long ' 512 Bytes

' Asignación inversa de Marcos Físicos a Páginas Virtuales (-1 si el marco está libre)
Public MarcoAsignadoA_VPN(0 To NUM_MARCOS_FISICOS - 1) As Long

' Cola FIFO para reemplazo de marcos (0 a 7)
Public ColaFIFO_Marcos(0 To NUM_MARCOS_FISICOS - 1) As Long
Public TotalMarcosOcupados As Long
Public TickGlobalMemoria As Long

' Métricas de Paginación
Public TotalAccesosMemoria As Long
Public TotalPageFaults As Long
Public UltimoPageFaultOcurrido As Boolean
Public MensajePageFault As String
Public UltimaVPN_Accedida As Long
Public UltimoPFN_Asignado As Long

' Inicializa la Memoria Física, la Memoria Virtual y la Tabla de Páginas
Public Sub ResetMemoriaVirtual()
    Dim i As Long, j As Long
    
    TickGlobalMemoria = 0
    TotalAccesosMemoria = 0
    TotalPageFaults = 0
    UltimoPageFaultOcurrido = False
    MensajePageFault = "Sin fallos de página"
    UltimaVPN_Accedida = -1
    UltimoPFN_Asignado = -1
    TotalMarcosOcupados = 0
    
    ' Limpiar RAM Física (256 bytes)
    For i = 0 To 255
        RAM_Fisica(i) = 0
    Next i
    
    ' Limpiar Disco / Swap (512 bytes)
    For i = 0 To 511
        Disco_Swap(i) = 0
    Next i
    
    ' Reset de Marcos Físicos
    For i = 0 To NUM_MARCOS_FISICOS - 1
        MarcoAsignadoA_VPN(i) = -1
        ColaFIFO_Marcos(i) = -1
    Next i
    
    ' Inicializar Tabla de Páginas (todas en Swap al inicio)
    For i = 0 To NUM_PAGINAS_VIRTUALES - 1
        TablaPaginas(i).VPN = i
        TablaPaginas(i).Presente = 0
        TablaPaginas(i).PFN = -1
        TablaPaginas(i).Dirty = 0
        TablaPaginas(i).Referenciado = 0
        TablaPaginas(i).TiempoCarga = 0
    Next i
End Sub

' ------------------------------------------------------------------------------
' UNIDAD MMU: Traducción de Dirección Virtual a Física con Atención de Page Fault
' ------------------------------------------------------------------------------
Public Function MMU_Traducir(ByVal dirVirtual As Long, ByVal esEscritura As Boolean) As Long
    Dim vpn As Long
    Dim offset As Long
    Dim pfn As Long
    
    dirVirtual = dirVirtual And 511 ' Rango de 0 a 511
    vpn = dirVirtual \ TAM_PAGINA
    offset = dirVirtual Mod TAM_PAGINA
    
    TickGlobalMemoria = TickGlobalMemoria + 1
    TotalAccesosMemoria = TotalAccesosMemoria + 1
    UltimaVPN_Accedida = vpn
    
    ' Verificar si la página está presente en RAM física (Hit de Paginación)
    If TablaPaginas(vpn).Presente = 1 Then
        UltimoPageFaultOcurrido = False
        pfn = TablaPaginas(vpn).PFN
        TablaPaginas(vpn).Referenciado = 1
        If esEscritura Then
            TablaPaginas(vpn).Dirty = 1
        End If
        MMU_Traducir = (pfn * TAM_PAGINA) + offset
        Exit Function
    End If
    
    ' --------------------------------------------------------------------------
    ' EVENTO: PAGE FAULT (Página no presente en RAM)
    ' --------------------------------------------------------------------------
    TotalPageFaults = TotalPageFaults + 1
    UltimoPageFaultOcurrido = True
    
    pfn = AtenderPageFault(vpn)
    UltimoPFN_Asignado = pfn
    
    TablaPaginas(vpn).Referenciado = 1
    If esEscritura Then
        TablaPaginas(vpn).Dirty = 1
    End If
    
    MMU_Traducir = (pfn * TAM_PAGINA) + offset
End Function

' ------------------------------------------------------------------------------
' ATENCIÓN DE FALLO DE PÁGINA (PAGE FAULT HANDLER)
' ------------------------------------------------------------------------------
Private Function AtenderPageFault(ByVal vpnDemandada As Long) As Long
    Dim marcoLibre As Long
    Dim marcoVictima As Long
    Dim vpnVictima As Long
    Dim i As Long
    Dim dirBaseRAM As Long
    Dim dirBaseSwap As Long
    
    marcoLibre = -1
    
    ' 1. Buscar si hay algún marco físico libre en RAM (0 a 7)
    For i = 0 To NUM_MARCOS_FISICOS - 1
        If MarcoAsignadoA_VPN(i) = -1 Then
            marcoLibre = i
            Exit For
        End If
    Next i
    
    ' 2. Si no hay marcos libres, ejecutar ALGORITMO DE REEMPLAZO (FIFO)
    If marcoLibre = -1 Then
        ' El primer marco ingresado en la cola FIFO es la víctima
        marcoVictima = ColaFIFO_Marcos(0)
        vpnVictima = MarcoAsignadoA_VPN(marcoVictima)
        
        MensajePageFault = "PAGE FAULT: RAM llena. Desalojando VPN " & vpnVictima & _
                           " de Marco " & marcoVictima & " para cargar VPN " & vpnDemandada
        
        ' Si la página víctima fue modificada (Dirty = 1), guardar en Swap (Swap-out)
        If TablaPaginas(vpnVictima).Dirty = 1 Then
            dirBaseRAM = marcoVictima * TAM_PAGINA
            dirBaseSwap = vpnVictima * TAM_PAGINA
            For i = 0 To TAM_PAGINA - 1
                Disco_Swap(dirBaseSwap + i) = RAM_Fisica(dirBaseRAM + i)
            Next i
        End If
        
        ' Actualizar tabla de páginas de la víctima (ya no está presente)
        TablaPaginas(vpnVictima).Presente = 0
        TablaPaginas(vpnVictima).PFN = -1
        TablaPaginas(vpnVictima).Dirty = 0
        TablaPaginas(vpnVictima).Referenciado = 0
        
        ' Desplazar la cola FIFO hacia la izquierda
        For i = 0 To NUM_MARCOS_FISICOS - 2
            ColaFIFO_Marcos(i) = ColaFIFO_Marcos(i + 1)
        Next i
        
        marcoLibre = marcoVictima
    Else
        MensajePageFault = "PAGE FAULT: Asignando marco libre " & marcoLibre & " a VPN " & vpnDemandada
        TotalMarcosOcupados = TotalMarcosOcupados + 1
    End If
    
    ' 3. Transferir el bloque de 32 bytes desde Disco / Swap al Marco de RAM (Swap-in)
    dirBaseRAM = marcoLibre * TAM_PAGINA
    dirBaseSwap = vpnDemandada * TAM_PAGINA
    For i = 0 To TAM_PAGINA - 1
        RAM_Fisica(dirBaseRAM + i) = Disco_Swap(dirBaseSwap + i)
    Next i
    
    ' 4. Actualizar asignación del marco y tabla de páginas
    MarcoAsignadoA_VPN(marcoLibre) = vpnDemandada
    TablaPaginas(vpnDemandada).Presente = 1
    TablaPaginas(vpnDemandada).PFN = marcoLibre
    TablaPaginas(vpnDemandada).Dirty = 0
    TablaPaginas(vpnDemandada).Referenciado = 1
    TablaPaginas(vpnDemandada).TiempoCarga = TickGlobalMemoria
    
    ' 5. Insertar el nuevo marco al final de la cola FIFO
    ColaFIFO_Marcos(TotalMarcosOcupados - 1) = marcoLibre
    
    AtenderPageFault = marcoLibre
End Function

' ------------------------------------------------------------------------------
' ACCESOS DIRECTOS A RAM FÍSICA (usados por la Jerarquía de Caché)
' ------------------------------------------------------------------------------
Public Function RAM_Fisica_LeerByteDirecto(ByVal dirFisica As Long) As Long
    dirFisica = dirFisica And 255
    RAM_Fisica_LeerByteDirecto = RAM_Fisica(dirFisica)
End Function

Public Sub RAM_Fisica_EscribirByteDirecto(ByVal dirFisica As Long, ByVal valor As Long)
    dirFisica = dirFisica And 255
    RAM_Fisica(dirFisica) = valor And 255
    
    ' Marcar Dirty en la página correspondiente
    Dim vpn As Long, pfn As Long
    pfn = dirFisica \ TAM_PAGINA
    vpn = MarcoAsignadoA_VPN(pfn)
    If vpn >= 0 And vpn < NUM_PAGINAS_VIRTUALES Then
        TablaPaginas(vpn).Dirty = 1
    End If
End Sub

' ------------------------------------------------------------------------------
' ESCRITURA DIRECTA AL ESPACIO VIRTUAL (Usado por el Ensamblador / Cargador)
' ------------------------------------------------------------------------------
Public Sub MemoriaVirtual_CargarByte(ByVal dirVirtual As Long, ByVal valor As Long)
    dirVirtual = dirVirtual And 511
    ' Escribir directamente en el almacenamiento de respaldo de Disco/Swap
    Disco_Swap(dirVirtual) = valor And 255
    
    ' Si la página ya está presente en RAM, sincronizar el dato
    Dim vpn As Long, offset As Long, pfn As Long
    vpn = dirVirtual \ TAM_PAGINA
    offset = dirVirtual Mod TAM_PAGINA
    If TablaPaginas(vpn).Presente = 1 Then
        pfn = TablaPaginas(vpn).PFN
        RAM_Fisica((pfn * TAM_PAGINA) + offset) = valor And 255
    End If
End Sub

' Lectura directa del espacio virtual (para inspección o ensamblador)
Public Function MemoriaVirtual_LeerByte(ByVal dirVirtual As Long) As Long
    dirVirtual = dirVirtual And 511
    Dim vpn As Long, offset As Long, pfn As Long
    vpn = dirVirtual \ TAM_PAGINA
    offset = dirVirtual Mod TAM_PAGINA
    If TablaPaginas(vpn).Presente = 1 Then
        pfn = TablaPaginas(vpn).PFN
        MemoriaVirtual_LeerByte = RAM_Fisica((pfn * TAM_PAGINA) + offset)
    Else
        MemoriaVirtual_LeerByte = Disco_Swap(dirVirtual)
    End If
End Function
