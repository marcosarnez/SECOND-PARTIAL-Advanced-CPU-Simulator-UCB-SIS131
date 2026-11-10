# Simulador Avanzado de Arquitectura de Computadoras: Buses Tri-Estado, Caché Multinivel (L1/L2), Memoria Virtual y Periféricos de E/S

> **Materia:** Arquitectura de Computadoras (SIS-131) — Semestre 2/2026  
> **Institución:** Universidad Católica Boliviana "San Pablo" — Sede Santa Cruz  
> **Docente:** Ing. Paulo César Loayza Carrasco  
> **Proyecto:** Evaluación Segundo Parcial — Simulación Avanzada y Sustentación Oral  

---

## 1. Identidad del Equipo y Logotipo

```
  =============================================================
   ██████╗ ███████╗██████╗ ██╗██████╗ ██╗   ██╗███████╗
  ██╔════╝ ██╔════╝██╔══██╗██║██╔══██╗██║   ██║██╔════╝
  ██║  ███╗█████╗  ██████╔╝██║██████╔╝██║   ██║███████╗
  ██║   ██║██╔══╝  ██╔══██╗██║██╔══██╗██║   ██║╚════██║
  ╚██████╔╝███████╗██║  ██║██║██████╔╝╚██████╔╝███████║
   ╚═════╝ ╚══════╝╚═╝  ╚═╝╚═╝╚═════╝  ╚═════╝ ╚══════╝
           DIGITAL COMPUTING ARCHITECTURE LAB - UCB
  =============================================================
```

* **Equipo Oficial:** GERIBUS COMPUTING CORP.
* **Integrantes y Roles Gitflow:**
  * **Líder & Arquitecto de Buses / Micro-palabra:** `feature/estudiante-1/modulo-buses-tristate`
  * **Ingeniero de Jerarquía de Memoria & Caché:** `feature/estudiante-2/modulo-cache-l1-l2`
  * **Especialista en Memoria Virtual & E/S:** `feature/estudiante-3/modulo-mmu-perifericos`

---

## 2. Visión General de la Arquitectura

El simulador modela un sistema computacional Von Neumann mononúcleo de 8 bits con jerarquía multinivel, paginación por demanda y entrada/salida aislada por puertos.

```mermaid
flowchart TD
    subgraph CPU["CPU Mononúcleo (8 bits)"]
        subgraph Registros["Registros de Datos y Control"]
            AX["AX (Acumulador)"]
            BX["BX (Base)"]
            CX["CX (Contador)"]
            DX["DX (E/S / Datos)"]
            PC["PC (Puntero 9b)"]
            IR["IR (Instrucción)"]
            Flags["Flags (ZF, CF, SF)"]
        end
        subgraph ALU_Sub["Unidad Aritmético-Lógica"]
            TempA["Temp A"]
            TempB["Temp B"]
            ALU_CORE["ALU"]
            ALU_OUT["ALU_OUT"]
        end
        subgraph Aduana["Aduanas de Frontera"]
            MAR["MAR (Direcciones)"]
            MDR["MDR (Buffer Bidireccional)"]
        end
        UC["Unidad de Control (Micro-palabra)"]
    end

    subgraph Buses_Sistema["Infraestructura de Buses"]
        BusInterno["Bus Interno de Datos (High-Z)"]
        BusDirExt["Bus de Direcciones Externo"]
        BusDatosExt["Bus de Datos Externo (High-Z)"]
        BusCtrlExt["Bus de Control (MEM_R, MEM_W, IO_R, IO_W)"]
    end

    subgraph JerarquiaMemoria["Jerarquía de Memoria"]
        subgraph Cache_Multi["Jerarquía de Caché"]
            L1I["Caché L1I (8B - Instrucciones)"]
            L1D["Caché L1D (8B - Datos / Write-Back)"]
            L2["Caché L2 Unificada (32B - 4 líneas)"]
        end
        subgraph MMU_Paginacion["Memoria Virtual y MMU"]
            MMU["MMU (Traducción VPN -> PFN)"]
            PageTable["Tabla de Páginas (16 Entradas)"]
            RAM_Fisica["RAM Física (256B / 8 Marcos de 32B)"]
            SwapDisk["Disco / Área de Swap (512B / 16 Páginas)"]
        end
    end

    subgraph Perifericos_ES["Periféricos de Entrada/Salida"]
        Teclado["Teclado Físico (Puertos 0x01 Status / 0x02 Data)"]
        ColaRAM["Cola FIFO en RAM (0x70..0x77, Punteros 0x78..0x7A)"]
        LCD["Pantalla LCD 8x1 (Puerto 0x03 / Video Latch en RAM 0xF8..0xFF)"]
    end

    %% Conexiones
    Registros <--> BusInterno
    ALU_Sub <--> BusInterno
    MAR <-- BusInterno
    MDR <--> BusInterno
    MAR --> BusDirExt
    MDR <--> BusDatosExt
    UC --> BusCtrlExt

    BusDirExt --> MMU
    MMU --> PageTable
    MMU --> Cache_Multi
    Cache_Multi <--> RAM_Fisica
    RAM_Fisica <--> SwapDisk
    Cache_Multi <--> BusDatosExt

    BusDirExt --> Perifericos_ES
    BusDatosExt <--> Perifericos_ES
```

---

## 3. Buses del Sistema y Regla de Exclusión Mutua

### Estados del Buffer Tri-Estado
Cada registro y dispositivo emisor cuenta con una compuerta con tres estados posibles:
1. **Estado 0 (Lógico Bajo):** Conduce $0\text{V}$ al bus con fuerza ($Enable = 1, Input = 0$).
2. **Estado 1 (Lógico Alto):** Conduce $5\text{V}$ al bus con fuerza ($Enable = 1, Input = 1$).
3. **Estado Z (Alta Impedancia / High-Z):** Desconectado eléctricamente ($Enable = 0$). Resistencia infinita, no perturba el bus.

### Ecuación Fundamental de Exclusión Mutua Externa
Para prevenir cortocircuitos lógicos en el Bus de Datos Externo, el simulador valida en cada ciclo:

$$\text{MEM\_READ} + \text{MEM\_WRITE} + \text{IO\_READ} + \text{IO\_WRITE} \le 1$$

* **Suma = 1:** Transferencia externa válida hacia la memoria o un puerto de E/S.
* **Suma = 0:** Operación interna pura en la CPU (ALU o registros; buses externos en reposo absoluto).
* **Suma > 1:** $\rightarrow$ **¡ERROR DE COLISIÓN DE BUS!** El simulador suspende la ejecución arrojando una excepción de cortocircuito.

---

## 4. Estructura de la Micro-Palabra de Control (Control Word)

Emitida por la Unidad de Control en cada micro-paso del reloj ($T$-State):

```
+-------------------+--------------------+----------+-----------+---------+----------+
|  EMISOR_ID (3b)   |   DESTINO_ID (3b)  |  MEM_R   |   MEM_W   |  IO_R   |   IO_W   |
+-------------------+--------------------+----------+-----------+---------+----------+
```

### Tabla de Identificadores (IDs de Registro)
| ID Binario | Decimal | Registro Emisor / Destino |
|:---:|:---:|:---|
| `000` | 0 | NINGUNO (Bus Interno en reposo / High-Z) |
| `001` | 1 | `AX` (Registro Acumulador) |
| `010` | 2 | `BX` (Registro Base) |
| `011` | 3 | `CX` (Contador de Bucles) |
| `100` | 4 | `DX` (Datos y Periféricos) |
| `101` | 5 | `MAR` (Memory Address Register) |
| `110` | 6 | `MDR` (Memory Data Register) |
| `111` | 7 | `ALU_OUT` / `PC` / `Temp` |

---

## 5. El Ciclo de Instrucción Unificado (T-States: $T_0$ a $T_6$)

| Paso | Fase | Microoperación | Bus Interno | Bus Dir. Ext. | Bus Datos Ext. | Señales de Control | Registros Afectados |
|:---:|:---:|:---|:---:|:---:|:---:|:---:|:---|
| **$T_0$** | FETCH | $\text{PC} \to \text{MAR}$ | PC expone Dir | Inactivo | High-Z | `MAR_Load = 1` | MAR carga dirección |
| **$T_1$** | FETCH | Solicitar OpCode, $\text{PC} = \text{PC} + 1$ | Reposo | MAR expone Dir | Caché/RAM entrega OpCode | `MEM_READ = 1` | PC incrementa, MDR absorbe OpCode |
| **$T_2$** | FETCH | $\text{MDR} \to \text{IR}$ | MDR expone OpCode | Inactivo | High-Z | `IR_Load = 1` | IR captura el OpCode |
| **$T_3$** | DECODE | Decodificación y lectura de operandos | Reposo | Inactivo | High-Z | Control Interno | UC activa decodificador |
| **$T_4$** | EXECUTE | Enrutamiento de direcciones u operandos | Operando $\to$ MAR/TempA | Inactivo | High-Z | `MAR_Load = 1` | MAR o TempA cargado |
| **$T_5$** | EXECUTE | Lectura/Escritura en bus o carga TempB | Operando $\to$ TempB/MDR | MAR expone Dir | Transferencia activa | `MEM_READ` o `IO_READ` | MDR o TempB cargado |
| **$T_6$** | STORE | Grabación final de resultado | MDR/ALU $\to$ Reg/RAM | MAR expone Dir | Dato expuesto | `MEM_WRITE` o `IO_WRITE` | Registro destino actualizado, $T \to 0$ |

---

## 6. Jerarquía de Memoria y Paginación

### Memoria Virtual y RAM Física
* **Memoria Virtual:** 512 Bytes divididos estrictamente en **16 Páginas Virtuales** (VPN 0 a 15) de 32 Bytes cada una.
* **Memoria Física (RAM):** 256 Bytes divididos en **8 Marcos Físicos** (PFN 0 a 7) de 32 Bytes cada uno.
* **Área de Swap (Disco):** 512 Bytes en archivo o array emulado para almacenar páginas desalojadas.

### Tabla de Páginas (Page Table - 16 Entradas)
| VPN (0..15) | Bit Presente ($P$) | Marco Físico ($PFN$) | Dirty Bit ($D$) | Bit Referencia ($R$) | Estado |
|:---:|:---:|:---:|:---:|:---:|:---|
| 0 | 1 | Marco 0 | 0 | 1 | Presente en RAM |
| 1 | 1 | Marco 1 | 1 | 1 | Modificada en RAM |
| ... | ... | ... | ... | ... | ... |
| 15 | 0 | - | 0 | 0 | En Disco (Swap) |

### Algoritmo de Atención de Page Fault
1. La MMU detecta acceso a una dirección virtual con $P = 0$.
2. Se detiene el avance de la CPU y se ilumina el indicador visual de **¡PAGE FAULT!** en rojo.
3. Se verifica si hay un marco físico libre en RAM.
4. **Si la RAM está llena (8 marcos ocupados):**
   * Se ejecuta el algoritmo **FIFO** seleccionando el marco más antiguo.
   * Si la página víctima tiene $D = 1$ (Dirty), se transfieren los 32 bytes al Área de Swap en Disco (**Swap-out**).
   * La entrada de la página víctima en la tabla se actualiza a $P = 0, PFN = -1$.
5. Se lee el bloque de 32 bytes de la página solicitada desde el Área de Swap (**Swap-in**) y se graba en el marco libre.
6. Se actualiza la tabla de páginas ($P = 1, PFN = \text{nuevo}, D = 0, R = 1$) y se reanuda la instrucción.

### Jerarquía de Caché Multinivel
* **L1I (Instrucciones):** 8 Bytes (1 línea). Caching exclusivo del flujo de ejecución del `PC`.
* **L1D (Datos):** 8 Bytes (1 línea). Caching de operandos leídos y escritos por `LOAD`/`STORE`. Política **Write-Back** con Dirty bit.
* **L2 (Unificada):** 32 Bytes organizados en 4 líneas de 8 Bytes con política de reemplazo FIFO.
* **Métricas en Pantalla:** Indicadores en tiempo real de Hit/Miss, contadores y tasa de aciertos acumulada (*Hit Rate %*).

---

## 7. Periféricos de Entrada/Salida: Teclado y Pantalla LCD 8x1

### Mapa de Memoria y Puertos
* **`Puerto 0x01` (KBD Status):** Bit 0 = 1 si hay tecla lista.
* **`Puerto 0x02` (KBD Data):** Código ASCII de la tecla pulsada. Al leerse, resetea el Puerto 0x01 a 0 automáticamente.
* **`Puerto 0x03` (LCD Display):** Escritura de caracteres a la pantalla LCD con retención en Cerrojo (*Video Latch*).
* **Cola Circular FIFO en RAM (`0x70`..`0x77`):** Capacidad de 8 bytes.
  * `0x78 = KB_HEAD` (Puntero de lectura).
  * `0x79 = KB_TAIL` (Puntero de escritura).
  * `0x7A = KB_COUNT` (Número de caracteres, protección contra desbordamiento).
* **VRAM en RAM (`0xF8`..`0xFF`):** Buffer de video de 8 bytes mapeado directamente a la pantalla con soporte de **Backspace (ASCII 0x08)**.

---

## 8. Conjunto de Instrucciones (ISA Extendida)

| Mnemónico | Opcode | Bytes | Operandos | Descripción | Banderas |
|:---|:---:|:---:|:---|:---|:---:|
| `HLT` | `00` | 1 | — | Detiene el reloj del procesador | — |
| `MOV r, imm` | `01` | 3 | reg, imm8 | Carga inmediato en registro (AX..DX) | — |
| `MOV r, r` | `02` | 3 | reg, reg | Copia un registro en otro | — |
| `LOAD r, [dir]`| `03` | 3 | reg, dir9 | Lee Memoria Virtual hacia registro pasando por Caché L1D/L2 | — |
| `STORE [dir], r`|`04` | 3 | dir9, reg | Escribe registro en Memoria Virtual vía L1D Write-Back | — |
| `ADD r, imm` | `10` | 3 | reg, imm8 | Suma inmediato al registro | ZF, CF, SF |
| `ADD r, r` | `11` | 3 | reg, reg | Suma dos registros | ZF, CF, SF |
| `SUB r, imm` | `12` | 3 | reg, imm8 | Resta inmediato del registro | ZF, CF, SF |
| `SUB r, r` | `13` | 3 | reg, reg | Resta dos registros | ZF, CF, SF |
| `INC r` | `14` | 2 | reg | Incrementa registro en 1 | ZF, CF, SF |
| `DEC r` | `15` | 2 | reg | Decrementa registro en 1 | ZF, CF, SF |
| `CMP r, imm` | `16` | 3 | reg, imm8 | Resta sin modificar registro | ZF, CF, SF |
| `CMP r, r` | `17` | 3 | reg, reg | Resta sin modificar registro | ZF, CF, SF |
| `AND r, r` | `18` | 3 | reg, reg | Operación AND bit a bit | ZF, SF (CF=0) |
| `OR r, r` | `19` | 3 | reg, reg | Operación OR bit a bit | ZF, SF (CF=0) |
| `XOR r, r` | `1A` | 3 | reg, reg | Operación XOR bit a bit | ZF, SF (CF=0) |
| `NOT r` | `1B` | 2 | reg | Operación NOT bit a bit | ZF, SF (CF=0) |
| `JMP dir` | `20` | 3 | dir9 | Salto incondicional al espacio virtual | — |
| `JZ dir` | `21` | 3 | dir9 | Salto condicional si ZF = 1 | — |
| `JNZ dir` | `22` | 3 | dir9 | Salto condicional si ZF = 0 | — |
| `JC dir` | `23` | 3 | dir9 | Salto condicional si CF = 1 | — |
| `IN r, [puerto]`|`30` | 3 | reg, puerto8 | Entrada desde periférico (E/S Aislada) | — |
| `OUT [puerto], r`|`31`| 3 | puerto8, reg | Salida hacia periférico con Latch | — |

---

## 9. Catálogo de los 3 Programas Oficiales

### Programa 1: Aritmética y Bucles con AX, BX, CX, DX
Calcula una serie aritmética acumulativa utilizando los 4 registros generales:
* `AX`: Acumulador de sumas.
* `BX`: Factor de incremento.
* `CX`: Contador de iteraciones del bucle (`DEC CX`, `JNZ`).
* `DX`: Contador de micro-operaciones y registro de control.

### Programa 2: E/S Interactiva Teclado -> Pantalla LCD 8x1
Implementa el ciclo de terminal interactiva:
1. Sondea el `Puerto 0x01` hasta detectar pulsación (`Bit 0 == 1`).
2. Lee el código ASCII del `Puerto 0x02`.
3. Emite la instrucción `OUT [0x03], DX` para proyectar el carácter en el display LCD.
4. Soporta borrado dinámico mediante tecla Backspace (`0x08`).

### Programa 3: Accesos Dispersos, Page Faults y Fallos de Caché
Genera accesos alternados a través de más de 8 páginas virtuales diferentes (`0x010`, `0x040`, `0x080`, `0x0C0`, `0x100`, `0x140`, `0x180`, `0x1C0`, `0x1E0`):
* Desborda los 8 marcos físicos de la RAM.
* Demuestra visualmente el desalojo FIFO de marcos hacia el Swap.
* Dispara eventos de **¡PAGE FAULT!** resaltados en rojo.
* Demuestra penalizaciones de latencia y fallos consecutivos de caché L1 y L2.

---

## 10. Manual de Uso del Simulador

1. Abrir `simulador-cpu.xlsm` y asegurarse de habilitar las macros.
2. En la pestaña **CPU**, observar los módulos de la placa:
   * **Buses y Micro-palabra:** Columna central con valores y estado High-Z.
   * **Jerarquía de Caché:** Indicadores de L1I, L1D, L2 y Hit Rate.
   * **Memoria Virtual:** Tabla de 16 páginas y semáforo de Page Fault.
   * **Periféricos:** Display LCD 8x1 y campo de ingreso de teclas.
3. Cargar cualquiera de los 3 programas pulsando los botones:
   * **LOAD PROG 1 (Bucles AX..DX)**
   * **LOAD PROG 2 (Teclado -> LCD)**
   * **LOAD PROG 3 (Page Faults & Caché)**
4. Presionar **STEP** para avanzar micro-paso a micro-paso ($T_0 \to T_1 \to \dots \to T_6$).
5. Presionar **RUN** para ejecución continua. El control deslizante o celda **Delay (ms)** permite ajustar la velocidad.
6. Presionar **PAUSE** para detener el reloj temporalmente o **RESET** para limpiar todo el hardware.
