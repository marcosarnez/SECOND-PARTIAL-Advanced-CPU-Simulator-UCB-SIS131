# Advanced Computer Architecture Simulator: Tri-State Buses, Multi-level Cache (L1/L2), Paged Virtual Memory, and I/O Peripherals

> **Course:** Computer Architecture (SIS-131) — Semester 2/2026  
> **Institution:** Universidad Católica Boliviana "San Pablo" — Santa Cruz Campus  
> **Professor:** Eng. Paulo César Loayza Carrasco  
> **Project:** Second Midterm Evaluation — Advanced Engineering Simulation and Oral Defense  

---

## 1. Team Identity and Official Logo

<div align="center">
  <img src="assets/logo_cochi_al_horno.jpg" alt="COCHI AL HORNO - Hardware & Design - Computer Architecture Group" width="340"/>
  <h3>COCHI AL HORNO — Computer Architecture Group</h3>
  <p><em>Hardware & Design Lab • Universidad Católica Boliviana "San Pablo"</em></p>
</div>

* **Official Team Name:** COCHI AL HORNO (Computer Architecture Group)
* **Members & Gitflow Responsibilities:**
  * **Marcos Arnez (`@marcosarnez`):** Tri-State Buses, CPU Core (AX..DX & Control Word), UI Cockpit & Two-Pass Assembler (`feature/marcos/...`)
  * **Fernando Arias (`@fernandoqarius`):** Multi-level Cache (L1I/L1D/L2), Paged Virtual Memory & Swap, I/O Peripherals (`feature/companero/...`)

---

## 2. Architectural Overview

The simulator models an advanced 8-bit single-core, single-threaded Von Neumann computing system featuring multi-level caching, demand-paged virtual memory, and isolated port-mapped I/O.

```mermaid
flowchart TD
    subgraph CPU["Single-Core CPU (8-bit)"]
        subgraph Registers["Data & Control Registers"]
            AX["AX (Accumulator)"]
            BX["BX (Base)"]
            CX["CX (Counter)"]
            DX["DX (I/O / Data)"]
            PC["PC (9-bit Pointer)"]
            IR["IR (Instruction)"]
            Flags["Flags (ZF, CF, SF)"]
        end
        subgraph ALU_Sub["Arithmetic Logic Unit (ALU)"]
            TempA["Temp A"]
            TempB["Temp B"]
            ALU_CORE["ALU"]
            ALU_OUT["ALU_OUT"]
        end
        subgraph Gateway["Border Gateway Registers"]
            MAR["MAR (Address Register)"]
            MDR["MDR (Bidirectional Data Buffer)"]
        end
        UC["Control Unit (Micro-word Generator)"]
    end

    subgraph SystemBuses["System Bus Infrastructure"]
        BusInterno["Internal Data Bus (High-Z)"]
        BusDirExt["External Address Bus"]
        BusDatosExt["External Data Bus (High-Z)"]
        BusCtrlExt["Control Bus (MEM_R, MEM_W, IO_R, IO_W)"]
    end

    subgraph MemoryHierarchy["Memory Hierarchy"]
        subgraph Cache_Multi["Multi-level Cache"]
            L1I["L1I Cache (8B - Instructions)"]
            L1D["L1D Cache (8B - Data / Write-Back)"]
            L2["L2 Unified Cache (32B - 4 Lines)"]
        end
        subgraph MMU_Paging["Virtual Memory & MMU"]
            MMU["MMU (VPN -> PFN Translation)"]
            PageTable["Page Table (16 Entries)"]
            RAM_Fisica["Physical RAM (256B / 8 Frames of 32B)"]
            SwapDisk["Swap Space / Virtual Disk (512B / 16 Pages)"]
        end
    end

    subgraph IO_Peripherals["Input / Output Peripherals"]
        Keyboard["Physical Keyboard (Ports 0x01 Status / 0x02 Data)"]
        RAM_Queue["Circular FIFO Queue in RAM (0x70..0x77, Ptrs 0x78..0x7A)"]
        LCD["8x1 LCD Screen (Port 0x03 / Video Latch in RAM 0xF8..0xFF)"]
    end

    %% Wiring
    Registers <--> BusInterno
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

    BusDirExt --> IO_Peripherals
    BusDatosExt <--> IO_Peripherals
```

---

## 3. System Buses & Mutual Exclusion Law

### Tri-State Buffer States
Every driving register or device features a 3-state output buffer:
1. **State 0 (Logic Low):** Actively drives $0\text{V}$ to the bus line ($Enable = 1, Input = 0$).
2. **State 1 (Logic High):** Actively drives $5\text{V}$ to the bus line ($Enable = 1, Input = 1$).
3. **State Z (High Impedance / High-Z):** Electrically disconnected ($Enable = 0$). Infinite resistance, leaving the bus free for another talker.

### External Mutual Exclusion Equation
To prevent logical short circuits on the External Data Bus, the simulator validates every clock state:

$$\text{MEM\_READ} + \text{MEM\_WRITE} + \text{IO\_READ} + \text{IO\_WRITE} \le 1$$

* **Sum = 1:** Valid external transfer to memory or an I/O device.
* **Sum = 0:** Pure internal CPU operation (ALU or register transfers; external buses idle).
* **Sum > 1:** $\rightarrow$ **BUS COLLISION ERROR!** Halts execution immediately with an electrical short-circuit exception.

---

## 4. Control Micro-Word (Control Word)

Emitted by the Control Unit at each clock micro-step ($T$-State):

```
+-------------------+--------------------+----------+-----------+---------+----------+
|  EMITTER_ID (3b)  |  RECEIVER_ID (3b)  |  MEM_R   |   MEM_W   |  IO_R   |   IO_W   |
+-------------------+--------------------+----------+-----------+---------+----------+
```

### Identifier Table (Register IDs)
| Binary ID | Decimal | Emitter / Receiver Component |
|:---:|:---:|:---|
| `000` | 0 | NONE (Internal bus remains in High-Z) |
| `001` | 1 | `AX` (Accumulator Register) |
| `010` | 2 | `BX` (Base Register) |
| `011` | 3 | `CX` (Loop Counter Register) |
| `100` | 4 | `DX` (Data & I/O Register) |
| `101` | 5 | `MAR` (Memory Address Register) |
| `110` | 6 | `MDR` (Memory Data Register) |
| `111` | 7 | `ALU_OUT` / `PC` / `Temp` |

---

## 5. Unified Instruction Cycle ($T$-States: $T_0$ to $T_6$)

| Step | Phase | Micro-operation | Internal Bus | Ext Address Bus | Ext Data Bus | Control Signals | Affected Registers |
|:---:|:---:|:---|:---:|:---:|:---:|:---:|:---|
| **$T_0$** | FETCH | $\text{PC} \to \text{MAR}$ | PC drives address | Inactive | High-Z | `MAR_Load = 1` | MAR loads address |
| **$T_1$** | FETCH | Fetch OpCode, $\text{PC} = \text{PC} + 1$ | Idle | MAR drives address | Cache/RAM delivers OpCode | `MEM_READ = 1` | PC increments, MDR captures OpCode |
| **$T_2$** | FETCH | $\text{MDR} \to \text{IR}$ | MDR drives OpCode | Inactive | High-Z | `IR_Load = 1` | IR latches OpCode |
| **$T_3$** | DECODE | Decode OpCode & read operands | Idle | Inactive | High-Z | Internal UC | UC configures datapath |
| **$T_4$** | EXECUTE | Route operand addresses | Operand $\to$ MAR/TempA | Inactive | High-Z | `MAR_Load = 1` | MAR or TempA loaded |
| **$T_5$** | EXECUTE | Read/Write bus or load TempB | Operand $\to$ TempB/MDR | MAR drives address | Active transfer | `MEM_READ` or `IO_READ` | MDR or TempB loaded |
| **$T_6$** | STORE | Latch result into destination | MDR/ALU $\to$ Reg/RAM | MAR drives address | Driven data | `MEM_WRITE` or `IO_WRITE` | Destination register updated, $T \to 0$ |

---

## 6. Memory Hierarchy & Virtual Paging

### Virtual Space & Physical RAM
* **Virtual Memory:** 512 Bytes partitioned into **16 Virtual Pages** (VPN 0 to 15) of 32 Bytes each.
* **Physical RAM:** 256 Bytes partitioned into **8 Physical Page Frames** (PFN 0 to 7) of 32 Bytes each.
* **Swap Space (Disk):** 512-byte secondary storage area holding paged-out memory blocks.

### Page Table (16 Entries)
| VPN (0..15) | Present Bit ($P$) | Physical Frame ($PFN$) | Dirty Bit ($D$) | Reference Bit ($R$) | Status |
|:---:|:---:|:---:|:---:|:---:|:---|
| 0 | 1 | Frame 0 | 0 | 1 | Resident in RAM |
| 1 | 1 | Frame 1 | 1 | 1 | Modified in RAM |
| ... | ... | ... | ... | ... | ... |
| 15 | 0 | - | 0 | 0 | In Disk (Swap) |

### Page Fault Handling Sequence
1. MMU intercepts an address whose virtual page has $P = 0$.
2. CPU clock pauses and the visual **PAGE FAULT** indicator turns bright red.
3. If all 8 physical frames are occupied:
   * The **FIFO** replacement algorithm selects the oldest resident frame.
   * If the victim page has $D = 1$ (Dirty), its 32 bytes are written back to disk (**Swap-out**).
   * The victim entry is updated to $P = 0, PFN = -1$.
4. The 32 bytes of the demanded page are loaded from disk into the free frame (**Swap-in**).
5. The Page Table entry is updated ($P = 1, PFN = \text{new}, D = 0, R = 1$) and execution resumes.

### Multi-level Cache Hierarchy
* **L1I (Instructions):** 8 Bytes (1 line). Dedicated caching for PC instruction stream.
* **L1D (Data):** 8 Bytes (1 line). Dedicated caching for `LOAD`/`STORE`. Uses **Write-Back** with Dirty bit.
* **L2 (Unified):** 32 Bytes structured as 4 lines of 8 Bytes with FIFO replacement.
* **Real-time Metrics:** Hit/Miss dynamic labels, hit/miss counters, and Hit Rate percentage displays.

---

## 7. I/O Peripherals: Keyboard and 8x1 LCD Screen

### Port & Memory Map
* **`Port 0x01` (KBD Status):** Bit 0 = 1 when a key is available.
* **`Port 0x02` (KBD Data):** ASCII code of the pressed key. Reading automatically clears Port 0x01.
* **`Port 0x03` (LCD Display):** Direct character output latched to prevent display flickering.
* **Circular RAM Queue (`0x70`..`0x77`):** 8-byte FIFO keyboard buffer.
  * `0x78 = KB_HEAD` (Read pointer).
  * `0x79 = KB_TAIL` (Write pointer).
  * `0x7A = KB_COUNT` (Item counter with overflow discard protection).
* **VRAM (`0xF8`..`0xFF`):** 8-byte video buffer mapped to display with **Backspace (ASCII 0x08)** support.

---

## 8. Extended Instruction Set Architecture (ISA)

| Mnemonic | Opcode | Bytes | Operands | Description | Flags |
|:---|:---:|:---:|:---|:---|:---:|
| `HLT` | `00` | 1 | — | Halts CPU clock | — |
| `MOV r, imm` | `01` | 3 | reg, imm8 | Load immediate value into register (AX..DX) | — |
| `MOV r, r` | `02` | 3 | reg, reg | Transfer data between registers | — |
| `LOAD r, [dir]`| `03` | 3 | reg, dir9 | Load virtual memory into register via L1D/L2 Cache | — |
| `STORE [dir], r`|`04` | 3 | dir9, reg | Store register to virtual memory via L1D Write-Back | — |
| `ADD r, imm` | `10` | 3 | reg, imm8 | Add immediate to register | ZF, CF, SF |
| `ADD r, r` | `11` | 3 | reg, reg | Add two registers | ZF, CF, SF |
| `SUB r, imm` | `12` | 3 | reg, imm8 | Subtract immediate from register | ZF, CF, SF |
| `SUB r, r` | `13` | 3 | reg, reg | Subtract two registers | ZF, CF, SF |
| `INC r` | `14` | 2 | reg | Increment register | ZF, CF, SF |
| `DEC r` | `15` | 2 | reg | Decrement register | ZF, CF, SF |
| `CMP r, imm` | `16` | 3 | reg, imm8 | Subtract without saving result | ZF, CF, SF |
| `CMP r, r` | `17` | 3 | reg, reg | Subtract without saving result | ZF, CF, SF |
| `AND r, r` | `18` | 3 | reg, reg | Bitwise AND | ZF, SF (CF=0) |
| `OR r, r` | `19` | 3 | reg, reg | Bitwise OR | ZF, SF (CF=0) |
| `XOR r, r` | `1A` | 3 | reg, reg | Bitwise XOR | ZF, SF (CF=0) |
| `NOT r` | `1B` | 2 | reg | Bitwise NOT | ZF, SF (CF=0) |
| `JMP dir` | `20` | 3 | dir9 | Unconditional branch to virtual address | — |
| `JZ dir` | `21` | 3 | dir9 | Branch if Zero (ZF = 1) | — |
| `JNZ dir` | `22` | 3 | dir9 | Branch if Not Zero (ZF = 0) | — |
| `JC dir` | `23` | 3 | dir9 | Branch if Carry (CF = 1) | — |
| `IN r, [port]` | `30` | 3 | reg, port8 | Read from isolated I/O port | — |
| `OUT [port], r`| `31` | 3 | port8, reg | Write to latched I/O port | — |

---

## 9. Official Program Catalog

### Program 1: Arithmetic & Loops using AX, BX, CX, DX
Performs cumulative series calculation using all 4 general-purpose registers:
* `AX`: Accumulator sum register.
* `BX`: Increment factor.
* `CX`: Loop iteration counter (`DEC CX`, `JNZ`).
* `DX`: Micro-operation counter and control tracker.

### Program 2: Interactive I/O Keyboard -> 8x1 LCD Screen
Executes real-time interactive terminal loop:
1. Polls `Port 0x01` until keystroke is ready (`Bit 0 == 1`).
2. Reads ASCII byte from `Port 0x02`.
3. Issues `OUT [0x03], DX` to project the character to the LCD display.
4. Handles dynamic backspace deletion (`ASCII 0x08`).

### Program 3: Dispersed Memory Accesses, Page Faults & Cache Misses
Accesses non-contiguous addresses across more than 8 distinct virtual pages (`0x010`, `0x040`, `0x080`, `0x0C0`, `0x100`, `0x140`, `0x180`, `0x1C0`, `0x1E0`):
* Overflows the 8 physical RAM frames.
* Visually demonstrates FIFO eviction to swap space.
* Triggers real-time **PAGE FAULT** events highlighted in red.
* Exhibits multi-level cache miss penalties.

---

## 10. Simulator User Guide

1. Open `simuladorgood-cpu.xlsm` and enable macros.
2. Select the **CPU** worksheet:
   * **Buses & Control Word:** Observe values and dynamic High-Z status.
   * **Cache Hierarchy:** Track L1I, L1D, and L2 Hit/Miss counters.
   * **Virtual Memory:** View the 16-entry Page Table and Page Fault beacon.
   * **Peripherals:** Enter keystrokes and inspect the 8x1 LCD output.
3. Click any program preset button:
   * **LOAD PROG 1 (Loops AX..DX)**
   * **LOAD PROG 2 (Keyboard -> LCD)**
   * **LOAD PROG 3 (Page Faults & Cache)**
4. Click **STEP** to execute one micro-state at a time ($T_0 \to T_1 \to \dots \to T_6$).
5. Click **RUN** for continuous execution (speed adjustable via **Delay (ms)**).
6. Click **PAUSE** to freeze the clock or **RESET** to clear all hardware state.
