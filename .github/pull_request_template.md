## 📌 Pull Request: Implementación de Módulo / Feature

### 🎯 Destino del Merge
- **Base Branch:** `developer`
- **Feature Branch:** `feature/<estudiante>/<modulo>`

---

### 📝 Resumen del Cambio
*(Describe de forma concisa qué subsistema o funcionalidad añade o corrige este PR).*

- **Subsistema afectado:**
  - [ ] Buses & Señales de Control (`Buses.bas`)
  - [ ] Núcleo CPU & Datapath (`CPU.bas`, `Ciclo.bas`)
  - [ ] Jerarquía de Caché L1I/L1D/L2 (`Cache.bas`)
  - [ ] Memoria Virtual & Paginación Swap (`MemoriaVirtual.bas`, `MemoriaRAM.bas`)
  - [ ] Periféricos E/S Teclado & LCD (`Perifericos.bas`)
  - [ ] Ensamblador & Programas Demo (`Ensamblador.bas`, `ProgramasDemo.bas`)
  - [ ] Interfaz de Usuario / Cockpit Excel (`UI.bas`, `SetupTableroParcial2.bas`)

---

### ⚙️ Criterios de Aceptación Técnicos (Checklist de Rúbrica)
- [ ] **Exclusión Mutua de Buses:** Se verifica que $MEM\_R + MEM\_W + IO\_R + IO\_W \le 1$.
- [ ] **Alta Impedancia (High-Z):** Todo componente desconectado del bus expone `-1` (`HIGH_Z`).
- [ ] **Aduanas MAR / MDR:** No existen asignaciones de memoria directas sin pasar por MAR y MDR.
- [ ] **Caché Multinivel:** Hits y Misses contabilizados adecuadamente sin datos corruptos.
- [ ] **Paginación & Page Fault:** MMU traduce VPN a PFN; los fallos de página leen/escriben en Swap correctamente.
- [ ] **Periféricos E/S:** El teclado usa cola FIFO y el LCD retiene el dato mediante cerrojo (*Latch*).
- [ ] **No Bloqueo de UI:** Las hojas permanecen desprotegidas (`Locked = False`) y la ejecución responde a `DoEvents`.

---

### 🧪 Evidencia de Pruebas Automatizadas
*(Pega aquí la salida de `test_parcial2_all.ps1` ejecutado en tu entorno local).*

```text
[PROG 1] Aritmética y Store:       [PASS]
[PROG 2] Teclado a Pantalla LCD:   [PASS]
[PROG 3] Fallos de Página y Swap:  [PASS]
```

---

### 👥 Revisión de Pares
- **Autor:** @<github-username-autor>
- **Revisor Asignado:** @<github-username-par>
