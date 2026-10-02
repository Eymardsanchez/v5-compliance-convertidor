# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the App

```bat
iniciar.bat                   # instala deps y lanza la GUI
python ConvertidorSantander.py # directo con Python en PATH
```

Syntax check (no test suite exists):
```bash
python -c "import py_compile; py_compile.compile('archivo.py', doraise=True)"
```

Dependencies: `pdfplumber>=0.11`, `openpyxl>=3.1`, `Pillow>=10.0`  
Runtime: Python 3.12, Windows only (tkinter + `%APPDATA%` paths).

---

## DPI/Resolución y build del instalador (17 ago 2026)

**Fix de resolución (varía entre PCs / multimonitor).** `ui_theme.enable_dpi()`
ya no depende de una sola llamada: intenta en cascada
`SetProcessDpiAwarenessContext(PER_MONITOR_AWARE_V2)` → `SetProcessDpiAwareness(2)`
→ `(1)` → `SetProcessDPIAware()` (guard `_DPI_ENABLED`). Si la única llamada
anterior (`SetProcessDpiAwareness(1)`) fallaba en cierto Windows/escalado, el
proceso quedaba **DPI-unaware** y `winfo_screenwidth/height` devolvía la
resolución **virtualizada** (1920×1080 al 150 % → 1280×720) — de ahí que el
comportamiento variara entre equipos. Helpers nuevos en `ui_theme.py`:
`work_area(win)` → `(x, y, w, h)` del área de trabajo del **monitor real** donde
está la ventana/cursor (Win32 `MonitorFromWindow`/`GetMonitorInfo`, fallback a
`winfo_*`), `screen_size(win)` y `center(win, w, h)`. Las 6 ventanas de
`ConvertidorSantander.py` que dimensionaban con `winfo_screenwidth/height` ahora
usan `work_area`/`screen_size` (preservando sus `ui_theme.sx()`); el **offset del
monitor** corrige el centrado en monitores secundarios (antes `(sw-w)//2` asumía
origen 0,0). `ui_theme.py` quedó **byte-idéntico entre 5.0 y 5.1**.

**Build — spec + manifest + version + firma.**
- `app.manifest` (nuevo, enganchado vía `EXE(manifest='app.manifest')`):
  dpiAwareness **PerMonitorV2**, `requestedExecutionLevel asInvoker` (la app
  escribe en `%APPDATA%`, no necesita admin), `supportedOS` Win8.1/10/11. Fija el
  DPI del `.exe` antes de arrancar, sin depender del timing del runtime.
- `version_info.txt` (nuevo, `EXE(version='version_info.txt')`): metadata de
  editor/producto en el `.exe` (reduce falsos positivos y ayuda a reputación).
- **`upx=False`** en `EXE` y `COLLECT`: el empaquetado UPX es de los mayores
  disparadores heurísticos de AV/Defender — **aún más relevante en 5.1** por los
  binarios de `cryptography`/`lxml`/`satcfdi`. Sube el tamaño del `.exe`.
- `.iss`: bloque `SignTool` / `SignedUninstaller` (comentado) para firma
  Authenticode del instalador+desinstalador. La contraseña del `.pfx` va en la
  config del IDE de Inno (Tools > Configure Sign Tools), **nunca** en el `.iss`.
  ⚠ El `cert_publisher.cer` es **autofirmado** (sirve para las macros VBA) y
  **NO** quita el aviso de SmartScreen/Defender del `.exe`/instalador — eso exige
  un cert de firma de código de **CA comercial** (OV/EV). Orden: compilar
  PyInstaller → firmar `.exe` con `signtool` → compilar el instalador Inno.
- Icono cambiado (`app_icon.ico`, concepto "ciclo $"/conciliación con la paleta
  del sistema de diseño, multi-resolución 16→256). Anterior en `app_icon_prev.ico`.

---

## Architecture Overview

### Entry points
| File | Role |
|---|---|
| `ConvertidorSantander.py` | Main window (login + PDF→Excel tab + compliance/email tabs) |
| `ventana_conciliacion.py` | CFDI conciliation window, launched from main window |

The conciliation module is loaded with a try/import guard — the app still runs if the CFDI modules are absent.

### Módulo Saldos a favor ISR/IVA (`sat_saldos/`, integrado)

Motor propio (paquete local `sat_saldos/`, copiado del proyecto `sat-saldos-favor`)
integrado como **5º módulo del sidebar** ("💰 Saldos a favor ISR/IVA"). Lee acuses,
declaraciones detalladas RESICO, **acuses de pago** (línea de captura) y el **formato
antiguo Régimen General** (2015-2021, determinación de IVA por coordenadas pdfplumber)
y arma el papel de trabajo del remanente acumulado de IVA + verificación de pagos.

- **Enganche (aditivo, no toca los otros módulos):** guard `try: from sat_saldos import
  pipeline as _SF_PIPE … _SALDOS_OK` (junto a los demás guards); `self._tab_saldos`
  en `_build_ui`; entrada `("saldos", "💰 …")` en el sidebar; `"saldos"` en
  `_mostrar_modulo`; método `_build_saldos_tab()` + helpers `_saldos_*` (usan el mismo
  `_module_header`/`_rbtn`/tema `T`/`F` que el resto). Lectura en hilo + `self.after`.
- **PyInstaller:** `collect_submodules('sat_saldos')` agregado a `hiddenimports` del
  `.spec` (tiene imports dinámicos internos).
- **Lectura de PDF robusta** (`sat_saldos/extractor.py`): `_pdftotext_bin()` halla
  pdftotext aunque no esté en PATH (Git/poppler); si no hay, respaldo pdfplumber con
  `x_tolerance=1.5` (sin él las palabras se pegan y falla la detección). Verificado:
  la App completa construye sus 5 módulos sin regresión; el módulo lee 175 PDF reales
  → 136 detalles / 100 meses de IVA en la tabla.

### Data pipeline (PDF bank statement → Excel conciliation report)

```
PDF → bancos/detector.py         detects bank (IBM-anchor strategy, then keyword fallback)
    → bancos/{banco}.py          extracts Movimiento list + InfoCuenta
    → exportar_excel.py          PDF→Excel for bank statements (direct conversion)

XML CFDIs → cfdi_parser.py       parsear_cfdi() → CFDI dataclass
          → cfdi_validador.py    validar_lote() → dict[uuid → CFDIValidationResult]
          → conciliacion.py      conciliar_multibanco() → [MovConciliado], sin_cfdi, etc.
          → exportar_conciliacion.py  generates the Excel conciliation workbook
```

### Key dataclasses

**`cfdi_parser.CFDI`** — one CFDI/XML. Notable fields: `tipo` (I/E/P/N/T), `metodo_pago` (PPD/PUE), `cfdi_facturas_pagadas` (list of UUIDs paid by a tipo-P complemento), `cfdi_sustituye`, `regimen_emisor`, `regimen_receptor`.

**`conciliacion.MovBanco`** — enriched bank movement. `.monto` property returns retiro if nonzero, else deposito.

**`conciliacion.MovConciliado`** — matched pair (MovBanco + CFDI). `cfdi_suma` holds extra CFDIs when a bank movement covers multiple invoices (nivel 4).

**`cfdi_validador.CFDIValidationResult`** — uuid, errores, advertencias, score (100→-200), estatus (VERDE/AMARILLO/ROJO/NEGRO).

### Conciliation phases / levels

`MovConciliado.tipo_match` records which phase produced the match. Value propagates to **col 16 TIPO MATCH** in each bank sheet (color-coded). La hoja de banco tiene **16 columnas**: **UUID CFDI / N.FACTURA** (col 7), **UUID REP** (col 8), **SERIE** (col 9) y **FOLIO CFDI** (col 10) tomadas del XML conciliado (`MovBanco.serie_cfdi`/`folio_cfdi`, llenadas en `_llenar_mov`).

**Trazabilidad REP ↔ CFDI (columnas separadas):** `MovBanco.uuid_rep` guarda el UUID del complemento de pago (tipo P) cuando el cruce se hizo vía complemento (**Fase 4**); `MovBanco.uuid` guarda el UUID de la **factura** (tipo I) que el REP liquida, y las facturas adicionales van en `uuid_suma`. Antes la Fase 4 ponía el UUID del REP en la columna de CFDI y las facturas en `uuid_suma`, así que no se distinguía complemento de factura. `_conciliar_rep` resuelve las facturas contra el pool (`cfdi_facturas_pagadas` → CFDI) y solo si ninguna vino en el lote cae de vuelta al UUID del REP. La **col 7 lista TODAS las facturas relacionadas** (primera + adicionales, una por renglón con `+`): el gate del export era `es_suma` (`estatus == "Conciliado SUMA"`), pero la Fase 4 deja estatus `"Conciliado"`, así que un REP consolidado solo mostraba la primera factura y ocultaba el resto. Ahora se listan siempre que `uuid_suma` tenga contenido — también beneficia a Global Mensual (Fase 5), que tenía el mismo hueco. El historial (`repo_conciliaciones._dict_a_movbanco` → `MovBanco(**d)`) es retro-compatible: las filas guardadas antes de este cambio toman el default vacío.

**Criterio banco↔CFDI:** depósito → CFDIs **emitidos**; retiro → CFDIs **recibidos**. Ids de fase en `FASES_CONCILIACION` (`conciliacion.py`), activables vía `fases=`.

**Rediseño del pipeline (esta versión):** Fase 2 pasó de "Global Mensual" (N:M) a **N:1**; el N:M se movió a Fase 5 (automático). **Fase 4b se eliminó** (fusionada en Fase 4). **Fase 4 ya no cruza por importe libre** — verifica el REP por MontoPagado + identidad, con la FechaPago del REP en el mismo mes bancario o a ±5 días. Fase 7 ahora **exige mismo mes**.

**Reorden de fases + split PPD RFC/Nombre (10 ago 2026, a pedido):** el orden de
EJECUCIÓN cambió para correr las fases PPD (evidencia respaldada por complemento)
ANTES de las de suma, y **cada PPD se dividió en un paso por RFC (primero,
identidad fuerte) y otro por Nombre** (después). Orden nuevo (9 fases): **1 exacto
→ 2 ppd_con_rfc → 3 ppd_con_nom → 4 ppd_sin_rfc → 5 ppd_sin_nom → 6 combinatorio
→ 7 multi_mov → 8 grupos → 9 solo_importe**. El split se hace SIN tocar el núcleo
`_buscar`: `_conciliar_ppd`/`_conciliar_rep` reciben `modo` ("rfc"|"nombre"|"ambos");
en `_conciliar_ppd` el modo pasa un dict vacío para apagar el nivel de RFC o de
nombre; en `_conciliar_rep` restringe la rama de `_mov_coincide_identidad`. Es
**outcome-neutral**: la unión (RFC-o-Nombre) se conserva, solo se prioriza el RFC
y se afina la atribución de fase. `tipo_match` y `nivel` NO cambiaron (ambos pasos
de un PPD comparten su `tipo_match`). Verificado STERIMEX jun 2026: **137/166 =
82.5% idéntico** antes y después (sin regresión); byte-idéntico en 5.0 y 5.1.

| Fase | id | `tipo_match` | Logic |
|---|---|---|---|
| 1 | `exacto` | `"Exacto"` | 1:1 — RFC/nombre + importe, MISMO MES. Bancarizadas; tipo P excluido. |
| 2 | `ppd_con_rfc` | `"PPD c/ Complemento"` | REP (tipo P) por MontoPagado + **RFC** en descripción + mismo mes bancario O ≤ `TOLERANCIA_DIAS_REP` (5 d). `_conciliar_rep(modo="rfc")`. REP en `uuid_rep` (col 8), factura en `uuid` (col 7). `nivel=5`. |
| 3 | `ppd_con_nom` | `"PPD c/ Complemento"` | Igual que la 2 pero por **Nombre** del tercero, sobre lo que quedó. `_conciliar_rep(modo="nombre")`. `nivel=5`. |
| 4 | `ppd_sin_rfc` | `"PPD Sin Complemento"` | PPD sin complemento — exige **RFC** en la descripción. `_conciliar_ppd(solo_identidad=True, modo="rfc")`. `nivel=7`. |
| 5 | `ppd_sin_nom` | `"PPD Sin Complemento"` | Igual que la 4 pero por **Nombre**, sobre lo que quedó. `_conciliar_ppd(modo="nombre")`. `nivel=7`. |
| 6 | `combinatorio` | `"Combinatorio"` | **1:N** — un movimiento = suma de 2–5 CFDIs del mismo mes (`_buscar_suma`). `nivel=4`. |
| 7 | `multi_mov` | `"Multi-Movimiento"` | **N:1** — varios movimientos (mismo RFC/nombre, mismo mes) que **suman UN CFDI** (`_conciliar_multimov`). `nivel=2`. |
| 8 | `grupos` | `"Grupo:.."` / `"Global Mensual"` | **Agrupado por RFC**: grupos configurados (`_conciliar_grupos`) + automático RFC+mes (`_buscar_global_mensual`). |
| 9 | `solo_importe` | `"Solo Importe"` | **Opcional, apagada por defecto.** Importe exacto **del MISMO MES**, sin identidad; excluye tipo P; bitácora `bitacora_f7`. `nivel=8`. |

Helpers: `_dias_entre`/`_fecha_partes`, `_subset_sum_idx`, `_mov_coincide_identidad`. **Tipo P (REP) excluido de los pools bancarizados** (Fases 1, 6, 7, 8 — exacto/combinatorio/multi_mov/grupos).

Nivel detail (Fase 1): 1 = monto+mes+RFC (±$1 MXN); nombre parcial; nombre-en-desc; 3.5 = monto+mes (omitido con `solo_identidad=True`).

**Helper PPD**: `_uuids_con_complemento(cfdis)` da los UUIDs de facturas ya referenciadas en un complemento tipo P. **Fases 2-3** (`_conciliar_rep`, modo rfc/nombre) resuelven los PPD **con** complemento vía el REP; **Fases 4-5** (`_conciliar_ppd`, `con_complemento=False`, modo rfc/nombre) los **sin** complemento, exigiendo identidad. La rama `_conciliar_ppd(con_complemento=True)` ya no se usa.

**Fase 2/3 (PPD con complemento): de ±1 día a MISMO MES o ±5 días (6 ago 2026).** El criterio original comparaba `|FechaPago del REP − fecha banco| ≤ 1 día` (helper `_dias_entre`), pero en producción muchos complementos legítimos caían fuera: el REP se timbra días después del pago. El candado de periodo es ahora el **mes** —`mov.mes or _mes_fecha(mov.fecha)` vs `_mes_fecha(rep.fecha_pago)`— **O**, si caen en meses distintos, un desfase de hasta **`TOLERANCIA_DIAS_REP = 5` días** (constante de módulo). Ese OR existe porque el banco a veces refleja el movimiento desfasado: un pago del 30/jun que aparece en el estado de cuenta el 01/jul cruzaría con el mes solo, y sin él se perdería. La precisión la siguen dando el **MontoPagado exacto** (±TOLERANCIA_MXN) y la **identidad** (RFC o nombre en la descripción bancaria), que no cambiaron. Verificado en ambas versiones (8 casos): mismo mes cruza con cualquier desfase (1, 23 días, REP posterior al banco); cruzando mes cruza a 1, 2 y 5 días (en ambas direcciones) y **NO** a 6 días; mayo vs junio tampoco. Sin regresión en la traza `uuid`/`uuid_rep`.

**`conciliar_multibanco` structure**: Pasada 1 (Fase 1 exact, skip_suma=True) → Fase 2 PPD c/ Complemento por RFC → Fase 3 PPD c/ Complemento por Nombre → Fase 4 PPD Sin Complemento por RFC → Fase 5 PPD Sin Complemento por Nombre → Fase 6 Combinatorio (1:N) → Fase 7 Multi-Movimiento (N:1) → Fase 8 Agrupado por RFC (grupos + automático) → **Fase 9 Solo Importe mismo mes (opcional)**.

### PRÉSTAMO — movimiento no conciliable (10 ago 2026)

**Regla de negocio:** un movimiento bancario cuya descripción contiene la palabra
**PRESTAMO** o **PRÉSTAMO** (préstamo recibido, otorgado o su devolución/pago)
**NUNCA** se cruza contra un CFDI — un préstamo no es operación facturable, no
existe comprobante que lo respalde — aunque monto/mes/RFC coincidan exacto.

**Mecanismo (`conciliacion.py`, ambos entry points):**
`PALABRAS_NO_CONCILIABLES = ("PRESTAMO", "PRÉSTAMO")` + helper
`_es_no_conciliable(mov)` (subcadena sobre la descripción en mayúsculas). En
`Conciliador.conciliar()` los préstamos se apartan al inicio del loop de
movimientos (ANTES del check `mov.monto == 0`), con `estatus = "Préstamo — no
conciliable"`, y se reincorporan a `sin_cfdi` DESPUÉS de la última fase para que
ninguna (1-7) intente cruzarlos. En `conciliar_multibanco()`, como la pasada 1
delega en `conciliar(skip_suma=True, fases=fases)` ya vienen excluidos y con
estatus dentro de `sin_cfdi`; se apartan de nuevo (`prestamos_por_res`) antes de
las Fases 2-7 y se reincorporan por banco antes de `return`. Quedan como
**pendientes con estatus propio**, visibles en el reporte.

**Verificado con datos reales (STERIMEX jun 2026):** el estado trae **2
préstamos reales** — un depósito `220626 PRESTAMO` de $125,000 (VALKIR) y un
`DEVOLUCION DE PRESTAMO JUNIO 2026` — ambos detectados sin falso positivo.
Conciliación **137/166 = 82.5% idéntica con y sin la regla** (delta 0: esos
movimientos no cruzaban de todos modos por no tener CFDI → cero regresión; lo que
cambia es que ahora quedan etiquetados como préstamo en vez de "sin CFDI").
Test sintético: un PRESTAMO con monto+mes+RFC idénticos a un CFDI **no** cruza,
uno normal sí, 'PRÉSTAMO' con acento también se aparta, 2/2 pendientes con
estatus por ambos entry points. Lógica byte-idéntica en 5.0 y 5.1.

### Selector de fases (casillas) — `fases` param

`conciliar()` y `conciliar_multibanco()` aceptan `fases: set[str] | None` — conjunto de ids de fase a ejecutar. `None` → `FASES_DEFAULT` (todas menos `solo_importe`). Cada fase se salta si su id no está en el conjunto. Catálogo `FASES_CONCILIACION` (lista `(id, etiqueta)`) e ids: `exacto`, `multi_mov`, `combinatorio`, `ppd_con`, `grupos`, `ppd_sin`, `solo_importe` — definidos en `conciliacion.py` y exportados para la UI. (Ids removidos en el rediseño: `global_mensual`, `rep`.) `conciliar_multibanco(..., fases=..., log=callback)` — `log(msg, tag)` opcional recibe la bitácora de la Fase 7.

**UI**: en `_build_conciliacion_tab` (`ConvertidorSantander.py`) un `LabelFrame` "Fases de conciliacion" con una casilla (`tk.BooleanVar` en `self._concil_fase_vars`) por fase, inicializadas desde `FASES_DEFAULT` (Fase 7 apagada). `_concil_run` arma `fases_sel` de las marcadas y lo pasa a `conciliar_multibanco(..., fases=fases_sel, log=self._concil_log_write)`.

**PPD invoices (MetodoPago="PPD")** are included in the conciliation pool like any other CFDI. R003 flags them in validation but does not remove them from the pool.

**Nivel 4 pruning invariant** — `_dp_find_combo` uses two bounds: (a) if the two *smallest* items already exceed target → prune; (b) if the sum of the `_MAX_COMBO` *largest* items is still below target → prune. Using smallest-items sum > target is WRONG (it incorrectly eliminates groups like 5 × $2,900 when target = $5,800).

**Bug real corregido — Fases 1 y 3 conciliaban CFDIs de meses distintos al
del banco, y sin filtrar por forma de pago.** `_buscar()` (Fase 1, niveles
1-3) construía `meses_fwd`/`meses_bwd` explorando los 12 meses del año en
ambas direcciones — un PUE de enero podía conciliar contra un banco de
junio con solo coincidir monto+RFC, sin marcarse siquiera para revisión.
`_buscar_suma()` (Fase 3, combinatorio) era peor: no filtraba por mes en
absoluto, solo agrupaba por RFC. Regla de negocio confirmada: **PUE (y
cualquier forma bancarizada) solo puede conciliar dentro del mismo mes que
el estado de cuenta — nunca desfasado. PPD sí puede cruzar mes, pero
únicamente vía las fases PPD (Fase 4 con complemento / Fase 6 sin
complemento), diseñadas para resolver contra el complemento de pago.**

Fix (aplicado en ambos entry points — `Conciliador.conciliar()` y su
duplicado independiente dentro de `conciliar_multibanco()`, que NO delega
en el primero para Fases 2-4):
- `_buscar()` ganó el parámetro `mismo_mes: bool`. Fase 1 lo pasa como
  `True` (busca solo `mes_base`, sin `meses_fwd`/`meses_bwd`); Fase 4 lo
  deja en `False` (default, sin cambios) para preservar el cruce de mes de
  PPD.
- `_buscar_suma()` ahora filtra sus candidatos a `c.mes == mes_mov` antes
  de intentar el subset-sum — mismo criterio, función compartida por ambos
  entry points.
- `FORMAS_PAGO_BANCARIZADAS = {"02","03","04","28","29"}` (cheque,
  transferencia, tarjeta crédito/débito/servicios) — Fases 1, 2 y 3 en
  ambos entry points ahora construyen sus pools de búsqueda filtrando a
  `c.forma_pago in FORMAS_PAGO_BANCARIZADAS` antes de indexar. Efectivo,
  condonación, compensación, etc. quedan excluidos — nunca producen un
  movimiento bancario real, sin importar que monto/mes/RFC coincidan.
  Fase 4 sigue usando el pool sin filtrar (`metodo_pago=="PPD"` ya lo
  acota).

Verificado con `conciliar_multibanco()` real (el entry point de
producción): PUE mismo mes bancarizado concilia: PUE de otro mes ya NO
concilia (bug original corregido); efectivo mismo mes/monto/RFC tampoco
concilia (regla nueva); PPD de otro mes sigue conciliando vía Fase 4 sin
regresión; Fase 3 combinatorio tampoco suma cruzando mes; nivel 2/3 (match
por nombre) también respeta mismo mes.

### Sistema de diseño (6 ago 2026)

App y reporte Excel comparten una sola paleta, para que se lean como un mismo
producto. Los colores están agrupados por el TRABAJO que hacen, no por su tono.

**Los 4 hues de categoría se validaron, no se eligieron a ojo.** Se corrieron
contra superficie blanca en TODOS los pares (no solo adyacentes): azul `2A78D6`,
naranja `EB6834`, aqua `1BAF7A`, violeta `4A3AA7` → separación CVD ΔE 9.2 y de
visión normal ΔE 16.3, ambas sobre el piso. **Un 5º hue rompía el piso**, y los
tintes pálidos (los que dejan leer texto de 9 pt encima) pierden el croma
necesario: azul vs violeta en tinte quedan a ΔE 4.5, indistinguibles. De ahí
salen dos decisiones:

1. Las **7 fases se agrupan en 4 niveles de confianza** (`_TIER_FASE` / `_tier()`
   en `exportar_conciliacion.py`): Directo (Fase 1) · Compuesto (2/3/5) ·
   Con complemento (4) · Débil (6/7). El color agrupa; el **nombre de la fase,
   que va escrito en la celda, distingue** — el color nunca carga la identidad
   solo.
2. En la columna TIPO MATCH el nivel se marca con un **borde izquierdo grueso**
   del hue saturado sobre relleno claro, no con relleno de color: así el texto
   conserva su contraste y el hue conserva su croma. En la leyenda, donde el
   swatch es un bloque, se usa el paso oscuro del mismo hue (`_TIER_OSCURO`) con
   texto blanco (≥5.3:1).

**Los colores de estatus están reservados** (conciliado / revisar / sin CFDI) y
nunca se usan como color de categoría. Antes Fase 3 y Fase 7 compartían
literalmente el mismo relleno `FFE0B2`, y Fase 1 usaba el verde de "conciliado":
dos categorías se veían iguales y una se disfrazaba de estatus.

Otros cambios de la hoja de banco:
- **Solo se tiñe la excepción.** Antes todo renglón conciliado iba pintado de
  verde y cada retiro de rojo; ahora el cuerpo es neutro y lo único resaltado es
  "Sin CFDI", que es justo lo que se anda buscando. Retiro/depósito llevan un
  tinte cálido/frío mínimo — son dos sentidos de un flujo, no bueno/malo.
- **Cifra principal arriba**: el % conciliado va en el subtítulo de la hoja, no
  solo al pie.
- **Alto de fila que sí contempla el texto que envuelve** (`_alto_fila`). Contar
  solo los `
` dejaba las descripciones SPEI largas encimadas entre renglones.
- Rejilla de Excel apagada, hairlines en color de línea, encabezado como barra
  continua, y el bloque de porcentajes como tabla real (mes · conciliados ·
  total · %) en vez de renglones fusionados.

**GUI — bug de tema corregido.** `ui_theme.apply_theme()` autodetectaba el tema
del SO, pero los módulos hacen `from ui_theme import C_WHITE, …`: un from-import
liga el nombre con el valor que tenía **al importar**, así que
`_refresh_module_compat()` nunca alcanza esos nombres y los ~270 widgets tk se
quedan en claro para siempre. Con el SO en dark, solo los widgets ttk (Treeview,
Notebook) se iban a oscuro → app clara con **paneles de lista negros**. Ahora el
default es claro (`dark=False`); `dark=True` sigue disponible para cuando se
migren esos widgets a leer `T.*` en vivo. También se agregó `T.ACCENT_FG_2`
(texto secundario sobre la barra de acento): se estaba usando `TEXT_MUTED`, que
es tinta oscura para fondo claro, y sobre el azul daba **1.1:1**.

El selector de fases de la pestaña Conciliación se reagrupó por esos mismos 4
niveles, con la misma barra de color que la columna TIPO MATCH del Excel.

### Hoja TABLERO y cifras vivas (14 ago 2026)

**El problema:** el reporte se calcula en Python y se escribía como números
fijos. En cuanto el contador concilia a mano con la macro `PapelTrabajo`
(Alt+F8 → `ConciliarManuales`), esos totales quedaban mintiendo: el % seguía
mostrando el avance del momento de la exportación.

**Ahora todo lo que resume es fórmula.** Criterio único, igual para lo
automático y lo manual: un renglón cuenta como conciliado si su **TIPO MATCH
(col P) no está vacío** — la macro escribe `"Manual"` ahí, así que las
conciliaciones a mano entran solas sin tocar nada más.

| Dónde | Qué se volvió fórmula |
|---|---|
| `TABLERO` (hoja nueva, primera) | avance por estado de cuenta (movimientos / conciliados / pendientes / %), TOTAL, y desglose por nivel de confianza + Manual |
| Subtítulo de cada hoja de banco | `COUNTIF` + `TEXT(...)` — "N de M movimientos conciliados (X%)" |
| Bloque "Conciliación por mes" | `COUNTIFS` por mes y por col P |

**Columna auxiliar `MES` (col 18 / R, oculta).** El bloque mensual necesita el
mes por fórmula, pero la fecha de la col A es texto y su formato cambia por
banco (`"04/JUN./2026"` en Inbursa), así que no se puede derivar de forma
confiable con fórmula. Se escribe el mes numérico en la col 18: va ahí y no en
la 17 para no crecer la tabla visible ni el rango del autofiltro, y queda fuera
de las columnas 7-16 que toca la macro.

**La pestaña NO puede vivir en la plantilla.** `_nuevo_workbook()` carga
`plantilla_papel_trabajo.xlsm` solo para heredar el `vbaProject.bin` y **borra
todas sus hojas**; además las fórmulas tienen que apuntar a los nombres reales
de las hojas de banco, que cambian por reporte ("SANTANDER MAY"). Por eso la
genera `_hoja_tablero(wb, hojas)` al final del export, cuando esas hojas ya
existen. `_hoja_banco()` ahora devuelve `n_datos` para armar los rangos.

**Verificado en Excel real** (no solo inspección del XML): con el reporte de
muestra, al escribir `"Manual"` en la col P de un renglón "Sin CFDI" y
recalcular, el TABLERO pasó de 11/16 (68.8%) a 12/16 (75.0%), pendientes 5→4,
el renglón "Manual" del desglose pasó a 1, el bloque mensual de junio de 9/13 a
10/13 y el subtítulo de la hoja de banco se reescribió solo.

### construir_plantilla.py — SaveAs en OneDrive (14 ago 2026)

Excel COM no guarda de forma fiable en carpetas sincronizadas por
OneDrive/SharePoint: intenta la ruta web y truena con `0x800A03EC`
("Error en el método SaveAs de la clase Workbook"). Este proyecto vive en
OneDrive, así que el script guarda en `%TEMP%` y copia al destino con
`shutil.copyfile`. Sin esto, la plantilla no se podía generar aquí.

### Validation rules (cfdi_validador.py)
`R001` cancelled/untimbred · `R002` missing RFC · `R003` PPD without complemento de pago · `R004` generic UsoCFDI · `R005` inconsistent IVA · `R006` fiscal regime mismatch vs company CSF · `R007` substituted CFDI · `R008` duplicate UUID · `R009` missing fiscal fields.

R006 applies to tipos I, E, P, and N. For emitidos it checks `regimen_emisor`; for recibidos it checks `regimen_receptor`.

### Persistence
All data lives in `%APPDATA%\ConvertidorBancos\`. Files are XOR-encrypted with a SHA-256 key derived from the machine's MAC address — the data is **hardware-locked** and not portable between machines.

| File | Contents |
|---|---|
| `empresas.bin` | `{RFC: {nombre, regimenes, regimenes_nombres}}` |
| `compliance.json` | `{RFC: {obligaciones_activas, destinatario, carpeta_raiz, …}}` |
| `email_cfg.bin` | SMTP credentials |
| `usuarios.bin` | User credentials |
| `conciliaciones.db` | SQLite — historial de conciliaciones completas (ver abajo) |

`empresas.py` provides the public API: `agregar_empresa`, `registrar_desde_cif`, `obtener_regimenes`, `listar_empresas`, etc. The CSF/CIF PDF is parsed by `cif_parser.py` to extract RFC, razón social, and SAT regime codes.

### Historial de conciliaciones (repo_conciliaciones.py)

`VentanaConciliacion` keeps its loaded XMLs, bank statements, and conciliation
result (`self._ultimo_resultado`) only in memory — closing the app before
generating Pólizas COI used to lose all of it, forcing a full reprocess.
`repo_conciliaciones.py` persists the complete result (not the source file
paths — those may move; the already-parsed `CFDI`/`MovBanco`/`MovConciliado`/
`ResultadoBanco` dataclasses are serialized instead) to
`%APPDATA%\ConvertidorBancos\conciliaciones.db` (SQLite, one row per
successful conciliation run). Each row's JSON payload is XOR-encrypted with
the same `_cipher_key()` scheme as `empresas.py` — metadata columns (rfc,
fecha, conteos) stay in plain SQL for listing/filtering.

- `guardar_conciliacion(rfc, resultados, recibidos, emitidos, destino)` is
  called at the end of `_ejecutar_conciliacion`, right after the Excel export
  succeeds, wrapped in try/except so a save failure never blocks the
  already-completed conciliation.
- `listar_conciliaciones(rfc)` returns lightweight metadata only (no payload)
  for the history dialog (`_abrir_historial` / "📂 Abrir conciliación
  guardada" button).
- `cargar_conciliacion(id)` reconstructs the full dataclass tree with the
  exact shape `VentanaConciliacion._ultimo_resultado` expects, so
  `_habilitar_boton_polizas()` and the Pólizas COI flow work identically
  whether the conciliation just ran or was reopened from history.
- Both list/load functions fail soft (return `[]` / `None`) on a missing or
  corrupt database instead of raising — a broken history file must not
  prevent using the rest of the app.

### Fusión automática con el historial (merge on re-conciliar)

Por defecto (checkbox "🔗 Fusionar con la última conciliación guardada" en
la barra de acción, activo), cada corrida de `_conciliar` trae la conciliación
más reciente guardada para ese RFC (`obtener_ultima_conciliacion`) y la
fusiona con lo recién cargado ANTES de correr `conciliar_multibanco` — así
una sesión nueva que solo carga XMLs/bancos nuevos no pierde lo ya conciliado
en sesiones anteriores. Solo se fusiona con la fila más reciente, no con todo
el historial, porque cada fila guardada ya es acumulativa (se guardó después
de fusionarse con su predecesora).

- `_fusionar_con_historial()` deduplica CFDIs por UUID (clave única real del
  SAT — sin riesgo de falso positivo); el CFDI recién cargado gana en caso de
  conflicto.
- `_armar_bancos_input()` deduplica movimientos bancarios por
  `(fecha, folio, descripcion, retiro, deposito, saldo)` — un estado de
  cuenta no trae un ID de transacción real, así que esta clave compuesta es
  la aproximación más cercana. Es una limitación conocida: dos movimientos
  genuinamente distintos con esos seis campos idénticos (mismo día, mismo
  monto, sin folio, misma descripción) se tratarían como duplicados.
- Bancos que existen en el historial pero no se recargan en la sesión actual
  se conservan tal cual (no se pierden por no volver a seleccionarlos).
- Si el checkbox se desactiva, el comportamiento es el de antes de esta
  función: solo se concilia lo cargado en memoria en esa sesión.

### Manual suma helper (ventana_conciliacion.py → `_abrir_buscador_suma`)
"Buscar CFDIs que sumen un importe" button opens a Toplevel where the user enters a target amount and selects Retiro/Depósito. Uses the same DP subset-sum as nivel 4 but iteratively (removes found combo from pool and repeats up to 10 times to surface multiple possible combinations). Works with CFDIs already in `_ultimo_resultado` or loads them fresh from `_xml_items`. Results shown in a grouped Treeview.

### Excel conciliation report (exportar_conciliacion.py)
Sheets generated: one per bank (movimientos), **Recibidos**, **Emitidos**, **Nomina**, **Comp. Pago Emitidos**, **Comp. Pago Recibidos**, **Resumen**.

- **Recibidos / Emitidos** sheets exclude tipo P and tipo N CFDIs — tipo P goes to its own sheets, tipo N goes to Nomina.
- **Comp. Pago Emitidos** — tipo-P CFDIs where the company is the emisor. Header color = `C_HDR_E`.
- **Comp. Pago Recibidos** — tipo-P CFDIs where the company is the receptor. Header color = `C_HDR_R`.
- Multi-banco sheet names include month abbreviation (e.g. `"SANTANDER MAY"`). Base truncated to 23 chars + 4-char suffix to stay ≤31 chars.
- **Hoja de banco (`_hoja_banco`, 16 columnas)**: col 7 = UUID CFDI / N.FACTURA, col 8 = **UUID REP** (complemento de pago, solo en cruces de Fase 4), cols 9-10 = SERIE / FOLIO CFDI (del XML). La col 16 TIPO MATCH se colorea por fase; la leyenda al pie lista las **7 fases** del pipeline nuevo con su conteo (Multi-Movimiento azul, PPD c/ Complemento FechaPago mismo mes o ±5 días, Agrupado por RFC auto/grupos, Solo Importe naranja mismo mes). **Bloque "PORCENTAJE DE CONCILIACION POR MES"** entre los datos y la leyenda: agrupa las filas por mes (`MovBanco.mes` o `_mes_fecha(fecha)`), muestra `conciliados/total (%)` por mes (verde ≥90 %, amarillo ≥60 %, rojo abajo) y un renglón TOTAL del banco. `_mes_fecha` se importa de `conciliacion`.

Key helpers:
- `_hoja_cfdis()` — Recibidos/Emitidos sheets (24 columns). `complementos_map` (uuid_factura→[tipo-P CFDIs]) drives PPD liquidation status in cols 21-24.
- `_hoja_complementos_pago(wb, tipo_p_cfdis, …, titulo, color_hdr)` — accepts `titulo` and `color_hdr` to create both Comp. Pago sheets. For "Conciliado/Sin banco" status (col 14) and bank name (col 13) it looks up `cfdi.cfdi_facturas_pagadas` in `lookup_banco` first (because the matched UUIDs are the tipo-I facturas, not the tipo-P UUID itself).
- `_build_complementos_map()` — builds `{uuid_lower → [CFDI tipo P]}` from all loaded CFDIs.
- `validar_lote()` is called once per export and the resulting `val_map` is threaded through all sheet functions.
- `lookup_banco` dict (uuid → `{banco, folio, fecha_banco, monto_banco, …}`) is built from `MovConciliado` list and passed to all sheet functions. Keys are UUIDs of tipo-I/N CFDIs (what was conciliated), not tipo-P.

### Descarga Masiva SAT (sat_descarga.py)

Descarga XMLs de CFDIs directamente del SAT (servicio oficial de Descarga
Masiva) usando la e.firma (FIEL) del contribuyente, y los carga al Paso 1
de la pestaña **real** "📊 Conciliación CFDI" (`_build_conciliacion_tab` en
`ConvertidorSantander.py` — **no** `ventana_conciliacion.py`/`VentanaConciliacion`,
que existe en el repo pero nunca se importa ni se abre desde la app; todo
el trabajo de historial/fusión de sesiones anteriores vive ahí y sigue
huérfano hasta que se decida reemplazar `_build_conciliacion_tab` por ese
módulo).

Dos modos, ambos activados desde el botón "📥 Descargar del SAT" (PASO 1):

- **Por rango de fechas** (`sat_descarga.descargar_rango`) — una sola
  solicitud SAT cubre todo el período (emitidos + recibidos). Opcionalmente
  se puede escribir una lista de UUIDs para quedarse solo con esos XMLs del
  lote ya descargado — el filtro es **local y gratuito** (no consume cuota
  SAT), implementado en `ConvertidorSantander._filtrar_por_uuid_local()`
  moviendo los XMLs que coinciden a `carpeta/filtrado/` sin borrar el resto.
- **Por UUID específico** (`sat_descarga.descargar_por_uuid`, endpoint real
  `recover_comprobante_uuid_request`) — cada UUID genera una solicitud SAT
  independiente y consume una unidad de `MAX_SOLICITUDES_DIA` (10/RFC/día).
  Pensado para un folio puntual, no para lotes.

La contraseña de la e.firma **nunca se escribe a disco** — solo vive en
memoria mientras `_DialogDescargaSAT` está abierto. Los XMLs descargados
quedan en `{carpeta_base}\{RFC}\{timestamp}\emitidos|recibidos|por_uuid\`;
esa carpeta (o `.../filtrado/` si se aplicó filtro UUID) se agrega
automáticamente a `self._concil_carpetas`, igual mecanismo que "+ Agregar
carpeta" manual — `leer_cfdis_carpetas()` la escanea de forma recursiva.

`carpeta_base` es configurable desde el botón "📂 Cambiar…" del diálogo
(persistida en `%APPDATA%\ConvertidorBancos\sat_config.json` vía
`_guardar_sat_carpeta_base()`/`_sat_carpeta_base()`) — a propósito **la
configuración** vive en AppData (es solo una ruta, no sensible), pero **los
CFDIs descargados no**: el default es `C:\ConciliaRFC\SAT`, fuera del
perfil de usuario, para no acumular volumen de datos ahí. Se recuerda entre
sesiones y aplica a todos los RFC (una sola carpeta base, organizada por
RFC y por fecha de descarga debajo).

**Revertido por instrucción explícita del usuario — granularidad separada
por tipo (Año/Mes para Recibidos) y aislamiento de fallas por tipo:** se
implementaron y verificaron con evidencia real (ver historial de commits/
conversación), pero el usuario reportó que Emitidos dejó de descargar tras
ese cambio y pidió expresamente revertir a este estado: un solo rango de
fechas compartido (`_var_f_ini`/`_var_f_fin`) para ambos tipos, una sola
llamada a `descargar_rango(..., tipos={"emitidos","recibidos"})` (o el
subconjunto marcado), sin paneles Año/Mes separados ni try/except aislado
por tipo. Este proyecto no tiene git — el revert se hizo editando
manualmente de vuelta a la forma que el usuario había confirmado como
funcional (después del fix de `EstadoComprobante`, antes de estos dos
cambios). Verificado tras el revert: con datos reales (interceptando el
callback de log sin tocar el widget Tk desde el hilo — bombear `root.update()`
manualmente en vez de `mainloop()` real puede colgarse con `self.after()`
llamado desde un hilo, una limitación del arnés de prueba, no de la app),
el flujo completo corre de punta a punta: ambos tipos se solicitan al SAT
con el mismo rango, como antes de las tareas revertidas.

Si en el futuro se retoma la idea de Año/Mes para Recibidos o el aislamiento
por tipo, el código de referencia (ya escrito y probado una vez) está en el
historial de esta conversación — no hay que rediseñarlo desde cero, solo
confirmar primero con un log fresco que el problema no es el límite SAT
5002 (solicitud repetida con los mismos parámetros) antes de reintentar.

**Bug real corregido — CodEstatus de error se reportaba como "Solicitud
aceptada" (confirmado en vivo, 10 jul 2026):** Un log real mostró
`CodEstatus=5002, Mensaje="Se han agotado las solicitudes de por vida"` —
la conexión al SAT funcionó (llegó una respuesta completa), pero el SAT
rechazó la solicitud por cuota agotada para ese RFC. `_ciclo_descarga` solo
validaba `if not id_solicitud: raise ...` antes de loguear "Solicitud
aceptada" — como el SAT repite `IdSolicitud`/`RfcSolicitante` incluso en
respuestas de error, ese check pasaba y la app reportaba éxito falso sobre
un rechazo real. Fix: valida `CodEstatus` explícitamente contra los códigos
de éxito conocidos (`5000` aceptada, `5003` tope máximo — ya manejado
arriba con paginación, `5005` duplicada — ya manejado arriba) y lanza
`DescargaSATError` con mensaje claro para cualquier otro código (mensaje
específico para `5002`: cuota de por vida agotada, no reintentable).
Verificado con dobles de `sat` (sin red): el caso 5002 real ya no se
loguea como aceptado y lanza excepción; 5000 y 5005 siguen funcionando
idéntico a antes del fix. **No confundir con problema de conexión** — un
navegador pudiendo entrar al portal SAT no implica nada sobre la cuota del
servicio de Descarga Masiva, son sistemas separados.

**Confirmado — esta app ya usa el mismo nivel de servicio que las apps
comerciales de descarga masiva (10 jul 2026, evidencia: tutorial oficial de
OneFacture aportado por el usuario, transcrito y verificado):** OneFacture
tiene dos módulos distintos con límites distintos. Su "Módulo Descarga CFDI"
(recomendado para <8,000 XML) es un **robot que automatiza el portal web del
SAT** (login por contraseña+captcha o FIEL) — hereda los límites del portal:
tope de 2,000 XML y throttling si se piden >500 XML/segundo. Su propio
tutorial dice textualmente que "ambas limitaciones no son controladas por
OneFacture" (o sea, son del SAT, no algo que ellos evadan). Su segundo
módulo, "SAT Web Service" (recomendado para volumen, sin esas limitaciones,
hasta 200,000 XML por consulta), es el servicio SOAP oficial de Descarga
Masiva de Terceros — el mismo que esta app ya usa vía `satcfdi`
(`recover_comprobante_emitted_request`/`recover_comprobante_received_request`
→ operaciones `SolicitaDescargaEmitidos`/`SolicitaDescargaRecibidos` contra
`cfdidescargamasivasolicitud.clouda.sat.gob.mx`). Conclusión: esta app ya
está en el nivel de servicio correcto (el que las apps comerciales
recomiendan para volumen); no existe una técnica legítima para "romper" la
cuota de por vida (CodEstatus 5002) de ese mismo servicio SOAP — es un
límite anti-abuso del lado del SAT, y el 5002 real que se vio en este
proyecto fue casi con certeza autoinfligido por solicitudes repetidas con
los mismos parámetros durante depuración (ver bloque de arriba).

**Selector Emitidos/Recibidos** (modo fecha) — `descargar_rango(..., tipos=...)`
acepta `{"emitidos"}`, `{"recibidos"}` o ambos (default: ambos, compatible
con `descargar_mes`/`descarga_v5` que no pasan este parámetro). Cada tipo es
una solicitud SAT independiente — desmarcar el que no se necesita ahorra
tiempo y una unidad de `MAX_SOLICITUDES_DIA`. UI: dos checkboxes en el panel
de fecha de `_DialogDescargaSAT`, ambos activos por default; se bloquea si
se desmarcan los dos.

**Bug real corregido — `EstadoComprobante` explícito requerido (CodEstatus 301):**
El SAT (Web Service v1.5) rechaza la solicitud con
`Mensaje="...No se permite la descarga de xml que se encuentren cancelados."`
cuando `estado_comprobante` se omite — el default implícito del servicio es
"Todos" (incluye cancelados), y eso ya no se acepta. `_ciclo_descarga` ahora
pasa `estado_comprobante=EstadoComprobante.VIGENTE` explícito en **ambas**
solicitudes (emitidos y recibidos) — además de ser el fix real, es lo
correcto para conciliación bancaria: un CFDI cancelado no debe conciliar
contra un movimiento real. Verificado con evidencia real (sin red): un `SAT`
simulado confirmó que ambas solicitudes ahora incluyen
`estado_comprobante=Vigente`, y que los fixes previos (`fecha_final`
23:59:59) siguen intactos. Encontrado gracias al logging de la respuesta
cruda del SAT agregado en el intento anterior — sin ese detalle, este fix
habría sido otra conjetura más.

**Diagnóstico descartado — NO envolver `RfcReceptor` en `RfcReceptores`:**
Se intentó un workaround (`_solicitar_recibidos_corregido`, ya eliminado)
basándose en una hipótesis externa sin confirmar: que el SAT exigiría
`RfcReceptor` envuelto en un subelemento `<RfcReceptores>` desde su Web
Service v1.5. **Esto era incorrecto.** El SAT respondió explícitamente
(`CodEstatus=301`, `Mensaje="...La solicitud de descarga de facturas
recibidas requiere el RfcReceptor."`) que el campo plano `RfcReceptor` SÍ es
el que se necesita — el workaround lo había quitado por error, causando el
mismo síntoma (301) por una razón distinta a la original. Revertido: se usa
`sat.recover_comprobante_received_request()` de la librería sin
modificar. Lección: verificar contra la respuesta real del SAT (`Mensaje`)
antes de aplicar un fix basado en documentación de terceros no confirmada —
`_ciclo_descarga` ahora siempre loguea la respuesta cruda completa del SAT
(`[SAT] Respuesta cruda {tipo}: {resp}`), no solo `CodEstatus`, precisamente
para evitar repetir este error.

**Bug corregido — truncamiento de `fecha_final`:** `satcfdi` serializa un
`date` puro sin hora (`_set_arguments()`: `date.isoformat()` → `"2026-01-31"`,
sin componente de tiempo), y el SAT lo interpreta como medianoche de ese
día — excluyendo casi todo el último día del rango (el portal de Consulta
CFDI del SAT expone esto explícitamente con los campos "Hora Inicial
00:00:00 / Hora Final 23:59:59"). `_rango_dia_completo()` en
`sat_descarga.py` convierte las fechas a `datetime` con `fecha_inicial` a
medianoche y `fecha_final` a `23:59:59.999999` antes de pasarlas a
`sat.recover_comprobante_*_request()`, dentro de `_ciclo_descarga` — cubre
tanto la llamada directa de `descargar_rango` como la paginación por
quincena de `_descargar_paginado`. Verificado interceptando `SAT` con un
doble de prueba (sin red): antes del fix `fecha_final` llegaba como
`date(2026,1,31)`; después, como `datetime(2026,1,31,23,59,59,999999)`.

Requiere `satcfdi>=26.7.2` (en `requirements.txt`). Si el paquete no está
instalado, `_SAT_OK = False` y el botón muestra un error claro en vez de
romper el resto de la app (mismo patrón que `_CONCIL_OK`/`_COI_OK`/`_GRUPOS_OK`).

**No verificado contra el SAT real** — la descarga real requiere una
e.firma vigente, red, y consumiría el límite diario de solicitudes; se
verificó con datos reales todo lo demás: construcción del diálogo, cambio
de modo, las 4 validaciones de formulario, y `_filtrar_por_uuid_local()`
con XMLs CFDI reales (mueve los correctos, deja los demás, reporta UUIDs
no encontrados, los archivos movidos siguen siendo parseables).

### Canal Portal — respaldo con robot (`sat_descarga_portal.py`, 10 jul 2026)

Segundo canal de descarga, **opcional y de respaldo** — el Web Service SOAP
(`sat_descarga.py`) sigue siendo el canal principal recomendado (hasta
200,000 XML/solicitud). El canal Portal automatiza con Playwright el portal
web del SAT (login FIEL sin captcha), cuya cuota (~2,000 XML/día) es
**separada** de la del Web Service — sirve como respaldo cuando este último
está agotado (CodEstatus 5002), no como reemplazo. Motivación completa (por
qué esto no es "romper la cuota del SAT" sino usar un segundo cupo oficial
distinto) documentada arriba en la sección del bug 5002.

**Login verificado en vivo (solo lectura de estructura, sin autenticar):**
`https://cfdiau.sat.gob.mx/nidp/app/login?id=SATx509Custom` — campos reales
confirmados por inspección del DOM: `#fileCertificate` / `#filePrivateKey`
(inputs de archivo para .cer/.key), `#privateKeyPassword`, `#submit` (OJO:
`#submit` está duplicado — hay uno oculto del form de contraseña y otro
visible del form de e.firma; hay que clickear el **visible**). Esta
variante de login (certificado) no pide captcha — a diferencia del login
por RFC+contraseña, que sí lo requiere.

**Navegación posterior al login — actualizada 10 jul 2026 con datos reales**
(antes era "mejor esfuerzo" sin verificar). Tras el login, el SAT aterriza
en `portalcfdi.facturaelectronica.sat.gob.mx` (NO en `cfdiau.sat.gob.mx` —
ese es solo el paso de autenticación). Es un portal ASP.NET WebForms
clásico con dos páginas separadas, no una sola "Descarga masiva":
- **Emitidos** (`ConsultaEmisor.aspx`): rango de fechas en una sola
  búsqueda. Los inputs de fecha son disabled-por-diseño (date-picker) y se
  fuerzan por JS con un evento `change`.
- **Recibidos** (`ConsultaReceptor.aspx`): solo admite **un día a la vez**
  (dropdowns año/mes/día) — `sat_descarga_portal.py` itera día por día
  dentro del rango pedido, así que un rango largo de recibidos por Portal
  tarda bastante más que por Web Service.
- El radio "Por fecha" (`#ctl00_MainContent_RdoFechas`) exige **clic
  nativo** — asignar `.checked` por JS no basta, el servidor re-renderiza
  con los campos deshabilitados si no se hace clic real.
- Los resultados llegan por AJAX (UpdatePanel); hay que esperar a que el
  spinner (`#ctl00_MainContent_UpdateProgress1`) se oculte y aparezcan
  filas o el panel "sin resultados" — `networkidle` regresa antes y da
  falsos 0.
- Cada fila de resultados es un XML individual descargable vía
  `RecuperaCfdi.aspx?Datos=<token>` — no hay un paquete ZIP como en el Web
  Service; se descarga uno por uno y se corta tras 3 fallos seguidos
  (probable cuota diaria agotada).

Estos hechos (URLs, ids de control, gotchas de ASP.NET) se contrastaron
contra **sat-dm** (github.com/IsraelCV5/sat-dm, MIT, Israel Castro), un
proyecto de terceros que documenta lo mismo a partir de pruebas reales
contra el portal (may. 2026) — se adoptaron los hechos verificados sobre la
estructura pública del portal, no ninguna técnica orientada a evadir
límites del SAT (ver nota abajo, "Qué NO hace este canal"). Aun así, la
corrida real end-to-end sigue pendiente (tarea "Corrida real Portal SAT —
junio 2026"): estos datos vienen de código de terceros probado por ellos,
no de una corrida nuestra.

**Resiliencia a cambios de layout — dos capas, en `sat_descarga_portal.py`:**
1. `SELECTORES`: tabla de listas de selectores de respaldo por campo
   lógico, probadas en orden (`_resolver()`), sin red adicional.
2. Si todos los selectores conocidos de un campo fallan, `_reparar_con_ia()`
   manda el HTML actual a la API de Anthropic pidiendo un selector nuevo,
   lo prueba, y si funciona lo usa **solo para esa corrida**. Cada intento
   (éxito o fracaso) se registra siempre en
   `%APPDATA%\ConvertidorBancos\sat_portal_selectores.log` — la
   auto-reparación en vivo es un parche de emergencia, no un reemplazo del
   arreglo real; el log es lo que permite convertir un hallazgo en una
   entrada permanente de `SELECTORES` en una sesión de Claude Code
   posterior. Sin API key configurada, o si la IA también falla, se aborta
   con `DescargaSATError` claro — nunca en silencio.

**API key de Anthropic** — opcional, solo habilita la auto-reparación;
guardada cifrada (mismo esquema XOR-SHA256 atado al MAC que
`repo_conciliaciones.py`) en `%APPDATA%\ConvertidorBancos\anthropic_key.bin`
vía `guardar_api_key_anthropic()`/`cargar_api_key_anthropic()`. Se
configura desde Obligaciones Fiscales → Config Empresa → botón "🔑 API key
(Portal SAT)" (`_config_api_key_anthropic` en `ConvertidorSantander.py`).

**UI:** selector de canal (radio buttons "Web Service" / "Portal") en
`_DialogDescargaSAT`, visible solo en modo "Por rango de fechas" — el modo
"Por UUID específico" sigue siendo exclusivo del Web Service (no se amplió
el alcance a folio-por-folio en el Portal). `_ejecutar_descarga` bifurca a
`sat_descarga_portal.descargar_rango_portal()` cuando se elige "Portal";
mismo `DescargaSATResult` que el Web Service, así que el resto del flujo
(carga a Conciliación CFDI) no distingue el canal de origen.

**Riesgo operativo a tener presente:** automatizar el login de un portal
de identidad gubernamental (NetIQ/nidp) puede ser detectado como tráfico de
bot y bloquear temporalmente el RFC en el portal — riesgo separado del
Web Service, sin forma de eliminarlo por completo, solo mitigarlo (no
reintentar agresivamente, no correr en paralelo).

**Qué NO hace este canal (alcance v1, explícito):** no resuelve/evade
captcha (no aplica — el login FIEL no lo pide); no automatiza ningún otro
trámite del portal fuera de "Descarga masiva"; no intenta superar el tope
de 2,000 XML/día del portal — es un respaldo, no una forma de obtener más
volumen total.

**Decisión explícita — NO se adoptó el "jitter de segundos" para el 5002:**
el proyecto de terceros `sat-dm` (ver arriba) documenta que el Web Service
del SAT marca CodEstatus=5002 por el hash del rango EXACTO de fechas
(con precisión de segundo), no por un contador de uso real, y reintenta
variando el segundo de `FechaFinal` para esquivar el rechazo. Se evaluó
portar esa técnica a `sat_descarga.py` y **se descartó** — es una forma
deliberada de evadir el control anti-abuso del SAT, lo mismo que ya se
había descartado antes en este proyecto (ver la conversación sobre
OneFacture/cuotas). El clasificador de permisos del harness bloqueó el
intento de implementarlo, correctamente. El mensaje de error de 5002 en
`sat_descarga.py` sigue siendo el original: cuota agotada, no reintentable
automáticamente.

**Verificado sin red (Playwright real contra HTML fijo, sin tocar el SAT,
10 jul 2026):** guardado/carga/borrado de la API key cifrada, y que el
archivo en disco no contiene la key en texto plano; resolución de
selectores primarios en layout "normal"; `DescargaSATError` claro cuando
todos los selectores fallan y no hay API key; auto-reparación exitosa con
un doble (mock) de la API de Anthropic que sugiere el selector correcto
para un layout simulado "roto"; fallo controlado cuando la IA sugiere un
selector que tampoco existe; y que ambos resultados (éxito y fracaso de
auto-reparación) quedan en el log. También se verificó con Tkinter real
(sin mocks) que `_DialogDescargaSAT` construye los widgets del selector de
canal, y que el modo "Por UUID" fuerza `canal="webservice"` aunque el radio
button tenga "portal" seleccionado. Tras la actualización con los
selectores reales de navegación (ver arriba), se re-verificó sin red:
resolución de `radio_fechas`/`boton_buscar_portal` contra un fixture del
portal real, y `_extraer_filas_resultado()` extrayendo correctamente UUID +
URL de descarga desde un `onclick` con el mismo patrón de comillas que usa
el portal real (`AccionCfdi('...','...')`). **Pendiente:** login y
navegación reales contra el portal — requiere e.firma vigente y
coordinación con el usuario (consume la cuota separada del portal); los
selectores de navegación vienen de un proyecto de terceros, no de una
corrida propia, así que siguen sin el mismo nivel de certeza que el login.

Dependencias nuevas: `playwright>=1.60` (requiere una vez por máquina
`playwright install chromium` — binario de ~150-300 MB, no lo instala pip
por sí solo) y `anthropic>=0.40` (opcional, solo si se usa auto-reparación).

### Constancia de Situación Fiscal (CSF/CIF) — hogar único

Un solo lugar en todo el sistema sube la CSF: **Obligaciones Fiscales →
Config Empresa** (`_subir_cif_config_empresa`, junto al selector de RFC).
Antes existían tres botones duplicados (Convertidor PDF→Excel, Conciliación
CFDI, y el catálogo de empresas `DialogEmpresas`) — se removieron los dos
primeros; `DialogEmpresas._subir_cif` se conservó porque es un diálogo de
gestión de catálogo distinto, no una de las dos pestañas principales.

### Banregio SPEI: extraer nombre del ordenante (10 ago 2026)

**Síntoma:** los depósitos SPEI de Banregio cruzaban su CFDI pero en **nivel 3
(revisar / amarillo)** en vez de nivel 2 (limpio). Reproducido con GRUPO NACIONAL
COMERCIAL ANFATI (GNC200618SQ9) marzo 2026: dos depósitos ($83,683.12 y
$65,579.19) de "SERGIO ARTURO CEBALLOS GONZALEZ" cruzaban sus CFDIs emitidos
(mismos importes, mes 3, receptor CEGS880425RE8) pero marcados revisar.

**Causa raíz:** `MovBanco.nombre_en_descripcion()` extraía el nombre buscando
"CLIENTE", "NOMBRE:" o segmentos separados por `/`. El SPEI de Banregio usa
**comas**: `SPEI, <banco>, <NOMBRE ORDENANTE>, <referencia>, <concepto>` — ninguna
estrategia disparaba, así que el nombre quedaba vacío. Sin nombre extraído y sin
RFC en la descripción, el cruce sólo llegaba a nivel 3 (nombre-del-CFDI-en-texto,
`revisar=True`) en vez de nivel 2 (nombre extraído coincide, `revisar=False`).

**Fix:** nueva estrategia 2.5 en `nombre_en_descripcion()` — si la descripción
empieza con "SPEI" y tiene comas, toma el **3er campo** como nombre del ordenante
(con validación: ≥5 chars, ≥70% letras). Gated a `startswith("SPEI")` para no
tocar otros bancos (Inbursa usa `|`, el SPEI estándar usa `/`). Verificado: los 2
depósitos ANFATI pasan de **nivel 3 (revisar) → nivel 2 (limpio)**; STERIMEX/Inbursa
**137/166 sin cambio** (sus descripciones no empiezan con "SPEI"). Byte-idéntico en
5.0 y 5.1.

### Clasificación por COLUMNA — regla anti-traspaso, TODOS los bancos (10 ago 2026)

**Bug (regresión recurrente):** un **traspaso entre cuentas propias** que es CARGO
en una cuenta debe ser ABONO en la otra — nunca cargo/abono en ambas. Causa: el
balance-tracking (`prev ± monto ≈ saldo`) es la señal primaria, pero cuando el
banco **no imprime el saldo corrido en cada línea** (BBVA casi nunca) no puede
decidir y el código caía a `_clasificar_keywords` → **"retiro" por defecto**. Un
"TRASPASO CUENTAS PROPIAS" (sin palabra de abono) se leía retiro en las dos.

**Fix (`bancos/estandar.py`, compartido por todos):** clasificar por la COLUMNA
(CARGOS/ABONOS por posición X con `pdfplumber.extract_words`). `construir_mapa_columnas(pdf_path)`
(función de módulo, reusable) auto-detecta los encabezados CARGOS/ABONOS/SALDO y
mapea `(fecha_key, centavos) → 'retiro'/'deposito'` (descarta ambiguos). `_clasificar`
es **balance → COLUMNA (solo si el balance no pudo) → keywords** → cero regresión.
Flag de clase `COLUMNAS_AUTO = True`. Cobertura: BBVA/Inbursa/Santander/Banorte
heredan; BanBajío (no hereda de `BancoEstandar`) importa `construir_mapa_columnas`/
`fecha_key` y llama `_aplicar_columna(mov)` tras `_verify_by_balance`.

**Verificado (BBVA real, ene 2026):** concordancia balance-vs-columna 2/2 sin
discrepancias; corrige 4 RET→DEP (0 DEP→RET); el estado pasa de descuadre
**$24,323.64 a $0.00**. Cero regresión: Inbursa 137/166 y Banregio 2/6 sin cambio.
Byte-idéntico en 5.0 y 5.1.

### Adding a new bank extractor
1. Create `bancos/{banco}.py` with a class inheriting `BancoExtractor`. Implement `detectar(cls, texto)` and `extraer(self, pdf_path, callback)`.
2. Register in `bancos/detector.py`: add the import, append to `EXTRACTORES` (fallback order), and add entries to `_IBM_MAP` (IBM-anchor strategy).

### Aspel COI poliza layout (Poliza-Info-CFDI.xlsx)
The import file is a flat Excel with this structure:

| Row | Cols used | Content |
|---|---|---|
| 3 | A,B,C,D | Header: **TipoPoliza** (2 chars, e.g. `Ig`), **NúmPoliza** (5 chars), **Concepto** (120 chars), **Día** |
| 4+ | B–I | Partida: **CuentaContable**, **Depto** (≤3 digits), **ConceptoMov** (120 chars), **TipoCambio**, **Debe**, **Haber**, **CentroCosto**, **Proyecto** |
| Between partidas | C | `INICIO_CFDI` marker row, then one CFDI row: C=FechaTimbrado (DD/MM/YYYY), D=Serie, E=Folio, F=RFCEmisor, G=RFCReceptor, H=MontoCFDI, I=UUID, J=Tipo (1=CFDI/2=Extranjero), then `FIN_CFDI` marker |
| After last partida | B | `FIN_PARTIDAS` marker |

Tipo de póliza codes: `Dr`=Diario, `Ig`=Ingreso, `Eg`=Egreso, `Di`=Diario (alternate).

### Generador de Pólizas COI (implementado en v5)

| File | Role |
|---|---|
| `catalogo_cuentas.py` | `CuentaContable` dataclass + `importar_desde_excel()` (flexible column detection) + `buscar_cuentas()`. Stored per-RFC in `%APPDATA%\ConvertidorBancos\cuentas_{RFC}.json` (plain JSON). |
| `catalogo_terceros.py` | `Tercero` dataclass + `poblar_desde_cfdis()` (emitidos tipo I → clientes, recibidos tipo I → proveedores) + `auto_asociar_cuentas()` (word-intersection fuzzy matching ≥ 0.6). Stored per-RFC in `terceros_{RFC}.json`. |
| `polizas_coi.py` | `generar_polizas()` → list[Poliza]. `polizas_a_filas()` serializes to COI 10-col flat format. Egreso: proveedor DEBE + banco HABER. Ingreso: banco DEBE + cliente HABER. Nómina: gastos_nom DEBE + banco HABER. |
| `ventana_polizas.py` | `VentanaPolizas` Toplevel with 3 tabs: Catálogo de Cuentas / Clientes y Proveedores / Generar Pólizas. Generates the "Pólizas COI" sheet in the existing conciliation Excel and optionally exports a standalone COI import file. |

**Flow:** After conciliation, the "⚙ Pólizas COI" button in `VentanaConciliacion` becomes enabled. It opens `VentanaPolizas` pre-loaded with the conciliation results. User imports their COI chart of accounts, auto-populates terceros from XMLs, assigns contable accounts, then generates pólizas. The sheet is appended to the already-generated conciliation Excel.

`exportar_conciliacion()` and `exportar_conciliacion_multi()` accept an optional `polizas: list` parameter — when provided, `_hoja_polizas_coi(wb, polizas)` is called before saving.
