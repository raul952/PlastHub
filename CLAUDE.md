# PlastHub / Factory "Amaru I" — Plastics Consolidation Hub (Lima Norte, Perú)

Working folder for the **PlasHub** Stage-1 plastics consolidation hub (company **Neo Amaru Global S.A.C.**):
warehouse layout, capacity modelling, and supporting project documents.

## What this is

A consolidation hub that receives post-industrial recyclable plastics by truck, classifies/weighs/stores
them by polymer, and consolidates **~20 t per material** before dispatching a 40-ft container.
**Stage 1 only** — no shredding, washing, or pelletizing. Flow: RECIBE → CLASIFICA → PESA → ALMACENA →
CONSOLIDA 20 t → DESPACHA contenedor 40′.

## Validated operation facts (from Project Profile v2 + slide deck)

- **Target volume:** **80 t/mo** base · 100 t/mo optimistic · **break-even 61 t/mo**.
- **Inbound:** ~10 t/week · 1–2 trucks/week · 8–12 bolsones/truck.
- **Bolsones:** 1 m³, ~1 t, stacked **2-high** (sporadic zones: optional 3-high racking).
- **Equipment:** **3-t counterbalance forklift is mandatory** (1-t bolsones can't be handled by pallet jacks); 1.5×1.5 m platform scale. Min aisle 3.5 m.
- **Economics:** weighted margin **$128.7/t** (~49%).
- **Site CONFIRMED (croquis):** Cooperativa Agraria Parque Porcino, Ventanilla (Callao). **Area 2 = 1,800 m² = 50 m (frente, sobre la rampa) × 36 m (fondo) × ~6 m height.** Hosts **Stage 1 (distribution, front) + Stage 2 (processing, back) + offices (2 floors ~70 m²)**; the ramp/dock (~600 m²) is the front strip of Area 2. (Lot total ~4,200 m² = Area 2 1,800 + Area 1 1,248 + Area 3 1,152, the latter two already rented to others.) LOI: Phase 1 rents ~600 m² of Area 2; Phase 2 the full 1,800.
- **Lease structure CONFIRMED (LOI):** Phase 1 (M1–M6): M1 grace ($0) · M2–M6 **$2,000/mo** (600 m² × $3 + $200 dock). Phase 2 (M7+, mandatory): **$4,200/mo** consolidated (1,800 m²). Upfront at signing: **$6,000** (2 mo deposit + 1 mo advance, "2×1"); deposit ~$4,000 recoverable. Contract: **24 months** min.
- **7 material types — taxonomy to confirm:** Excel = 5 frequent (PET, HDPE, PP, PE Film, Polvillo de Flake) + 2 sporadic (ABS/PS, Nylon/PA); slide deck = 4 frequent + 3 sporadic (incl. "Mix especiales").

### Open flags (carry forward)
- ⚠ **IGV on rent** — landlord is *persona natural*; confirm whether 18% applies (Phase 1 = $2,360/mo · Phase 2 = $4,956/mo if it does).
- ⚠ **LOI sent ~Jun 2026, awaiting countersignature** from Jaime Ernesto Gracia Arevalo + Ing. Armando Gómez Milla. Verify both are registered *titulares* (check SUNARP) before signing the final contract.
- 🔴 **Equipment gap** — CAPEX budgets pallet jacks, not the required forklift. Add rent (~$800–1,500/mo) or used unit (~$10–20K).
- 🔴 **Cash-flow sheet (3b) bug** — supplier payments hard-coded at −$21,200 all months (should be ~−$11,360 from M2); shows false insolvency. Fix before fundraising.
- ⚠ **EC-RS (MINAM/SIGERSOL)** registration is the critical path — 60–90 days, start immediately. RMT tax regime. Notarized pacto de socios + transfer-pricing docs (SUNAT related-party risk: Alex is supplier+client+partner).

## Files

| File | Purpose |
|---|---|
| `index.html` | **Landing page / centro de documentos** — hub con identidad Hyperion: herramientas, documentos (con miniaturas + **lector in-app** modal para leer antes de descargar), conceptos, audit, **manual**, stats y próximos pasos. Punto de entrada. |
| `PlasHub_Manual_de_Uso.docx` | **Manual de uso (Word, separado de la app)** — documenta CADA parámetro editable de las 4 pestañas (qué es / cómo se usa / a qué afecta), un **guion para la reunión de socios**, cómo leer los resultados (GO/NO-GO, break-even de caja, pozo, matriz), glosario y banderas de honestidad. Generado con `Context/_build_manual.py`. **Las explicaciones coinciden con los íconos ℹ inline** de cada herramienta (`.ttip` con `data-t`). *(La antigua `manual.html` in-app fue eliminada.)* |
| `plashub-etapa2.html` | **Etapa 2 (procesamiento) — interactivo.** Distribución/capacidad de la línea de proceso + modelo financiero de *timing*: elige el mes de lanzamiento (4/6/8/12 o cualquiera) y ve cómo cambian flujo de caja, capital requerido y retorno a 24 meses. Editable + autosave + export/import. Supuestos gruesos a validar. |
| `conceptos-utiles.html` | **Guía del Consultor** — conceptos de finanzas/estrategia/operaciones con las cifras del proyecto + el porqué de cada decisión y supuesto. |
| `tareas-maestras.js` | **Lista maestra de tareas (única fuente de verdad)** — 30 tareas {etapa GEN/E1/E2, tipo, prioridad, done}. API de estado en localStorage (`plashub_tareas_v1`). La cargan `plashub-tareas.html` y la sección Permisos del Plan de Negocio. |
| `plashub-tareas.html` | **Tareas y Permisos — vista maestra.** Todas las tareas/permisos clasificados por General/E1/E2, con filtros (etapa/tipo/prioridad/estado), progreso y checklist. Sincronizado (mismo estado) con el Plan de Negocio. |
| `plashub-plan-accion.html` | *(redirige)* → `plashub-tareas.html` (la lista maestra reemplazó al antiguo plan de acción). |
| `PlasHub_Project_Profile_v4.xlsx` | Modelo financiero **v4** (= v3 corregido, cash-flow arreglado) — versión a distribuir. |
| `audit.html` | **Audit del proyecto** — auditoría de las apps + auditoría franca (McKinsey) del plan de negocio/operaciones, por severidad, con recomendaciones y veredicto GO condicional. |
| `doc-previews.js` | Contenido HTML fiel de cada documento (generado desde los .docx/.xlsx) para el lector in-app de la landing. Regenerar con `Context/_build_readers.py`. |
| `plashub-app.html` | **App unificada — 4 pestañas:** `#d1` Etapa 1 · Distribución del Almacén · `#d2` Etapa 2 · Distribución de Planta · `#pn` Plan de Negocio (E1+E2) · `#tp` Tareas & Permisos (lista maestra). Cada pestaña embebe su módulo aislado (iframe), conservando edición/autosave. **Entrada principal.** |
| `plashub-distribucion-e2.html` | *(módulo, pestaña #d2)* **Etapa 2 — Distribución y Capacidad de Planta.** Plano de la planta de procesamiento (flujo "I", 2 líneas + buffer Escurrido). **Sizing FUNCIONAL**: zonas de soporte (oficina/lab QA/PTAR/welfare) por su driver (staff/familias/caudal) — no escalan con la nave; buffers MP/PT por días×volumen objetivo; despacho por contenedor. Override manual por zona + columna Driver + aviso de fiteo. **Edificio compartido 36×50 m (preset por defecto)**: control "Profundidad E1" reserva el frente para la Etapa 1 (zona fantasma, sigue operando); la Etapa 2 ocupa el fondo; línea divisoria E1/E2 en el plano. Editable + autosave + Export/Import. |
| `plashub-etapa2.html` | *(superado)* versión previa Etapa 2 (layout esquemático + finanzas). Su layout pasó a `plashub-distribucion-e2.html` y sus finanzas a la sección 10 del Plan de Negocio. Ya no se enlaza desde el app. |
| `plashub-distribucion.html` | *(módulo, embebido en la app)* **Distribución + modelo de capacidad (validado v2, español)**. Sliders + **campos editables** (por material: nombre/color/frecuencia y **bols/sem** = bolsones/semana, con añadir/eliminar; densidades, objetivos, **dimensiones físicas del bolsón** bolsonW/D/H, minZoneDepth, días hábiles, geometría) → plano, tabla de zonas y capacidad en vivo. **Almacén por polímero dimensionado por bolsones reales** (stock pico = bols/sem × días hold ÷ 7; profundidad mínima física ≥2 m; aviso si la nave queda justa). Zonas de soporte por **driver funcional** (headcount/equipo) + override manual (↺). **Tablero operativo** (posiciones, ritmo/día del montacargas, ¿nave suficiente?, uso del almacén). **Total dinámico bols/t vs objetivo.** Preset **36×50 m** + panel "Utilización del terreno" con split E1/E2 en el mismo edificio. Autoguardado + Restablecer + Export/Import JSON. |
| `plashub-plan-negocio.html` | **Plan de negocio interactivo (E1 + E2).** Sub-nav (Plan E1 · Plan E2 · Permisos). **Panel de Decisión GO/NO-GO** arriba del todo (7 condiciones auto 🤖 + manual ✋ → veredicto). E1: precios/mix, CAPEX A/B y OPEX A/B itemizados, dashboard. **Modelo financiero realista**: timing de **aportes de socios** (Alex M0 / Raúl M1, editable), **Mes 0 de setup** (CAPEX+OPEX antes de operar), **lag de pago a proveedor** y **estrés de cobro** (retraso de Alex en meses dados), **prepago de Alex** como palanca de capital. **Slider de volumen continuo** (20–220 t) con mini-dashboard en vivo + 3 escenarios como atajos. **Matriz de sensibilidad** precio×volumen. **P&L con vista Ejecutiva (M0/M1/M3/M6/M12/Total) o Detallada (12 meses)**. **Break-even de CAJA** (vol. mín. sin caja negativa; "no alcanzable" si el capital no basta — se vuelve alcanzable al subir capital). **E2: CAPEX y OPEX itemizados editables** + **margen sincronizado desde E1** (botón ↺ Sync E1). **Checklist de permisos por etapa** (sincronizado con Tareas). **§4 con layout de dos columnas: controles sticky a la izquierda + dashboards a la derecha** (mover el slider actualiza los resultados sin hacer scroll). **Todo editable** + **Autoguardado + Restablecer + Exportar/Importar JSON.** Cifras estimadas (≈) y **gate de "precios verificados"**. Imprimible A4. |
| `PlasHub_Business_Plan.docx` | Word version of the business plan (investor/lender pack, Spanish). |
| `PlasHub_Project_Profile_v3_corrected.xlsx` | Profile v2 **with the cash-flow bug fixed** (row 14 now uses each month's own price/flete). Caja M12 −$34,718 (was −$142,958). Open in Excel to recalc. Original v2 preserved in `Context/`. |
| `PlasHub_Permits_Legal_Tracker.xlsx` | Critical-path tracker: EC-RS, RMT, pacto de socios, lease, ITSE/INDECI, EIA, equipment, capital. Owners, durations, dependencies, status. |
| `MOU_Suministro_Alex_PlasHub_BORRADOR.docx` | **Draft** minimum-volume supply + commercial + governance MOU for Alex (Spanish). For lawyer/notary review — not legal advice. |
| `Amaru-I-Project-Charter-v2.docx` | **Corrected Project Charter** with validated Stage-1 numbers + a v1→v2 correction log. (Original v1 is in `Context/`.) |
| `Context/` | Source documents: `PlasHub_Project_Profile_v2.xlsx` (authoritative), `Requerimientos_Almacen_Peru.pptx`, `Factory "Amaru I" - Proyecto Piloto.docx`, `Amaru-I-Project-Charter.docx` (**v1, outdated numbers**). Plus `_extract.py` / `_build_charter.py` helpers. |
| `.claude/launch.json` | `plashub-static` — serves this folder via `npx http-server`. |

## Key model finding

Capacity = **min(storage-buffer limit, handling limit)** vs the **80 t/mo** target.
- At **~900 m²** (lease target): **handling-bound** at ~528 t/mo (1 forklift, 1 shift) — i.e. heavily
  over-provisioned for throughput by design. The extra m² buy staging comfort, buffer resilience and
  Stage-2 room, **not** more tonnes. More throughput = 2nd forklift / 2nd shift, not more floor.
- At **~350 m²**: too small for the validated spec (150 m² staging + 7 streams) — storage-bound below break-even.
- **Pallet-jack mode** collapses capacity (~25 t/mo) and storage to single-layer — demonstrates the equipment gap.

## Stage-1 financial headlines (from the business-plan model)

- **Blended margin** ~49% (~$127/t): weighted buy $132/t, sell $259/t. Flake anchor $150/$300 (50%).
- **Break-even volume:** ~82 t/mo (Lean A) · ~144 t/mo (Full B).
- **Base 100 t/mo (Lean A):** steady EBITDA ≈ +$2.0K/mo, op. break-even M3, ROI ~69%/yr on CAPEX, CAPEX payback > M12.
- **Conservative 40 t/mo:** loss-making (≈ −$4.7K/mo) — below break-even; a warning case, not viable standalone.
- **Optimistic 200 t/mo (Full B):** EBITDA ≈ +$6.3K/mo, payback ~M10, ROI ~142%.
- 🔴 **The real gap is working capital, not CAPEX:** buy-now/collect-later drives the bank negative in early months even in the profitable Base case. **Confirmed capital now $27K (Alex $16K + Raúl $11K, both at start).** Close any remaining gap via deposit negotiation + Alex prepay + socio bridge.
- **Pricing honesty:** global "recycled HDPE $930–1,310/MT" = washed/pelletized regrind (Stage-2 product), NOT the Stage-1 baled-feedstock/consolidated-lot tier modelled here.

## Preview locally

`npx http-server . -p 3017` then open `http://localhost:3017/plashub-layout.html` — or just open the HTML file directly.
(This session is rooted in the sibling Hyperion folder; for a clean PlastHub-rooted session use that as the working dir.)

## Stages (per Project Profile)
- **Stage 1** (this project): consolidation, ~80 t/mo, ~900 m².
- **Stage 2:** basic processing (compacting/packing), 400–500+ t/mo.
- **Stage 3:** regional replication (5+ countries), ~6,000 m², advanced processing.

## Brand colours (Hyperion)
```
dark #252A31 · orange #DF6901 · teal #15616D · light #F2F2F2 · muted #CACEDA
```
