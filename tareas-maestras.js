/* ============================================================
   PlasHub — LISTA MAESTRA DE TAREAS (única fuente de verdad)
   Usada por: plashub-tareas.html (vista completa) y la sección
   "Permisos" del Plan de Negocio. Estado compartido en localStorage.
   ============================================================ */
window.PLASHUB_TASKKEY = 'plashub_tareas_v1';

// etapa: GEN | E1 | E2   ·   tipo: Permiso | Operación | Financiero | Layout
// crit: true=🔴 crítica / false=🟡 importante   ·   done: ya implementado en la app
window.PLASHUB_TASKS = [
  /* ---------- GENERAL (empresa / financiero / legal de base) ---------- */
  {id:'gen-vol-alex', etapa:'GEN', tipo:'Financiero', crit:true, resp:'Raúl + Alex', urg:'Antes de firmar nave',
    title:'Compromiso de volumen de Alex por escrito', detail:'Carta/contrato de volumen mínimo (≥82 t/mes, 12 meses). Sin esto el modelo financiero es una hipótesis.'},
  {id:'gen-gap', etapa:'GEN', tipo:'Financiero', crit:true, resp:'CFO / Alex', urg:'Antes de operar',
    title:'Cerrar la brecha de capital (Etapa 1 ~$45–50K)', detail:'Incluye el pozo de capital de trabajo, no solo el CAPEX. Negociar garantía 1 mes + prepago de Alex + puente entre socios.'},
  {id:'gen-pacto', etapa:'GEN', tipo:'Permiso', crit:true, resp:'Abogado + socios', urg:'Antes de operar',
    title:'Pacto de socios notarizado', detail:'Equity 60/40, política de dividendos, precio de transferencia, resolución de conflictos. Firmar ante SUNARP antes de la primera operación.'},
  {id:'gen-rmt', etapa:'GEN', tipo:'Permiso', crit:true, resp:'Contador', urg:'Antes de la 1ª venta',
    title:'Inscripción / cambio a Régimen MYPE Tributario (RMT)', detail:'Ventas año 1 > límite RER → RMT obligatorio. Verificar en SUNAT antes de facturar.'},
  {id:'gen-transfer', etapa:'GEN', tipo:'Financiero', crit:true, resp:'Contador (partes vinculadas)', urg:'Antes de operar',
    title:'Documentación de precios de transferencia (Alex)', detail:'Alex es proveedor+cliente+socio. Documentar precios de mercado verificables con facturas para SUNAT.'},
  {id:'gen-forklift', etapa:'GEN', tipo:'Financiero', crit:true, resp:'CFO / Alex', urg:'Antes de fundraising',
    title:'Decidir y cotizar montacarga 3t (renta vs. compra)', detail:'Renta $800–1,500/mo (OPEX) o usado $10–20K (CAPEX). Incorporar al modelo financiero y recalcular antes de presentar a inversores.'},
  {id:'gen-cashflow', etapa:'GEN', tipo:'Financiero', crit:true, done:true, resp:'CFO', urg:'Hecho',
    title:'Corregir bug del flujo de caja (hoja 3b)', detail:'Pagos a proveedor fijados en −$21,200 → fórmula por mes. Caja M12 pasa de −$142,958 a −$34,718. Corregido en Profile v3/v4.'},
  {id:'gen-seguro', etapa:'GEN', tipo:'Permiso', crit:false, resp:'Asesor', urg:'Antes de operar',
    title:'Seguro de responsabilidad civil / contenido', detail:'Cobertura integral (incendio, contenido). Carga de fuego alta por plásticos — confirmar cláusulas.'},

  /* ---------- ETAPA 1 — Consolidación / almacén ---------- */
  {id:'e1-ecrs', etapa:'E1', tipo:'Permiso', crit:true, resp:'Abogado ambiental', urg:'Esta semana · 60–90 d',
    title:'EC-RS — registro de comercializadora de residuos (MINAM/SIGERSOL)', detail:'RUTA CRÍTICA. Sin este registro la operación es ilegal. Abrir cuenta SIGERSOL + Plan de Manejo de Residuos. Iniciar el día 1 y paralelizar con la búsqueda de nave.'},
  {id:'e1-itse', etapa:'E1', tipo:'Permiso', crit:true, resp:'Asesor + Jefe de Planta', urg:'Requisito para operar',
    title:'Certificado ITSE (INDECI)', detail:'Inspección Técnica de Seguridad en Edificaciones. Requiere demarcación, extintores y señalética listos.'},
  {id:'e1-zonif', etapa:'E1', tipo:'Permiso', crit:false, resp:'Asesor legal', urg:'Según predio',
    title:'Compatibilidad de uso / zonificación municipal', detail:'Verificar que el predio permite almacenar plásticos (residuo no peligroso).'},
  {id:'e1-licfunc', etapa:'E1', tipo:'Permiso', crit:false, resp:'Asesor legal', urg:'Tras zonificación + ITSE',
    title:'Licencia de funcionamiento municipal', detail:'Depende de zonificación e ITSE.'},
  {id:'e1-porton', etapa:'E1', tipo:'Operación', crit:true, resp:'Gerente Operaciones', urg:'Antes de operar',
    title:'Protocolo de portón único / acceso secundario', detail:'Exigir acceso secundario (≥3.5 m) al negociar la nave; si solo hay frontal, ventanas de tiempo exclusivas (mañana recepción / tarde despacho). KPI: conflictos de portón = 0.'},
  {id:'e1-bulbo', etapa:'E1', tipo:'Layout', crit:true, done:true, resp:'Supervisor de planta', urg:'Día 1',
    title:'Bulbo de giro en cabecera del pasillo central', detail:'Caja libre ≥4×4 m demarcada; el montacargas no retrocede con carga. Señalado en el plano de Etapa 1.'},
  {id:'e1-indeci', etapa:'E1', tipo:'Layout', crit:false, resp:'Supervisor de planta', urg:'Antes de ITSE',
    title:'Demarcación y señalética INDECI + extintores', detail:'Corredor peatonal amarillo 1.2 m, espejos convexos en esquinas ciegas, 9 extintores PQS/ABC 6 kg ≤15 m, señalética de salidas/EPP.'},
  {id:'e1-clasif', etapa:'E1', tipo:'Operación', crit:false, resp:'Operaciones', urg:'Mes 1',
    title:'Protocolo anti-saturación de Clasificación', detail:'WIP máx. 3 bolsones + 1 en báscula; cola en Recepción; spots numerados; objetivo <15 min/bolsón.'},
  {id:'e1-slotting', etapa:'E1', tipo:'Operación', crit:false, resp:'Supervisor de planta', urg:'Mes 2',
    title:'Reclasificación de slotting PET ↔ HDPE', detail:'Medir frecuencia real 2 meses; si el ahorro > $150/mes, reubicar en período de baja actividad; actualizar plano + mapa físico.'},

  /* ---------- ETAPA 2 — Procesamiento ---------- */
  {id:'e2-eia', etapa:'E2', tipo:'Permiso', crit:true, resp:'Asesor ambiental', urg:'90–180 d (ruta crítica)',
    title:'EIA / DIA ante SENACE (procesamiento)', detail:'El procesamiento suele gatillar EIA — más largo que el EC-RS. Iniciar con mucha anticipación.'},
  {id:'e2-efluentes', etapa:'E2', tipo:'Permiso', crit:true, resp:'Ing. ambiental', urg:'Diseño E2',
    title:'Autorización de vertimiento / efluentes (ANA)', detail:'El efluente de la PTAR debe cumplir los LMP del DS 003-2010-MINAM antes de conectar a desagüe/pozo.'},
  {id:'e2-pt', etapa:'E2', tipo:'Layout', crit:true, done:true, resp:'Arquitecto / Socio', urg:'Antes de cotizar obra',
    title:'Reubicar Almacén PT junto a Despacho (flujo "I")', detail:'Flujo recto Recepción→MP→Proceso→PT→Despacho, sin backtracking. Actualizado en el plano de Etapa 2.'},
  {id:'e2-lineas', etapa:'E2', tipo:'Layout', crit:true, done:true, resp:'Ing. de planta', urg:'Antes de cotizar equipo',
    title:'Plano con 2 líneas paralelas reales', detail:'2 líneas con pasillo central; confirmar huella ~1,800 m². Actualizado en el plano de Etapa 2.'},
  {id:'e2-electrica', etapa:'E2', tipo:'Permiso', crit:false, resp:'Ing. eléctrico', urg:'Diseño E2',
    title:'Ampliación de potencia eléctrica', detail:'El procesamiento es intensivo en energía; tramitar con la concesionaria.'},
  {id:'e2-ecrs2', etapa:'E2', tipo:'Permiso', crit:false, resp:'Asesor ambiental', urg:'Diseño E2',
    title:'Actualización del EC-RS por nueva actividad', detail:'El procesamiento cambia el alcance del registro EC-RS.'},
  {id:'e2-edificacion', etapa:'E2', tipo:'Permiso', crit:false, resp:'Asesor legal', urg:'Diseño E2',
    title:'Licencia de edificación / ampliación (obra civil)', detail:'Por la obra civil de la planta de procesamiento.'},
  {id:'e2-itse2', etapa:'E2', tipo:'Permiso', crit:false, resp:'Asesor', urg:'Antes de operar E2',
    title:'ITSE actualizado por nueva maquinaria', detail:'Reinspección por las líneas de proceso (energía, agua, fuego).'},
  {id:'e2-pellet', etapa:'E2', tipo:'Operación', crit:false, resp:'QA / Comercial', urg:'Antes de vender pellet',
    title:'Ficha técnica / certificación de calidad del pellet', detail:'Necesaria para contratos de offtake y para vender al tramo de precio alto.'},
  {id:'e2-lodos', etapa:'E2', tipo:'Permiso', crit:false, resp:'Ing. ambiental', urg:'Diseño E2',
    title:'Manejo de lodos de PTAR (declaración / disposición)', detail:'Los lodos del tratamiento de agua son un residuo a declarar/disponer.'},
  {id:'e2-ptar', etapa:'E2', tipo:'Operación', crit:false, resp:'Ing. ambiental', urg:'Diseño E2',
    title:'Validar dimensionamiento de la PTAR', detail:'Caudal (m³/h) × líneas × horas vs. capacidad de la PTAR cotizada; ampliar huella si hay déficit.'},
  {id:'e2-buffer', etapa:'E2', tipo:'Layout', crit:false, done:true, resp:'Ing. de proceso', urg:'Diseño E2',
    title:'Buffer de escurrido Lavado → Secado', detail:'Zona ≥20–30 m² que absorbe el diferencial de velocidad; evita exceso de humedad en la peletizadora. Añadido al plano de Etapa 2.'},
  {id:'e2-balance', etapa:'E2', tipo:'Operación', crit:false, resp:'Proveedor de equipo', urg:'Diseño E2',
    title:'Balance de línea real + revisión de Despacho', detail:'Throughput nominal, MTBF y MTTR por equipo; el cuello de botella define la capacidad. Despacho ≥150 m² (contenedor + paletizado).'},
];

window.PLASHUB_ETAPAS = {GEN:'General (empresa / financiero / legal)', E1:'Etapa 1 — Consolidación', E2:'Etapa 2 — Procesamiento'};

window.plashubInitState = function(){
  let st=null; try{st=JSON.parse(localStorage.getItem(window.PLASHUB_TASKKEY));}catch(e){}
  if(!st){ st={}; window.PLASHUB_TASKS.forEach(t=>{ if(t.done) st[t.id]=true; });
    try{localStorage.setItem(window.PLASHUB_TASKKEY,JSON.stringify(st));}catch(e){} }
  return st;
};
window.plashubSaveState = function(st){ try{localStorage.setItem(window.PLASHUB_TASKKEY,JSON.stringify(st));}catch(e){} };
