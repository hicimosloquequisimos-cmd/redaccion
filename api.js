/* =========================================================
   api.js · Conexión con el servidor (Apps Script conectado al Gmail)
   Si config.js no tiene la dirección del servidor, la página funciona
   en MODO DEMOSTRACIÓN: con datos de ejemplo guardados en este navegador.
   ========================================================= */
(function () {
  'use strict';
  const URL_SERVIDOR = (window.REDACCION || {}).servidor || '';
  const DEMO = !URL_SERVIDOR;

  async function llamarServidor(accion, datos = {}) {
    const r = await fetch(URL_SERVIDOR, { method: 'POST', headers: { 'Content-Type': 'text/plain;charset=utf-8' }, body: JSON.stringify({ accion, ...datos }), redirect: 'follow' });
    if (!r.ok) throw new Error('No hay conexión con el servidor (' + r.status + ')');
    const j = await r.json();
    if (!j.ok) { const e = new Error(j.error || 'Error'); e.sesion = j.sesion; throw e; }
    return j;
  }

  // =======================================================
  //  MODO DEMOSTRACIÓN (imita al servidor real)
  // =======================================================
  const CLAVE = window.SIMULACRO ? 'redaccion-simulacro-v1' : 'redaccion-demo-v3';
  let db = null;
  const ahora = () => new Date().toISOString();
  // Almacenamiento de la demo: IndexedDB (admite cientos de MB); si no está disponible, localStorage
  const IDB = (() => {
    let base = null;
    const abrir = () => base || (base = new Promise((ok, mal) => { try { const r = indexedDB.open('redaccion-demo', 1); r.onupgradeneeded = () => r.result.createObjectStore('kv'); r.onsuccess = () => ok(r.result); r.onerror = () => mal(r.error); } catch (e) { mal(e); } }));
    const op = (modo, fn) => abrir().then(d => new Promise((ok, mal) => { const t = d.transaction('kv', modo); const st = t.objectStore('kv'); const r = fn(st); t.oncomplete = () => ok(r && r.result); t.onerror = () => mal(t.error); }));
    return { leer: k => op('readonly', st => st.get(k)), poner: (k, v) => op('readwrite', st => st.put(v, k)), borrar: k => op('readwrite', st => st.delete(k)) };
  })();
  let escribiendo = Promise.resolve();
  const guardarDB = () => {
    const copia = JSON.stringify(db);
    escribiendo = escribiendo.then(() => IDB.poner(CLAVE, copia)).catch(() => { try { localStorage.setItem(CLAVE, copia); } catch (e) { /* sin espacio: solo en memoria */ } });
  };
  const RUB = { poemas: 'Poesía', cuentos: 'Cuento', microrrelato: 'Microrrelato', cronica: 'Crónica', ensayo: 'Ensayo', formas_hibridas: 'Formas híbridas', fragmento: 'Fragmento de obra mayor', dramaturgia: 'Dramaturgia', cortometraje_performance: 'Cortometraje y performance', historieta_comic: 'Historieta o cómic', narrativa_redes: 'Narrativa para redes', narrativas_digitales: 'Narrativas digitales', obra_sonora_podcast: 'Obra sonora y podcast', videopoema_videoensayo: 'Videopoema y videoensayo' };
  const ESTADOS = { por_asignar: 'Por asignar', en_evaluacion: 'En evaluación', tercer_evaluador: 'Necesita tercer evaluador', revision_directora: 'Revisión de la Directora', ajustes: 'Aceptado con ajustes · avisar al autor', esperando_autor: 'Esperando versión corregida', version_recibida: 'Versión corregida recibida · enviar a la Directora', devuelto: 'Devuelto por la Directora', aceptado: 'Aceptado · avisar al autor', correcciones_directora: 'Correcciones de la Directora · enviar al autor', aprobado: 'Aprobado · enviar a plataforma', listo_publicar: 'Listo para publicar', publicado: 'Publicado', no_aceptado: 'No aceptado · avisar al autor', cerrado: 'Cerrado (no aceptado)' };
  const DICT = { aceptada: 'Aceptada', ajustes: 'Aceptada con ajustes', no: 'No aceptada' };

  async function semilla() {
    const D = window.Docx;
    if (window.SIMULACRO) { db = await window.SIMULACRO.semilla(D); guardarDB(); return; }
    const poema = await D.desdeTexto('La orilla que baja\n\nCuando el agua se retira\nla orilla aprende a hablar:\n\n        dice piedra,\n        dice garza,\n        dice lo que el río calló.\n\nMe enseñaron que esto no era poesía.\nYo lo anoto igual,\ncon la boca llena de barro.');
    const cuento = await D.desdeTexto('El planchón de las seis\n\nEl planchón salía a las seis, aunque nadie supiera muy bien de qué reloj. Don Abel lo empujaba con la vara como quien empuja una conversación vieja, y los pasajeros se quedaban callados para no espantar al río.\n\nAquella mañana subió una muchacha con un cuaderno forrado en plástico. No miró la otra orilla: miró el agua, y escribió.\n\n—¿Qué anota? —preguntó Don Abel.\n\n—Lo que usted no dice —contestó ella.');
    const cronica = await D.desdeTexto('Mercado de los sábados\n\nA las cinco de la mañana el mercado ya tiene voz. Primero la de los camiones, después la de las vendedoras que acomodan el bocachico en hielo y, al final, la de los que regatean por costumbre más que por necesidad.');
    const ficha = (n, c, t, g) => D.desdeTexto(`Ficha de autor\nNombre completo: ${n}\nNombre de publicación: \nCorreo electrónico: ${c}\nCiudad: Aguachica\nTítulo de la obra: ${t}\nGénero o formato: ${g}`);
    const licencia = await D.desdeTexto('Licencia de uso no exclusiva\n(firmada por la persona autora · ejemplo)');
    const b = x => D.aBase64(x);
    const archivos = {
      f1: { nombre: 'LA_ORILLA_QUE_BAJA.docx', b64: b(poema) }, f2: { nombre: 'FICHA_LA_ORILLA_QUE_BAJA.docx', b64: b(await ficha('María Ejemplo Pérez', 'maria.ejemplo@correo.com', 'La orilla que baja', 'Poesía')) }, f3: { nombre: 'LICENCIA_LA_ORILLA_QUE_BAJA.docx', b64: b(licencia) },
      f4: { nombre: 'EL_PLANCHON_DE_LAS_SEIS.docx', b64: b(cuento) }, f5: { nombre: 'FICHA_EL_PLANCHON.docx', b64: b(await ficha('Andrés Ejemplo Ruiz', 'andres.ejemplo@correo.com', 'El planchón de las seis', 'Cuento')) }, f6: { nombre: 'LICENCIA_EL_PLANCHON.docx', b64: b(licencia) },
      f7: { nombre: 'MERCADO.docx', b64: b(cronica) }, f8: { nombre: 'FICHA_MERCADO.docx', b64: b(await ficha('Lucía Ejemplo Gómez', 'lucia.ejemplo@correo.com', 'Mercado de los sábados', 'Crónica')) }, f9: { nombre: 'LICENCIA_MERCADO.docx', b64: b(licencia) },
    };
    const equipo = [['Sebastian Montealegre Vergel', 'admin,plataforma,evaluador'], ['Luisa Mariana López', 'admin,directora'], ['Carmen Alicia Guagua Tovar', 'auxiliar'], ['Luis Hernán Romero Reyes', 'evaluador'], ['Juan José Árevalo Gallardo', 'evaluador'], ['Liz Alexandra Vega Rivera', 'evaluador'], ['Cristian Andrés Rangel Vera', 'evaluador'], ['Mariangel Arrieta', 'evaluador'], ['Sharon Michell Bandera Cárdenas', 'evaluador'], ['Eilyn Paz Patiño', 'evaluador']];
    const hace = d => new Date(Date.now() - d * 864e5).toISOString();
    db = {
      usuarios: equipo.map(([nombre, roles], i) => ({ id: 'U' + String(i + 1).padStart(3, '0'), nombre, correo: '', roles: roles.split(','), clave: 'demo', activo: true, debeCambiar: false, avisos: true })),
      archivos,
      hilos: [
        { hiloId: 'h1', asunto: 'ENVÍO | POESÍA | La orilla que baja', de: 'María Ejemplo Pérez <maria.ejemplo@correo.com>', fecha: hace(0.2), noLeido: true, mensajes: [{ id: 'm1', de: 'María Ejemplo Pérez <maria.ejemplo@correo.com>', fecha: hace(0.2), cuerpo: 'Hola, equipo de Hicimos lo que quisimos:\n\nEnvío mi poema para la gaceta. Adjunto la obra, la ficha de autor y la licencia firmada.\n\nGracias.', adjuntos: [{ idx: 0, nombre: archivos.f1.nombre, fid: 'f1' }, { idx: 1, nombre: archivos.f2.nombre, fid: 'f2' }, { idx: 2, nombre: archivos.f3.nombre, fid: 'f3' }] }] },
        { hiloId: 'h5', asunto: 'Tienes 3 notificaciones nuevas', de: 'Instagram <no-reply@mail.instagram.com>', fecha: hace(0.5), noLeido: true, mensajes: [{ id: 'm5', de: 'Instagram <no-reply@mail.instagram.com>', fecha: hace(0.5), cuerpo: 'A hicimosloquequisimos le gustaron 3 publicaciones nuevas.', adjuntos: [] }] },
        { hiloId: 'h6', asunto: 'Boletín institucional · Semana cultural', de: 'Comunicaciones UPC <comunicaciones@correo.com>', fecha: hace(2), noLeido: false, mensajes: [{ id: 'm6', de: 'Comunicaciones UPC <comunicaciones@correo.com>', fecha: hace(2), cuerpo: 'Te invitamos a la semana cultural. Adjuntamos la pieza gráfica.', adjuntos: [{ nombre: 'afiche-semana-cultural.png', tipo: 'image/png', datos: '' }] }] },
        { hiloId: 'h2', asunto: '¿Hasta cuándo reciben textos?', de: 'Pedro Ejemplo <pedro.ejemplo@correo.com>', fecha: hace(1), noLeido: true, mensajes: [{ id: 'm2', de: 'Pedro Ejemplo <pedro.ejemplo@correo.com>', fecha: hace(1), cuerpo: 'Buenas, quería saber hasta cuándo está abierta la convocatoria. Gracias.', adjuntos: [] }] },
        { hiloId: 'h3', asunto: 'ENVÍO | CUENTO | El planchón de las seis', de: 'Andrés Ejemplo Ruiz <andres.ejemplo@correo.com>', fecha: hace(9), noLeido: false, mensajes: [{ id: 'm3', de: 'Andrés Ejemplo Ruiz <andres.ejemplo@correo.com>', fecha: hace(9), cuerpo: 'Adjunto mi cuento, ficha y licencia.', adjuntos: [{ idx: 0, nombre: archivos.f4.nombre, fid: 'f4' }, { idx: 1, nombre: archivos.f5.nombre, fid: 'f5' }, { idx: 2, nombre: archivos.f6.nombre, fid: 'f6' }] }] },
        { hiloId: 'h4', asunto: 'ENVÍO | CRÓNICA | Mercado de los sábados', de: 'Lucía Ejemplo Gómez <lucia.ejemplo@correo.com>', fecha: hace(20), noLeido: false, mensajes: [{ id: 'm4', de: 'Lucía Ejemplo Gómez <lucia.ejemplo@correo.com>', fecha: hace(20), cuerpo: 'Les envío mi crónica.', adjuntos: [{ idx: 0, nombre: archivos.f7.nombre, fid: 'f7' }, { idx: 1, nombre: archivos.f8.nombre, fid: 'f8' }, { idx: 2, nombre: archivos.f9.nombre, fid: 'f9' }] }] },
      ],
      envios: [
        { id: 'ENV-001', creado: hace(9), hiloId: 'h3', autorNombre: 'Andrés Ejemplo Ruiz', autorCorreo: 'andres.ejemplo@correo.com', titulo: 'El planchón de las seis', genero: 'Cuento', rubrica: 'cuentos', estado: 'en_evaluacion', ronda: 1, textoId: 'f4', textoNombre: archivos.f4.nombre, fichaId: 'f5', licenciaId: 'f6', versiones: [{ n: 1, id: 'f4', nombre: archivos.f4.nombre, fecha: hace(9) }], otros: [], avisos: { recibido: hace(9) }, actualizado: hace(2) },
        { id: 'ENV-002', creado: hace(20), hiloId: 'h4', autorNombre: 'Lucía Ejemplo Gómez', autorCorreo: 'lucia.ejemplo@correo.com', titulo: 'Mercado de los sábados', genero: 'Crónica', rubrica: 'cronica', estado: 'revision_directora', decision: 'aceptada', ronda: 1, textoId: 'f7', textoNombre: archivos.f7.nombre, fichaId: 'f8', licenciaId: 'f9', versiones: [{ n: 1, id: 'f7', nombre: archivos.f7.nombre, fecha: hace(20) }], otros: [], avisos: { recibido: hace(20) }, actualizado: hace(1) },
      ],
      asignaciones: [
        { id: 'A001', envioId: 'ENV-001', evaluadorId: 'U001', rubrica: 'cuentos', ronda: 1, estado: 'pendiente', asignada: hace(8), niveles: [], anotaciones: [] },
        { id: 'A002', envioId: 'ENV-001', evaluadorId: 'U004', rubrica: 'cuentos', ronda: 1, estado: 'entregada', asignada: hace(8), entregada: hace(2), modo: 'pagina', niveles: [4, 4, 5, 4, 4], puntaje: 4.2, dictamen: 'ajustes', observaciones: 'El cierre funciona muy bien. Revisar la puntuación de los diálogos.', anotaciones: [{ p: 2, ini: 0, fin: 11, tipo: 'comentario', comentario: '¿El planchón o la planchada? Unificar.' }, { p: 2, ini: 46, fin: 47, tipo: 'reemplazar', nuevo: ' —', comentario: 'Raya, no guion' }] },
        { id: 'A003', envioId: 'ENV-002', evaluadorId: 'U005', rubrica: 'cronica', ronda: 1, estado: 'entregada', asignada: hace(18), entregada: hace(3), modo: 'pagina', niveles: [5, 4, 5, 5, 4, 5], puntaje: 4.7, dictamen: 'aceptada', observaciones: 'Crónica viva y precisa. Lista para publicar.', anotaciones: [{ p: 2, ini: 2, fin: 7, tipo: 'comentario', comentario: 'Gran arranque.' }] },
        { id: 'A004', envioId: 'ENV-002', evaluadorId: 'U006', rubrica: 'cronica', ronda: 1, estado: 'entregada', asignada: hace(18), entregada: hace(1), modo: 'pagina', niveles: [4, 5, 5, 4, 5, 5], puntaje: 4.7, dictamen: 'aceptada', observaciones: 'Me gustaría un poco más de contexto del mercado, pero funciona.', anotaciones: [] },
      ],
      plantillas: [
        { clave: 'recibido', nombre: 'Recibido', asunto: 'Recibimos tu texto · Hicimos lo que quisimos', cuerpo: 'Hola:\n\nRecibimos tu envío «{titulo}». Gracias por confiarnos lo que escribiste.\n\nTu texto pasa ahora a evaluación doble ciego: dos lectores del equipo lo leerán sin saber quién lo escribió. Te escribiremos cuando tengamos una respuesta.\n\nUn abrazo,\nEquipo editorial · Hicimos lo que quisimos' },
        { clave: 'aceptado', nombre: 'Aceptado', asunto: 'Tu texto fue aceptado · Hicimos lo que quisimos', cuerpo: 'Hola:\n\n¡Buenas noticias! Tu texto «{titulo}» fue aceptado para publicarse en Hicimos lo que quisimos.\n\nPronto te avisaremos cuando esté en línea.\n\nUn abrazo,\nEquipo editorial · Hicimos lo que quisimos' },
        { clave: 'ajustes', nombre: 'Aceptado con ajustes', asunto: 'Tu texto fue aceptado con ajustes · Hicimos lo que quisimos', cuerpo: 'Hola:\n\nTu texto «{titulo}» fue aceptado con ajustes. Te adjuntamos las sugerencias de los evaluadores: puedes ver los comentarios y los cambios propuestos en Word.\n\nTú decides qué enmiendas aceptar. Cuando tengas la versión corregida, respóndenos a este mismo correo con el archivo.\n\nUn abrazo,\nEquipo editorial · Hicimos lo que quisimos' },
        { clave: 'directora', nombre: 'Correcciones de la Dirección', asunto: 'Correcciones a tu texto · Hicimos lo que quisimos', cuerpo: 'Hola:\n\nLa Dirección editorial leyó tu texto «{titulo}» antes de publicarlo e hizo algunas correcciones y sugerencias. Te las adjuntamos en Word: las verás como comentarios y control de cambios.\n\nRevísalas con calma: tú decides qué aceptar. Cuando estés de acuerdo, respóndenos a este mismo correo con la versión final (o dinos que así está bien).\n\nUn abrazo,\nEquipo editorial · Hicimos lo que quisimos' },
        { clave: 'no', nombre: 'No aceptado', asunto: 'Sobre tu texto · Hicimos lo que quisimos', cuerpo: 'Hola:\n\nGracias por enviarnos «{titulo}». Esta vez tu texto no fue seleccionado para la gaceta.\n\nNo dejes de escribir: la convocatoria sigue abierta y nos encantaría volver a leerte.\n\nUn abrazo,\nEquipo editorial · Hicimos lo que quisimos' },
        { clave: 'libre', nombre: 'Mensaje libre', asunto: 'Hicimos lo que quisimos', cuerpo: 'Hola:\n\n\n\nUn abrazo,\nEquipo editorial · Hicimos lo que quisimos' },
      ],
      historial: [{ fecha: hace(9), envioId: 'ENV-001', usuario: 'Carmen Alicia Guagua Tovar', accion: 'Envío registrado', detalle: '' }, { fecha: hace(8), envioId: 'ENV-001', usuario: 'Carmen Alicia Guagua Tovar', accion: 'Evaluadores asignados', detalle: 'Sebastian Montealegre Vergel, Luis Hernán Romero Reyes · rúbrica: Cuento' }],
      ajustes: { nombreRemitente: 'Hicimos lo que quisimos', urlPagina: '', diasRecordatorio: 7, avisosActivos: true },
      enviados: [], sesiones: {},
      notificaciones: [
        { id: 'N1', usuarioId: 'U001', fecha: hace(8), titulo: 'Tienes una evaluación nueva · ENV-001', texto: 'Cuento · rúbrica de Cuento', ruta: 'evaluacion/A001', leida: false },
        { id: 'N2', usuarioId: 'U003', fecha: hace(2), titulo: 'Evaluación entregada · ENV-001', texto: 'Luis Hernán Romero Reyes: Aceptada con ajustes (4,2)', ruta: 'envio/ENV-001', leida: false },
        { id: 'N3', usuarioId: 'U003', fecha: hace(1), titulo: 'ENV-002: Revisión de la Directora', texto: 'Mercado de los sábados', ruta: 'envio/ENV-002', leida: false },
        { id: 'N4', usuarioId: 'U002', fecha: hace(1), titulo: 'Obra aceptada para revisar · ENV-002', texto: 'Mercado de los sábados', ruta: 'envio/ENV-002', leida: false },
      ],
    };
    guardarDB();
  }

  async function cargarDemo() {
    if (db) return;
    try { const g = await IDB.leer(CLAVE); db = g ? JSON.parse(g) : null; } catch (e) { db = null; }
    if (!db) { try { db = JSON.parse(localStorage.getItem(CLAVE) || 'null'); } catch (e) { db = null; } }
    if (!db) await semilla();
  }
  const sesion = t => { const id = db.sesiones[t]; const u = db.usuarios.find(x => x.id === id); if (!u) { const e = new Error('Tu sesión terminó. Vuelve a entrar.'); e.sesion = false; throw e; } return u; };
  const pub = u => ({ id: u.id, nombre: u.nombre, correo: u.correo, roles: u.roles, activo: u.activo, debeCambiar: u.debeCambiar, avisos: u.avisos, tieneFirma: !!u.firmaId });
  const tiene = (u, r) => u.roles.includes(r);
  const exigir = (u, ...rs) => { if (!rs.some(r => tiene(u, r))) throw new Error('Tu perfil no tiene permiso para esto.'); };
  const hist = (envioId, u, accion, detalle = '') => db.historial.push({ fecha: ahora(), envioId, usuario: u ? u.nombre : 'Sistema', accion, detalle });
  const sigId = (pre, lista) => pre + String(Math.max(0, ...lista.map(x => +(String(x.id).match(/(\d+)$/) || [0, 0])[1])) + 1).padStart(3, '0');
  const respondio = e => { if (e.estado !== 'esperando_autor') return false; const av = e.avisos || {}; const desde = [av.ajustes, av.directora].filter(Boolean).sort().pop() || ''; const h = db.hilos.find(x => x.hiloId === e.hiloId); return !!h && h.mensajes.some(m => m.fecha > desde && !m.nuestro && m.adjuntos.length); };
  const textoEstado = e => e.estado === 'revision_directora' && (e.versiones || []).length > 1 ? 'Revisión de la Directora · versión corregida' : ESTADOS[e.estado];
  const VIS = { directora: ['revision_directora', 'ajustes', 'esperando_autor', 'version_recibida', 'devuelto', 'correcciones_directora', 'aceptado', 'aprobado', 'listo_publicar', 'publicado'], plataforma: ['listo_publicar', 'publicado'] };
  const puedeVer = (e, u) => tiene(u, 'auxiliar') || ['directora', 'plataforma'].some(r => tiene(u, r) && VIS[r].includes(e.estado));
  const notificar = (u, titulo, texto, ruta) => { db.notificaciones = db.notificaciones || []; db.notificaciones.push({ id: 'N' + Date.now() + Math.random().toString(36).slice(2, 5), usuarioId: u.id, fecha: ahora(), titulo, texto: texto || '', ruta: ruta || '', leida: false }); db.enviados.push({ fecha: ahora(), para: u.nombre, texto: titulo, tipo: 'aviso' }); };
  const aviso = (rol, titulo, ruta, texto) => db.usuarios.filter(u => tiene(u, rol) && u.activo).forEach(u => notificar(u, titulo, texto, ruta));
  function envioPara(e, u) {
    const asig = db.asignaciones.filter(a => a.envioId === e.id).map(a => ({ ...a, evaluador: (db.usuarios.find(x => x.id === a.evaluadorId) || {}).nombre }));
    return { ...e, estadoTexto: textoEstado(e), respondio: tiene(u, 'auxiliar') && respondio(e), asignaciones: tiene(u, 'auxiliar') ? asig : tiene(u, 'directora') ? asig.filter(a => a.estado === 'entregada') : [] /* plataforma no ve rúbricas */, historial: db.historial.filter(h => h.envioId === e.id) };
  }
  const versionEn = (e, f) => { const vs = e.versiones.filter(v => !f || new Date(v.fecha) <= new Date(new Date(f).getTime() + 60000)); return (vs.slice(-1)[0] || e.versiones[0] || {}).id || e.textoId; };
  const CARPETAS = { aceptada: 'Aceptados', ajustes: 'Aceptados con ajustes', no: 'No aceptados' };
  function resolverRonda(e) {
    const asig = db.asignaciones.filter(a => a.envioId === e.id && +a.ronda === +e.ronda);
    if (!asig.length || asig.some(a => a.estado !== 'entregada')) return;
    const d = asig.map(a => a.dictamen); let r;
    if (d.length >= 3) { const c = {}; d.forEach(x => c[x] = (c[x] || 0) + 1); r = Object.keys(c).find(k => c[k] >= 2) || 'ajustes'; }
    else if (d.length === 1 || d[0] === d[1]) r = d[0];
    else if (d.includes('no')) r = 'desacuerdo'; else r = 'ajustes';
    e.estado = { desacuerdo: 'tercer_evaluador', aceptada: 'revision_directora', ajustes: 'ajustes', no: 'no_aceptado' }[r];
    e.decision = r; e.actualizado = ahora(); if (CARPETAS[r]) e.archivo = CARPETAS[r];
    hist(e.id, null, 'Resultado de la evaluación', r === 'desacuerdo' ? 'Los evaluadores no coinciden' : DICT[r]);
    aviso('auxiliar', `${e.id}: ${ESTADOS[e.estado]}`, 'envio/' + e.id, e.titulo);
    if (r === 'aceptada' || r === 'ajustes') aviso('directora', `${e.id}: ${ESTADOS[e.estado]}`, 'envio/' + e.id, e.titulo);
  }

  const ACC = {
    perfiles: () => ({ perfiles: db.usuarios.filter(u => u.activo).map(u => ({ id: u.id, nombre: u.nombre })) }),
    entrar: p => { const u = db.usuarios.find(x => x.id === p.id); if (!u || u.clave !== p.clave) throw new Error('Contraseña incorrecta. (En la demostración, la contraseña de todos es «demo».)'); const t = 'tk' + Math.random().toString(36).slice(2); db.sesiones[t] = u.id; return { token: t, usuario: pub(u) }; },
    salir: p => { delete db.sesiones[p.token]; return {}; },
    yo: (p, u) => ({ usuario: pub(u) }),
    miFirma: (p, u) => ({ firma: u.firmaId && db.archivos[u.firmaId] ? db.archivos[u.firmaId].b64 : null }),
    guardarFirma: (p, u) => { exigir(u, 'admin'); const t = db.usuarios.find(x => x.id === p.usuarioId); if (!t) throw new Error('No existe ese perfil.'); if (!p.b64) throw new Error('Falta la imagen.'); if (!/^iVBORw0KGgo/.test(p.b64)) throw new Error('La firma debe ser una imagen PNG.'); const id = 'firma' + t.id + Date.now(); db.archivos[id] = { nombre: `Firma · ${t.nombre}.png`, b64: p.b64 }; t.firmaId = id; return { usuario: pub(t) }; },
    quitarFirma: (p, u) => { exigir(u, 'admin'); const t = db.usuarios.find(x => x.id === p.usuarioId); if (!t) throw new Error('No existe ese perfil.'); t.firmaId = ''; return { usuario: pub(t) }; },
    firmaDe: (p, u) => { exigir(u, 'admin'); const t = db.usuarios.find(x => x.id === p.usuarioId); return { firma: t && t.firmaId && db.archivos[t.firmaId] ? db.archivos[t.firmaId].b64 : null }; },
    cambiarClave: (p, u) => { if (u.clave !== p.actual) throw new Error('La contraseña actual no es correcta.'); if (!p.nueva || p.nueva.length < 8) throw new Error('La contraseña nueva debe tener al menos 8 caracteres.'); u.clave = p.nueva; u.debeCambiar = false; return { usuario: pub(u) }; },
    resumen: (p, u) => { const r = {}; if (tiene(u, 'evaluador')) r.pendientes = db.asignaciones.filter(a => a.evaluadorId === u.id && a.estado !== 'entregada').length; if (tiene(u, 'auxiliar')) { r.porHacer = db.envios.filter(e => ['por_asignar', 'tercer_evaluador'].includes(e.estado)).length; r.porHacerEval = db.envios.filter(e => ['ajustes', 'devuelto', 'correcciones_directora', 'aceptado', 'aprobado', 'no_aceptado', 'version_recibida'].includes(e.estado) || respondio(e)).length; } if (tiene(u, 'directora')) r.porRevisar = db.envios.filter(e => e.estado === 'revision_directora').length; if (tiene(u, 'plataforma')) r.listos = db.envios.filter(e => e.estado === 'listo_publicar').length; return { resumen: r }; },
    bandeja: (p, u) => { exigir(u, 'auxiliar'); const doc = n => /\.(docx?|pdf|odt|rtf|txt|pages)$/i.test(n); return { soloDocumentos: !p.todo, hilos: db.hilos.filter(h => p.todo || h.mensajes.some(m => m.adjuntos.some(a => doc(a.nombre)))).slice().sort((a, b) => b.fecha.localeCompare(a.fecha)).map(h => ({ hiloId: h.hiloId, asunto: h.asunto, de: h.de, fecha: h.fecha, fragmento: h.mensajes.slice(-1)[0].cuerpo.slice(0, 180), adjuntos: h.mensajes.flatMap(m => m.adjuntos.map(a => a.nombre)), mensajes: h.mensajes.length, noLeido: h.noLeido, envioId: (db.envios.find(e => e.hiloId === h.hiloId) || {}).id || null })) }; },
    hilo: (p, u) => { exigir(u, 'auxiliar'); const h = db.hilos.find(x => x.hiloId === p.hiloId); h.noLeido = false; return { asunto: h.asunto, envioId: (db.envios.find(e => e.hiloId === h.hiloId) || {}).id || null, mensajes: h.mensajes.map(m => ({ ...m, adjuntos: m.adjuntos.map(a => ({ idx: a.idx, nombre: a.nombre, tam: (db.archivos[a.fid] || {}).b64 ? Math.round(db.archivos[a.fid].b64.length * .75) : 0 })) })) }; },
    adjunto: (p, u) => { exigir(u, 'auxiliar'); const m = db.hilos.flatMap(h => h.mensajes).find(x => x.id === p.mensajeId); const a = m.adjuntos[p.idx]; return { nombre: a.nombre, b64: db.archivos[a.fid].b64 }; },
    registrar: (p, u) => {
      exigir(u, 'auxiliar'); if (!p.titulo) throw new Error('Falta el título de la obra.'); if (!p.texto) throw new Error('Elige cuál adjunto es el texto (la obra).');
      const fidDe = x => { if (!x) return null; const m = db.hilos.flatMap(h => h.mensajes).find(y => y.id === x.mensajeId); return m.adjuntos[x.idx].fid; };
      const id = sigId('ENV-', db.envios); const tf = fidDe(p.texto);
      const e = { id, creado: ahora(), hiloId: p.hiloId || '', autorNombre: p.autorNombre || '', autorCorreo: p.autorCorreo || '', titulo: p.titulo, genero: p.genero || '', rubrica: p.rubrica || '', estado: 'por_asignar', ronda: 1, textoId: tf, textoNombre: db.archivos[tf].nombre, fichaId: fidDe(p.ficha), licenciaId: fidDe(p.licencia), versiones: [{ n: 1, id: tf, nombre: db.archivos[tf].nombre, fecha: ahora() }], otros: [], avisos: {}, actualizado: ahora() };
      db.envios.push(e); hist(id, u, 'Envío registrado', p.titulo); return { envio: envioPara(e, u) };
    },
    editarEnvio: (p, u) => { exigir(u, 'auxiliar'); const e = db.envios.find(x => x.id === p.id); ['titulo', 'genero', 'rubrica', 'autorNombre', 'autorCorreo', 'nota'].forEach(k => { if (p[k] != null) e[k] = p[k]; }); hist(e.id, u, 'Datos editados'); return { envio: envioPara(e, u) }; },
    evaluadores: (p, u) => { exigir(u, 'auxiliar'); return { evaluadores: db.usuarios.filter(x => x.activo && tiene(x, 'evaluador')).map(x => ({ id: x.id, nombre: x.nombre })) }; },
    asignar: (p, u) => {
      exigir(u, 'auxiliar'); const e = db.envios.find(x => x.id === p.envioId);
      if (!(p.evaluadores || []).length) throw new Error('Elige al menos un evaluador.'); if (!p.rubrica) throw new Error('Elige la rúbrica.');
      const ya = db.asignaciones.filter(a => a.envioId === e.id && +a.ronda === +e.ronda).map(a => a.evaluadorId);
      p.evaluadores.filter(x => !ya.includes(x)).forEach(uid => { db.asignaciones.push({ id: sigId('A', db.asignaciones), envioId: e.id, evaluadorId: uid, rubrica: p.rubrica, ronda: e.ronda, estado: 'pendiente', asignada: ahora(), niveles: [], anotaciones: [] }); notificar(db.usuarios.find(x => x.id === uid), `Tienes una evaluación nueva · ${e.id}`, `${e.genero || ''} · rúbrica de ${RUB[p.rubrica]}`, 'evaluacion/' + db.asignaciones[db.asignaciones.length - 1].id); });
      e.rubrica = p.rubrica; e.estado = 'en_evaluacion'; e.actualizado = ahora();
      hist(e.id, u, 'Evaluadores asignados', p.evaluadores.map(id => db.usuarios.find(x => x.id === id).nombre).join(', ') + ' · rúbrica: ' + RUB[p.rubrica]);
      return { envio: envioPara(e, u) };
    },
    quitarAsignacion: (p, u) => { exigir(u, 'auxiliar'); const a = db.asignaciones.find(x => x.id === p.id); if (a.estado === 'entregada') throw new Error('Esa evaluación ya fue entregada; no se puede quitar.'); db.asignaciones = db.asignaciones.filter(x => x !== a); const e = db.envios.find(x => x.id === a.envioId); if (!db.asignaciones.some(x => x.envioId === e.id && +x.ronda === +e.ronda)) e.estado = 'por_asignar'; hist(e.id, u, 'Evaluador retirado'); return { envio: envioPara(e, u) }; },
    responderAutor: (p, u) => {
      exigir(u, 'auxiliar'); const e = db.envios.find(x => x.id === p.envioId); if (!p.cuerpo) throw new Error('El mensaje está vacío.');
      p.adjuntos = (p.adjuntos || []).filter(x => x.tipo !== 'rubrica'); // las rúbricas son internas
      db.enviados.push({ fecha: ahora(), para: e.autorCorreo, texto: p.asunto, tipo: 'autor', adjuntos: (p.adjuntos || []).length });
      e.avisos = e.avisos || {}; e.avisos[p.tipo || 'libre'] = ahora();
      if (p.tipo === 'aceptado' && e.estado === 'aceptado') e.estado = 'aprobado';
      if (p.tipo === 'directora' && e.estado === 'correcciones_directora') e.estado = 'esperando_autor';
      if (p.tipo === 'ajustes' && e.estado === 'ajustes') e.estado = 'esperando_autor';
      if (p.tipo === 'no' && e.estado === 'no_aceptado') e.estado = 'cerrado';
      // Demo: la autora «responde» enseguida con su versión corregida, para poder practicar
      if (e.estado === 'esperando_autor') { const h = db.hilos.find(x => x.hiloId === e.hiloId); const f0 = db.archivos[e.textoId]; if (h && f0) { const fid = 'vr' + Date.now(); const nom = e.titulo + ' (versión corregida).docx'; db.archivos[fid] = { nombre: nom, b64: f0.b64 }; const t = new Date(Date.now() + 2000).toISOString(); h.mensajes.push({ id: 'mr' + Date.now(), de: h.de, fecha: t, cuerpo: 'Hola, equipo:\n\nGracias por los comentarios. Les envío la versión corregida de mi texto.\n\nUn abrazo.', adjuntos: [{ idx: 0, nombre: nom, fid }] }); h.fecha = t; h.noLeido = true; } }
      e.actualizado = ahora(); hist(e.id, u, 'Correo al autor', (p.tipo || 'libre') + ((p.adjuntos || []).length ? ` · ${p.adjuntos.length} adjunto(s)` : ''));
      return { envio: envioPara(e, u), demo: true };
    },
    subirVersion: (p, u) => { exigir(u, 'auxiliar'); const e = db.envios.find(x => x.id === p.envioId); const fid = 'v' + Date.now(); db.archivos[fid] = { nombre: p.nombre || 'version.docx', b64: p.b64 }; e.versiones.push({ n: e.versiones.length + 1, id: fid, nombre: p.nombre, fecha: ahora() }); e.textoId = fid; if (e.estado === 'esperando_autor') e.estado = 'version_recibida'; e.actualizado = ahora(); hist(e.id, u, 'Nueva versión del texto', 'Versión ' + e.versiones.length); return { envio: envioPara(e, u) }; },
    moverEnvio: (p, u) => { exigir(u, 'auxiliar'); const e = db.envios.find(x => x.id === p.envioId); if (p.a === 'directora') { e.estado = 'revision_directora'; aviso('directora', `Versión corregida para revisar · ${e.id}`, 'envio/' + e.id, e.titulo); } else if (p.a === 'nueva_ronda') { e.ronda = (+e.ronda || 1) + 1; e.estado = 'por_asignar'; } else if (p.a === 'cerrar') e.estado = 'cerrado'; else if (p.a === 'plataforma' || p.a === 'listo_publicar') { if ((!e.fichaId || !e.licenciaId) && !p.sinFichaOLicencia) throw new Error('Falta la ficha o la licencia del autor. Súbelas o confirma que quieres enviarlo así.'); e.estado = 'listo_publicar'; const v = e.versiones.slice(-1)[0]; aviso('plataforma', `Listo para publicar · ${e.id}`, 'envio/' + e.id, `${e.titulo} · versión final ${v.n}${e.fichaId ? ', ficha' : ''}${e.licenciaId ? ' y licencia' : ''}${p.nota ? ' · ' + p.nota : ''}`); if (p.nota) hist(e.id, u, 'Nota para plataforma', p.nota); } e.actualizado = ahora(); hist(e.id, u, 'Estado cambiado', ESTADOS[e.estado]); return { envio: envioPara(e, u) }; },
    plantillas: (p, u) => { exigir(u, 'auxiliar'); return { plantillas: db.plantillas }; },
    guardarPlantilla: (p, u) => { exigir(u, 'auxiliar'); const t = db.plantillas.find(x => x.clave === p.clave); if (t) Object.assign(t, { asunto: p.asunto, cuerpo: p.cuerpo, nombre: p.nombre || t.nombre }); else db.plantillas.push({ clave: p.clave, nombre: p.nombre, asunto: p.asunto, cuerpo: p.cuerpo }); return { plantillas: db.plantillas }; },
    envios: (p, u) => { exigir(u, 'auxiliar', 'directora', 'plataforma'); return { envios: db.envios.filter(e => puedeVer(e, u)).map(e => ({ id: e.id, titulo: e.titulo, genero: e.genero, autorNombre: e.autorNombre, estado: e.estado, estadoTexto: textoEstado(e), respondio: tiene(u, 'auxiliar') && respondio(e), versionN: (e.versiones || []).length, creado: e.creado, actualizado: e.actualizado, ronda: e.ronda, decision: e.decision, archivo: e.archivo || '', corrDir: !!(e.dirCorr && e.dirCorr.fecha && !e.dirCorr.aplicadas), evaluadores: db.asignaciones.filter(a => a.envioId === e.id && +a.ronda === +e.ronda).map(a => ({ nombre: db.usuarios.find(x => x.id === a.evaluadorId).nombre, estado: a.estado, dictamen: a.dictamen, puntaje: a.puntaje, firmada: !!a.firmadaId })) })).sort((a, b) => String(b.actualizado).localeCompare(a.actualizado)) }; },
    envio: (p, u) => { const e = db.envios.find(x => x.id === p.id); if (!e || !puedeVer(e, u)) throw new Error('Tu perfil no puede ver este envío.'); return { envio: envioPara(e, u) }; },
    archivo: (p, u) => {
      let fid;
      if (p.asignacionId) { const a = db.asignaciones.find(x => x.id === p.asignacionId); const e = db.envios.find(x => x.id === a.envioId); if (!(a.evaluadorId === u.id || tiene(u, 'auxiliar') || (tiene(u, 'directora') && puedeVer(e, u)))) throw new Error('No puedes ver este archivo.'); fid = p.cual === 'texto' ? versionEn(e, a.asignada) : a[p.cual + 'Id']; }
      else { const e = db.envios.find(x => x.id === p.envioId); if (!puedeVer(e, u)) throw new Error('No puedes ver este archivo.'); fid = p.cual === 'version' ? (e.versiones.find(v => +v.n === +p.n) || {}).id : p.cual === 'directora' ? (e.dirCorr || {}).archivoId : e[p.cual + 'Id']; }
      if (!fid || !db.archivos[fid]) throw new Error('Ese archivo no existe.');
      return { nombre: db.archivos[fid].nombre, b64: db.archivos[fid].b64 };
    },
    misEvaluaciones: (p, u) => { exigir(u, 'evaluador'); return { evaluaciones: db.asignaciones.filter(a => a.evaluadorId === u.id).map(a => { const e = db.envios.find(x => x.id === a.envioId) || {}; return { id: a.id, envioId: a.envioId, titulo: e.titulo, genero: e.genero, rubrica: a.rubrica, estado: a.estado, asignada: a.asignada, entregada: a.entregada, dictamen: a.dictamen, puntaje: a.puntaje }; }) }; },
    evaluacion: (p, u) => { const a = db.asignaciones.find(x => x.id === p.id); if (!a || a.evaluadorId !== u.id) throw new Error('Esa evaluación no es tuya.'); const e = db.envios.find(x => x.id === a.envioId); const v = e.versiones.slice(-1)[0]; return { asignacion: a, envio: { id: e.id, titulo: e.titulo, genero: e.genero, textoNombre: v.nombre, ronda: e.ronda }, evaluador: u.nombre }; },
    guardarBorrador: (p, u) => { const a = db.asignaciones.find(x => x.id === p.id); if (!a || a.evaluadorId !== u.id) throw new Error('Esa evaluación no es tuya.'); if (a.estado === 'entregada') throw new Error('Ya fue entregada.'); ['niveles', 'anotaciones', 'observaciones', 'modo', 'enlace'].forEach(k => { if (p[k] !== undefined) a[k] = p[k]; }); a.estado = 'borrador'; return { guardado: ahora() }; },
    entregar: (p, u) => {
      const a = db.asignaciones.find(x => x.id === p.id); if (!a || a.evaluadorId !== u.id) throw new Error('Esa evaluación no es tuya.'); if (a.estado === 'entregada') throw new Error('Ya la habías entregado.');
      if (!['aceptada', 'ajustes', 'no'].includes(p.dictamen)) throw new Error('Falta el dictamen.');
      Object.assign(a, { niveles: p.niveles || [], puntaje: p.puntaje, dictamen: p.dictamen, observaciones: p.observaciones || '', anotaciones: p.anotaciones || [], modo: p.modo, enlace: p.enlace || '', estado: 'entregada', entregada: ahora() });
      const e = db.envios.find(x => x.id === a.envioId);
      const gu = (b64, n) => { const id = 'g' + Math.random().toString(36).slice(2); db.archivos[id] = { nombre: n, b64 }; return id; };
      if (p.rubricaDocx) a.rubricaId = gu(p.rubricaDocx, `${a.envioId} · rúbrica${p.firmadaAuto ? ' FIRMADA' : ''} · ${u.nombre}.docx`);
      if (p.correccionDocx) a.correccionId = gu(p.correccionDocx, `${a.envioId} · texto corregido · ${u.nombre}.docx`);
      if (p.firmadaAuto && a.rubricaId) Object.assign(a, { firmadaId: a.rubricaId, firmadaFecha: ahora(), firmadaPor: 'Firma digital guardada en su perfil' });
      if (p.subidaRubrica) a.subidaRubricaId = gu(p.subidaRubrica.b64, p.subidaRubrica.nombre);
      if (p.subidaCorreccion) a.subidaCorreccionId = gu(p.subidaCorreccion.b64, p.subidaCorreccion.nombre);
      hist(e.id, u, 'Evaluación entregada', `${DICT[a.dictamen]} · ${String(a.puntaje).replace('.', ',')}`);
      aviso('auxiliar', `Evaluación entregada · ${e.id}`, 'envio/' + e.id, `${u.nombre}: ${DICT[a.dictamen]} (${String(a.puntaje).replace('.', ',')})`);
      resolverRonda(e); return { asignacion: a };
    },
    subirFirmada: (p, u) => {
      const a = db.asignaciones.find(x => x.id === p.asignacionId); if (!a) throw new Error('No existe esa evaluación.');
      if (a.evaluadorId !== u.id && !tiene(u, 'auxiliar')) throw new Error('Solo quien evaluó o la Auxiliar pueden subir la rúbrica firmada.');
      if (a.estado !== 'entregada') throw new Error('Primero hay que entregar la evaluación.');
      const e = db.envios.find(x => x.id === a.envioId); const ev = db.usuarios.find(x => x.id === a.evaluadorId).nombre;
      const id = 'g' + Math.random().toString(36).slice(2); db.archivos[id] = { nombre: `${e.id} · rúbrica FIRMADA · ${ev}${(String(p.nombre).match(/\.[a-z0-9]+$/i) || ['.pdf'])[0]}`, b64: p.b64 };
      Object.assign(a, { firmadaId: id, firmadaFecha: ahora(), firmadaPor: u.nombre });
      hist(e.id, u, 'Rúbrica firmada subida', ev);
      if (a.evaluadorId === u.id) aviso('auxiliar', `Rúbrica firmada · ${e.id}`, 'envio/' + e.id, `${ev} subió su rúbrica firmada`);
      return { asignacion: a };
    },
    reclasificar: (p, u) => {
      exigir(u, 'auxiliar'); const e = db.envios.find(x => x.id === p.envioId); if (!e) throw new Error('No existe ese envío.');
      if (!CARPETAS[p.decision]) throw new Error('Elige el nuevo resultado.');
      if (['listo_publicar', 'publicado'].includes(e.estado)) throw new Error('Este texto ya está en plataforma: no se puede cambiar su resultado.');
      if (!String(p.nota || '').trim()) throw new Error('Escribe el motivo del cambio.');
      const antes = e.decision; e.decision = p.decision; e.archivo = CARPETAS[p.decision];
      e.estado = { aceptada: 'revision_directora', ajustes: 'ajustes', no: 'no_aceptado' }[p.decision]; e.actualizado = ahora();
      hist(e.id, u, 'Resultado cambiado', `${DICT[antes] || 'Sin resultado'} → ${DICT[p.decision]} · ${p.nota}`);
      if (p.decision !== 'no') aviso('directora', `${e.id}: ${ESTADOS[e.estado]}`, 'envio/' + e.id, `La Auxiliar cambió el resultado a «${DICT[p.decision]}»: ${p.nota}`);
      return { envio: envioPara(e, u) };
    },
    borradorDirectora: (p, u) => { exigir(u, 'directora'); const e = db.envios.find(x => x.id === p.envioId); e.dirCorr = e.dirCorr || {}; e.dirCorr.borrador = { anotaciones: p.anotaciones || [], nota: p.nota || '', version: e.versiones.slice(-1)[0].n }; return { guardado: ahora() }; },
    borrarEnvio: (p, u) => { exigir(u, 'auxiliar', 'admin'); const e = db.envios.find(x => x.id === p.id); if (!e) throw new Error('No existe ese envío.'); if (String(p.confirmar || '').trim().toUpperCase() !== e.id) throw new Error('Escribe el código del envío para confirmar.'); const ids = db.asignaciones.filter(x => x.envioId === e.id).map(x => x.id); db.asignaciones = db.asignaciones.filter(x => x.envioId !== e.id); db.historial = db.historial.filter(h => h.envioId !== e.id); db.notificaciones = (db.notificaciones || []).filter(n => n.ruta !== 'envio/' + e.id && !ids.some(id => n.ruta === 'evaluacion/' + id)); db.envios = db.envios.filter(x => x !== e); const h = db.hilos.find(x => x.envioId === e.id); if (h) h.envioId = null; return { borrado: e.id }; },
    decidir: (p, u) => { exigir(u, 'directora'); const e = db.envios.find(x => x.id === p.envioId); if (p.decision === 'corregirAprobar') { if (!(p.anotaciones || []).length) throw new Error('Haz al menos una corrección (o usa «Aprobar para publicar»).'); const v = e.versiones.slice(-1)[0]; let archivoId = ''; if (p.correccionDocx) { archivoId = 'd' + Date.now(); db.archivos[archivoId] = { nombre: `${e.id} · correcciones de la Dirección.docx`, b64: p.correccionDocx }; } if (p.finalDocx) { const fid = 'fin' + Date.now(); db.archivos[fid] = { nombre: `${e.id} · texto FINAL con correcciones de la Dirección.docx`, b64: p.finalDocx }; e.versiones.push({ n: v.n + 1, id: fid, nombre: 'Texto final con correcciones de la Dirección.docx', fecha: ahora(), deDirectora: true }); e.textoId = fid; } e.dirCorr = { version: v.n, anotaciones: p.anotaciones || [], nota: p.nota || '', archivoId, fecha: ahora(), aplicadas: true }; e.estado = 'aceptado'; hist(e.id, u, 'Aprobada con correcciones de la Directora', `${p.anotaciones.length} corrección(es) aplicadas${p.nota ? ' · ' + p.nota : ''}`); aviso('auxiliar', `Aprobada con correcciones de la Directora · ${e.id}`, 'envio/' + e.id, 'Avísale a la autora y envíala a plataforma: el texto final ya tiene las correcciones.'); e.actualizado = ahora(); return { envio: envioPara(e, u) }; } if (p.decision === 'corregir') { if (!(p.anotaciones || []).length && !p.nota) throw new Error('Haz al menos una corrección o escribe una nota.'); let archivoId = ''; if (p.correccionDocx) { archivoId = 'd' + Date.now(); db.archivos[archivoId] = { nombre: `${e.id} · correcciones de la Dirección.docx`, b64: p.correccionDocx }; } e.dirCorr = { version: e.versiones.slice(-1)[0].n, anotaciones: p.anotaciones || [], nota: p.nota || '', archivoId, fecha: ahora() }; e.estado = 'correcciones_directora'; hist(e.id, u, 'Correcciones de la Directora', `${(p.anotaciones || []).length} corrección(es)${p.nota ? ' · ' + p.nota : ''}`); aviso('auxiliar', `Correcciones de la Directora · ${e.id}`, 'envio/' + e.id, 'Envíaselas al autor para que las revise.'); e.actualizado = ahora(); return { envio: envioPara(e, u) }; } if (p.decision === 'confirmar') { e.estado = 'aceptado'; hist(e.id, u, 'Aceptación confirmada', p.nota || ''); } else { e.estado = 'devuelto'; e.nota = p.nota || ''; hist(e.id, u, 'Devuelto a la Auxiliar', p.nota || ''); } aviso('auxiliar', `${p.decision === 'confirmar' ? 'La Directora confirmó' : 'La Directora devolvió'} · ${e.id}`, 'envio/' + e.id, p.nota || e.titulo); e.actualizado = ahora(); return { envio: envioPara(e, u) }; },
    marcarPublicado: (p, u) => { exigir(u, 'plataforma'); const e = db.envios.find(x => x.id === p.envioId); e.estado = 'publicado'; e.publicadoUrl = p.url || ''; e.archivo = 'Publicados'; hist(e.id, u, 'Publicado en la gaceta', p.url || ''); aviso('auxiliar', `Publicado · ${e.id}`, 'envio/' + e.id, e.titulo); return { envio: envioPara(e, u) }; },
    usuarios: (p, u) => { exigir(u, 'admin'); return { usuarios: db.usuarios.map(pub) }; },
    guardarUsuario: (p, u) => {
      exigir(u, 'admin'); const d = p.usuario; if (!d.nombre) throw new Error('Falta el nombre.'); if (!(d.roles || []).length) throw new Error('Elige al menos un perfil.');
      let x = d.id && db.usuarios.find(y => y.id === d.id), temporal = null;
      if (x && !d.roles.includes('admin') && x.roles.includes('admin') && db.usuarios.filter(y => y.roles.includes('admin') && y.activo).length <= 1) throw new Error('Debe quedar al menos una persona con Administración.');
      if (!x) { temporal = 'rio-' + Math.floor(1000 + Math.random() * 9000); x = { id: sigId('U', db.usuarios), clave: temporal, debeCambiar: true }; db.usuarios.push(x); }
      Object.assign(x, { nombre: d.nombre, correo: d.correo || '', roles: d.roles, activo: d.activo !== false, avisos: d.avisos !== false });
      return { usuario: pub(x), claveTemporal: temporal };
    },
    borrarUsuario: (p, u) => { exigir(u, 'admin'); if (p.id === u.id) throw new Error('No puedes borrar tu propio perfil.'); const pend = db.asignaciones.filter(a => a.evaluadorId === p.id && a.estado !== 'entregada').length; if (pend) throw new Error(`Tiene ${pend} evaluación(es) pendiente(s). Reasígnalas antes, o desactiva el perfil.`); db.usuarios = db.usuarios.filter(x => x.id !== p.id); return {}; },
    restablecerClave: (p, u) => { exigir(u, 'admin'); const x = db.usuarios.find(y => y.id === p.id); const t = 'rio-' + Math.floor(1000 + Math.random() * 9000); x.clave = t; x.debeCambiar = true; return { claveTemporal: t }; },
    notificaciones: (p, u) => { const n = (db.notificaciones || []).filter(x => x.usuarioId === u.id).sort((a, b) => b.fecha.localeCompare(a.fecha)); return { notificaciones: n.slice(0, 60), sinLeer: n.filter(x => !x.leida).length }; },
    marcarLeidas: (p, u) => { (db.notificaciones || []).forEach(x => { if (x.usuarioId === u.id && (!p.ids || p.ids.includes(x.id))) x.leida = true; }); return {}; },
    ajustes: (p, u) => ({ ajustes: db.ajustes }),
    guardarAjustes: (p, u) => { exigir(u, 'admin'); Object.assign(db.ajustes, p.ajustes || {}); return { ajustes: db.ajustes }; },
  };
  const PUBLICAS = ['perfiles', 'entrar'];

  async function llamarDemo(accion, datos = {}) {
    await cargarDemo();
    await new Promise(r => setTimeout(r, 120));
    const fn = ACC[accion]; if (!fn) throw new Error('Acción desconocida: ' + accion);
    const u = PUBLICAS.includes(accion) ? null : sesion(datos.token);
    const r = JSON.parse(JSON.stringify(fn(datos, u) || {}));
    guardarDB();
    return { ok: true, ...r };
  }

  window.API = {
    demo: DEMO,
    llamar: DEMO ? llamarDemo : llamarServidor,
    reiniciarDemo: async () => { await escribiendo; try { await IDB.borrar(CLAVE); } catch (e) { } try { localStorage.removeItem(CLAVE); } catch (e) { } db = null; await cargarDemo(); await escribiendo; },
    esperarGuardado: () => escribiendo,
    enviadosDemo: () => (db ? db.enviados.slice().reverse() : []),
  };
})();
