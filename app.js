/* =========================================================
   Redacción · Hicimos lo que quisimos
   Página de evaluación editorial (perfiles: Administración, Directora,
   Auxiliar, Jefe de plataforma y Evaluadores).
   ========================================================= */
(function () {
  'use strict';
  const $ = (s, c = document) => c.querySelector(s);
  const $$ = (s, c = document) => Array.from(c.querySelectorAll(s));
  const esc = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  const app = $('#app');

  const ROLES = { admin: 'Administración', directora: 'Directora general', auxiliar: 'Auxiliar', plataforma: 'Jefe de plataforma', evaluador: 'Evaluador/a' };
  const DICT = { aceptada: 'Aceptada', ajustes: 'Aceptada con ajustes', no: 'No aceptada' };
  const GENEROS = ['Poesía', 'Cuento', 'Microrrelato', 'Crónica', 'Ensayo', 'Formas híbridas', 'Fragmento de obra mayor', 'Dramaturgia', 'Historieta o cómic', 'Cortometraje o performance', 'Videopoema o videoensayo', 'Obra sonora o podcast', 'Narrativa para redes', 'Narrativa digital'];
  const RUB_DE_GENERO = { 'poes': 'poemas', 'poema': 'poemas', 'cuento': 'cuentos', 'micro': 'microrrelato', 'crónica': 'cronica', 'cronica': 'cronica', 'ensayo': 'ensayo', 'híbrid': 'formas_hibridas', 'hibrid': 'formas_hibridas', 'fragmento': 'fragmento', 'novela': 'fragmento', 'drama': 'dramaturgia', 'teatro': 'dramaturgia', 'historieta': 'historieta_comic', 'cómic': 'historieta_comic', 'comic': 'historieta_comic', 'corto': 'cortometraje_performance', 'performance': 'cortometraje_performance', 'videopoema': 'videopoema_videoensayo', 'videoensayo': 'videopoema_videoensayo', 'podcast': 'obra_sonora_podcast', 'sonora': 'obra_sonora_podcast', 'redes': 'narrativa_redes', 'digital': 'narrativas_digitales' };
  const rubricaDeGenero = g => { const n = String(g || '').toLowerCase(); const k = Object.keys(RUB_DE_GENERO).find(x => n.includes(x)); return k ? RUB_DE_GENERO[k] : ''; };

  let S = { token: null, usuario: null, resumen: {}, notif: { lista: [], sinLeer: 0 }, rubricas: [], plantillasRub: {} };
  try { const g = JSON.parse(sessionStorage.getItem('redaccion-sesion') || 'null'); if (g) Object.assign(S, g); } catch (e) { }
  const guardarSesion = () => { try { sessionStorage.setItem('redaccion-sesion', JSON.stringify({ token: S.token, usuario: S.usuario })); } catch (e) { } };
  const tiene = r => S.usuario && S.usuario.roles.includes(r);

  // ---------- Utilidades ----------
  function toast(msg, error) {
    const t = $('#toast'); t.textContent = msg; t.classList.toggle('error', !!error); t.classList.add('visible');
    clearTimeout(t._t); t._t = setTimeout(() => t.classList.remove('visible'), error ? 5000 : 2800);
  }
  async function api(accion, datos = {}) {
    try { return await API.llamar(accion, { token: S.token, ...datos }); }
    catch (e) {
      if (e.sesion === false) { S.token = null; S.usuario = null; guardarSesion(); location.hash = '#entrar'; }
      throw e;
    }
  }
  const fecha = f => { if (!f) return ''; const d = new Date(f); return d.toLocaleDateString('es-CO', { day: 'numeric', month: 'short', year: 'numeric' }); };
  const fechaHora = f => { if (!f) return ''; const d = new Date(f); return d.toLocaleString('es-CO', { day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit' }); };
  const hace = f => { const m = Math.round((Date.now() - new Date(f)) / 60000); if (m < 1) return 'ahora'; if (m < 60) return `hace ${m} min`; const h = Math.round(m / 60); if (h < 24) return `hace ${h} h`; const d = Math.round(h / 24); return d === 1 ? 'ayer' : `hace ${d} días`; };
  const dias = f => Math.floor((Date.now() - new Date(f)) / 864e5);
  const punt = n => n == null || n === '' ? '—' : String((+n).toFixed(1)).replace('.', ',');
  const estadoChip = (e, t) => `<span class="estado e-${esc(e)}">${esc(t || e)}</span>`;
  const dictChip = d => d ? `<span class="dict d-${esc(d)}">${esc(DICT[d] || d)}</span>` : '';
  function descargar(nombre, bytes, tipo = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
    const url = URL.createObjectURL(new Blob([bytes], { type: tipo }));
    const a = document.createElement('a'); a.href = url; a.download = nombre; document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 4000);
  }
  async function descargarArchivo(datos, nombreSugerido) {
    try { const r = await api('archivo', datos); descargar(nombreSugerido || r.nombre, Docx.deBase64(r.b64), r.tipo || 'application/octet-stream'); }
    catch (e) { toast(e.message, true); }
  }
  const leerArchivoLocal = f => new Promise((ok, mal) => { const r = new FileReader(); r.onload = () => ok({ nombre: f.name, tipo: f.type, b64: String(r.result).split(',')[1] }); r.onerror = mal; r.readAsDataURL(f); });
  function modal(html, alAbrir) {
    const m = $('#modal'); $('.modal', m).innerHTML = html; m.hidden = false;
    const cerrar = () => { m.hidden = true; $('.modal', m).innerHTML = ''; };
    m.onclick = e => { if (e.target === m) cerrar(); };
    $$('[data-cerrar]', m).forEach(b => b.onclick = cerrar);
    if (alAbrir) alAbrir($('.modal', m), cerrar);
    const f = $('input, textarea, select, button', m); if (f) f.focus();
    return cerrar;
  }
  const confirmar = (titulo, texto, si = 'Sí, continuar') => new Promise(ok => modal(`<h2>${esc(titulo)}</h2><p class="suave">${texto}</p><div class="botones" style="margin-top:1rem"><button class="btn btn--rojo" data-si>${esc(si)}</button><button class="btn" data-cerrar>Cancelar</button></div>`, (m, cerrar) => { $('[data-si]', m).onclick = () => { cerrar(); ok(true); }; $$('[data-cerrar]', m).forEach(b => b.addEventListener('click', () => ok(false))); }));
  async function cargarRubricas() {
    if (S.rubricas.length) return S.rubricas;
    S.rubricas = await (await fetch('rubricas.json')).json(); return S.rubricas;
  }
  const rubrica = c => S.rubricas.find(r => r.clave === c);
  async function plantillaRubrica(clave) {
    if (!S.plantillasRub[clave]) { const url = rubrica(clave).archivo; const r = await fetch(url); S.plantillasRub[clave] = /\.json$/.test(url) ? Docx.deBase64((await r.json()).b64) : new Uint8Array(await r.arrayBuffer()); }
    return S.plantillasRub[clave];
  }
  const calcular = (niveles, n) => { const v = (niveles || []).filter(Boolean); if (v.length < n) return { puntaje: null, dictamen: null }; const p = Math.round(v.reduce((a, b) => a + b, 0) / n * 10) / 10; return { puntaje: p, dictamen: p >= 4.5 ? 'aceptada' : p >= 3.5 ? 'ajustes' : 'no' }; };

  // ---------- Estructura general ----------
  function pestanas() {
    const t = [['inicio', 'Inicio']];
    if (tiene('auxiliar')) t.push(['bandeja', 'Correo'], ['envios', 'Envíos', S.resumen.porHacer], ['evaluados', 'Textos evaluados', S.resumen.porHacerEval], ['plantillas', 'Mensajes']);
    if (tiene('evaluador')) t.push(['evaluaciones', 'Mis evaluaciones', S.resumen.pendientes]);
    if (tiene('directora')) t.push(['revisar', 'Para revisar', S.resumen.porRevisar]);
    if (tiene('plataforma')) t.push(['publicar', 'Listos para publicar', S.resumen.listos]);
    if (tiene('admin')) t.push(['perfiles', 'Perfiles'], ['ajustes', 'Ajustes']);
    return t;
  }
  function marco(contenido, actual) {
    const u = S.usuario;
    return `${API.demo ? `<div class="demo">Modo demostración · nada se envía de verdad · contraseña de todos: <b>demo</b> · <button data-reiniciar>Reiniciar datos</button><button data-simulados>Ver correos simulados</button></div>` : ''}
    <header class="barra">
      <div class="barra__fila">
        <a class="logo" href="#inicio">Hicimos <b>lo que</b> <i>quisimos</i><small>Redacción</small></a>
        <span class="barra__espacio"></span>
        <button class="btn btn--icono campana" data-campana aria-label="Notificaciones">${svgCampana}${S.notif.sinLeer ? `<span class="num">${S.notif.sinLeer > 99 ? '99+' : S.notif.sinLeer}</span>` : ''}</button>
        <button class="usuario-btn" data-usuario aria-label="Mi cuenta"><span class="avatar">${esc((u.nombre || '?')[0])}</span><span class="nombre">${esc(u.nombre.split(' ')[0])}</span></button>
      </div>
      <nav class="pestanas" aria-label="Secciones">${pestanas().map(([r, n, num]) => `<a class="pestana" href="#${r}" ${actual === r ? 'aria-current="page"' : ''}>${n}${num ? `<span class="num">${num}</span>` : ''}</a>`).join('')}</nav>
    </header>
    <main id="principal">${contenido}</main>`;
  }
  const svgCampana = '<svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6 8a6 6 0 1 1 12 0c0 7 3 8 3 8H3s3-1 3-8"/><path d="M10.3 20a1.9 1.9 0 0 0 3.4 0"/></svg>';
  function conectarMarco() {
    const r = $('[data-reiniciar]'); if (r) r.onclick = async () => { await API.reiniciarDemo(); S.token = null; S.usuario = null; guardarSesion(); toast('Datos de demostración reiniciados'); location.hash = '#entrar'; ruta(); };
    const sim = $('[data-simulados]'); if (sim) sim.onclick = () => { const l = API.enviadosDemo(); modal(`<h2>Correos simulados</h2><p class="suave">En la versión real, estos mensajes salen del Gmail de la gaceta.</p>${l.length ? l.slice(0, 40).map(x => `<div class="archivo"><span><b>${x.tipo === 'autor' ? 'Al autor' : 'Aviso'}</b> → ${esc(x.para)}: ${esc(x.texto)}${x.adjuntos ? ` · 📎${x.adjuntos}` : ''}</span><small class="meta">${hace(x.fecha)}</small></div>`).join('') : '<div class="vacio">Todavía no se ha enviado nada.</div>'}<div class="botones" style="margin-top:1rem"><button class="btn" data-cerrar>Cerrar</button></div>`); };
    $('[data-campana]').onclick = e => { e.stopPropagation(); abrirNotificaciones(); };
    $('[data-usuario]').onclick = e => { e.stopPropagation(); abrirMenu(); };
  }
  function cerrarFlotantes() { $$('.panel-notif, .menu').forEach(x => x.remove()); }
  document.addEventListener('click', e => { if (!e.target.closest('.panel-notif, .menu')) cerrarFlotantes(); });
  function abrirMenu() {
    if ($('.menu')) return cerrarFlotantes(); cerrarFlotantes();
    const m = document.createElement('div'); m.className = 'menu';
    m.innerHTML = `<div style="padding:.6rem .8rem"><b>${esc(S.usuario.nombre)}</b><div class="meta" style="margin-top:.2rem">${S.usuario.roles.map(r => ROLES[r]).join(' · ')}</div></div><a href="#cuenta">Cambiar contraseña</a><button data-salir>Salir</button>`;
    document.body.appendChild(m);
    $('[data-salir]', m).onclick = async () => { try { await api('salir'); } catch (e) { } S.token = null; S.usuario = null; guardarSesion(); cerrarFlotantes(); location.hash = '#entrar'; };
    $('a', m).onclick = cerrarFlotantes;
  }
  async function actualizarNotificaciones() {
    if (!S.token) return;
    try {
      const [n, r] = await Promise.all([api('notificaciones'), api('resumen')]);
      const antes = S.notif.sinLeer; S.notif = { lista: n.notificaciones, sinLeer: n.sinLeer }; S.resumen = r.resumen;
      const c = $('[data-campana]'); if (c) c.innerHTML = svgCampana + (n.sinLeer ? `<span class="num">${n.sinLeer > 99 ? '99+' : n.sinLeer}</span>` : '');
      $$('.pestana').forEach(a => { const p = pestanas().find(x => '#' + x[0] === a.getAttribute('href')); if (p) a.innerHTML = p[1] + (p[2] ? `<span class="num">${p[2]}</span>` : ''); });
      document.title = (n.sinLeer ? `(${n.sinLeer}) ` : '') + 'Redacción · Hicimos lo que quisimos';
      if (n.sinLeer > antes && antes !== undefined && S._notifIniciada) toast('🔔 ' + (n.notificaciones.find(x => !x.leida) || {}).titulo);
      S._notifIniciada = true;
    } catch (e) { }
  }
  setInterval(actualizarNotificaciones, 60000);
  document.addEventListener('visibilitychange', () => { if (!document.hidden) actualizarNotificaciones(); });
  function abrirNotificaciones() {
    if ($('.panel-notif')) return cerrarFlotantes(); cerrarFlotantes();
    const p = document.createElement('div'); p.className = 'panel-notif';
    const l = S.notif.lista;
    p.innerHTML = `<header><b>Notificaciones</b>${S.notif.sinLeer ? '<button class="btn btn--chico" data-todas>Marcar todas como leídas</button>' : ''}</header>${l.length ? l.map(n => `<a class="notif ${n.leida ? '' : 'nueva'}" href="#${esc(n.ruta || 'inicio')}" data-id="${esc(n.id)}"><b>${esc(n.titulo)}</b>${n.texto ? `<span>${esc(n.texto)}</span>` : ''}<small>${hace(n.fecha)}</small></a>`).join('') : '<div class="vacio" style="margin:1rem">No tienes notificaciones.</div>'}`;
    document.body.appendChild(p);
    $$('.notif', p).forEach(a => a.onclick = async () => { cerrarFlotantes(); try { await api('marcarLeidas', { ids: [a.dataset.id] }); } catch (e) { } actualizarNotificaciones(); });
    const t = $('[data-todas]', p); if (t) t.onclick = async () => { await api('marcarLeidas'); await actualizarNotificaciones(); cerrarFlotantes(); abrirNotificaciones(); };
  }

  // ---------- Rutas ----------
  const RUTAS = {};
  async function ruta() {
    cerrarFlotantes(); $('#modal').hidden = true; quitarHerramientas();
    const h = location.hash.replace(/^#/, '') || 'inicio';
    const [nombre, ...resto] = h.split('/'); const arg = resto.join('/');
    if (!S.token || !S.usuario) return vistaEntrar();
    if (S.usuario.debeCambiar && nombre !== 'cuenta') { location.hash = '#cuenta'; return; }
    const fn = RUTAS[nombre] || RUTAS.inicio;
    app.innerHTML = marco('<p class="suave">Cargando…</p>', nombre); conectarMarco();
    try { const html = await fn.render(arg); if (location.hash.replace(/^#/, '').split('/')[0] !== nombre && nombre !== 'inicio') return; app.innerHTML = marco(html, nombre); conectarMarco(); if (fn.conectar) await fn.conectar(arg); window.scrollTo(0, 0); }
    catch (e) { app.innerHTML = marco(`<div class="vacio"><span class="mano">Algo salió mal</span>${esc(e.message)}</div>`, nombre); conectarMarco(); }
  }
  window.addEventListener('hashchange', ruta);

  // ---------- Entrar ----------
  async function vistaEntrar() {
    app.innerHTML = `${API.demo ? '<div class="demo">Modo demostración · contraseña de todos: <b>demo</b></div>' : ''}<div class="entrada"><div class="entrada__caja">
      <p class="meta">Gaceta Hicimos lo que quisimos · Equipo editorial</p>
      <h1>Redacción</h1>
      <p class="suave" style="margin-bottom:1.8rem">Redacción: recepción, evaluación y publicación de obras.</p>
      <form class="tarjeta" id="f-entrar">
        <label class="campo"><span>Tu perfil</span><select name="id" required><option value="">Cargando perfiles…</option></select></label>
        <label class="campo"><span>Contraseña</span><input type="password" name="clave" required autocomplete="current-password"></label>
        <button class="btn btn--rojo" style="width:100%">Entrar</button>
        <p class="suave" style="font-size:.88rem;margin:1rem 0 0">¿Olvidaste tu contraseña? Pídele a Sebastian o a Luisa que te asignen una nueva.</p>
      </form></div></div>`;
    try {
      const r = await API.llamar('perfiles');
      $('#f-entrar [name=id]').innerHTML = '<option value="">Elige tu nombre</option>' + r.perfiles.map(p => `<option value="${esc(p.id)}">${esc(p.nombre)}</option>`).join('');
      let ult = null; try { ult = localStorage.getItem('redaccion-ultimo'); } catch (e) { } if (ult) $('#f-entrar [name=id]').value = ult;
    } catch (e) { toast('No se pudo conectar con el servidor: ' + e.message, true); }
    $('#f-entrar').onsubmit = async e => {
      e.preventDefault(); const f = new FormData(e.target); const b = $('button', e.target); b.disabled = true;
      try {
        const r = await API.llamar('entrar', { id: f.get('id'), clave: f.get('clave') });
        S.token = r.token; S.usuario = r.usuario; guardarSesion(); try { localStorage.setItem('redaccion-ultimo', f.get('id')); } catch (e) { }
        await actualizarNotificaciones();
        location.hash = r.usuario.debeCambiar ? '#cuenta' : '#inicio'; ruta();
      } catch (er) { toast(er.message, true); b.disabled = false; }
    };
  }

  // ---------- Cuenta ----------
  RUTAS.cuenta = {
    render: () => `<div style="max-width:520px">${S.usuario.debeCambiar ? `<div class="tarjeta" style="border-color:var(--canaguate)"><p class="mano" style="font-size:1.5rem;margin:0">¡Bienvenida/o!</p><p style="margin:0">Es tu primera vez o te asignaron una contraseña temporal. Crea tu propia contraseña para seguir.</p></div>` : ''}
      <h1>Cambiar contraseña</h1>
      <form class="tarjeta" id="f-clave">
        <label class="campo"><span>${S.usuario.debeCambiar ? 'Contraseña que te dieron' : 'Contraseña actual'}</span><input type="password" name="actual" required autocomplete="current-password"></label>
        <label class="campo"><span>Contraseña nueva (mínimo 8 caracteres)</span><input type="password" name="nueva" required minlength="8" autocomplete="new-password"></label>
        <label class="campo"><span>Repite la contraseña nueva</span><input type="password" name="otra" required minlength="8" autocomplete="new-password"></label>
        <button class="btn btn--rojo">Guardar contraseña</button>
      </form>
      ${tiene('evaluador') && !S.usuario.debeCambiar ? `<h2 style="margin-top:2rem">Mi firma</h2><div class="tarjeta" id="mi-firma"><p class="suave">Cargando…</p></div>` : ''}</div>`,
    conectar: () => {
      const caja = $('#mi-firma');
      const pintarFirma = async () => {
        if (!caja) return;
        let f = null; try { f = (await api('miFirma')).firma; } catch (e) { }
        caja.innerHTML = `<p class="suave">Cuando entregas una evaluación, tu firma <b>se pone sola en tu rúbrica</b> y queda firmada, sin imprimir ni escanear. La carga Sebastian o Luisa en «Perfiles».</p>
          ${f ? `<div class="firma-vista"><img src="data:image/png;base64,${f}" alt="Tu firma"></div>` : '<div class="vacio" style="padding:1rem">Todavía no tienes firma guardada. Envíale a Sebastian o a Luisa una foto de tu firma (en hoja blanca) para que la carguen.</div>'}`;
      };
      pintarFirma();
      $('#f-clave').onsubmit = async e => {
        e.preventDefault(); const f = new FormData(e.target);
        if (f.get('nueva') !== f.get('otra')) return toast('Las contraseñas nuevas no coinciden.', true);
        try { const r = await api('cambiarClave', { actual: f.get('actual'), nueva: f.get('nueva') }); S.usuario = r.usuario; guardarSesion(); toast('Contraseña guardada'); location.hash = '#inicio'; }
        catch (er) { toast(er.message, true); }
      };
    },
  };

  // ---------- Inicio ----------
  RUTAS.inicio = {
    render: async () => {
      await actualizarNotificaciones();
      const r = S.resumen; const tarj = [];
      if (tiene('auxiliar')) tarj.push(['#envios', r.porHacer, 'Envíos por asignar'], ['#evaluados', r.porHacerEval, 'Textos evaluados que necesitan algo de ti'], ['#bandeja', '✉', 'Revisar el correo de la gaceta']);
      if (tiene('evaluador')) tarj.push(['#evaluaciones', r.pendientes, 'Evaluaciones pendientes']);
      if (tiene('directora')) tarj.push(['#revisar', r.porRevisar, 'Obras para revisar']);
      if (tiene('plataforma')) tarj.push(['#publicar', r.listos, 'Listas para publicar']);
      if (tiene('admin')) tarj.push(['#perfiles', '☺', 'Administrar perfiles']);
      const nuevas = S.notif.lista.filter(n => !n.leida).slice(0, 5);
      return `<p class="meta" >${S.usuario.roles.map(x => ROLES[x]).join(' · ')}</p>
        <h1>Hola, ${esc(S.usuario.nombre.split(' ')[0])}</h1>
        <div class="rejilla" style="margin:1.5rem 0">${tarj.map(([h, n, t]) => `<a class="cifra-g" href="${h}"><span class="meta">${esc(t)}</span><b>${n ?? 0}</b></a>`).join('')}</div>
        <h2>Lo último</h2>${nuevas.length ? `<div class="tarjeta" style="padding:0">${nuevas.map(n => `<a class="notif nueva" href="#${esc(n.ruta)}" data-id="${esc(n.id)}"><b>${esc(n.titulo)}</b>${n.texto ? `<span>${esc(n.texto)}</span>` : ''}<small>${hace(n.fecha)}</small></a>`).join('')}</div>` : '<div class="vacio"><span class="mano">Todo al día</span>No tienes notificaciones nuevas.</div>'}`;
    },
    conectar: () => { $$('main .notif').forEach(a => a.onclick = () => api('marcarLeidas', { ids: [a.dataset.id] }).then(actualizarNotificaciones).catch(() => { })); },
  };

  // =========================================================
  //  AUXILIAR: correo
  // =========================================================
  RUTAS.bandeja = {
    render: async () => {
      let todo = false; try { todo = localStorage.getItem('redaccion-correo-todo') === '1'; } catch (e) { }
      const r = await api('bandeja', { todo });
      return `<div class="cabecera"><div><p class="meta">Gmail de la gaceta</p><h1>Correo</h1></div><button class="btn btn--chico" onclick="location.reload()">Actualizar</button></div>
      <div class="modo" role="group" aria-label="Qué correos ver" style="margin-bottom:.6rem"><button data-filtro="docs" aria-pressed="${!todo}">📄 Solo documentos</button><button data-filtro="todo" aria-pressed="${todo}">Todo el correo</button></div>
      <p class="suave" style="font-size:.88rem;margin:0 0 1rem">${todo ? 'Estás viendo todo el correo de la gaceta.' : 'Solo aparecen los correos que traen documentos adjuntos (Word, PDF…): los envíos de textos. El resto sigue en Gmail.'}</p>
      ${r.hilos.length ? `<ul class="lista">${r.hilos.map(h => `<li><a class="fila ${h.noLeido ? 'nuevo' : ''}" href="#hilo/${encodeURIComponent(h.hiloId)}">
        <div class="fila__cuerpo"><div class="meta">${esc(h.de.replace(/<.*>/, ''))} · ${hace(h.fecha)}${h.mensajes > 1 ? ` · ${h.mensajes} mensajes` : ''}</div>
        <div class="fila__titulo">${esc(h.asunto || '(sin asunto)')}</div><div class="fila__sub">${esc(h.fragmento)}</div>
        ${h.adjuntos.length ? `<div class="meta" style="margin-top:.3rem">📎 ${h.adjuntos.length} adjunto(s)</div>` : ''}</div>
        ${h.envioId ? `<span class="estado e-listo_publicar">${esc(h.envioId)}</span>` : (h.adjuntos.length ? '<span class="estado e-por_asignar">Sin registrar</span>' : '')}</a></li>`).join('')}</ul>` : `<div class="vacio">${todo ? 'No hay correos.' : '<span class="mano">Nada nuevo</span>No hay correos con documentos adjuntos.'}</div>`}`;
    },
    conectar: () => $$('[data-filtro]').forEach(b => b.onclick = () => { try { localStorage.setItem('redaccion-correo-todo', b.dataset.filtro === 'todo' ? '1' : '0'); } catch (e) { } ruta(); }),
  };
  RUTAS.hilo = {
    render: async id => {
      const h = await api('hilo', { hiloId: id }); S._hilo = { ...h, hiloId: id };
      const adj = h.mensajes.flatMap(m => m.adjuntos.map(a => ({ ...a, mensajeId: m.id })));
      const primero = h.mensajes[0];
      const correo = (primero.de.match(/<([^>]+)>/) || [0, primero.de])[1].trim();
      const nombre = primero.de.replace(/<.*>/, '').replace(/"/g, '').trim();
      const tituloAsunto = (h.asunto.split('|')[2] || '').trim();
      const generoAsunto = (h.asunto.split('|')[1] || '').trim();
      const rol = a => /ficha/i.test(a.nombre) ? 'ficha' : /licencia/i.test(a.nombre) ? 'licencia' : 'texto';
      let yaTexto = false; const roles = adj.map(a => { let r = rol(a); if (r === 'texto') { if (yaTexto) r = 'otro'; yaTexto = true; } return r; });
      await cargarRubricas();
      return `<div class="migas"><a href="#bandeja">Correo</a> › ${esc(h.asunto)}</div>
      <div class="dos"><div>
        <h1>${esc(h.asunto || '(sin asunto)')}</h1>
        ${h.mensajes.map(m => `<div class="mensaje"><div class="meta">${esc(m.de)} · ${fechaHora(m.fecha)}</div><pre>${esc(m.cuerpo)}</pre>
          ${m.adjuntos.map(a => `<div class="archivo"><span>📎 ${esc(a.nombre)}</span><button class="btn btn--chico" data-bajar="${esc(m.id)}|${a.idx}">Descargar</button></div>`).join('')}</div>`).join('')}
      </div><div>
        ${h.envioId ? `<div class="tarjeta"><p class="meta">Este correo ya es un envío</p><h2>${esc(h.envioId)}</h2><a class="btn btn--claro" href="#envio/${esc(h.envioId)}">Ver el envío</a></div>
          <div class="tarjeta"><h3>¿El autor mandó una versión corregida?</h3><p class="suave">Elige el archivo y súbelo como nueva versión del envío.</p>${adj.map(a => `<div class="archivo"><span>${esc(a.nombre)}</span><button class="btn btn--chico" data-version="${esc(a.mensajeId)}|${a.idx}">Usar como versión</button></div>`).join('') || '<p class="suave">Este hilo no tiene adjuntos.</p>'}</div>`
        : adj.length ? `<form class="tarjeta" id="f-registrar"><p class="meta" >Nuevo envío</p><h2>Registrar como envío</h2>
          <p class="suave" style="font-size:.92rem">Indica qué es cada adjunto. La obra se guarda sin datos del autor: los evaluadores solo verán el texto.</p>
          ${adj.map((a, i) => `<div class="adjunto-rol"><span style="overflow-wrap:anywhere">${esc(a.nombre)}</span><select name="rol${i}"><option value="texto" ${roles[i] === 'texto' ? 'selected' : ''}>Obra (texto)</option><option value="ficha" ${roles[i] === 'ficha' ? 'selected' : ''}>Ficha de autor</option><option value="licencia" ${roles[i] === 'licencia' ? 'selected' : ''}>Licencia</option><option value="otro" ${roles[i] === 'otro' ? 'selected' : ''}>Otro archivo</option><option value="">No guardar</option></select></div>`).join('')}
          <div style="margin-top:1rem"><button type="button" class="btn btn--chico" data-leer-ficha>Llenar con los datos de la ficha</button></div>
          <div style="margin-top:1rem">
          <label class="campo"><span>Título de la obra</span><input type="text" name="titulo" required value="${esc(tituloAsunto)}"></label>
          <label class="campo"><span>Género o formato</span><input type="text" name="genero" list="generos" value="${esc(generoAsunto)}"><datalist id="generos">${GENEROS.map(g => `<option>${g}</option>`).join('')}</datalist></label>
          <label class="campo"><span>Rúbrica sugerida</span><select name="rubrica"><option value="">Elegir después</option>${S.rubricas.map(r => `<option value="${r.clave}" ${rubricaDeGenero(generoAsunto) === r.clave ? 'selected' : ''}>${esc(r.nombre)}</option>`).join('')}</select></label>
          <label class="campo"><span>Nombre del autor (solo lo ven Auxiliar, Directora y Plataforma). Si tiene seudónimo: «Nombre (firma: Seudónimo)»; la carpeta en Drive usa el seudónimo.</span><input type="text" name="autorNombre" value="${esc(nombre)}"></label>
          <label class="campo"><span>Correo del autor</span><input type="email" name="autorCorreo" value="${esc(correo)}"></label>
          <label class="chequeo"><input type="checkbox" name="avisar" checked> Enviarle al autor el mensaje de «Recibido»</label></div>
          <button class="btn btn--rojo" style="width:100%;margin-top:.8rem">Registrar envío</button></form>`
        : '<div class="tarjeta"><p class="suave">Este correo no trae adjuntos. Puedes responderlo directamente desde Gmail.</p></div>'}
      </div></div>`;
    },
    conectar: id => {
      const h = S._hilo; const adj = h.mensajes.flatMap(m => m.adjuntos.map(a => ({ ...a, mensajeId: m.id })));
      $$('[data-bajar]').forEach(b => b.onclick = async () => { const [m, i] = b.dataset.bajar.split('|'); try { const r = await api('adjunto', { mensajeId: m, idx: +i }); descargar(r.nombre, Docx.deBase64(r.b64), r.tipo || 'application/octet-stream'); } catch (e) { toast(e.message, true); } });
      $$('[data-version]').forEach(b => b.onclick = async () => { const [m, i] = b.dataset.version.split('|'); try { const r = await api('adjunto', { mensajeId: m, idx: +i }); await api('subirVersion', { envioId: h.envioId, nombre: r.nombre, b64: r.b64, tipo: r.tipo, desdeCorreo: API.demo ? undefined : { mensajeId: m, idx: +i } }); toast('Versión nueva guardada'); location.hash = '#envio/' + h.envioId; } catch (e) { toast(e.message, true); } });
      const f = $('#f-registrar'); if (!f) return;
      f.genero.oninput = () => { const r = rubricaDeGenero(f.genero.value); if (r) f.rubrica.value = r; };
      $('[data-leer-ficha]').onclick = async () => {
        const i = adj.findIndex((a, k) => f['rol' + k].value === 'ficha'); if (i < 0) return toast('Marca primero cuál adjunto es la ficha.', true);
        if (!/\.docx$/i.test(adj[i].nombre)) return toast('La ficha no es un Word (.docx); llena los datos a mano.', true);
        try {
          const r = await api('adjunto', { mensajeId: adj[i].mensajeId, idx: adj[i].idx });
          const ps = (await Docx.leerParrafos(Docx.deBase64(r.b64))).map(p => p.texto.trim()).filter(Boolean);
          const dato = re => { for (let k = 0; k < ps.length; k++) { const m = ps[k].match(new RegExp('^' + re + '\\s*:?\\s*(.*)$', 'i')); if (m) { if (m[1] && !/^\(|^tal como|^poesía, cuento|^seudónimo|^estudios/i.test(m[1])) return m[1]; const sig = ps[k + 1]; if (sig && !/^(nombre|ocupaci|correo|ciudad|t[ií]tulo|g[eé]nero|difusi)/i.test(sig) && !/^(Seudónimo o forma|Estudios, oficio|Tal como aparece|Poesía, cuento|¿Circuló)/i.test(sig)) return sig; } } return ''; };
          const n = dato('Nombre completo'), pub = dato('Nombre de publicaci[oó]n'), c = dato('Correo electr[oó]nico'), t = dato('T[ií]tulo de la obra'), g = dato('G[eé]nero o formato');
          if (n) f.autorNombre.value = n + (pub ? ` (firma: ${pub})` : ''); if (c) f.autorCorreo.value = c; if (t) f.titulo.value = t; if (g) { f.genero.value = g; f.genero.oninput(); }
          toast(n || t ? 'Datos copiados de la ficha. Revísalos.' : 'No encontré datos en la ficha; llénalos a mano.', !(n || t));
        } catch (e) { toast('No pude leer la ficha: ' + e.message, true); }
      };
      f.onsubmit = async e => {
        e.preventDefault(); const d = new FormData(f); const b = $('button.btn--rojo', f); b.disabled = true;
        const de = rol => { const i = adj.findIndex((a, k) => d.get('rol' + k) === rol); return i < 0 ? null : { mensajeId: adj[i].mensajeId, idx: adj[i].idx }; };
        const otros = adj.map((a, k) => d.get('rol' + k) === 'otro' ? { mensajeId: a.mensajeId, idx: a.idx } : null).filter(Boolean);
        try {
          const r = await api('registrar', { hiloId: h.hiloId, titulo: d.get('titulo'), genero: d.get('genero'), rubrica: d.get('rubrica'), autorNombre: d.get('autorNombre'), autorCorreo: d.get('autorCorreo'), texto: de('texto'), ficha: de('ficha'), licencia: de('licencia'), otros });
          if (d.get('avisar')) { const pl = (await api('plantillas')).plantillas.find(x => x.clave === 'recibido'); if (pl) await api('responderAutor', { envioId: r.envio.id, tipo: 'recibido', asunto: rellenar(pl.asunto, r.envio), cuerpo: rellenar(pl.cuerpo, r.envio) }); }
          toast(`Envío ${r.envio.id} registrado${d.get('avisar') ? ' y autor avisado' : ''}`); location.hash = '#envio/' + r.envio.id;
        } catch (er) { toast(er.message, true); b.disabled = false; }
      };
    },
  };
  const rellenar = (t, e, obs) => String(t || '').replace(/\{titulo\}/g, e.titulo || '').replace(/\{autor\}/g, (e.autorNombre || '').split('(')[0].trim()).replace(/\{codigo\}/g, e.id || '').replace(/\{observaciones\}/g, obs || '');
  // Comentarios de los evaluadores para la autora, sin nombres (las rúbricas y los puntajes no se envían)
  const observacionesPara = e => (e.asignaciones || []).filter(a => a.estado === 'entregada' && +a.ronda === +e.ronda && String(a.observaciones || '').trim())
    .map((a, i) => `Lector/a ${i + 1}: ${String(a.observaciones).trim()}`).join('\n\n');
  // Si el mensaje de «No aceptado» o «Con ajustes» no trae {observaciones}, se agregan antes de la despedida
  function conObservaciones(cuerpo, tipo, obs) {
    if (!obs || !['no', 'ajustes'].includes(tipo) || /\{observaciones\}/.test(cuerpo)) return cuerpo;
    const bloque = `Estos son los comentarios de quienes leyeron tu texto:\n\n${obs}`;
    const i = cuerpo.search(/\n\s*(Un abrazo|Saludos|Cordialmente)/i);
    return i > -1 ? cuerpo.slice(0, i) + '\n\n' + bloque + cuerpo.slice(i) : cuerpo + '\n\n' + bloque;
  }

  // =========================================================
  //  ENVÍOS (Auxiliar, Directora, Plataforma)
  // =========================================================
  // «Envíos»: textos registrados desde el correo que están entrando a evaluación o evaluándose
  const GRUPOS_AUX = [
    ['Por asignar evaluadores', ['por_asignar']],
    ['Necesitan tercer evaluador', ['tercer_evaluador']],
    ['En evaluación', ['en_evaluacion']],
  ];
  function filaEnvio(e) {
    return `<li><a class="fila" href="#envio/${esc(e.id)}"><div class="fila__cuerpo"><div class="meta">${esc(e.id)} · ${esc(e.genero || 'sin género')} · ${hace(e.actualizado || e.creado)}</div>
      <div class="fila__titulo">${esc(e.titulo)}</div>
      ${e.evaluadores && e.evaluadores.length ? `<div class="fila__sub">${e.evaluadores.map(v => `${esc(v.nombre)} ${v.estado === 'entregada' ? '✓ ' + esc(DICT[v.dictamen] || '') : '· pendiente'}`).join(' · ')}</div>` : ''}</div>${estadoChip(e.estado, e.estadoTexto)}</a></li>`;
  }
  RUTAS.envios = {
    render: async () => {
      const { envios } = await api('envios');
      return `<div class="cabecera"><div><p class="meta">Todos los envíos</p><h1>Envíos</h1></div><a class="btn btn--chico" href="#bandeja">Ir al correo</a></div>
        <p class="suave" style="margin:-.4rem 0 1rem">Los textos que registraste desde el correo y están en evaluación. Cuando los evaluadores entregan, pasan a <a href="#evaluados">Textos evaluados</a>.</p>
        ${GRUPOS_AUX.map(([t, est]) => { const l = envios.filter(e => est.includes(e.estado)); return l.length ? `<h2 style="margin-top:1.4rem">${t} <span class="meta">${l.length}</span></h2><ul class="lista">${l.map(filaEnvio).join('')}</ul>` : ''; }).join('') || `<div class="vacio"><span class="mano">Nada en evaluación</span>Registra los envíos nuevos desde el <a href="#bandeja">Correo</a>.${envios.length ? ' Los que ya tienen resultado están en <a href="#evaluados">Textos evaluados</a>.' : ''}</div>`}`;
    },
  };
  // ---------- Textos evaluados (Auxiliar): resultado de la evaluación ----------
  const RESULTADOS = [
    ['aceptada', 'Aceptados', 'd-aceptada'],
    ['ajustes', 'Aceptados con ajustes', 'd-ajustes'],
    ['dircorr', 'Correcciones de la Directora', 'd-ajustes'],
    ['aprobado', 'Aprobados para publicar', 'd-aceptada'],
    ['no', 'No aceptados', 'd-no'],
    ['publicado', 'Publicados', 'd-aceptada'],
  ];
  // Los publicados pasan a «Publicados»; los que están con correcciones de la Directora (por enviar o esperando a la autora), a su propia casilla
  const TE_TOCA = ['ajustes', 'devuelto', 'correcciones_directora', 'aceptado', 'aprobado', 'no_aceptado', 'version_recibida'];
  const teToca = e => TE_TOCA.includes(e.estado) || !!e.respondio;
  const chipRes = e => grupoDe(e) === 'aprobado' ? '<span class="dict d-aceptada">✓ Aprobado por la Directora</span>' : dictChip(e.decision);
  const chipResp = e => e.respondio ? '<span class="estado" style="border-color:var(--buchon);color:var(--buchon)">📩 Respondió con nueva versión</span>' : '';
  const grupoDe = e => e.estado === 'publicado' ? 'publicado'
    : ['aceptado', 'aprobado', 'listo_publicar'].includes(e.estado) ? 'aprobado'
    : (e.estado === 'correcciones_directora' || (['esperando_autor', 'version_recibida'].includes(e.estado) && e.corrDir)) ? 'dircorr' : e.decision;
  RUTAS.evaluados = {
    render: async () => {
      const { envios } = await api('envios');
      const con = envios.filter(e => RESULTADOS.some(r => r[0] === grupoDe(e)));
      let filtro = 'todos'; try { filtro = localStorage.getItem('redaccion-evaluados') || 'todos'; } catch (e) { }
      if (!['todos', ...RESULTADOS.map(r => r[0])].includes(filtro)) filtro = 'todos';
      const cuenta = k => con.filter(e => grupoDe(e) === k).length;
      const fila = e => `<li><a class="fila ${teToca(e) ? 'nuevo' : ''}" href="#envio/${esc(e.id)}"><div class="fila__cuerpo"><div class="meta">${esc(e.id)} · ${esc(e.genero || 'sin género')} · ${fecha(e.actualizado || e.creado)}</div>
        <div class="fila__titulo">${esc(e.titulo)}</div>
        <div class="fila__sub">${(e.evaluadores || []).filter(v => v.estado === 'entregada').map(v => `${esc(v.nombre)}: ${punt(v.puntaje)} · ${esc(DICT[v.dictamen] || '')} ${v.firmada ? '✍✓' : '✍✗'}`).join(' — ')}</div>
        <div class="meta" style="margin-top:.3rem">${(() => { const ev = (e.evaluadores || []).filter(v => v.estado === 'entregada'); const f = ev.filter(v => v.firmada).length; return `Rúbricas firmadas: ${f}/${ev.length}${f < ev.length ? ' ⚠' : ' ✓'}`; })()}${e.archivo ? ` · 📁 Archivo › ${esc(e.archivo)}` : ''}${grupoDe(e) === 'aprobado' && e.decision && e.decision !== 'aceptada' ? ` · Venía de: ${esc(DICT[e.decision] || e.decision)}${e.versionN > 1 ? ` (versión corregida ${e.versionN})` : ''}` : ''}</div></div>
        <div style="display:flex;flex-direction:column;align-items:flex-end;gap:.35rem">${teToca(e) ? '<span class="estado" style="background:var(--canaguate);color:#1b1b1b;border-color:var(--canaguate)">Te toca</span>' : ''}${chipResp(e)}${chipRes(e)}${estadoChip(e.estado, e.estadoTexto)}${['listo_publicar', 'publicado'].includes(e.estado) ? '' : `<button type="button" class="btn btn--chico" data-mover-res="${esc(e.id)}">Mover a…</button>`}</div></a></li>`;
      S._evaluados = con;
      return `<div class="cabecera"><div><p class="meta">Resultado de la evaluación</p><h1>Textos evaluados</h1><p class="suave">Los textos que ya pasaron por los evaluadores, según su resultado. Los que tienen correcciones de la Directora en curso van en su propia casilla; cuando la Directora aprueba un texto (venga de aceptado, de ajustes o de sus correcciones) pasa a «Aprobados para publicar», y cuando se publica, a «Publicados». La segunda etiqueta muestra en qué va cada uno ahora.</p></div></div>
        ${con.some(e => teToca(e)) ? `<div class="tarjeta" style="border-color:var(--canaguate)"><b>Te toca:</b> ${con.filter(e => teToca(e)).map(e => `<a href="#envio/${esc(e.id)}">${esc(e.id)} · ${esc(e.titulo)}</a> <span class="meta">(${esc(e.respondio ? 'la autora respondió con una versión nueva' : e.estadoTexto)})</span>`).join(' · ')}</div>` : ''}
        <div class="rejilla" style="margin:1rem 0 1.4rem">${RESULTADOS.map(([k, t]) => `<button type="button" class="cifra-g" data-res="${k}" style="text-align:left;cursor:pointer;${filtro === k ? 'border-color:var(--acento);background:rgba(91,141,239,.12)' : ''}"><span class="meta">${t}</span><b>${cuenta(k)}</b></button>`).join('')}</div>
        <div class="modo" role="group" aria-label="Filtrar" style="margin-bottom:1.2rem"><button data-res="todos" aria-pressed="${filtro === 'todos'}">Todos (${con.length})</button>${RESULTADOS.map(([k, t]) => `<button data-res="${k}" aria-pressed="${filtro === k}">${t}</button>`).join('')}</div>
        ${con.length ? RESULTADOS.filter(([k]) => filtro === 'todos' || filtro === k).map(([k, t]) => { const l = con.filter(e => grupoDe(e) === k).sort((x, y) => (teToca(y) ? 1 : 0) - (teToca(x) ? 1 : 0)); return `<h2 style="margin-top:1.2rem">${t} <span class="meta">${l.length}</span></h2>${l.length ? `<ul class="lista">${l.map(fila).join('')}</ul>` : '<div class="vacio" style="padding:1.2rem">Ninguno por ahora.</div>'}`; }).join('')
          : '<div class="vacio"><span class="mano">Aún no hay textos evaluados</span>Aquí aparecen cuando los evaluadores entregan sus rúbricas.</div>'}`;
    },
    conectar: () => {
      $$('[data-res]').forEach(b => b.onclick = () => { try { localStorage.setItem('redaccion-evaluados', b.dataset.res); } catch (e) { } ruta(); });
      $$('[data-mover-res]').forEach(b => b.onclick = ev => { ev.preventDefault(); ev.stopPropagation(); modalReclasificar(S._evaluados.find(x => x.id === b.dataset.moverRes)); });
    },
  };

  RUTAS.revisar = {
    render: async () => {
      const { envios } = await api('envios');
      const p = envios.filter(e => e.estado === 'revision_directora'), o = envios.filter(e => e.estado !== 'revision_directora');
      return `<div class="cabecera"><div><p class="meta">Directora general</p><h1>Para revisar</h1></div></div>
        ${p.length ? `<ul class="lista">${p.map(filaEnvio).join('')}</ul>` : '<div class="vacio"><span class="mano">Nada por revisar</span>Te avisamos con la campanita cuando llegue una obra aceptada.</div>'}
        ${o.length ? `<h2 style="margin-top:2rem">Seguimiento</h2><ul class="lista">${o.map(filaEnvio).join('')}</ul>` : ''}`;
    },
  };
  RUTAS.publicar = {
    render: async () => {
      const { envios } = await api('envios');
      const l = envios.filter(e => e.estado === 'listo_publicar'), p = envios.filter(e => e.estado === 'publicado');
      return `<div class="cabecera"><div><p class="meta">Jefe de plataforma</p><h1>Listos para publicar</h1></div></div>
        ${l.length ? `<ul class="lista">${l.map(filaEnvio).join('')}</ul>` : '<div class="vacio"><span class="mano">Nada pendiente</span>Aquí llegan las obras que completaron todo el proceso.</div>'}
        ${p.length ? `<h2 style="margin-top:2rem">Publicados</h2><ul class="lista">${p.map(filaEnvio).join('')}</ul>` : ''}`;
    },
  };

  // ---------- Detalle de un envío ----------
  RUTAS.envio = {
    render: async id => {
      await cargarRubricas();
      const { envio: e } = await api('envio', { id }); S._envio = e;
      const aux = tiene('auxiliar'), dir = tiene('directora'), pla = tiene('plataforma');
      const ronda = e.asignaciones.filter(a => +a.ronda === +e.ronda);
      const entregadas = e.asignaciones.filter(a => a.estado === 'entregada');
      const acciones = [];
      if (aux) {
        if (['por_asignar', 'tercer_evaluador', 'devuelto', 'en_evaluacion'].includes(e.estado)) acciones.push(`<button class="btn btn--rojo" data-asignar>${e.estado === 'tercer_evaluador' ? 'Asignar tercer evaluador' : 'Asignar evaluadores'}</button>`);
        const tipoAviso = { ajustes: 'ajustes', aceptado: 'aceptado', no_aceptado: 'no', correcciones_directora: 'directora' }[e.estado];
        if (e.estado === 'aprobado') acciones.push('<button class="btn btn--verde" data-plataforma>Enviar a plataforma</button>');
        acciones.push(`<button class="btn ${tipoAviso ? 'btn--ama' : ''}" data-escribir="${tipoAviso || ''}">${tipoAviso ? 'Avisar al autor' : 'Escribir al autor'}</button>`);
        if (['esperando_autor', 'ajustes', 'devuelto', 'version_recibida'].includes(e.estado)) acciones.push(`<button class="btn" data-subir-version>${e.estado === 'version_recibida' ? 'Subir otra versión' : 'Subir versión corregida'}</button>`);
        if (['esperando_autor', 'devuelto', 'version_recibida'].includes(e.estado)) acciones.push('<button class="btn btn--claro" data-mover="directora">Enviar a la Directora</button><button class="btn" data-mover="nueva_ronda">Nueva ronda de evaluación</button>');
        if (['aceptada', 'ajustes', 'no'].includes(e.decision) && !['listo_publicar', 'publicado'].includes(e.estado)) acciones.push('<button class="btn btn--chico" data-cambiar-res>Cambiar resultado</button>');
        acciones.push('<button class="btn btn--chico" data-editar>Editar datos</button>');
      }
      if (dir && e.estado === 'revision_directora') acciones.push(`<button class="btn btn--verde" data-decidir="confirmar">Aprobar para publicar</button><a class="btn btn--rojo" href="#corregir/${esc(e.id)}">Hacer correcciones${e.dirCorr && e.dirCorr.borrador && e.dirCorr.borrador.anotaciones.length ? ' (borrador)' : ''}</a><button class="btn" data-decidir="devolver">Devolver a la Auxiliar</button>`);
      if (tiene('admin') && !aux) acciones.push('<button class="btn btn--chico" data-borrar-admin style="color:var(--rojo);border-color:var(--rojo);margin-left:auto">Borrar envío</button>');
      if (pla && e.estado === 'listo_publicar') acciones.push('<button class="btn btn--verde" data-publicado>Marcar como publicada</button>');
      const r = rubrica(e.rubrica);
      return `<div class="migas"><a href="#${aux ? (DICT[e.decision] ? 'evaluados' : 'envios') : dir ? 'revisar' : 'publicar'}">${aux ? (DICT[e.decision] ? 'Textos evaluados' : 'Envíos') : dir ? 'Para revisar' : 'Listos para publicar'}</a> › ${esc(e.id)}</div>
      <div class="cabecera"><div><p class="meta">${esc(e.id)} · ${esc(e.genero || 'sin género')}${e.ronda > 1 ? ` · ronda ${e.ronda}` : ''}</p><h1>${esc(e.titulo)}</h1>${estadoChip(e.estado, e.estadoTexto)}</div></div>
      ${e.nota && e.estado === 'devuelto' ? `<div class="tarjeta" style="border-color:var(--canaguate)"><p class="meta">Nota de la Directora</p><p style="margin:0">${esc(e.nota)}</p></div>` : ''}
      ${aux && e.respondio ? `<div class="tarjeta" style="border-color:var(--buchon)"><b>📩 La autora respondió con un documento.</b> Ábrelo en el correo y pulsa <b>Usar como versión</b>; después, <b>Enviar a la Directora</b>.<div class="botones" style="margin-top:.6rem"><a class="btn btn--verde" href="#hilo/${encodeURIComponent(e.hiloId)}">Abrir el correo</a></div></div>` : ''}
      ${aux && e.estado === 'version_recibida' ? `<div class="tarjeta" style="border-color:var(--buchon)"><b>✓ Versión corregida guardada (versión ${(e.versiones || []).length}).</b> Revísala en «Archivos» y pulsa <b>Enviar a la Directora</b> para que la apruebe.</div>` : ''}
      ${dir && e.estado === 'revision_directora' && (e.versiones || []).length > 1 ? `<div class="tarjeta" style="border-color:var(--acento)"><b>Esta es la versión corregida</b> (versión ${(e.versiones || []).length}) que mandó la autora después de ${e.dirCorr && e.dirCorr.fecha ? 'tus correcciones' : 'los ajustes de los evaluadores'}. Si está bien, pulsa <b>Aprobar para publicar</b>: pasará a «Aprobados para publicar» y la Auxiliar la enviará a plataforma. Las versiones anteriores siguen en «Archivos».</div>` : ''}
      ${acciones.length ? `<div class="botones" style="margin-bottom:1.4rem">${acciones.join('')}</div>` : ''}
      <div class="dos"><div>
        ${e.dirCorr && e.dirCorr.fecha ? `<div class="tarjeta" style="border-color:var(--acento)"><p class="meta">Correcciones de la Directora · versión ${e.dirCorr.version} · ${fecha(e.dirCorr.fecha)}</p>${e.dirCorr.aplicadas ? '<p style="margin:.2rem 0 .5rem"><b>✓ Aplicadas en el texto final</b> (la última versión en «Archivos»).</p>' : ''}${e.dirCorr.nota ? `<p style="white-space:pre-wrap">${esc(e.dirCorr.nota)}</p>` : ''}<div class="botones">${e.dirCorr.anotaciones.length ? `<button class="btn btn--chico btn--claro" data-ver-dir>Ver correcciones (${e.dirCorr.anotaciones.length})</button>` : ''}${e.dirCorr.archivoId ? '<button class="btn btn--chico" data-archivo="directora">Texto corregido en Word</button>' : ''}</div></div>` : ''}
        ${!aux && !dir ? '' : `<h2>Evaluaciones</h2>`}
        ${!aux && !dir ? '' : entregadas.length ? `<div class="lado-a-lado">${entregadas.map((a, i) => tarjetaEvaluacion(a, i)).join('')}</div>` : '<div class="vacio">Todavía no hay evaluaciones entregadas.</div>'}
        ${aux && ronda.some(a => a.estado !== 'entregada') ? `<h3 style="margin-top:1.4rem">Pendientes</h3>${ronda.filter(a => a.estado !== 'entregada').map(a => `<div class="archivo"><span>${esc(a.evaluador)} · asignada ${hace(a.asignada)}${dias(a.asignada) >= 7 ? ' ⚠' : ''} · ${a.estado === 'borrador' ? 'ya empezó' : 'sin empezar'}</span><button class="btn btn--chico" data-quitar="${esc(a.id)}">Quitar</button></div>`).join('')}` : ''}
      </div><div>
        <div class="tarjeta"><h3>Datos</h3><dl class="datos">
          <dt>Código</dt><dd>${esc(e.id)}</dd><dt>Rúbrica</dt><dd>${esc(r ? r.nombre : 'Sin elegir')}</dd>
          ${aux || dir || pla ? `<dt>Autor/a</dt><dd>${esc(e.autorNombre || '—')}</dd><dt>Correo</dt><dd>${esc(e.autorCorreo || '—')}</dd>` : ''}
          <dt>Recibido</dt><dd>${fecha(e.creado)}</dd>${e.decision && DICT[e.decision] ? `<dt>Resultado</dt><dd>${dictChip(e.decision)}</dd>` : ''}${e.archivo ? `<dt>Carpeta</dt><dd>Archivo › ${esc(e.archivo)}</dd>` : ''}${e.carpetaUrl ? `<dt>Drive</dt><dd><a href="${esc(e.carpetaUrl)}" target="_blank" rel="noopener">Abrir la carpeta del envío ↗</a></dd>` : ''}${e.publicadoUrl ? `<dt>Publicado</dt><dd><a href="${esc(e.publicadoUrl)}" target="_blank" rel="noopener">${esc(e.publicadoUrl)}</a></dd>` : ''}
        </dl></div>
        <div class="tarjeta"><h3>Archivos</h3>
          ${(e.versiones || []).map(v => `<div class="archivo"><span>${(['listo_publicar', 'publicado', 'aprobado'].includes(e.estado) && v.n === e.versiones.length ? '<b>Texto final</b> · ' : '') + (e.versiones.length > 1 ? `Versión ${v.n} · ` : 'Obra · ')}${esc(v.nombre)}</span><button class="btn btn--chico" data-archivo="version|${v.n}">Descargar</button></div>`).join('')}
          ${e.fichaId ? '<div class="archivo"><span>Ficha de autor</span><button class="btn btn--chico" data-archivo="ficha">Descargar</button></div>' : ''}
          ${e.licenciaId ? '<div class="archivo"><span>Licencia de uso</span><button class="btn btn--chico" data-archivo="licencia">Descargar</button></div>' : ''}
          ${(e.otros || []).map((o, i) => `<div class="archivo"><span>${esc(o.nombre)}</span><button class="btn btn--chico" data-archivo="otro|${i}">Descargar</button></div>`).join('')}
        </div>
        ${(e.historial || []).length ? `<div class="tarjeta"><h3>Historial</h3><ul class="historial">${e.historial.slice().reverse().map(h => `<li><b>${esc(h.accion)}</b>${h.detalle ? ` · ${esc(h.detalle)}` : ''}<div class="meta">${esc(h.usuario)} · ${fechaHora(h.fecha)}</div></li>`).join('')}</ul></div>` : ''}
      </div></div>`;
    },
    conectar: () => {
      const e = S._envio;
      $$('[data-archivo]').forEach(b => b.onclick = () => { const [cual, n] = b.dataset.archivo.split('|'); descargarArchivo({ envioId: e.id, cual, n }); });
      $$('[data-ver-eval]').forEach(b => b.onclick = () => verEvaluacion(e, e.asignaciones.find(a => a.id === b.dataset.verEval)));
      $$('[data-rub-word]').forEach(b => b.onclick = () => wordRubrica(e, e.asignaciones.find(a => a.id === b.dataset.rubWord)));
      $$('[data-corr-word]').forEach(b => b.onclick = () => wordCorreccion(e, e.asignaciones.find(a => a.id === b.dataset.corrWord)));
      $$('[data-subido]').forEach(b => b.onclick = () => { const [aid, cual] = b.dataset.subido.split('|'); descargarArchivo({ asignacionId: aid, cual }); });
      $$('[data-quitar]').forEach(b => b.onclick = async () => { if (!await confirmar('¿Quitar a este evaluador?', 'Se borra su asignación y lo que haya avanzado.')) return; try { await api('quitarAsignacion', { id: b.dataset.quitar }); toast('Asignación quitada'); ruta(); } catch (er) { toast(er.message, true); } });
      const as = $('[data-asignar]'); if (as) as.onclick = () => modalAsignar(e);
      const es = $('[data-escribir]'); if (es) es.onclick = () => modalEscribir(e, es.dataset.escribir);
      const sv = $('[data-subir-version]'); if (sv) sv.onclick = () => modalVersion(e);
      const ed = $('[data-editar]'); if (ed) ed.onclick = () => modalEditar(e);
      $$('[data-mover]').forEach(b => b.onclick = async () => { try { await api('moverEnvio', { envioId: e.id, a: b.dataset.mover }); toast(b.dataset.mover === 'directora' ? 'Enviado a la Directora' : 'Lista para una nueva ronda'); ruta(); } catch (er) { toast(er.message, true); } });
      $$('[data-decidir]').forEach(b => b.onclick = () => {
        const confirma = b.dataset.decidir === 'confirmar';
        modal(`<h2>${confirma ? 'Aprobar para publicar' : 'Devolver a la Auxiliar'}</h2><p class="suave">${confirma ? 'La Auxiliar le avisará al autor y luego enviará el texto final, con la ficha y la licencia, a plataforma.' : 'Explica qué falta o qué debe revisarse.'}</p><label class="campo"><span>Nota ${confirma ? '(opcional)' : ''}</span><textarea name="nota"></textarea></label><div class="botones"><button class="btn ${confirma ? 'btn--verde' : 'btn--rojo'}" data-si>${confirma ? 'Aprobar' : 'Devolver'}</button><button class="btn" data-cerrar>Cancelar</button></div>`, (m, cerrar) => {
          $('[data-si]', m).onclick = async () => { const nota = $('[name=nota]', m).value; if (!confirma && !nota.trim()) return toast('Escribe la nota.', true); try { await api('decidir', { envioId: e.id, decision: b.dataset.decidir, nota }); cerrar(); toast(confirma ? 'Aprobado. La Auxiliar ya fue avisada.' : 'Devuelto a la Auxiliar'); ruta(); } catch (er) { toast(er.message, true); } };
        });
      });
      const ba = $('[data-borrar-admin]'); if (ba) ba.onclick = () => modalBorrarEnvio(e);
      $$('[data-rub-completa]').forEach(b => b.onclick = () => { const a = e.asignaciones.find(x => x.id === b.dataset.rubCompleta); rubricaCompleta(rubrica(a.rubrica), a); });
      $$('[data-firmada]').forEach(b => b.onclick = () => descargarArchivo({ asignacionId: b.dataset.firmada, cual: 'firmada' }));
      $$('[data-subir-firmada]').forEach(b => b.onclick = () => modalFirmada(b.dataset.subirFirmada));
      const cr = $('[data-cambiar-res]'); if (cr) cr.onclick = () => modalReclasificar(e);
      const vd = $('[data-ver-dir]'); if (vd) vd.onclick = () => verCorrDirectora(e);
      const pf = $('[data-plataforma]'); if (pf) pf.onclick = () => modalPlataforma(e);
      const pu = $('[data-publicado]'); if (pu) pu.onclick = () => modal(`<h2>Marcar como publicada</h2><p class="suave">Pega el enlace de la obra en la gaceta (opcional).</p><label class="campo"><span>Enlace en la gaceta</span><input type="url" name="url" placeholder="https://hicimosloquequisimos-cmd.github.io/obra/…"></label><div class="botones"><button class="btn btn--verde" data-si>Marcar como publicada</button><button class="btn" data-cerrar>Cancelar</button></div>`, (m, cerrar) => {
        $('[data-si]', m).onclick = async () => { try { await api('marcarPublicado', { envioId: e.id, url: $('[name=url]', m).value }); cerrar(); toast('¡Publicada!'); ruta(); } catch (er) { toast(er.message, true); } };
      });
    },
  };
  // Rúbrica como en el formato de la gaceta: criterios × niveles, con ☒ en lo que marcó el evaluador
  function matrizRubrica(r, niveles) {
    const corto = t => String(t || '').replace(/^\d\s·\s/, '');
    return `<div class="matriz-env"><table class="matriz"><thead><tr><th>Criterio</th>${r.niveles.map((n, j) => `<th title="${esc(n)}">${j + 1}</th>`).join('')}</tr></thead><tbody>
      ${r.criterios.map((c, k) => `<tr><td>${esc(c.nombre)}</td>${c.niveles.map((d, j) => `<td class="${niveles[k] === j + 1 ? 'marcada' : ''}" title="${esc(d)}">${niveles[k] === j + 1 ? '☒' : '☐'}</td>`).join('')}</tr>`).join('')}
      </tbody></table></div><p class="meta" style="margin:.3rem 0 0">${r.niveles.map((n, j) => `${j + 1} ${esc(corto(n))}`).join(' · ')}</p>`;
  }
  function rubricaCompleta(r, a) {
    modal(`<h2>Rúbrica · ${esc(r.nombre)}</h2><p class="suave">${esc(a.evaluador || 'Evaluador/a')} · puntaje <b>${punt(a.puntaje)}</b> · ${esc(DICT[a.dictamen] || '')}. Lo que marcó está resaltado con ☒.</p>
      <div class="matriz-env"><table class="matriz matriz--completa"><thead><tr><th>Criterio</th>${r.niveles.map(n => `<th>${esc(n)}</th>`).join('')}</tr></thead><tbody>
      ${r.criterios.map((c, k) => `<tr><td><b>${esc(c.nombre)}</b></td>${c.niveles.map((d, j) => `<td class="${a.niveles[k] === j + 1 ? 'marcada' : ''}"><span class="casilla">${a.niveles[k] === j + 1 ? '☒' : '☐'}</span> ${esc(d)}</td>`).join('')}</tr>`).join('')}
      </tbody></table></div>
      ${a.observaciones ? `<p class="meta" style="margin-top:1rem">Observaciones</p><p style="white-space:pre-wrap">${esc(a.observaciones)}</p>` : ''}
      <div class="botones" style="margin-top:1rem"><button class="btn" data-cerrar>Cerrar</button></div>`, () => { $('.modal', $('#modal')).style.width = 'min(1100px, 100%)'; });
  }
  function estadoFirma(a) {
    if (a.firmadaId) return `<div class="firma si">✓ Rúbrica firmada${a.firmadaId === a.rubricaId ? ' (firma digital)' : ''}${a.firmadaFecha ? ` · ${fecha(a.firmadaFecha)}` : ''}${a.firmadaPor && a.firmadaPor !== a.evaluador && a.firmadaId !== a.rubricaId ? ` · subida por ${esc(a.firmadaPor)}` : ''} <button class="btn btn--chico" data-firmada="${esc(a.id)}">Ver firmada</button>${tiene('auxiliar') ? ` <button class="btn btn--chico" data-subir-firmada="${esc(a.id)}">Reemplazar</button>` : ''}</div>`;
    return `<div class="firma no">⚠ Falta la rúbrica firmada${tiene('auxiliar') ? ` <button class="btn btn--chico" data-subir-firmada="${esc(a.id)}">Subir rúbrica firmada</button>` : ''}</div>`;
  }
  function modalFirmada(asignacionId, alTerminar) {
    modal(`<h2>Subir rúbrica firmada</h2><p class="suave">Sube la rúbrica <b>firmada</b> por quien evaluó: un PDF, una foto o un Word con firma digital. Queda guardada en la carpeta del envío en Drive.</p>
      <label class="campo"><span>Archivo (PDF, imagen o Word)</span><input type="file" name="f" accept=".pdf,.jpg,.jpeg,.png,.heic,.docx,.doc"></label>
      <div class="botones"><button class="btn btn--rojo" data-si>Subir</button><button class="btn" data-cerrar>Cancelar</button></div>`, (m, cerrar) => {
      $('[data-si]', m).onclick = async () => {
        const f = $('[name=f]', m).files[0]; if (!f) return toast('Elige el archivo.', true);
        if (f.size > 15e6) return toast('El archivo pesa demasiado (máximo 15 MB).', true);
        const b = $('[data-si]', m); b.disabled = true; b.textContent = 'Subiendo…';
        try { const d = await leerArchivoLocal(f); await api('subirFirmada', { asignacionId, ...d }); cerrar(); toast('Rúbrica firmada guardada ✓'); if (alTerminar) alTerminar(); else ruta(); }
        catch (er) { toast(er.message, true); b.disabled = false; b.textContent = 'Subir'; }
      };
    });
  }
  function tarjetaEvaluacion(a, i) {
    const r = rubrica(a.rubrica);
    return `<div class="tarjeta"><div class="meta">${tiene('auxiliar') || tiene('directora') ? esc(a.evaluador) : 'Evaluador/a ' + (i + 1)} · ${fecha(a.entregada)}${a.ronda > 1 ? ` · ronda ${a.ronda}` : ''}</div>
      <div style="display:flex;align-items:baseline;gap:.8rem;margin:.4rem 0 .6rem"><span style="font-size:1.9rem;font-weight:650;line-height:1">${punt(a.puntaje)}</span>${dictChip(a.dictamen)}</div>
      ${r && a.niveles && a.niveles.length ? matrizRubrica(r, a.niveles) + `<button class="btn btn--chico" style="margin-top:.5rem" data-rub-completa="${esc(a.id)}">Ver rúbrica completa</button>` : `<p class="suave">${a.modo === 'word' ? 'Calificó con su propio archivo de Word: la rúbrica está en «Su rúbrica (archivo)».' : ''}</p>`}
      ${estadoFirma(a)}
      ${a.observaciones ? `<p class="meta" style="margin-top:.8rem">Observaciones</p><p style="white-space:pre-wrap">${esc(a.observaciones)}</p>` : ''}
      <div class="botones" style="margin-top:.8rem">
        ${(a.anotaciones || []).length ? `<button class="btn btn--chico btn--claro" data-ver-eval="${esc(a.id)}">Ver correcciones (${a.anotaciones.length})</button>` : ''}
        ${a.niveles && a.niveles.length ? `<button class="btn btn--chico" data-rub-word="${esc(a.id)}">Rúbrica en Word</button>` : ''}
        ${(a.anotaciones || []).length ? `<button class="btn btn--chico" data-corr-word="${esc(a.id)}">Texto corregido en Word</button>` : ''}
        ${a.subidaRubricaId ? `<button class="btn btn--chico" data-subido="${esc(a.id)}|subidaRubrica">Su rúbrica (archivo)</button>` : ''}
        ${a.subidaCorreccionId ? `<button class="btn btn--chico" data-subido="${esc(a.id)}|subidaCorreccion">Su corrección (archivo)</button>` : ''}
      </div></div>`;
  }
  async function textoDe(a) {
    const r = await api('archivo', { asignacionId: a.id, cual: 'texto' });
    if (!/\.docx$/i.test(r.nombre)) throw new Error('El texto no es un Word (.docx).');
    return Docx.leerParrafos(Docx.deBase64(r.b64));
  }
  async function verEvaluacion(e, a) {
    try {
      const ps = await textoDe(a);
      modal(`<h2>Correcciones de ${esc(a.evaluador || 'evaluador/a')}</h2><p class="suave">${a.anotaciones.length} anotación(es). Toca una marca para ver su nota.</p><div class="hoja" style="max-height:60vh;overflow:auto;font-size:1rem">${pintarHoja(ps, a.anotaciones)}</div><div style="margin-top:1rem">${listaNotas(ps, a.anotaciones, false)}</div><div class="botones" style="margin-top:1rem"><button class="btn" data-corr>Descargar en Word</button><button class="btn" data-cerrar>Cerrar</button></div>`, m => {
        $('.modal', $('#modal')).style.width = 'min(900px, 100%)';
        $('[data-corr]', m).onclick = () => wordCorreccion(e, a);
        conectarMarcas(m, a.anotaciones);
      });
    } catch (er) { toast(er.message, true); }
  }
  async function wordRubrica(e, a) {
    if (a.rubricaId) return descargarArchivo({ asignacionId: a.id, cual: 'rubrica' });
    try {
      const fecha = new Date(a.entregada || Date.now()).toLocaleDateString('es-CO');
      const bytes = await Docx.llenarRubrica(await plantillaRubrica(a.rubrica), { titulo: e.titulo, codigo: e.id, fecha, evaluador: a.evaluador || S.usuario.nombre, enlace: a.enlace, niveles: a.niveles, puntaje: a.puntaje, dictamen: a.dictamen, observaciones: a.observaciones });
      descargar(`${e.id} · rúbrica · ${a.evaluador || 'evaluador'}.docx`, bytes);
    } catch (er) { toast('No se pudo crear la rúbrica: ' + er.message, true); }
  }
  async function wordCorreccion(e, a) {
    try {
      const ps = await textoDe(a);
      const n = (e.asignaciones || []).filter(x => x.estado === 'entregada').findIndex(x => x.id === a.id) + 1 || 1;
      const bytes = await Docx.crearCorregido({ parrafos: ps, anotaciones: a.anotaciones, autor: `Evaluador/a ${n} · Hicimos lo que quisimos`, iniciales: 'E' + n, titulo: `${e.id} · ${e.titulo}` });
      descargar(`${e.id} · texto corregido · ${a.evaluador || 'evaluador ' + n}.docx`, bytes);
    } catch (er) { toast('No se pudo crear el Word: ' + er.message, true); }
  }
  async function modalAsignar(e) {
    const { evaluadores } = await api('evaluadores');
    const ya = e.asignaciones.filter(a => +a.ronda === +e.ronda).map(a => a.evaluadorId);
    const sel = new Set();
    const sugerida = e.rubrica || rubricaDeGenero(e.genero);
    modal(`<h2>${e.estado === 'tercer_evaluador' ? 'Tercer evaluador' : 'Asignar evaluadores'}</h2><p class="suave">${e.estado === 'tercer_evaluador' ? 'Los dos evaluadores no coinciden: elige una persona más.' : 'Elige dos personas. Solo verán el texto, sin datos del autor.'}</p>
      <label class="campo"><span>Rúbrica</span><select name="rubrica">${S.rubricas.map(r => `<option value="${r.clave}" ${sugerida === r.clave ? 'selected' : ''}>${esc(r.nombre)}</option>`).join('')}</select></label>
      <span class="meta">Evaluadores</span><div class="chips" style="margin:.5rem 0 1rem">${evaluadores.map(v => `<button type="button" class="chip" data-ev="${esc(v.id)}" aria-pressed="false" ${ya.includes(v.id) ? 'disabled title="Ya asignado"' : ''}>${esc(v.nombre)}${ya.includes(v.id) ? ' ✓' : ''}</button>`).join('')}</div>
      <div class="botones"><button class="btn btn--rojo" data-si>Asignar</button><button class="btn" data-cerrar>Cancelar</button></div>`, (m, cerrar) => {
      $$('[data-ev]', m).forEach(b => b.onclick = () => { const v = b.dataset.ev; sel.has(v) ? sel.delete(v) : sel.add(v); b.setAttribute('aria-pressed', sel.has(v)); });
      $('[data-si]', m).onclick = async () => { if (!sel.size) return toast('Elige al menos un evaluador.', true); try { await api('asignar', { envioId: e.id, evaluadores: [...sel], rubrica: $('[name=rubrica]', m).value }); cerrar(); toast('Asignado. Les llegó el aviso.'); ruta(); } catch (er) { toast(er.message, true); } };
    });
  }
  async function modalEscribir(e, tipo) {
    const { plantillas } = await api('plantillas');
    const entregadas = e.asignaciones.filter(a => a.estado === 'entregada' && +a.ronda === +e.ronda);
    const inicial = plantillas.find(p => p.clave === (tipo || 'libre')) || plantillas[0];
    modal(`<h2>Escribir al autor</h2><p class="suave">Se envía desde el Gmail de la gaceta${e.hiloId ? ', en el mismo hilo del envío' : ''} a <b>${esc(e.autorCorreo || 'sin correo')}</b>.</p>
      <label class="campo"><span>Mensaje listo</span><select name="plantilla">${plantillas.map(p => `<option value="${esc(p.clave)}" ${p.clave === inicial.clave ? 'selected' : ''}>${esc(p.nombre)}</option>`).join('')}</select></label>
      <label class="campo"><span>Asunto</span><input type="text" name="asunto"></label>
      <label class="campo"><span>Mensaje</span><textarea name="cuerpo" style="min-height:14rem"></textarea></label>
      ${e.dirCorr && e.dirCorr.archivoId ? `<label class="chequeo"><input type="checkbox" name="adj" value="directora|" ${tipo === 'directora' || (tipo === 'aceptado' && e.dirCorr.aplicadas) ? 'checked' : ''}> Correcciones de la Directora (Word con comentarios y control de cambios)</label>` : ''}
      ${entregadas.some(a => (a.anotaciones || []).length || a.correccionId || a.subidaCorreccionId) ? `<span class="meta">Adjuntar correcciones (con comentarios y control de cambios, sin nombres de evaluadores)</span>${entregadas.map((a, i) => `<label class="chequeo"><input type="checkbox" name="adj" value="correccion|${esc(a.id)}" ${['ajustes', 'no'].includes(tipo) ? 'checked' : ''}> Correcciones del lector/a ${i + 1}</label>`).join('')}` : ''}
      <p class="suave" style="font-size:.85rem;margin-top:.6rem">🔒 Las rúbricas son internas del equipo: no se envían a la autora.</p>
      <div class="botones" style="margin-top:1rem"><button class="btn btn--rojo" data-si>Enviar</button><button class="btn" data-cerrar>Cancelar</button></div>`, (m, cerrar) => {
      const obs = observacionesPara(e);
      const poner = () => { const p = plantillas.find(x => x.clave === $('[name=plantilla]', m).value); $('[name=asunto]', m).value = rellenar(p.asunto, e, obs); $('[name=cuerpo]', m).value = rellenar(conObservaciones(p.cuerpo, p.clave, obs), e, obs);
        $$('[name=adj]', m).forEach(x => { if (x.value.startsWith('correccion|')) x.checked = ['ajustes', 'no'].includes(p.clave); }); };
      poner(); $('[name=plantilla]', m).onchange = poner;
      $('[data-si]', m).onclick = async () => {
        const b = $('[data-si]', m); b.disabled = true;
        const adjuntos = $$('[name=adj]:checked', m).map(x => { const [t, id] = x.value.split('|'); return { tipo: t, asignacionId: id }; });
        try {
          const r = await api('responderAutor', { envioId: e.id, tipo: $('[name=plantilla]', m).value, asunto: $('[name=asunto]', m).value, cuerpo: $('[name=cuerpo]', m).value, adjuntos });
          cerrar(); toast(r.demo ? 'Mensaje registrado (demostración: no se envió de verdad)' : 'Mensaje enviado al autor'); ruta();
        } catch (er) { toast(er.message, true); b.disabled = false; }
      };
    });
  }
  function modalVersion(e) {
    modal(`<h2>Subir versión corregida</h2><p class="suave">Si el autor respondió por correo, también puedes tomar el archivo desde el hilo en <a href="#hilo/${encodeURIComponent(e.hiloId || '')}">Correo</a>.</p><label class="campo"><span>Archivo de Word</span><input type="file" name="f" accept=".docx,.doc,.pdf,.odt"></label><div class="botones"><button class="btn btn--rojo" data-si>Subir</button><button class="btn" data-cerrar>Cancelar</button></div>`, (m, cerrar) => {
      $('[data-si]', m).onclick = async () => { const f = $('[name=f]', m).files[0]; if (!f) return toast('Elige el archivo.', true); try { const d = await leerArchivoLocal(f); await api('subirVersion', { envioId: e.id, ...d }); cerrar(); toast('Versión nueva guardada'); ruta(); } catch (er) { toast(er.message, true); } };
    });
  }
  function modalBorrarEnvio(e) {
    modal(`<h2>Borrar ${esc(e.id)}</h2><p class="suave">Se borra <b>todo</b> lo de «${esc(e.titulo)}»: evaluaciones, correcciones, historial y avisos. Los archivos van a la papelera de Drive y el correo vuelve a quedar sin registrar. <b>No se puede deshacer.</b></p><p class="suave">Úsalo para pruebas o envíos registrados por error. Si solo quieres archivarlo, mejor usa «Cerrar envío».</p>
      <label class="campo"><span>Para confirmar, escribe el código: <b>${esc(e.id)}</b></span><input type="text" name="conf" autocomplete="off"></label>
      <div class="botones"><button class="btn" data-si style="background:var(--rojo-osc);border-color:var(--rojo-osc);color:#fff">Borrar para siempre</button><button class="btn" data-cerrar>Cancelar</button></div>`, (m, cerrar) => {
      $('[data-si]', m).onclick = async () => {
        const c = $('[name=conf]', m).value.trim().toUpperCase(); if (c !== e.id) return toast('Escribe el código exacto: ' + e.id, true);
        try { await api('borrarEnvio', { id: e.id, confirmar: c }); cerrar(); toast(`${e.id} borrado`); location.hash = tiene('auxiliar') ? '#envios' : '#inicio'; } catch (er) { toast(er.message, true); }
      };
    });
  }
  function modalReclasificar(e) {
    const ops = [['aceptada', 'Aceptado', 'Pasa a la Directora para su revisión.'], ['ajustes', 'Aceptado con ajustes', 'Vuelve a ti para avisarle a la autora con las correcciones (también le llega a la Directora).'], ['no', 'No aceptado', 'Vuelve a ti para avisarle a la autora. No pasa por la Directora.']];
    modal(`<h2>Cambiar resultado</h2><p class="suave">Úsalo solo si «${esc(e.titulo)}» quedó mal clasificado. El texto pasa al grupo nuevo, sigue el camino que le corresponde y su carpeta se mueve en Drive. Queda registrado en el historial.</p>
      ${ops.map(([k, t, d]) => `<label class="chequeo" style="align-items:flex-start"><input type="radio" name="res" value="${k}" ${e.decision === k ? 'disabled' : ''}> <span><b>${t}</b>${e.decision === k ? ' (actual)' : ''}<br><span class="suave" style="font-size:.85rem">${d}</span></span></label>`).join('')}
      <label class="campo" style="margin-top:.8rem"><span>Motivo del cambio</span><textarea name="nota" placeholder="Ej.: la Directora pidió reconsiderarlo tras leer las rúbricas."></textarea></label>
      <div class="botones"><button class="btn btn--rojo" data-si>Cambiar</button><button class="btn" data-cerrar>Cancelar</button></div>`, (m, cerrar) => {
      $('[data-si]', m).onclick = async () => {
        const r = $('[name=res]:checked', m); if (!r) return toast('Elige el nuevo resultado.', true);
        try { await api('reclasificar', { envioId: e.id, decision: r.value, nota: $('[name=nota]', m).value }); cerrar(); toast('Resultado cambiado'); ruta(); } catch (er) { toast(er.message, true); }
      };
    });
  }
  function modalEditar(e) {
    modal(`<h2>Editar datos</h2>
      <label class="campo"><span>Título</span><input type="text" name="titulo" value="${esc(e.titulo)}"></label>
      <label class="campo"><span>Género</span><input type="text" name="genero" value="${esc(e.genero)}" list="generos2"><datalist id="generos2">${GENEROS.map(g => `<option>${g}</option>`).join('')}</datalist></label>
      <label class="campo"><span>Rúbrica</span><select name="rubrica"><option value="">Sin elegir</option>${S.rubricas.map(r => `<option value="${r.clave}" ${e.rubrica === r.clave ? 'selected' : ''}>${esc(r.nombre)}</option>`).join('')}</select></label>
      <label class="campo"><span>Autor/a</span><input type="text" name="autorNombre" value="${esc(e.autorNombre)}"></label>
      <label class="campo"><span>Correo del autor</span><input type="email" name="autorCorreo" value="${esc(e.autorCorreo)}"></label>
      <div class="botones"><button class="btn btn--rojo" data-si>Guardar</button><button class="btn" data-cerrar>Cancelar</button><span style="margin-left:auto"></span>${e.estado !== 'cerrado' ? '<button class="btn btn--chico" data-cerrar-envio>Cerrar envío</button>' : ''}<button class="btn btn--chico" data-borrar-envio style="color:var(--rojo);border-color:var(--rojo)">Borrar envío</button></div>`, (m, cerrar) => {
      $('[data-si]', m).onclick = async () => { const d = {}; ['titulo', 'genero', 'rubrica', 'autorNombre', 'autorCorreo'].forEach(k => d[k] = $(`[name=${k}]`, m).value); try { await api('editarEnvio', { id: e.id, ...d }); cerrar(); toast('Guardado'); ruta(); } catch (er) { toast(er.message, true); } };
      const be = $('[data-borrar-envio]', m); if (be) be.onclick = () => { cerrar(); modalBorrarEnvio(e); };
      const ce = $('[data-cerrar-envio]', m); if (ce) ce.onclick = async () => { cerrar(); if (await confirmar('¿Cerrar este envío?', 'Queda archivado y sale de las listas de trabajo.')) { await api('moverEnvio', { envioId: e.id, a: 'cerrar' }); ruta(); } };
    });
  }

  async function textoVersion(e, n) {
    const r = await api('archivo', { envioId: e.id, cual: 'version', n });
    if (!/\.docx$/i.test(r.nombre)) throw new Error('Esa versión no es un Word (.docx), así que no se puede corregir en la página.');
    return Docx.leerParrafos(Docx.deBase64(r.b64));
  }
  async function verCorrDirectora(e) {
    try {
      const ps = await textoVersion(e, e.dirCorr.version), ans = e.dirCorr.anotaciones;
      modal(`<h2>Correcciones de la Directora</h2><p class="suave">Versión ${e.dirCorr.version} · ${ans.length} corrección(es). Toca una marca para ver su nota.</p><div class="hoja" style="max-height:60vh;overflow:auto">${pintarHoja(ps, ans)}</div><div style="margin-top:1rem">${listaNotas(ps, ans, false)}</div><div class="botones" style="margin-top:1rem">${e.dirCorr.archivoId ? '<button class="btn" data-bajar>Descargar en Word</button>' : ''}<button class="btn" data-cerrar>Cerrar</button></div>`, m => {
        $('.modal', $('#modal')).style.width = 'min(900px, 100%)';
        const b = $('[data-bajar]', m); if (b) b.onclick = () => descargarArchivo({ envioId: e.id, cual: 'directora' });
        conectarMarcas(m, ans);
      });
    } catch (er) { toast(er.message, true); }
  }
  function modalPlataforma(e) {
    const v = (e.versiones || []).slice(-1)[0] || {};
    const falta = !e.fichaId || !e.licenciaId;
    const ok = x => x ? '<b style="color:var(--buchon)">✓</b>' : '<b style="color:var(--rojo)">✗ falta</b>';
    modal(`<h2>Enviar a plataforma</h2><p class="suave">Sebastian recibe el aviso y ve en su perfil todo lo necesario para publicar.</p>
      <div class="tarjeta"><div class="archivo"><span>${ok(1)} Texto final · versión ${v.n || 1} · ${esc(v.nombre || '')}</span></div><div class="archivo"><span>${ok(e.fichaId)} Ficha de autor</span></div><div class="archivo"><span>${ok(e.licenciaId)} Licencia de uso</span></div></div>
      ${falta ? '<p class="suave">Si el autor la mandó después, súbela desde el hilo del correo o pídesela antes de enviar.</p><label class="chequeo"><input type="checkbox" name="igual"> Enviar de todas formas</label>' : ''}
      <label class="campo"><span>Nota para plataforma (opcional)</span><textarea name="nota" placeholder="Ej.: firma como «M. Ejemplo»; publicar con la ilustración que mandó."></textarea></label>
      <div class="botones"><button class="btn btn--verde" data-si>Enviar a plataforma</button><button class="btn" data-cerrar>Cancelar</button></div>`, (m, cerrar) => {
      $('[data-si]', m).onclick = async () => {
        const igual = $('[name=igual]', m); if (falta && !igual.checked) return toast('Falta la ficha o la licencia. Marca «Enviar de todas formas» si quieres seguir.', true);
        try { await api('moverEnvio', { envioId: e.id, a: 'plataforma', nota: $('[name=nota]', m).value, sinFichaOLicencia: falta }); cerrar(); toast('Enviado a plataforma'); ruta(); } catch (er) { toast(er.message, true); }
      };
    });
  }
  // Aplica las correcciones al texto (como «Aceptar todos los cambios» en Word). Los comentarios no cambian el texto.
  function aplicarCorrecciones(parrafos, anotaciones) {
    return parrafos.map((par, pi) => {
      let t = par.texto;
      anotaciones.filter(a => a.p === pi && a.tipo !== 'comentario').sort((x, y) => y.ini - x.ini || (x.tipo === 'insertar' ? 1 : -1))
        .forEach(a => {
          if (a.tipo === 'borrar') t = t.slice(0, a.ini) + t.slice(a.fin);
          else if (a.tipo === 'reemplazar') t = t.slice(0, a.ini) + (a.nuevo || '') + t.slice(a.fin);
          else if (a.tipo === 'insertar' && a.nuevo) t = t.slice(0, a.ini) + a.nuevo + t.slice(a.ini);
        });
      return t.replace(/ {2,}/g, ' ');
    }).filter((t, i) => !(t.trim() === '' && parrafos[i].texto.trim() !== '')); // un verso tachado completo no deja línea vacía
  }
  // ---------- La Directora corrige el texto ----------
  RUTAS.corregir = {
    render: async id => {
      const { envio: e } = await api('envio', { id });
      if (!tiene('directora') || e.estado !== 'revision_directora') throw new Error('Este texto no está en revisión de la Directora.');
      const v = e.versiones.slice(-1)[0];
      const bor = e.dirCorr && e.dirCorr.borrador && +e.dirCorr.borrador.version === +v.n ? e.dirCorr.borrador : null;
      E = { env: e, version: v.n, anotaciones: bor ? bor.anotaciones.slice() : [], nota: bor ? bor.nota : '', parrafos: null, textoError: null, entregada: false };
      try { E.parrafos = await textoVersion(e, v.n); } catch (er) { E.textoError = er.message; }
      E.repintar = pintarCorrDirectora;
      E.guardar = () => api('borradorDirectora', { envioId: e.id, anotaciones: E.anotaciones, nota: E.nota }).then(() => { const g = $('[data-guardado]'); if (g) g.textContent = 'Borrador guardado ✓'; });
      return `<div class="migas"><a href="#envio/${esc(e.id)}">${esc(e.id)}</a> › Correcciones</div>
        <div class="cabecera"><div><p class="meta">${esc(e.id)} · versión ${v.n} · Directora general</p><h1>Corregir «${esc(e.titulo)}»</h1></div></div>
        <div id="zona-corr"></div>`;
    },
    conectar: () => pintarCorrDirectora(),
  };
  function pintarCorrDirectora() {
    const z = $('#zona-corr'); if (!z) return;
    z.innerHTML = `<div class="ayuda"><b>Cómo funciona:</b> marca tus correcciones igual que en Word: selecciona una palabra o frase y elige <b>Comentar</b>, <b>Tachar</b> o <b>Cambiar por…</b>, o toca un punto para <b>Insertar</b>. Al terminar, elige: <b>Enviar a la Auxiliar</b> (la autora revisa tus correcciones y responde) o <b>Corregir y aprobar para publicar</b> (tus correcciones se aplican y la obra sigue a plataforma con el texto final corregido).</div>
      ${E.textoError ? `<div class="vacio">${esc(E.textoError)}</div>` : `<div class="corrector"><div class="hoja" id="hoja">${pintarHoja(E.parrafos, E.anotaciones)}</div><aside class="margen" id="margen">${listaNotas(E.parrafos, E.anotaciones, true)}</aside></div>`}
      <label class="campo" style="margin-top:1rem"><span>Nota general para la autora (opcional)</span><textarea data-nota-dir>${esc(E.nota)}</textarea></label>
      <div class="resultado"><span class="meta" data-guardado>${E.anotaciones.length} corrección(es)</span><span style="flex:1"></span>${E.parrafos && E.anotaciones.length ? '<button class="btn" data-previa>Ver en Word</button>' : ''}<button class="btn" data-enviar-dir>Enviar a la Auxiliar (la autora revisa)</button><button class="btn btn--verde" data-corr-aprobar>Corregir y aprobar para publicar</button></div>`;
    conectarHoja();
    const t = $('[data-nota-dir]'); t.oninput = () => { E.nota = t.value; guardarBorradorPronto(); };
    const pv = $('[data-previa]'); if (pv) pv.onclick = async () => descargar(`${E.env.id} · correcciones de la Dirección.docx`, await Docx.crearCorregido({ parrafos: E.parrafos, anotaciones: E.anotaciones, autor: 'Dirección editorial · Hicimos lo que quisimos', iniciales: 'DE', titulo: `${E.env.id} · ${E.env.titulo}` }));
    $('[data-corr-aprobar]').onclick = async () => {
      if (!E.anotaciones.length) return toast('Haz al menos una corrección. Si no hay nada que corregir, usa «Aprobar para publicar».', true);
      if (!E.parrafos) return toast('Este texto no se puede corregir en la página.', true);
      const final = aplicarCorrecciones(E.parrafos, E.anotaciones);
      if (!await confirmar('¿Corregir y aprobar para publicar?', `Tus ${E.anotaciones.length} corrección(es) se aplican al texto y queda un <b>texto final ya corregido</b>. La Auxiliar le avisa a la autora y lo envía a plataforma; Sebastian recibe el texto final y el Word con tus correcciones marcadas.<br><br><span class="suave">Los comentarios no cambian el texto: quedan solo en el Word de correcciones.</span>`, 'Corregir y aprobar')) return;
      try {
        const d = { envioId: E.env.id, decision: 'corregirAprobar', anotaciones: E.anotaciones, nota: E.nota };
        d.correccionDocx = Docx.aBase64(await Docx.crearCorregido({ parrafos: E.parrafos, anotaciones: E.anotaciones, autor: 'Dirección editorial · Hicimos lo que quisimos', iniciales: 'DE', titulo: `${E.env.id} · ${E.env.titulo}` }));
        d.finalDocx = Docx.aBase64(await Docx.desdeTexto(final.join('\n')));
        await api('decidir', d); clearTimeout(tBorrador); E = null; toast('Aprobada con tus correcciones ✓'); location.hash = '#envio/' + d.envioId;
      } catch (er) { toast(er.message, true); }
    };
    $('[data-enviar-dir]').onclick = async () => {
      if (!E.anotaciones.length && !E.nota.trim()) return toast('Haz al menos una corrección o escribe una nota.', true);
      if (!await confirmar('¿Enviar las correcciones?', `${E.anotaciones.length} corrección(es). La Auxiliar recibe el aviso y se las manda a la autora; cuando ella responda con la versión final, el texto vuelve a ti.`, 'Enviar')) return;
      try {
        const d = { envioId: E.env.id, decision: 'corregir', anotaciones: E.anotaciones, nota: E.nota };
        if (E.anotaciones.length && E.parrafos) d.correccionDocx = Docx.aBase64(await Docx.crearCorregido({ parrafos: E.parrafos, anotaciones: E.anotaciones, autor: 'Dirección editorial · Hicimos lo que quisimos', iniciales: 'DE', titulo: `${E.env.id} · ${E.env.titulo}` }));
        await api('decidir', d); clearTimeout(tBorrador); E = null; toast('Correcciones enviadas a la Auxiliar'); location.hash = '#envio/' + d.envioId;
      } catch (er) { toast(er.message, true); }
    };
  }

  // ---------- Plantillas de mensajes ----------
  RUTAS.plantillas = {
    render: async () => {
      const { plantillas } = await api('plantillas'); S._plant = plantillas;
      return `<div class="cabecera"><div><p class="meta">Auxiliar</p><h1>Mensajes listos</h1><p class="suave">Los textos que se le envían a los autores. Puedes usar {titulo}, {autor}, {codigo} y {observaciones} (los comentarios de los evaluadores, sin nombres; si no lo pones en «No aceptado» o «Aceptado con ajustes», se agregan solos antes de la despedida): se reemplazan solos.</p></div></div>
        ${plantillas.map(p => `<form class="tarjeta" data-plant="${esc(p.clave)}"><h3>${esc(p.nombre)}</h3><label class="campo"><span>Asunto</span><input type="text" name="asunto" value="${esc(p.asunto)}"></label><label class="campo"><span>Mensaje</span><textarea name="cuerpo" style="min-height:12rem">${esc(p.cuerpo)}</textarea></label><button class="btn btn--chico btn--rojo">Guardar</button></form>`).join('')}`;
    },
    conectar: () => $$('[data-plant]').forEach(f => f.onsubmit = async e => { e.preventDefault(); try { await api('guardarPlantilla', { clave: f.dataset.plant, nombre: S._plant.find(p => p.clave === f.dataset.plant).nombre, asunto: f.asunto.value, cuerpo: f.cuerpo.value }); toast('Mensaje guardado'); } catch (er) { toast(er.message, true); } }),
  };

  // =========================================================
  //  EVALUADORES
  // =========================================================
  RUTAS.evaluaciones = {
    render: async () => {
      await cargarRubricas();
      const { evaluaciones } = await api('misEvaluaciones');
      const p = evaluaciones.filter(a => a.estado !== 'entregada'), h = evaluaciones.filter(a => a.estado === 'entregada');
      const f = a => `<li><a class="fila ${a.estado === 'pendiente' ? 'nuevo' : ''}" href="#evaluacion/${esc(a.id)}"><div class="fila__cuerpo"><div class="meta">${esc(a.envioId)} · ${esc(a.genero || '')} · ${a.estado === 'entregada' ? 'entregada ' + fecha(a.entregada) : 'asignada ' + hace(a.asignada)}</div><div class="fila__titulo">${esc(a.titulo)}</div><div class="fila__sub">Rúbrica: ${esc((rubrica(a.rubrica) || {}).nombre || a.rubrica)}</div></div>${a.estado === 'entregada' ? dictChip(a.dictamen) : `<span class="estado ${a.estado === 'borrador' ? 'e-en_evaluacion' : 'e-por_asignar'}">${a.estado === 'borrador' ? 'En curso' : 'Nueva'}</span>`}</a></li>`;
      return `<div class="cabecera"><div><p class="meta">Evaluación doble ciego</p><h1>Mis evaluaciones</h1></div></div>
        <h2>Pendientes <span class="meta">${p.length}</span></h2>${p.length ? `<ul class="lista">${p.map(f).join('')}</ul>` : '<div class="vacio"><span class="mano">¡Al día!</span>Cuando te asignen una obra te llega un aviso aquí y a tu correo.</div>'}
        ${h.length ? `<h2 style="margin-top:2rem">Entregadas</h2><ul class="lista">${h.map(f).join('')}</ul>` : ''}`;
    },
  };

  let E = null; // estado de la evaluación abierta
  RUTAS.evaluacion = {
    render: async id => {
      await cargarRubricas();
      const r = await api('evaluacion', { id });
      const a = r.asignacion; const rub = rubrica(a.rubrica);
      E = { id, a, env: r.envio, rub, evaluador: r.evaluador, niveles: (a.niveles || []).slice(), anotaciones: (a.anotaciones || []).slice(), observaciones: a.observaciones || '', enlace: a.enlace || '', modo: a.modo || 'pagina', parrafos: null, textoError: null, subidas: {}, entregada: a.estado === 'entregada', tab: 'texto' };
      try { const t = await api('archivo', { asignacionId: id, cual: 'texto' }); E.textoNombre = t.nombre; E.textoB64 = t.b64; if (/\.docx$/i.test(t.nombre)) E.parrafos = await Docx.leerParrafos(Docx.deBase64(t.b64)); else E.textoError = 'Este texto no es un Word (.docx), así que no se puede corregir en la página. Descárgalo y usa la opción «Con mis archivos de Word».'; }
      catch (er) { E.textoError = er.message; }
      return `<div class="migas"><a href="#evaluaciones">Mis evaluaciones</a> › ${esc(r.envio.id)}</div>
        <div class="cabecera"><div><p class="meta">${esc(r.envio.id)} · ${esc(r.envio.genero || '')} · rúbrica de ${esc(rub ? rub.nombre : a.rubrica)}</p><h1>${esc(r.envio.titulo)}</h1>${E.entregada ? `${dictChip(a.dictamen)} <span class="meta">Entregada ${fecha(a.entregada)} · puntaje ${punt(a.puntaje)}</span>` : ''}</div></div>
        ${E.entregada ? '' : `<div class="modo" role="group" aria-label="Cómo quieres evaluar"><button data-modo="pagina" aria-pressed="${E.modo === 'pagina'}">Evaluar en la página</button><button data-modo="word" aria-pressed="${E.modo === 'word'}">Con mis archivos de Word</button></div>`}
        <div id="zona-eval"></div>`;
    },
    conectar: () => { $$('[data-modo]').forEach(b => b.onclick = () => { E.modo = b.dataset.modo; $$('[data-modo]').forEach(x => x.setAttribute('aria-pressed', x === b)); guardarBorradorPronto(); pintarEvaluacion(); }); pintarEvaluacion(); },
  };

  function pintarEvaluacion() {
    const z = $('#zona-eval'); if (!z) return;
    const ro = E.entregada;
    const n = E.rub ? E.rub.criterios.length : 0;
    const { puntaje, dictamen } = calcular(E.niveles, n);
    const firma = ro ? `<div class="tarjeta" style="border-color:${E.a.firmadaId ? 'var(--borde)' : 'var(--canaguate)'}"><h3>Rúbrica firmada</h3>${E.a.firmadaId ? `<p>✓ ${E.a.firmadaId === E.a.rubricaId ? 'Tu rúbrica quedó firmada automáticamente con tu firma guardada' : 'Ya subiste tu rúbrica firmada'}${E.a.firmadaFecha ? ` (${fecha(E.a.firmadaFecha)})` : ''}.</p><div class="botones"><button class="btn btn--chico" data-ver-firmada>Ver mi rúbrica firmada</button><button class="btn btn--chico" data-mi-firmada>Reemplazarla</button></div>` : `<p class="suave">Para que la evaluación quede en regla, sube tu rúbrica con tu firma:</p><ol style="margin:.4rem 0 1rem 1.1rem;padding:0"><li>${E.niveles.filter(Boolean).length ? 'Descarga <b>Mi rúbrica llena (Word)</b>' : 'Usa la rúbrica que llenaste en Word'}.</li><li>Fírmala: imprímela y fírmala a mano, o pon tu firma digital.</li><li>Súbela aquí como PDF, foto o Word.</li></ol><div class="botones">${E.niveles.filter(Boolean).length ? '<button class="btn btn--chico" data-mi-rub>1 · Mi rúbrica llena (Word)</button>' : ''}<button class="btn btn--rojo btn--chico" data-mi-firmada>3 · Subir rúbrica firmada</button></div><p class="suave" style="font-size:.85rem;margin-top:.8rem">Para no tener que hacer esto cada vez, pídele a Sebastian o a Luisa que carguen tu firma en «Perfiles».</p>`}</div>` : '';
    const descargas = firma + `<div class="tarjeta"><h3>Archivos</h3><div class="botones"><button class="btn btn--chico" data-bajar-texto>Descargar el texto (Word original)</button><button class="btn btn--chico" data-bajar-rub>Descargar rúbrica en blanco (Word)</button>${ro && E.niveles.filter(Boolean).length ? '<button class="btn btn--chico" data-mi-rub>Mi rúbrica llena (Word)</button>' : ''}${ro && E.anotaciones.length ? '<button class="btn btn--chico" data-mi-corr>Mi texto corregido (Word)</button>' : ''}</div></div>`;
    if (E.modo === 'word' && !ro) {
      z.innerHTML = `<div class="ayuda">Descarga el texto y la rúbrica, trabájalos en Word en tu computador y súbelos aquí. Después marca tu puntaje y dictamen para que el sistema sepa a dónde enviar la obra.</div>${descargas}
        <div class="tarjeta"><h3>Subir mis archivos</h3>
          <label class="campo"><span>Rúbrica llena (Word o PDF)</span><input type="file" data-subir="rubrica" accept=".docx,.doc,.pdf"></label>${E.subidas.rubrica ? `<p class="suave">✓ ${esc(E.subidas.rubrica.nombre)}</p>` : ''}
          <label class="campo"><span>Texto corregido con comentarios y control de cambios (Word, opcional)</span><input type="file" data-subir="correccion" accept=".docx,.doc"></label>${E.subidas.correccion ? `<p class="suave">✓ ${esc(E.subidas.correccion.nombre)}</p>` : ''}
          <label class="campo"><span>Puntaje final (de 1 a 5)</span><input type="number" step="0.1" min="1" max="5" data-pw value="${E.puntajeWord ?? ''}"></label>
          <p class="suave">Dictamen según el puntaje: <b data-dw>${E.puntajeWord ? DICT[E.puntajeWord >= 4.5 ? 'aceptada' : E.puntajeWord >= 3.5 ? 'ajustes' : 'no'] : '—'}</b></p>
          <label class="campo"><span>Observaciones para quien escribe (opcional)</span><textarea data-obs>${esc(E.observaciones)}</textarea></label>
        </div><div class="resultado"><span class="meta">Listo para entregar</span><span style="flex:1"></span><button class="btn btn--rojo" data-entregar>Entregar evaluación</button></div>`;
    } else {
      z.innerHTML = `<div class="modo" role="tablist"><button data-tab="texto" aria-pressed="${E.tab === 'texto'}">1 · Texto y correcciones${E.anotaciones.length ? ` (${E.anotaciones.length})` : ''}</button><button data-tab="rubrica" aria-pressed="${E.tab === 'rubrica'}">2 · Rúbrica${E.niveles.filter(Boolean).length ? ` (${E.niveles.filter(Boolean).length}/${n})` : ''}</button></div>
        <div data-panel="texto" ${E.tab !== 'texto' ? 'hidden' : ''}>
          ${ro ? '' : '<div class="ayuda"><b>Cómo corregir:</b> selecciona una palabra o frase del texto y elige <b>Comentar</b>, <b>Tachar</b> o <b>Cambiar por…</b>. Para agregar algo, toca un punto del texto sin seleccionar y elige <b>Insertar aquí</b>. Todo se guarda solo y sale en Word como comentarios y control de cambios.</div>'}
          ${E.textoError ? `<div class="vacio">${esc(E.textoError)}</div>` : `<div class="corrector"><div class="hoja" id="hoja">${pintarHoja(E.parrafos, E.anotaciones)}</div><aside class="margen" id="margen">${listaNotas(E.parrafos, E.anotaciones, !ro)}</aside></div>`}
          ${ro ? '' : '<div class="botones" style="margin-top:1rem;justify-content:flex-end"><button class="btn btn--claro" data-ir-rub>Siguiente: rúbrica →</button></div>'}
        </div>
        <div data-panel="rubrica" ${E.tab !== 'rubrica' ? 'hidden' : ''}>
          ${E.rub && E.rub.nota ? `<p class="suave">${esc(E.rub.nota)}</p>` : ''}
          ${E.rub && E.rub.conEnlace ? `<label class="campo"><span>Enlace o archivo de la obra</span><input type="text" data-enlace value="${esc(E.enlace)}" ${ro ? 'disabled' : ''}></label>` : ''}
          ${E.rub ? E.rub.criterios.map((c, k) => `<div class="criterio"><h3><span>${esc(c.nombre)}</span>${E.niveles[k] ? `<span class="estado" style="border-color:var(--acento);color:#fff">${E.niveles[k]}</span>` : ''}</h3><div class="niveles">${c.niveles.map((d, j) => `<button type="button" class="nivel" data-crit="${k}" data-nivel="${j + 1}" aria-pressed="${E.niveles[k] === j + 1}" ${ro ? 'disabled' : ''}><b>${esc(E.rub.niveles[j])}</b>${esc(d)}</button>`).join('')}</div></div>`).join('') : '<div class="vacio">No se encontró la rúbrica.</div>'}
          <label class="campo"><span>Observaciones y ajustes sugeridos a quien escribe</span><textarea data-obs ${ro ? 'disabled' : ''}>${esc(E.observaciones)}</textarea></label>
          ${ro ? '' : (S.usuario.tieneFirma ? '<p class="firma si" style="margin:.4rem 0">✍ Tu firma guardada se pondrá sola en la rúbrica al entregar.</p>' : '<p class="firma no" style="margin:.4rem 0">✍ No tienes firma guardada: después de entregar tendrás que subir la rúbrica firmada (o pídele a Sebastian o a Luisa que carguen tu firma).</p>')}
          <div class="resultado"><div><span class="meta">Puntaje</span><br><b>${punt(puntaje)}</b></div><div><span class="meta">Dictamen</span><br>${dictamen ? dictChip(dictamen) : '<span class="suave">Califica todos los criterios</span>'}</div><span style="flex:1"></span>${ro ? '' : `<span class="meta" data-guardado></span><button class="btn btn--rojo" data-entregar ${puntaje == null ? 'disabled' : ''}>Entregar evaluación</button>`}</div>
        </div>${descargas}`;
    }
    conectarEvaluacion();
  }
  let tBorrador;
  function guardarBorradorPronto() {
    if (E.entregada) return; clearTimeout(tBorrador);
    tBorrador = setTimeout(async () => { if (E.guardar) { try { await E.guardar(); } catch (e) { } return; } try { await api('guardarBorrador', { id: E.id, niveles: E.niveles, anotaciones: E.anotaciones, observaciones: E.observaciones, modo: E.modo, enlace: E.enlace }); const g = $('[data-guardado]'); if (g) g.textContent = 'Guardado ✓'; } catch (e) { toast('No se pudo guardar el borrador: ' + e.message, true); } }, 900);
  }
  function conectarEvaluacion() {
    $$('[data-tab]').forEach(b => b.onclick = () => { E.tab = b.dataset.tab; quitarHerramientas(); pintarEvaluacion(); });
    const ir = $('[data-ir-rub]'); if (ir) ir.onclick = () => { E.tab = 'rubrica'; pintarEvaluacion(); window.scrollTo(0, 0); };
    $$('[data-crit]').forEach(b => b.onclick = () => { E.niveles[+b.dataset.crit] = +b.dataset.nivel; guardarBorradorPronto(); const y = window.scrollY; pintarEvaluacion(); window.scrollTo(0, y); });
    $$('[data-obs]').forEach(t => t.oninput = () => { E.observaciones = t.value; guardarBorradorPronto(); });
    const en = $('[data-enlace]'); if (en) en.oninput = () => { E.enlace = en.value; guardarBorradorPronto(); };
    const pw = $('[data-pw]'); if (pw) pw.oninput = () => { E.puntajeWord = pw.value ? +pw.value : null; $('[data-dw]').textContent = E.puntajeWord ? DICT[E.puntajeWord >= 4.5 ? 'aceptada' : E.puntajeWord >= 3.5 ? 'ajustes' : 'no'] : '—'; };
    $$('[data-subir]').forEach(i => i.onchange = async () => { const f = i.files[0]; if (!f) return; if (f.size > 15e6) return toast('El archivo pesa demasiado (máximo 15 MB).', true); E.subidas[i.dataset.subir] = await leerArchivoLocal(f); pintarEvaluacion(); });
    const bt = $('[data-bajar-texto]'); if (bt) bt.onclick = () => { if (E.textoB64) descargar(`${E.env.id} · texto.docx`.replace('.docx', (E.textoNombre.match(/\.[a-z0-9]+$/i) || ['.docx'])[0]), Docx.deBase64(E.textoB64)); };
    const br = $('[data-bajar-rub]'); if (br) br.onclick = async () => descargar(`Rúbrica · ${E.rub.nombre} · ${E.env.id}.docx`, await Docx.llenarRubrica(await plantillaRubrica(E.a.rubrica), { titulo: E.env.titulo, codigo: E.env.id, fecha: new Date().toLocaleDateString('es-CO'), evaluador: E.evaluador }));
    const mf = $('[data-mi-firmada]'); if (mf) mf.onclick = () => modalFirmada(E.id, () => ruta());
    const vf = $('[data-ver-firmada]'); if (vf) vf.onclick = () => descargarArchivo({ asignacionId: E.id, cual: 'firmada' });
    $$('[data-mi-rub]').forEach(b => b.onclick = () => wordRubrica({ id: E.env.id, titulo: E.env.titulo }, { ...E.a, evaluador: E.evaluador }));
    const mc = $('[data-mi-corr]'); if (mc) mc.onclick = async () => descargar(`${E.env.id} · texto corregido.docx`, await Docx.crearCorregido({ parrafos: E.parrafos, anotaciones: E.anotaciones, autor: E.evaluador, iniciales: iniciales(E.evaluador), titulo: `${E.env.id} · ${E.env.titulo}` }));
    const ent = $('[data-entregar]'); if (ent) ent.onclick = entregar;
    conectarHoja();
  }
  const repintar = () => { const y = window.scrollY; (E.repintar || pintarEvaluacion)(); window.scrollTo(0, y); };
  function conectarHoja() {
    const hoja = $('#hoja');
    if (hoja) {
      conectarMarcas(document, E.anotaciones);
      if (!E.entregada) {
        hoja.addEventListener('mouseup', () => setTimeout(mostrarHerramientas, 10));
        hoja.addEventListener('touchend', () => setTimeout(mostrarHerramientas, 350));
        document.onselectionchange = () => { if (window.matchMedia('(max-width:760px)').matches) { clearTimeout(document._sel); document._sel = setTimeout(mostrarHerramientas, 450); } };
      }
      $$('[data-borrar-nota]').forEach(b => b.onclick = ev => { ev.stopPropagation(); E.anotaciones.splice(+b.dataset.borrarNota, 1); guardarBorradorPronto(); repintar(); });
      $$('[data-editar-nota]').forEach(b => b.onclick = ev => { ev.stopPropagation(); editarNota(+b.dataset.editarNota); });
    }
  }
  const iniciales = n => String(n || 'EV').split(/\s+/).map(x => x[0]).slice(0, 3).join('').toUpperCase();

  // ---------- Hoja con marcas (corrector) ----------
  function pintarHoja(ps, ans) {
    if (!ps) return '';
    const orden = ans.map((a, i) => ({ ...a, _i: i })).sort((x, y) => x.p - y.p || x.ini - y.ini);
    const num = new Map(orden.map((a, k) => [a._i, k + 1]));
    return ps.map((par, pi) => {
      const mias = orden.filter(a => a.p === pi);
      const cortes = [...new Set([0, par.texto.length, ...mias.flatMap(a => [a.ini, a.fin ?? a.ini])])].sort((a, b) => a - b);
      let h = '';
      const ins = pos => mias.filter(a => a.tipo === 'insertar' && a.ini === pos).map(a => `<ins class="an" data-a="${a._i}" data-pos="${pos}">${esc(a.nuevo)}</ins><sup class="an-n" data-a="${a._i}">${num.get(a._i)}</sup>`).join('');
      for (let k = 0; k < cortes.length - 1; k++) {
        const x = cortes[k], y = cortes[k + 1]; h += ins(x);
        const trozo = runsEntre(par, x, y);
        const a = mias.find(z => z.tipo !== 'insertar' && z.ini <= x && z.fin >= y);
        if (!a) { h += `<span data-o="${x}">${trozo}</span>`; continue; }
        if (a.tipo === 'comentario') h += `<mark class="an" data-a="${a._i}" data-o="${x}">${trozo}</mark>`;
        else h += `<del class="an" data-a="${a._i}" data-o="${x}">${trozo}</del>`;
        if (y === a.fin) { if (a.tipo === 'reemplazar') h += `<ins class="an" data-a="${a._i}" data-pos="${y}">${esc(a.nuevo)}</ins>`; h += `<sup class="an-n" data-a="${a._i}">${num.get(a._i)}</sup>`; }
      }
      h += ins(par.texto.length);
      return `<p data-p="${pi}">${h || '<span data-o="0"></span>'}</p>`;
    }).join('');
  }
  function runsEntre(par, x, y) {
    let pos = 0, out = '';
    for (const r of par.runs) { const a = pos, z = pos + r.t.length; pos = z; const i = Math.max(a, x), j = Math.min(z, y); if (i < j) { let t = esc(r.t.slice(i - a, j - a)); if (r.b) t = `<b>${t}</b>`; if (r.i) t = `<i>${t}</i>`; out += t; } }
    return out;
  }
  const TIPOS = { comentario: 'Comentario', borrar: 'Tachar', reemplazar: 'Cambio', insertar: 'Inserción' };
  function listaNotas(ps, ans, editable) {
    if (!ps) return '';
    const orden = ans.map((a, i) => ({ ...a, _i: i })).sort((x, y) => x.p - y.p || x.ini - y.ini);
    if (!orden.length) return `<div class="vacio" style="padding:1.2rem">${editable ? 'Tus comentarios y cambios aparecerán aquí, como en el margen de Word.' : 'Sin anotaciones.'}</div>`;
    return orden.map((a, k) => { const frag = ps[a.p] ? ps[a.p].texto.slice(a.ini, a.fin) : ''; return `<div class="nota n-${a.tipo}" data-a="${a._i}"><div class="meta">${k + 1} · ${TIPOS[a.tipo]}</div>${frag ? `<q>${esc(frag.length > 90 ? frag.slice(0, 90) + '…' : frag)}</q>` : ''}${a.nuevo ? `<div>→ <b>${esc(a.nuevo)}</b></div>` : ''}${a.comentario ? `<div>${esc(a.comentario)}</div>` : ''}${editable ? `<div class="acciones"><button class="btn btn--chico" data-editar-nota="${a._i}">Editar</button><button class="btn btn--chico" data-borrar-nota="${a._i}">Quitar</button></div>` : ''}</div>`; }).join('');
  }
  function conectarMarcas(raiz, ans) {
    const marcar = i => { $$('.activo', raiz).forEach(x => x.classList.remove('activo')); $$(`[data-a="${i}"]`, raiz).forEach(x => x.classList.add('activo')); const n = $(`.nota[data-a="${i}"]`, raiz), m = $(`.hoja [data-a="${i}"]`, raiz); return { n, m }; };
    $$('.hoja [data-a]', raiz).forEach(el => el.addEventListener('click', ev => { if (window.getSelection().toString()) return; const { n } = marcar(el.dataset.a); if (n) n.scrollIntoView({ block: 'nearest', behavior: 'smooth' }); }));
    $$('.nota[data-a]', raiz).forEach(el => el.addEventListener('click', () => { const { m } = marcar(el.dataset.a); if (m) m.scrollIntoView({ block: 'center', behavior: 'smooth' }); }));
  }
  function posicion(nodo, off) {
    let el = nodo.nodeType === 3 ? nodo.parentElement : nodo;
    const p = el.closest('p[data-p]'); if (!p) return null;
    if (nodo.nodeType !== 3) { // el cursor quedó entre elementos
      const hijo = nodo.childNodes[off] || null;
      if (hijo && hijo.nodeType === 1 && hijo.dataset && hijo.dataset.o != null) return { p: +p.dataset.p, o: +hijo.dataset.o };
      const txt = E.parrafos[+p.dataset.p].texto.length; return { p: +p.dataset.p, o: off === 0 ? 0 : txt };
    }
    const ins = el.closest('ins[data-pos]'); if (ins) return { p: +p.dataset.p, o: +ins.dataset.pos };
    if (el.closest('sup')) { const s = el.closest('sup'); const prev = s.previousElementSibling; return { p: +p.dataset.p, o: prev && prev.dataset.pos ? +prev.dataset.pos : (prev && prev.dataset.o != null ? +prev.dataset.o + prev.textContent.length : 0) }; }
    const seg = el.closest('[data-o]'); if (!seg) return null;
    // desplazamiento dentro del segmento (puede tener <b>/<i> adentro)
    let n = 0; const w = document.createTreeWalker(seg, NodeFilter.SHOW_TEXT);
    while (w.nextNode()) { if (w.currentNode === nodo) return { p: +p.dataset.p, o: +seg.dataset.o + n + off }; n += w.currentNode.length; }
    return { p: +p.dataset.p, o: +seg.dataset.o };
  }
  let selActual = null;
  function quitarHerramientas() { $$('.herramientas').forEach(x => x.remove()); }
  function mostrarHerramientas() {
    if (!E || E.entregada || !$('#hoja')) return;
    const sel = window.getSelection(); quitarHerramientas();
    if (!sel.rangeCount) return;
    const rg = sel.getRangeAt(0);
    if (!$('#hoja').contains(rg.commonAncestorContainer)) return;
    const a = posicion(rg.startContainer, rg.startOffset), b = posicion(rg.endContainer, rg.endOffset);
    if (!a || !b) return;
    if (a.p !== b.p) { toast('Selecciona dentro de un solo párrafo (o verso).', true); return; }
    const ini = Math.min(a.o, b.o), fin = Math.max(a.o, b.o);
    selActual = { p: a.p, ini, fin };
    const caja = rg.getBoundingClientRect();
    const t = document.createElement('div'); t.className = 'herramientas';
    t.innerHTML = fin > ini ? '<button data-h="comentario">💬 Comentar</button><button data-h="borrar">Tachar</button><button data-h="reemplazar">Cambiar por…</button><button data-h="x" aria-label="Cerrar">✕</button>' : '<button data-h="insertar">＋ Insertar aquí</button><button data-h="comentario-punto">💬 Comentar aquí</button><button data-h="x" aria-label="Cerrar">✕</button>';
    document.body.appendChild(t);
    const top = Math.max(8, caja.top - t.offsetHeight - 8), left = Math.min(Math.max(8, caja.left + caja.width / 2 - t.offsetWidth / 2), window.innerWidth - t.offsetWidth - 8);
    t.style.top = (caja.top > 70 ? top : caja.bottom + 8) + 'px'; t.style.left = left + 'px';
    $$('[data-h]', t).forEach(btn => btn.onmousedown = btn.ontouchstart = ev => ev.preventDefault());
    $$('[data-h]', t).forEach(btn => btn.onclick = () => { const h = btn.dataset.h; quitarHerramientas(); if (h !== 'x') nuevaNota(h === 'comentario-punto' ? 'insertar-com' : h); });
  }
  function choca(nueva) {
    return E.anotaciones.some(x => x.p === nueva.p && (nueva.tipo === 'insertar'
      ? (x.tipo !== 'insertar' && nueva.ini > x.ini && nueva.ini < x.fin)
      : (x.tipo !== 'insertar' ? nueva.ini < x.fin && nueva.fin > x.ini : x.ini > nueva.ini && x.ini < nueva.fin)));
  }
  function nuevaNota(tipo) {
    if (!selActual) return;
    const s = { ...selActual };
    let nota = tipo === 'insertar-com' ? { p: s.p, ini: s.ini, fin: s.ini, tipo: 'insertar', nuevo: '' } : { p: s.p, ini: s.ini, fin: tipo === 'insertar' ? s.ini : s.fin, tipo };
    if (choca(nota)) return toast('Ya hay una marca en esa parte. Quítala o elige otro pedazo.', true);
    const frag = E.parrafos[s.p].texto.slice(s.ini, s.fin);
    const pideNuevo = tipo === 'reemplazar' || tipo === 'insertar';
    const pideCom = tipo !== 'insertar';
    modal(`<h2>${{ comentario: 'Comentar', borrar: 'Tachar', reemplazar: 'Cambiar por…', insertar: 'Insertar texto', 'insertar-com': 'Comentar aquí' }[tipo]}</h2>
      ${frag ? `<p class="suave">«${esc(frag.length > 160 ? frag.slice(0, 160) + '…' : frag)}»</p>` : ''}
      ${pideNuevo ? `<label class="campo"><span>${tipo === 'reemplazar' ? 'Texto nuevo' : 'Texto a insertar'}</span><input type="text" name="nuevo" autocomplete="off"></label>` : ''}
      <label class="campo"><span>Comentario ${tipo === 'comentario' || tipo === 'insertar-com' ? '' : '(opcional)'}</span><textarea name="com" style="min-height:5rem"></textarea></label>
      <div class="botones"><button class="btn btn--rojo" data-si>Guardar</button><button class="btn" data-cerrar>Cancelar</button></div>`, (m, cerrar) => {
      $('[data-si]', m).onclick = () => {
        const nuevo = $('[name=nuevo]', m) ? $('[name=nuevo]', m).value : ''; const com = $('[name=com]', m).value.trim();
        if (pideNuevo && !nuevo) return toast('Escribe el texto.', true);
        if ((tipo === 'comentario' || tipo === 'insertar-com') && !com) return toast('Escribe el comentario.', true);
        if (tipo === 'insertar-com') { nota = { p: s.p, ini: s.ini, fin: s.ini, tipo: 'insertar', nuevo: '', comentario: com }; }
        else { if (pideNuevo) nota.nuevo = (tipo === 'insertar' && s.ini > 0 && !/\s$/.test(E.parrafos[s.p].texto.slice(0, s.ini)) && !/^[\s,.;:!?)]/.test(nuevo) ? ' ' : '') + nuevo; if (com) nota.comentario = com; }
        E.anotaciones.push(nota); cerrar(); window.getSelection().removeAllRanges(); guardarBorradorPronto(); repintar();
      };
    });
  }
  function editarNota(i) {
    const a = E.anotaciones[i];
    modal(`<h2>Editar</h2>${a.tipo === 'reemplazar' || (a.tipo === 'insertar' && a.nuevo) ? `<label class="campo"><span>Texto</span><input type="text" name="nuevo" value="${esc((a.nuevo || '').trim())}"></label>` : ''}<label class="campo"><span>Comentario</span><textarea name="com">${esc(a.comentario || '')}</textarea></label><div class="botones"><button class="btn btn--rojo" data-si>Guardar</button><button class="btn" data-cerrar>Cancelar</button></div>`, (m, cerrar) => {
      $('[data-si]', m).onclick = () => { const n = $('[name=nuevo]', m); if (n) a.nuevo = (a.tipo === 'insertar' && /^\s/.test(a.nuevo || '') ? ' ' : '') + n.value; a.comentario = $('[name=com]', m).value.trim(); cerrar(); guardarBorradorPronto(); repintar(); };
    });
  }
  async function entregar() {
    const word = E.modo === 'word';
    const n = E.rub ? E.rub.criterios.length : 0;
    let { puntaje, dictamen } = calcular(E.niveles, n);
    if (word) {
      if (!E.subidas.rubrica) return toast('Sube tu rúbrica llena.', true);
      if (!E.puntajeWord || E.puntajeWord < 1 || E.puntajeWord > 5) return toast('Escribe el puntaje final (de 1 a 5).', true);
      puntaje = Math.round(E.puntajeWord * 10) / 10; dictamen = puntaje >= 4.5 ? 'aceptada' : puntaje >= 3.5 ? 'ajustes' : 'no';
    } else if (puntaje == null) return toast('Califica todos los criterios de la rúbrica.', true);
    const ok = await confirmar('¿Entregar la evaluación?', `Puntaje <b>${punt(puntaje)}</b> · <b>${DICT[dictamen]}</b>${word ? '' : ` · ${E.anotaciones.length} corrección(es) en el texto`}. Después de entregarla ya no se puede cambiar.`, 'Entregar');
    if (!ok) return;
    const b = $('[data-entregar]'); if (b) { b.disabled = true; b.textContent = 'Entregando…'; }
    try {
      const datos = { id: E.id, modo: E.modo, puntaje, dictamen, observaciones: E.observaciones, enlace: E.enlace };
      if (word) { datos.niveles = []; datos.anotaciones = []; datos.subidaRubrica = E.subidas.rubrica; if (E.subidas.correccion) datos.subidaCorreccion = E.subidas.correccion; }
      else {
        datos.niveles = E.niveles; datos.anotaciones = E.anotaciones;
        // Rúbrica llena (interna del equipo), con la firma guardada en el perfil si la hay
        let firmaPng = null; try { const f = (await api('miFirma')).firma; if (f) firmaPng = Docx.deBase64(f); } catch (e) { }
        datos.rubricaDocx = Docx.aBase64(await Docx.llenarRubrica(await plantillaRubrica(E.a.rubrica), { titulo: E.env.titulo, codigo: E.env.id, fecha: new Date().toLocaleDateString('es-CO'), evaluador: E.evaluador, enlace: E.enlace, niveles: E.niveles, puntaje, dictamen, observaciones: E.observaciones, firmaPng }));
        datos.firmadaAuto = !!firmaPng;
        // Para el autor: correcciones sin el nombre del evaluador
        if (E.anotaciones.length && E.parrafos) datos.correccionDocx = Docx.aBase64(await Docx.crearCorregido({ parrafos: E.parrafos, anotaciones: E.anotaciones, autor: 'Evaluación · Hicimos lo que quisimos', iniciales: 'HLQQ', titulo: `${E.env.id} · ${E.env.titulo}` }));
      }
      await api('entregar', datos);
      toast('¡Evaluación entregada! Ahora sube tu rúbrica firmada.'); location.hash = '#evaluacion/' + E.id; ruta();
    } catch (er) { toast(er.message, true); if (b) { b.disabled = false; b.textContent = 'Entregar evaluación'; } }
  }

  // =========================================================
  //  ADMINISTRACIÓN
  // =========================================================
  RUTAS.perfiles = {
    render: async () => {
      const { usuarios } = await api('usuarios'); S._us = usuarios;
      return `<div class="cabecera"><div><p class="meta">Administración</p><h1>Perfiles</h1><p class="suave">Crea, edita o borra perfiles. Si a alguien se le olvida la contraseña, dale una temporal con «Nueva contraseña».</p></div><button class="btn btn--rojo" data-nuevo>Nuevo perfil</button></div>
        <table class="tabla-perfiles"><thead><tr><th>Nombre</th><th>Perfiles</th><th>Correo para avisos</th><th>Firma</th><th></th></tr></thead><tbody>
        ${usuarios.map(u => `<tr><td><b>${esc(u.nombre)}</b>${u.activo ? '' : ' <span class="estado e-cerrado">Inactivo</span>'}${u.debeCambiar ? ' <span class="meta">· aún no cambia su contraseña</span>' : ''}</td><td>${u.roles.map(r => `<span class="estado" style="margin:.1rem">${ROLES[r]}</span>`).join(' ')}</td><td class="suave">${esc(u.correo || '—')}</td><td>${u.roles.includes('evaluador') ? (u.tieneFirma ? '<span class="estado e-listo_publicar">✓ Cargada</span>' : '<span class="estado e-por_asignar">Sin firma</span>') : '<span class="suave">—</span>'}</td><td><div class="botones">${u.roles.includes('evaluador') ? `<button class="btn btn--chico" data-firma-u="${esc(u.id)}">Firma</button>` : ''}<button class="btn btn--chico" data-editar-u="${esc(u.id)}">Editar</button><button class="btn btn--chico" data-clave-u="${esc(u.id)}">Nueva contraseña</button></div></td></tr>`).join('')}
        </tbody></table>`;
    },
    conectar: () => {
      $('[data-nuevo]').onclick = () => editarUsuario(null);
      $$('[data-editar-u]').forEach(b => b.onclick = () => editarUsuario(S._us.find(u => u.id === b.dataset.editarU)));
      $$('[data-firma-u]').forEach(b => b.onclick = () => modalFirmaPerfil(S._us.find(u => u.id === b.dataset.firmaU)));
      $$('[data-clave-u]').forEach(b => b.onclick = async () => {
        const u = S._us.find(x => x.id === b.dataset.claveU);
        if (!await confirmar(`¿Nueva contraseña para ${u.nombre}?`, 'Su contraseña actual deja de funcionar. Le daremos una temporal y, al entrar, tendrá que crear la suya.')) return;
        try { const r = await api('restablecerClave', { id: u.id }); await ruta(); mostrarClave(u.nombre, r.claveTemporal); } catch (er) { toast(er.message, true); }
      });
    },
  };
  async function modalFirmaPerfil(u) {
    let f = null; try { f = (await api('firmaDe', { usuarioId: u.id })).firma; } catch (e) { }
    modal(`<h2>Firma de ${esc(u.nombre)}</h2><p class="suave">Esta firma se pondrá sola en cada rúbrica que ${esc(u.nombre.split(' ')[0])} entregue, en la línea «Firma del evaluador(a)», y la rúbrica quedará firmada.</p>
      ${f ? `<div class="firma-vista"><img src="data:image/png;base64,${f}" alt="Firma"></div>` : '<div class="vacio" style="padding:1rem">Todavía no tiene firma cargada.</div>'}
      <label class="campo" style="margin-top:1rem"><span>${f ? 'Cambiar la firma' : 'Cargar la firma'} (imagen PNG, mejor con fondo blanco o transparente)</span><input type="file" accept="image/png" name="png"></label>
      <p class="suave" style="font-size:.85rem">Pídele que firme en una hoja blanca y te mande una foto bien iluminada; recórtala y guárdala como PNG.</p>
      <div class="botones"><button class="btn btn--rojo" data-si>Guardar firma</button>${f ? '<button class="btn" data-quitar>Quitar firma</button>' : ''}<button class="btn" data-cerrar>Cerrar</button></div>`, (m, cerrar) => {
      $('[data-si]', m).onclick = async () => {
        const file = $('[name=png]', m).files[0]; if (!file) return toast('Elige la imagen de la firma (PNG).', true);
        if (!/png$/i.test(file.type) && !/\.png$/i.test(file.name)) return toast('La firma debe ser una imagen PNG.', true);
        if (file.size > 1.5e6) return toast('La imagen pesa demasiado (máximo 1,5 MB).', true);
        try { const d = await leerArchivoLocal(file); await api('guardarFirma', { usuarioId: u.id, b64: d.b64 }); cerrar(); toast('Firma guardada ✓'); ruta(); } catch (er) { toast(er.message, true); }
      };
      const q = $('[data-quitar]', m); if (q) q.onclick = async () => { try { await api('quitarFirma', { usuarioId: u.id }); cerrar(); toast('Firma quitada'); ruta(); } catch (er) { toast(er.message, true); } };
    });
  }
  function mostrarClave(nombre, clave) {
    modal(`<h2>Contraseña temporal</h2><p class="suave">Dásela a <b>${esc(nombre)}</b> en persona o por mensaje privado. Al entrar, la página le pedirá crear su propia contraseña.</p><div class="clave-temporal">${esc(clave)}</div><div class="botones" style="margin-top:1rem"><button class="btn" data-copiar>Copiar</button><button class="btn btn--claro" data-cerrar>Listo</button></div>`, m => { $('[data-copiar]', m).onclick = () => navigator.clipboard.writeText(clave).then(() => toast('Copiada')); });
  }
  function editarUsuario(u) {
    const r = u ? u.roles : ['evaluador'];
    modal(`<h2>${u ? 'Editar perfil' : 'Nuevo perfil'}</h2>
      <label class="campo"><span>Nombre completo</span><input type="text" name="nombre" value="${esc(u ? u.nombre : '')}"></label>
      <label class="campo"><span>Correo para avisos (opcional)</span><input type="email" name="correo" value="${esc(u ? u.correo : '')}" placeholder="nombre@gmail.com"></label>
      <span class="meta">Perfiles</span>${Object.entries(ROLES).map(([k, v]) => `<label class="chequeo"><input type="checkbox" name="rol" value="${k}" ${r.includes(k) ? 'checked' : ''}> ${v}</label>`).join('')}
      <label class="chequeo" style="margin-top:.6rem"><input type="checkbox" name="avisos" ${!u || u.avisos ? 'checked' : ''}> Mandarle avisos por correo</label>
      <label class="chequeo"><input type="checkbox" name="activo" ${!u || u.activo ? 'checked' : ''}> Perfil activo (puede entrar)</label>
      <div class="botones" style="margin-top:1rem"><button class="btn btn--rojo" data-si>${u ? 'Guardar' : 'Crear perfil'}</button><button class="btn" data-cerrar>Cancelar</button>${u && u.id !== S.usuario.id ? '<button class="btn btn--chico" data-borrar style="margin-left:auto">Borrar perfil</button>' : ''}</div>`, (m, cerrar) => {
      $('[data-si]', m).onclick = async () => {
        const d = { id: u && u.id, nombre: $('[name=nombre]', m).value.trim(), correo: $('[name=correo]', m).value.trim(), roles: $$('[name=rol]:checked', m).map(x => x.value), avisos: $('[name=avisos]', m).checked, activo: $('[name=activo]', m).checked };
        try { const r2 = await api('guardarUsuario', { usuario: d }); cerrar(); toast('Perfil guardado'); if (u && u.id === S.usuario.id) { S.usuario = { ...S.usuario, ...r2.usuario }; guardarSesion(); } await ruta(); if (r2.claveTemporal) mostrarClave(d.nombre, r2.claveTemporal); }
        catch (er) { toast(er.message, true); }
      };
      const bb = $('[data-borrar]', m); if (bb) bb.onclick = async () => { cerrar(); if (!await confirmar(`¿Borrar el perfil de ${u.nombre}?`, 'No se puede deshacer. Sus evaluaciones entregadas se conservan.', 'Borrar')) return; try { await api('borrarUsuario', { id: u.id }); toast('Perfil borrado'); ruta(); } catch (er) { toast(er.message, true); } };
    });
  }
  RUTAS.ajustes = {
    render: async () => {
      const { ajustes: a } = await api('ajustes');
      return `<div style="max-width:620px"><p class="meta">Administración</p><h1>Ajustes</h1><form class="tarjeta" id="f-aj">
        <label class="campo"><span>Dirección de esta página (para los enlaces de los avisos)</span><input type="url" name="urlPagina" value="${esc(a.urlPagina || location.href.split('#')[0])}"></label>
        <label class="campo"><span>Nombre que aparece en los correos</span><input type="text" name="nombreRemitente" value="${esc(a.nombreRemitente)}"></label>
        <label class="campo"><span>Recordar a los evaluadores cada cuántos días</span><input type="number" min="1" max="60" name="diasRecordatorio" value="${esc(a.diasRecordatorio)}"></label>
        <label class="chequeo"><input type="checkbox" name="avisosActivos" ${String(a.avisosActivos) !== 'false' ? 'checked' : ''}> Mandar avisos por correo (la campanita funciona siempre)</label>
        <button class="btn btn--rojo" style="margin-top:1rem">Guardar</button></form>
        <p class="suave">Versión ${API.demo ? 'de demostración' : 'conectada al Gmail de la gaceta'}.</p></div>`;
    },
    conectar: () => $('#f-aj').onsubmit = async e => { e.preventDefault(); const f = e.target; try { await api('guardarAjustes', { ajustes: { urlPagina: f.urlPagina.value, nombreRemitente: f.nombreRemitente.value, diasRecordatorio: +f.diasRecordatorio.value || 7, avisosActivos: f.avisosActivos.checked } }); toast('Ajustes guardados'); } catch (er) { toast(er.message, true); } },
  };

  // ---------- Arranque ----------
  (async () => { if (S.token) await actualizarNotificaciones(); ruta(); })();
})();
