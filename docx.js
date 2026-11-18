/* =========================================================
   docx.js · Leer y escribir documentos de Word en el navegador
   Sin librerías externas. Sirve para:
   - leer el texto de un .docx (párrafos con negrita y cursiva)
   - crear el texto corregido con COMENTARIOS y CONTROL DE CAMBIOS reales de Word
   - llenar las rúbricas de la gaceta conservando su formato
   ========================================================= */
(function (raiz) {
  'use strict';

  // ---------- ZIP ----------
  const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
  const crc32 = b => { let c = 0xFFFFFFFF; for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; };
  const enc = new TextEncoder(), dec = new TextDecoder('utf-8');

  async function transformar(bytes, stream) {
    const s = new Blob([bytes]).stream().pipeThrough(stream);
    return new Uint8Array(await new Response(s).arrayBuffer());
  }
  const inflar = b => transformar(b, new DecompressionStream('deflate-raw'));
  const desinflar = b => transformar(b, new CompressionStream('deflate-raw'));

  async function leerZip(buf) {
    const b = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    const v = new DataView(b.buffer, b.byteOffset, b.byteLength);
    let e = b.length - 22;
    while (e >= 0 && v.getUint32(e, true) !== 0x06054b50) e--;
    if (e < 0) throw new Error('No es un archivo .docx válido');
    const total = v.getUint16(e + 10, true); let p = v.getUint32(e + 16, true);
    const archivos = {};
    for (let i = 0; i < total; i++) {
      const metodo = v.getUint16(p + 10, true), tam = v.getUint32(p + 20, true);
      const ln = v.getUint16(p + 28, true), lx = v.getUint16(p + 30, true), lc = v.getUint16(p + 32, true);
      const local = v.getUint32(p + 42, true);
      const nombre = dec.decode(b.subarray(p + 46, p + 46 + ln));
      const lnl = v.getUint16(local + 26, true), lxl = v.getUint16(local + 28, true);
      const datos = b.subarray(local + 30 + lnl + lxl, local + 30 + lnl + lxl + tam);
      archivos[nombre] = { metodo, datos };
      p += 46 + ln + lx + lc;
    }
    return {
      nombres: Object.keys(archivos),
      async bytes(n) { const a = archivos[n]; if (!a) return null; return a.metodo === 8 ? inflar(a.datos) : a.datos.slice(); },
      async texto(n) { const x = await this.bytes(n); return x ? dec.decode(x) : null; },
    };
  }

  async function escribirZip(entradas) { // [{nombre, datos: Uint8Array|string}]
    const partes = [], central = []; let off = 0;
    for (const en of entradas) {
      const nom = enc.encode(en.nombre);
      const crudo = typeof en.datos === 'string' ? enc.encode(en.datos) : en.datos;
      const comp = await desinflar(crudo);
      const crc = crc32(crudo);
      const h = new DataView(new ArrayBuffer(30));
      h.setUint32(0, 0x04034b50, true); h.setUint16(4, 20, true); h.setUint16(6, 0x0800, true); h.setUint16(8, 8, true);
      h.setUint32(14, crc, true); h.setUint32(18, comp.length, true); h.setUint32(22, crudo.length, true); h.setUint16(26, nom.length, true);
      partes.push(new Uint8Array(h.buffer), nom, comp);
      const c = new DataView(new ArrayBuffer(46));
      c.setUint32(0, 0x02014b50, true); c.setUint16(4, 20, true); c.setUint16(6, 20, true); c.setUint16(8, 0x0800, true); c.setUint16(10, 8, true);
      c.setUint32(16, crc, true); c.setUint32(20, comp.length, true); c.setUint32(24, crudo.length, true); c.setUint16(28, nom.length, true); c.setUint32(42, off, true);
      central.push(new Uint8Array(c.buffer), nom);
      off += 30 + nom.length + comp.length;
    }
    const tamC = central.reduce((s, x) => s + x.length, 0);
    const f = new DataView(new ArrayBuffer(22));
    f.setUint32(0, 0x06054b50, true); f.setUint16(8, entradas.length, true); f.setUint16(10, entradas.length, true); f.setUint32(12, tamC, true); f.setUint32(16, off, true);
    const todo = [...partes, ...central, new Uint8Array(f.buffer)];
    const out = new Uint8Array(todo.reduce((s, x) => s + x.length, 0)); let i = 0;
    for (const x of todo) { out.set(x, i); i += x.length; }
    return out;
  }

  // ---------- XML ----------
  const escXml = s => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const desXml = s => s.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(+n)).replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16))).replace(/&amp;/g, '&');
  const activo = (rPr, tag) => { const m = rPr.match(new RegExp(`<w:${tag}(?:\\s+w:val="([^"]*)")?\\s*/>`)); return !!m && !(m[1] && /^(0|false|off)$/i.test(m[1])); };

  // ---------- Leer texto ----------
  // Devuelve [{runs:[{t, b, i}], texto}] — un elemento por párrafo
  async function leerParrafos(bytes) {
    const zip = await leerZip(bytes);
    const xml = await zip.texto('word/document.xml');
    if (!xml) throw new Error('El archivo no tiene texto de Word');
    const cuerpo = xml.slice(xml.indexOf('<w:body'), xml.lastIndexOf('</w:body>'));
    const parrafos = [];
    const reP = /<w:p\/>|<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g; let m;
    while ((m = reP.exec(cuerpo))) {
      const px = m[0]; const runs = [];
      const reR = /<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g; let r;
      while ((r = reR.exec(px))) {
        const rx = r[0];
        const rPr = (rx.match(/<w:rPr>[\s\S]*?<\/w:rPr>/) || [''])[0];
        let t = '';
        const reT = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:t\/>|<w:tab\/>|<w:br(?:\s[^>]*)?\/>|<w:cr\/>|<w:noBreakHyphen\/>/g; let q;
        while ((q = reT.exec(rx))) {
          if (q[0].startsWith('<w:tab')) t += '\t';
          else if (q[0].startsWith('<w:br') || q[0].startsWith('<w:cr')) t += '\n';
          else if (q[0].startsWith('<w:noBreak')) t += '-';
          else if (q[1] !== undefined) t += desXml(q[1]);
        }
        if (t) runs.push({ t, b: activo(rPr, 'b'), i: activo(rPr, 'i') });
      }
      // juntar runs con el mismo formato
      const junt = [];
      for (const x of runs) { const u = junt[junt.length - 1]; if (u && u.b === x.b && u.i === x.i) u.t += x.t; else junt.push({ ...x }); }
      parrafos.push({ runs: junt, texto: junt.map(x => x.t).join('') });
    }
    return parrafos;
  }

  // ---------- Texto corregido con comentarios y control de cambios ----------
  // anotaciones: [{p, ini, fin, tipo:'comentario'|'borrar'|'reemplazar'|'insertar', nuevo, comentario}]
  function trozos(par, cortes) {
    // parte los runs del párrafo en los puntos de corte; devuelve [{ini, fin, runs:[...]}] entre cortes
    const pts = [...new Set([0, par.texto.length, ...cortes])].sort((a, b) => a - b);
    const seg = [];
    for (let k = 0; k < pts.length - 1; k++) {
      const a = pts[k], z = pts[k + 1]; if (a === z) continue;
      const rs = []; let pos = 0;
      for (const r of par.runs) {
        const ra = pos, rz = pos + r.t.length; pos = rz;
        const x = Math.max(a, ra), y = Math.min(z, rz);
        if (x < y) rs.push({ ...r, t: r.t.slice(x - ra, y - ra) });
      }
      seg.push({ ini: a, fin: z, runs: rs });
    }
    return seg;
  }
  const rPrX = (r, extra = '') => `<w:rPr>${r.b ? '<w:b/>' : ''}${r.i ? '<w:i/>' : ''}${extra}</w:rPr>`;
  function textoRun(t, tag = 't') {
    // respeta tabulaciones y saltos dentro del run
    return t.split(/(\t|\n)/).map(x => x === '\t' ? '<w:tab/>' : x === '\n' ? '<w:br/>' : x ? `<w:${tag} xml:space="preserve">${escXml(x)}</w:${tag}>` : '').join('');
  }

  async function crearCorregido({ parrafos, anotaciones = [], autor = 'Evaluador/a', iniciales = 'EV', titulo = '', fecha = new Date() }) {
    const iso = new Date(fecha).toISOString().replace(/\.\d+Z$/, 'Z');
    let id = 1; const comentarios = [];
    const cuerpo = parrafos.map((par, pi) => {
      const ans = anotaciones.filter(a => a.p === pi).sort((a, b) => a.ini - b.ini || (a.tipo === 'insertar' ? -1 : 1));
      const cortes = ans.flatMap(a => [a.ini, a.fin ?? a.ini]);
      const segs = trozos(par, cortes);
      let xml = '';
      const insertarEn = pos => ans.filter(a => a.tipo === 'insertar' && a.ini === pos).map(a => {
        let s = `<w:ins w:id="${id++}" w:author="${escXml(autor)}" w:date="${iso}"><w:r>${textoRun(a.nuevo || '')}</w:r></w:ins>`;
        if (a.comentario) { const c = id++; comentarios.push({ id: c, texto: a.comentario }); s = `<w:commentRangeStart w:id="${c}"/>${s}<w:commentRangeEnd w:id="${c}"/><w:r><w:rPr><w:rStyle w:val="Refdecomentario"/></w:rPr><w:commentReference w:id="${c}"/></w:r>`; }
        return s;
      }).join('');
      for (const sg of segs) {
        xml += insertarEn(sg.ini);
        const a = ans.find(x => x.tipo !== 'insertar' && x.ini <= sg.ini && (x.fin ?? x.ini) >= sg.fin);
        const base = sg.runs.map(r => `<w:r>${rPrX(r)}${textoRun(r.t)}</w:r>`).join('');
        if (!a) { xml += base; continue; }
        let pedazo;
        if (a.tipo === 'borrar' || a.tipo === 'reemplazar') {
          pedazo = `<w:del w:id="${id++}" w:author="${escXml(autor)}" w:date="${iso}">${sg.runs.map(r => `<w:r>${rPrX(r)}${textoRun(r.t, 'delText')}</w:r>`).join('')}</w:del>`;
          if (a.tipo === 'reemplazar' && sg.fin === a.fin) pedazo += `<w:ins w:id="${id++}" w:author="${escXml(autor)}" w:date="${iso}"><w:r>${rPrX(sg.runs[0] || {})}${textoRun(a.nuevo || '')}</w:r></w:ins>`;
        } else pedazo = base;
        if (a.comentario || a.tipo === 'comentario') {
          if (sg.ini === a.ini) { a._cid = id++; comentarios.push({ id: a._cid, texto: a.comentario || '' }); pedazo = `<w:commentRangeStart w:id="${a._cid}"/>` + pedazo; }
          if (sg.fin === a.fin) pedazo += `<w:commentRangeEnd w:id="${a._cid}"/><w:r><w:rPr><w:rStyle w:val="Refdecomentario"/></w:rPr><w:commentReference w:id="${a._cid}"/></w:r>`;
        }
        xml += pedazo;
      }
      if (!segs.length) xml += insertarEn(0); else xml += insertarEn(par.texto.length);
      return `<w:p>${xml}</w:p>`;
    }).join('');

    const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
    const documento = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${NS}><w:body>${titulo ? `<w:p><w:pPr><w:pStyle w:val="Encabezado"/></w:pPr><w:r><w:t xml:space="preserve">${escXml(titulo)}</w:t></w:r></w:p>` : ''}${cuerpo}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1418" w:right="1418" w:bottom="1418" w:left="1418" w:header="709" w:footer="709" w:gutter="0"/></w:sectPr></w:body></w:document>`;
    const coments = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments ${NS}>${comentarios.map(c => `<w:comment w:id="${c.id}" w:author="${escXml(autor)}" w:date="${iso}" w:initials="${escXml(iniciales)}">${String(c.texto).split('\n').map((l, k) => `<w:p><w:pPr><w:pStyle w:val="Textocomentario"/></w:pPr>${k === 0 ? '<w:r><w:rPr><w:rStyle w:val="Refdecomentario"/></w:rPr><w:annotationRef/></w:r>' : ''}<w:r><w:t xml:space="preserve">${escXml(l)}</w:t></w:r></w:p>`).join('')}</w:comment>`).join('')}</w:comments>`;
    const estilos = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia" w:cs="Georgia" w:eastAsia="Georgia"/><w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="es-CO"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="120" w:line="300" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Encabezado"><w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:b/><w:color w:val="1E2747"/><w:sz w:val="32"/></w:rPr></w:style>
<w:style w:type="paragraph" w:styleId="Textocomentario"><w:name w:val="annotation text"/><w:basedOn w:val="Normal"/><w:rPr><w:sz w:val="20"/></w:rPr></w:style>
<w:style w:type="character" w:styleId="Refdecomentario"><w:name w:val="annotation reference"/><w:rPr><w:sz w:val="16"/></w:rPr></w:style></w:styles>`;
    const tipos = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`;
    const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`;
    const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>`;
    const settings = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings ${NS}><w:trackRevisions/><w:defaultTabStop w:val="709"/></w:settings>`;
    return escribirZip([
      { nombre: '[Content_Types].xml', datos: tipos }, { nombre: '_rels/.rels', datos: rels },
      { nombre: 'word/document.xml', datos: documento }, { nombre: 'word/_rels/document.xml.rels', datos: docRels },
      { nombre: 'word/styles.xml', datos: estilos }, { nombre: 'word/comments.xml', datos: coments }, { nombre: 'word/settings.xml', datos: settings },
    ]);
  }

  // ---------- Rúbricas: llenar la plantilla original ----------
  function bloques(xml, abre, cierra, desde = 0, hasta = xml.length) {
    // devuelve [{ini, fin}] de elementos del primer nivel (maneja anidación)
    const out = []; let i = desde;
    const reA = new RegExp(`<${abre}(?=[\\s>/])`, 'g');
    while (i < hasta) {
      reA.lastIndex = i; const m = reA.exec(xml); if (!m || m.index >= hasta) break;
      let prof = 0, j = m.index;
      const reAll = new RegExp(`<${abre}(?=[\\s>/])|</${cierra}>`, 'g'); reAll.lastIndex = j;
      let k;
      while ((k = reAll.exec(xml))) {
        if (k[0].startsWith('</')) { prof--; if (prof === 0) { j = k.index + k[0].length; break; } }
        else { const fin = xml.indexOf('>', k.index); if (xml[fin - 1] === '/') { if (prof === 0) { j = fin + 1; break; } } else prof++; }
      }
      out.push({ ini: m.index, fin: j }); i = j;
    }
    return out;
  }
  const textoDe = x => (x.match(/<w:t(?:\s[^>]*)?>[\s\S]*?<\/w:t>/g) || []).map(t => desXml(t.replace(/<[^>]+>/g, ''))).join('');
  const marcar = celda => celda.replace('☐', '☒');
  function agregarTexto(celda, texto, { tam = 18, color = '1F2A4A', negrita = false } = {}) {
    const run = String(texto).split('\n').map((l, k) => `${k ? '<w:br/>' : ''}<w:t xml:space="preserve">${escXml(l)}</w:t>`).join('');
    const r = `<w:r><w:rPr>${negrita ? '<w:b/>' : ''}<w:color w:val="${color}"/><w:sz w:val="${tam}"/><w:szCs w:val="${tam}"/></w:rPr>${run}</w:r>`;
    const ps = bloques(celda, 'w:p', 'w:p');
    if (!ps.length) return celda.replace('</w:tc>', `<w:p>${r}</w:p></w:tc>`);
    const p = ps[ps.length - 1]; let px = celda.slice(p.ini, p.fin);
    px = px.endsWith('/>') ? px.replace(/\/>$/, `>${r}</w:p>`) : px.replace(/<\/w:p>$/, `${r}</w:p>`);
    return celda.slice(0, p.ini) + px + celda.slice(p.fin);
  }
  function reemplazarCeldas(tabla, fn) {
    // fn(filaIdx, celdaIdx, xmlCelda) -> xmlCelda
    let out = '', ult = 0;
    bloques(tabla, 'w:tr', 'w:tr').forEach((tr, fi) => {
      out += tabla.slice(ult, tr.ini);
      let fila = tabla.slice(tr.ini, tr.fin), fo = '', u2 = 0;
      bloques(fila, 'w:tc', 'w:tc').forEach((tc, ci) => { fo += fila.slice(u2, tc.ini) + fn(fi, ci, fila.slice(tc.ini, tc.fin)); u2 = tc.fin; });
      fo += fila.slice(u2); out += fo; ult = tr.fin;
    });
    return out + tabla.slice(ult);
  }

  // datos: {titulo, codigo, fecha, evaluador, enlace, niveles:[1..5 por criterio], puntaje, dictamen:'no'|'ajustes'|'aceptada', observaciones}
  async function llenarRubrica(plantilla, datos) {
    const zip = await leerZip(plantilla);
    let xml = await zip.texto('word/document.xml');
    const tablas = bloques(xml, 'w:tbl', 'w:tbl');
    const piezas = tablas.map(t => ({ ...t, xml: xml.slice(t.ini, t.fin) }));
    piezas.forEach((t, ti) => {
      const txt = textoDe(t.xml);
      if (ti === 0) { // datos del encabezado
        t.nuevo = reemplazarCeldas(t.xml, (fi, ci, c) => c); // se llena abajo (necesita etiqueta de la izquierda)
        const filas = bloques(t.xml, 'w:tr', 'w:tr');
        let out = '', ult = 0;
        filas.forEach(tr => {
          out += t.xml.slice(ult, tr.ini);
          const fila = t.xml.slice(tr.ini, tr.fin); const celdas = bloques(fila, 'w:tc', 'w:tc');
          let fo = '', u2 = 0;
          celdas.forEach((tc, ci) => {
            let c = fila.slice(tc.ini, tc.fin);
            const etiqueta = ci > 0 ? textoDe(fila.slice(celdas[ci - 1].ini, celdas[ci - 1].fin)) : '';
            if (ci % 2 === 1 && !textoDe(c).trim()) {
              const v = /t[ií]tulo/i.test(etiqueta) ? datos.titulo : /c[oó]digo/i.test(etiqueta) ? datos.codigo : /fecha/i.test(etiqueta) ? datos.fecha : /evaluador/i.test(etiqueta) ? datos.evaluador : /enlace|archivo|red/i.test(etiqueta) ? datos.enlace : '';
              if (v) c = agregarTexto(c, v);
            }
            fo += fila.slice(u2, tc.ini) + c; u2 = tc.fin;
          });
          out += fo + fila.slice(u2); ult = tr.fin;
        });
        t.nuevo = out + t.xml.slice(ult);
      } else if (/^\s*Criterio/.test(txt)) { // criterios
        let k = -1;
        t.nuevo = reemplazarCeldas(t.xml, (fi, ci, c) => {
          if (fi === 0) return c;
          const primera = ci === 0 ? textoDe(c) : null;
          if (ci === 0) { if (/^\s*Puntaje/i.test(primera)) k = 'puntaje'; else k = fi - 1; return c; }
          if (k === 'puntaje') return ci === 1 && datos.puntaje != null ? agregarTexto(c, `  ${String(datos.puntaje).replace('.', ',')}`, { tam: 22, color: 'D93A2B', negrita: true }) : c;
          return datos.niveles && Number(datos.niveles[k]) === ci ? marcar(c) : c;
        });
      } else if (/No aceptad/.test(txt) && /Aceptada con ajustes/.test(txt)) { // dictamen
        const idx = { no: 0, ajustes: 1, aceptada: 2 }[datos.dictamen];
        t.nuevo = reemplazarCeldas(t.xml, (fi, ci, c) => ci === idx ? marcar(c) : c);
      } else if (/Observaciones/.test(txt)) {
        t.nuevo = reemplazarCeldas(t.xml, (fi, ci, c) => datos.observaciones ? agregarTexto(c, '\n' + datos.observaciones, { tam: 18, color: '333333' }) : c);
      }
    });
    let out = '', ult = 0;
    for (const t of piezas) { out += xml.slice(ult, t.ini) + (t.nuevo ?? t.xml); ult = t.fin; }
    xml = out + xml.slice(ult);
    // Firma: imagen de la firma (si la hay) y nombre de quien evaluó
    let conFirma = false;
    if (datos.evaluador || datos.firmaPng) {
      const f = xml.indexOf('Firma del evaluador');
      if (f > -1) {
        const m = /_{6,}/.exec(xml.slice(f, f + 4000));
        if (m) {
          const a = f + m.index; let nuevo = escXml(datos.evaluador || '');
          if (datos.firmaPng && datos.firmaPng.length > 24) {
            const b = datos.firmaPng, ancho = (b[16] << 24 | b[17] << 16 | b[18] << 8 | b[19]) >>> 0, alto = (b[20] << 24 | b[21] << 16 | b[22] << 8 | b[23]) >>> 0;
            let cy = 560000, cx = Math.round(cy * (ancho / (alto || 1))); if (cx > 2200000) { cx = 2200000; cy = Math.round(cx * (alto / (ancho || 1))); }
            const dib = `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="9001" name="Firma del evaluador"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="9001" name="firma-evaluador.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdFirmaHLQQ"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
            nuevo = `</w:t></w:r>${dib}<w:r><w:t xml:space="preserve">  ${escXml(datos.evaluador || '')}`;
            conFirma = true;
          }
          xml = xml.slice(0, a) + nuevo + xml.slice(a + m[0].length);
        }
      }
    }
    const entradas = [];
    for (const n of zip.nombres) {
      let d = n === 'word/document.xml' ? xml : await zip.bytes(n);
      if (conFirma && n === 'word/_rels/document.xml.rels') d = (await zip.texto(n)).replace('</Relationships>', '<Relationship Id="rIdFirmaHLQQ" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/firma-evaluador.png"/></Relationships>');
      if (conFirma && n === '[Content_Types].xml') { d = await zip.texto(n); if (!/Extension="png"/i.test(d)) d = d.replace(/(<Types[^>]*>)/, '$1<Default Extension="png" ContentType="image/png"/>'); }
      entradas.push({ nombre: n, datos: d });
    }
    if (conFirma) entradas.push({ nombre: 'word/media/firma-evaluador.png', datos: datos.firmaPng });
    return escribirZip(entradas);
  }

  // ---------- Utilidades ----------
  const aBase64 = u8 => { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s); };
  const deBase64 = b64 => { const s = atob(b64); const u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u; };
  // Un .docx sencillo a partir de texto (para pruebas y para textos pegados en el correo)
  // ---------- Leer las correcciones que alguien hizo en Word ----------
  // Compara el Word corregido con el texto original y devuelve anotaciones con el mismo formato
  // que usa el corrector de la página: {tipo: borrar|reemplazar|insertar|comentario, p, ini, fin, nuevo, comentario}.
  // Lee el control de cambios (w:ins / w:del) y los comentarios; si no hay control de cambios, compara palabra por palabra.
  async function leerCorrecciones(bytes, original) {
    const zip = await leerZip(bytes);
    const xml = await zip.texto('word/document.xml');
    if (!xml) throw new Error('El archivo no tiene texto de Word.');
    const comXml = (await zip.texto('word/comments.xml')) || '';
    const textoDeRuns = x => { let t = ''; const re = /<w:(?:t|delText)(?:\s[^>]*)?>([\s\S]*?)<\/w:(?:t|delText)>|<w:tab\/>|<w:br(?:\s[^>]*)?\/>/g; let q; while ((q = re.exec(x))) t += q[0].startsWith('<w:tab') ? '\t' : q[0].startsWith('<w:br') ? '\n' : desXml(q[1]); return t; };
    const comentarios = {};
    const reC = /<w:comment\s[^>]*w:id="(\d+)"[^>]*>([\s\S]*?)<\/w:comment>/g; let mc;
    while ((mc = reC.exec(comXml))) comentarios[mc[1]] = mc[2].split(/<\/w:p>/).map(textoDeRuns).filter(Boolean).join('\n').trim();
    const cuerpo = xml.slice(xml.indexOf('<w:body'), xml.lastIndexOf('</w:body>'));
    const pars = []; const reP = /<w:p\/>|<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g; let m;
    let hayCambios = false;
    while ((m = reP.exec(cuerpo))) {
      const px = m[0];
      const pPr = (px.match(/<w:pPr>[\s\S]*?<\/w:pPr>/) || [''])[0];
      const marcaIns = /<w:rPr>[\s\S]*?<w:ins\b[\s\S]*?<\/w:rPr>/.test(pPr); // párrafo nuevo (Enter con control de cambios)
      const piezas = []; // {k: 'n'|'i'|'d', t} y marcas de comentario {c: id, ab: true|false}
      const sin = px.replace(/<w:pPr>[\s\S]*?<\/w:pPr>/, '');
      const reX = /<w:(?:ins|moveTo)\b[^>]*>([\s\S]*?)<\/w:(?:ins|moveTo)>|<w:(?:del|moveFrom)\b[^>]*>([\s\S]*?)<\/w:(?:del|moveFrom)>|<w:commentRangeStart\b[^>]*w:id="(\d+)"[^>]*\/>|<w:commentRangeEnd\b[^>]*w:id="(\d+)"[^>]*\/>|<w:r(?:\s[^>]*)?>[\s\S]*?<\/w:r>/g; let x;
      while ((x = reX.exec(sin))) {
        if (x[1] !== undefined) { const t = textoDeRuns(x[1]); if (t) { piezas.push({ k: 'i', t }); hayCambios = true; } }
        else if (x[2] !== undefined) { const t = textoDeRuns(x[2]); if (t) { piezas.push({ k: 'd', t }); hayCambios = true; } }
        else if (x[3] !== undefined) piezas.push({ c: x[3], ab: true });
        else if (x[4] !== undefined) piezas.push({ c: x[4], ab: false });
        else { if (/<w:commentReference\b/.test(x[0])) continue; const t = textoDeRuns(x[0]); if (t) piezas.push({ k: 'n', t }); }
      }
      pars.push({ piezas, marcaIns });
    }
    // Unir párrafos cuyo salto fue insertado: el salto es una inserción de «\n» dentro del párrafo original
    const unidos = [];
    for (const p of pars) {
      const u = unidos[unidos.length - 1];
      if (u && u.abierto) { u.piezas.push({ k: 'i', t: '\n' }, ...p.piezas); u.abierto = p.marcaIns; hayCambios = true; }
      else unidos.push({ piezas: p.piezas.slice(), abierto: p.marcaIns });
    }
    const vacio = p => !p.piezas.some(z => z.k === 'n' || z.k === 'd');
    // El texto original de cada párrafo (normal + borrado) debe coincidir con la versión que se corrigió
    let propios = unidos.filter(p => !(vacio(p) && p.piezas.every(z => z.k !== 'i')));
    const origDe = p => p.piezas.filter(z => z.k === 'n' || z.k === 'd').map(z => z.t).join('');
    const finalDe = p => p.piezas.filter(z => z.k === 'n' || z.k === 'i').map(z => z.t).join('');
    const norm = t => String(t).replace(/\s+/g, ' ').trim();
    const orig = original.map(p => p.texto);
    const anotaciones = [];
    // Emparejar párrafos del Word con los originales (por texto original), dejando que haya párrafos nuevos o vacíos
    const pares = emparejar(orig.map(norm), propios.map(p => norm(hayCambios ? origDe(p) : finalDe(p))));
    let nuevosAntes = '';
    pares.forEach(([oi, wi]) => {
      if (oi === null) { const fp = propios[wi]; const t = finalDe(fp); if (t.trim()) nuevosAntes += (nuevosAntes ? '\n' : '') + t; return; }
      if (wi === null) { if (orig[oi].trim()) anotaciones.push({ tipo: 'borrar', p: oi, ini: 0, fin: orig[oi].length }); return; }
      const wp = propios[wi];
      if (nuevosAntes) { anotaciones.push({ tipo: 'insertar', p: oi, ini: 0, nuevo: nuevosAntes + '\n' }); nuevosAntes = ''; }
      if (hayCambios) delCambios(wp, oi); else delDiff(orig[oi], finalDe(wp), oi, wp);
    });
    if (nuevosAntes && orig.length) { const ult = orig.length - 1; anotaciones.push({ tipo: 'insertar', p: ult, ini: orig[ult].length, nuevo: '\n' + nuevosAntes }); }
    // Con control de cambios: recorrer las piezas y convertirlas en anotaciones sobre el texto original
    function delCambios(wp, pi) {
      const base = orig[pi]; let pos = 0; const abiertos = {};
      const pz = wp.piezas;
      // ajustar si el texto original del Word no coincide exactamente (espacios): se usa la posición en el Word
      for (let i = 0; i < pz.length; i++) {
        const z = pz[i];
        if (z.c !== undefined) { if (z.ab) abiertos[z.c] = pos; else { const ini = abiertos[z.c] ?? pos; agregarComentario(pi, ini, pos, comentarios[z.c] || ''); } continue; }
        if (z.k === 'n') { pos += z.t.length; continue; }
        if (z.k === 'd') {
          const ini = pos; pos += z.t.length;
          let j = i + 1; while (j < pz.length && pz[j].c !== undefined) j++;
          if (j < pz.length && pz[j].k === 'i') { anotaciones.push({ tipo: 'reemplazar', p: pi, ini, fin: pos, nuevo: pz[j].t }); i = j; }
          else anotaciones.push({ tipo: 'borrar', p: pi, ini, fin: pos });
          continue;
        }
        if (z.k === 'i') {
          const prev = anotaciones[anotaciones.length - 1];
          if (prev && prev.p === pi && prev.tipo === 'insertar' && prev.ini === pos) prev.nuevo += z.t;
          else anotaciones.push({ tipo: 'insertar', p: pi, ini: pos, nuevo: z.t });
        }
      }
      if (pos > base.length) { /* el Word trae texto de más: se recorta a la longitud del original */
        anotaciones.forEach(a => { if (a.p === pi) { a.ini = Math.min(a.ini, base.length); if (a.fin !== undefined) a.fin = Math.min(a.fin, base.length); } });
      }
    }
    // Sin control de cambios: comparar palabra por palabra
    function delDiff(a, b, pi, wp) {
      const ta = tokens(a), tb = tokens(b); const ops = lcs(ta.map(t => t.t), tb.map(t => t.t));
      let i = 0, j = 0, k = 0;
      while (k < ops.length) {
        if (ops[k] === '=') { i++; j++; k++; continue; }
        const ini = i < ta.length ? ta[i].ini : a.length; let fin = ini, nuevo = '';
        while (k < ops.length && ops[k] !== '=') { if (ops[k] === '-') { fin = ta[i].fin; i++; } else { nuevo += tb[j].t; j++; } k++; }
        if (fin > ini && nuevo) anotaciones.push({ tipo: 'reemplazar', p: pi, ini, fin, nuevo });
        else if (fin > ini) anotaciones.push({ tipo: 'borrar', p: pi, ini, fin });
        else if (nuevo) anotaciones.push({ tipo: 'insertar', p: pi, ini, nuevo });
      }
      // comentarios: se ubican buscando el texto comentado en el original
      let pos = 0; const abiertos = {}; const fb = finalDe(wp);
      wp.piezas.forEach(z => {
        if (z.c !== undefined) { if (z.ab) abiertos[z.c] = pos; else { const frag = fb.slice(abiertos[z.c] ?? pos, pos); const at = frag ? a.indexOf(frag) : -1; const ini = at > -1 ? at : Math.min(abiertos[z.c] ?? 0, a.length); agregarComentario(pi, ini, at > -1 ? at + frag.length : ini, comentarios[z.c] || ''); } return; }
        if (z.k === 'n' || z.k === 'i') pos += z.t.length;
      });
    }
    function agregarComentario(pi, ini, fin, texto) {
      if (!texto) return;
      const igual = anotaciones.find(a => a.p === pi && a.tipo !== 'insertar' && a.ini === ini && a.fin === fin && !a.comentario);
      if (igual) { igual.comentario = texto; return; }
      if (fin > ini) anotaciones.push({ tipo: 'comentario', p: pi, ini, fin, comentario: texto });
      else anotaciones.push({ tipo: 'insertar', p: pi, ini, nuevo: '', comentario: texto });
    }
    const resumen = { cambios: anotaciones.filter(a => a.tipo !== 'comentario' && !(a.tipo === 'insertar' && !a.nuevo)).length, comentarios: anotaciones.filter(a => a.comentario).length, controlDeCambios: hayCambios };
    return { anotaciones, resumen };
  }
  const tokens = t => { const r = []; const re = /\s+|[^\s]+/g; let m; while ((m = re.exec(t))) r.push({ t: m[0], ini: m.index, fin: m.index + m[0].length }); return r; };
  function lcs(a, b) {
    const n = a.length, m = b.length; if (n * m > 4e6) return [...a.map(() => '-'), ...b.map(() => '+')];
    const d = Array.from({ length: n + 1 }, () => new Uint16Array(m + 1));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) d[i][j] = a[i] === b[j] ? d[i + 1][j + 1] + 1 : Math.max(d[i + 1][j], d[i][j + 1]);
    const ops = []; let i = 0, j = 0;
    while (i < n && j < m) { if (a[i] === b[j]) { ops.push('='); i++; j++; } else if (d[i + 1][j] >= d[i][j + 1]) { ops.push('-'); i++; } else { ops.push('+'); j++; } }
    while (i < n) { ops.push('-'); i++; } while (j < m) { ops.push('+'); j++; }
    return ops;
  }
  // Empareja párrafos por parecido (para que un párrafo agregado o quitado no descuadre el resto)
  function emparejar(a, b) {
    const sim = (x, y) => { if (x === y) return 1; if (!x || !y) return x === y ? 1 : 0; const wx = new Set(x.split(' ')), wy = y.split(' '); let c = 0; wy.forEach(w => { if (wx.has(w)) c++; }); return (2 * c) / (wx.size + wy.length); };
    const n = a.length, m = b.length; const S = Array.from({ length: n + 1 }, () => new Float32Array(m + 1));
    for (let i = n - 1; i >= 0; i--) for (let j = m - 1; j >= 0; j--) { const s = sim(a[i], b[j]); S[i][j] = Math.max(S[i + 1][j], S[i][j + 1], s >= 0.5 ? S[i + 1][j + 1] + s : -1); }
    const r = []; let i = 0, j = 0;
    while (i < n && j < m) { const s = sim(a[i], b[j]); if (s >= 0.5 && Math.abs(S[i][j] - (S[i + 1][j + 1] + s)) < 1e-6) { r.push([i, j]); i++; j++; } else if (S[i + 1][j] >= S[i][j + 1]) { r.push([i, null]); i++; } else { r.push([null, j]); j++; } }
    while (i < n) r.push([i++, null]); while (j < m) r.push([null, j++]);
    return r;
  }

  const desdeTexto = (texto, titulo = '') => crearCorregido({ parrafos: String(texto).split(/\r?\n/).map(l => ({ runs: l ? [{ t: l, b: false, i: false }] : [], texto: l })), titulo });

  const api = { _bloques: bloques, _textoDe: textoDe, leerZip, escribirZip, leerParrafos, leerCorrecciones, crearCorregido, llenarRubrica, aBase64, deBase64, desdeTexto };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else raiz.Docx = api;
})(typeof window !== 'undefined' ? window : globalThis);
