const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const net = require('net');
const { Client } = require('pg');
const { Client: SshClient } = require('ssh2');
const XLSX = require('xlsx');
const csvParse = require('csv-parse/sync');
require('dotenv').config();

const UPLOAD_DIR = path.join(__dirname, 'LIQUIDACION TERCOP');
const PORT = Number(process.env.APP_PORT || 3838);

if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

// Conexion a la base: por tunel SSH si SSH_HOST esta definido en .env, o directa si no
async function getDbClient() {
  const useTunnel = !!process.env.SSH_HOST;

  let c;
  if (useTunnel) {
    const ssh = new SshClient();
    const localPort = 15432 + 200 + (Math.floor(Math.random() * 800));

    await new Promise((resolve, reject) => {
      ssh.on('ready', resolve);
      ssh.on('error', reject);
      ssh.connect({
        host: process.env.SSH_HOST,
        port: Number(process.env.SSH_PORT),
        username: process.env.SSH_USERNAME,
        password: process.env.SSH_PASSWORD,
        keepaliveInterval: 15000,
        keepaliveCountMax: 3
      });
    });

    const server = net.createServer((sock) => {
      sock.on('error', () => sock.destroy());
      try {
        ssh.forwardOut(sock.remoteAddress, sock.remotePort || 0, process.env.DB_HOST, Number(process.env.DB_PORT), (err, stream) => {
          if (err) { sock.destroy(); return; }
          stream.on('error', () => stream.destroy());
          sock.pipe(stream);
          stream.pipe(sock);
        });
      } catch (e) {
        sock.destroy();
      }
    });

    await new Promise((resolve, reject) => {
      server.once('error', reject);
      server.listen(localPort, '127.0.0.1', () => { server.removeListener('error', reject); resolve(); });
    });

    c = new Client({
      host: '127.0.0.1',
      port: localPort,
      database: process.env.DB_NAME,
      user: process.env.DB_USER,
      password: process.env.DB_PASSWORD,
      connectionTimeoutMillis: 15000
    });
    c.closeTunnel = async () => {
      try { await c.end(); } catch (_) {}
      try { server.close(); } catch (_) {}
      try { ssh.end(); } catch (_) {}
    };
  } else {
    c = new Client({
      host: process.env.DB_HOST,
      port: Number(process.env.DB_PORT || 5432),
      database: process.env.DB_NAME,
      user: process.env.DB_USER,
      password: process.env.DB_PASSWORD,
      connectionTimeoutMillis: 15000
    });
    c.closeTunnel = async () => {
      try { await c.end(); } catch (_) {}
    };
  }

  c.on('error', () => {}); // evitar crash del proceso si la conexion se cae
  try {
    await c.connect();
  } catch (e) {
    console.error('[db] connect err :', e.message);
    try { await c.closeTunnel(); } catch (_) {}
    throw e;
  }
  return c;
}

const app = express();
app.use(express.json());

// Cache de resultados de comparacion (clave: pl|file, TTL 30 min)
const resultCache = new Map();
function cacheKey(pl, file) { return String(pl) + '|' + String(file); }
function cacheSet(pl, file, result) {
  resultCache.set(cacheKey(pl, file), { result, ts: Date.now() });
  if (resultCache.size > 50) {
    const now = Date.now();
    for (const [k, v] of resultCache) if (now - v.ts > 30 * 60 * 1000) resultCache.delete(k);
  }
}
function cacheGet(pl, file) {
  const e = resultCache.get(cacheKey(pl, file));
  if (!e) return null;
  if (Date.now() - e.ts > 30 * 60 * 1000) { resultCache.delete(cacheKey(pl, file)); return null; }
  return e.result;
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, UPLOAD_DIR),
  filename: (req, file, cb) => {
    const safe = file.originalname.replace(/[^\w.\- ]+/g, '_');
    cb(null, safe);
  }
});
const upload = multer({ storage });

function serveIndex(res) {
  fs.readFile(path.join(__dirname, 'index.html'), (err, data) => {
    if (err) { res.status(500).end('Error cargando index.html'); return; }
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(data);
  });
}

// Lista de procesos de liquidacion (PL)
app.get('/api/pls', async (req, res) => {
  try {
    const c = await getDbClient();
    const r = await c.query(`
      SELECT lqprclqid::text AS id, lqprcdsc AS descripcion, lqprcperliq AS periodo, lqprcestado AS estado
      FROM lqprc
      ORDER BY lqprclqid DESC
      LIMIT 500
    `);
    await c.closeTunnel();
    res.json(r.rows);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Archivos disponibles en la carpeta LIQUIDACION TERCOP
app.get('/api/files', (req, res) => {
  try {
    const files = fs.readdirSync(UPLOAD_DIR)
      .filter(f => /\.(csv|xlsx|xls)$/i.test(f))
      .map(f => ({ name: f, size: fs.statSync(path.join(UPLOAD_DIR, f)).size, modified: fs.statSync(path.join(UPLOAD_DIR, f)).mtime }));
    res.json(files);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Subir archivo TERCOP a la carpeta
app.post('/api/upload', (req, res) => {
  upload.single('file')(req, res, (err) => {
    if (err) return res.status(400).json({ error: err.message });
    if (!req.file) return res.status(400).json({ error: 'No se recibio archivo' });
    res.json({ ok: true, name: req.file.filename });
  });
});

const COL_MAP = [
  { field: 'documento', cols: ['documento'] },
  { field: 'legajo', cols: ['legajo', 'nrolegajo'] },
  { field: 'apellido', cols: ['apellido', 'apellidoynombre', 'nombre'] },
  { field: 'empresa', cols: ['empresa', 'empresacod'] },
  { field: 'estado', cols: ['estado'] },
  { field: 'dependencia', cols: ['dependencia', 'descripcion dependencia'] },
  { field: 'cargo', cols: ['cargo', 'descripcion', 'descripcioncargo'] },
  { field: 'concepto', cols: ['concepto1', 'cpt', 'concepto', 'codconcepto'] },
  { field: 'descripcion_concepto', cols: ['descripcion_concepto', 'descripcion'] },
  { field: 'unidades', cols: ['unidades', 'unidad', 'cantidad'] },
  { field: 'importe', cols: ['importe', 'remun', 'valor', 'importetotal'] },
  { field: 'cantidad_horas', cols: ['cantidad horas', 'horas', 'cantidadh'] }
];

function normHeader(h) {
  return String(h).toLowerCase().replace(/\s+/g, ' ').replace(/[^\w ]/g, '').trim();
}

function mapColumns(headers) {
  const idx = {};
  for (const m of COL_MAP) {
    for (const col of m.cols) {
      const key = normHeader(col);
      const found = headers.findIndex(h => normHeader(h) === key);
      if (found >= 0) { idx[m.field] = found; break; }
    }
  }
  return idx;
}

function num(v) {
  if (v === null || v === undefined || v === '') return 0;
  const s = String(v).replace(/,/g, '').trim();
  const n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function str(v) {
  if (v === null || v === undefined) return '';
  return String(v).trim();
}

// Tipo de concepto (conliqtip) -> descripcion
const TIPO_CONCEPTO = {
  1: 'Remunerativo',
  2: 'No Remunerativo',
  3: 'Asignaciones Familiares',
  4: 'Retenciones',
  5: 'Contribuciones',
  6: 'Totalizadores',
  7: 'Base de calculo'
};
function tipoConcepto(n) { return TIPO_CONCEPTO[Number(n)] || ''; }

// Repara lineas CSV que vienen con el registro completo envuelto en comillas
// (el archivo TERCOP duplica las comillas internas: "" -> ")
function repairWrappedCSVLines(text) {
  const lines = text.split('\n');
  let fixed = 0;
  const out = lines.map(l => {
    const lt = l.replace(/\r$/, '');
    if (lt.length > 2 && lt.startsWith('"') && lt.endsWith('"')) {
      fixed++;
      return lt.split('""').join('"').slice(1, -1) + '\r';
    }
    return l;
  });
  return { text: out.join('\n'), fixed };
}

// Parsear archivo TERCOP (CSV o XLSX) -> una fila por linea del archivo
function parseTercop(filePath) {
  let rows;
  if (/\.csv$/i.test(filePath)) {
    const buf = fs.readFileSync(filePath);
    const repaired = repairWrappedCSVLines(buf.toString('latin1'));
    rows = csvParse.parse(repaired.text, { columns: true, skip_empty_lines: true, bom: true });
  } else {
    const wb = XLSX.readFile(filePath, { cellDates: false });
    const ws = wb.Sheets[wb.SheetNames[0]];
    rows = XLSX.utils.sheet_to_json(ws, { defval: '', raw: true });
  }

  if (!rows.length) throw new Error('El archivo no tiene filas de datos');

  const idx = mapColumns(Object.keys(rows[0]));
  if (idx.legajo === undefined) {
    throw new Error('No se encontro la columna legajo en el archivo. Columnas detectadas: ' + Object.keys(rows[0]).join(', '));
  }

  const keys = Object.keys(rows[0]);
  const out = [];
  let conConcepto = false;
  for (const r of rows) {
    const legajo = str(r[keys[idx.legajo]]).replace(/\D/g, '');
    if (!legajo) continue;
    const conceptoExt = idx.concepto !== undefined ? str(r[keys[idx.concepto]]) : '';
    if (conceptoExt) conConcepto = true;
    out.push({
      legajo,
      documento: idx.documento !== undefined ? str(r[keys[idx.documento]]).replace(/\D/g, '') : '',
      apellido: idx.apellido !== undefined ? str(r[keys[idx.apellido]]) : '',
      empresa: idx.empresa !== undefined ? str(r[keys[idx.empresa]]) : '',
      estado: idx.estado !== undefined ? str(r[keys[idx.estado]]) : '',
      dependencia: idx.dependencia !== undefined ? str(r[keys[idx.dependencia]]) : '',
      cargo: idx.cargo !== undefined ? str(r[keys[idx.cargo]]) : '',
      descripcion_concepto: idx.descripcion_concepto !== undefined ? str(r[keys[idx.descripcion_concepto]]) : '',
      unidades: idx.unidades !== undefined ? num(r[keys[idx.unidades]]) : 0,
      importe: idx.importe !== undefined ? num(r[keys[idx.importe]]) : 0,
      cantidad_horas: idx.cantidad_horas !== undefined ? num(r[keys[idx.cantidad_horas]]) : 0,
      concepto_ext: conceptoExt
    });
  }
  return { rows: out, cols: keys, sinConcepto: !conConcepto };
}

// Fecha "abierta" en conliqext (0001-01-01 / 1900-01-01 no acotan)
function openTs(d) {
  if (!d) return null;
  const y = d.getFullYear();
  return (y >= 1901 && y <= 9999) ? d.getTime() : null;
}

// Mapeo de conceptos: conliqext (externo + regimen -> interno) y regsal (nombre regimen -> codigo)
async function getConceptMapping(c, periodo) {
  const regs = await c.query('SELECT regsalcod, regsalnom FROM regsal');
  const regByNom = new Map();
  for (const r of regs.rows) regByNom.set(normHeader(r.regsalnom), String(r.regsalcod));

  const m = await c.query(`SELECT conliqextcodexterno::text AS ext, conliqextcod::text AS con,
                                  conliqextregsalcod::text AS reg, conliqextfchdesde AS desde, conliqextfchhasta AS hasta
                           FROM conliqext`);
  const byKey = new Map(); // ext|reg -> [{con, desde, hasta}]
  for (const r of m.rows) {
    const key = r.ext + '|' + r.reg;
    if (!byKey.has(key)) byKey.set(key, []);
    byKey.get(key).push({ con: r.con, desde: openTs(r.desde), hasta: openTs(r.hasta) });
  }

  let periodTs = null;
  const per = String(periodo || '').trim();
  if (/^\d{6}$/.test(per)) periodTs = new Date(per.slice(0, 4) + '-' + per.slice(4) + '-01').getTime();

  // resolver externo+regimen -> concepto interno segun periodo
  function resolve(ext, regsalcod) {
    if (ext === undefined || ext === null || ext === '') return null;
    const list = byKey.get(String(ext) + '|' + String(regsalcod));
    if (!list || !list.length) return null;
    let best = null;
    for (const e of list) {
      if (periodTs !== null && e.desde !== null && periodTs < e.desde) continue;
      if (periodTs !== null && e.hasta !== null && periodTs > e.hasta) continue;
      if (!best || (e.desde !== null && (best.desde === null || e.desde > best.desde))) best = e;
    }
    return best ? best.con : null;
  }

  return { regByNom, resolve };
}

// Comparar contra la tabla correcta
app.post('/api/compare', async (req, res) => {
  const { pl, file } = req.body || {};
  if (!pl) return res.status(400).json({ error: 'Indique el numero de PL' });
  if (!file) return res.status(400).json({ error: 'Seleccione el archivo TERCOP' });

  const filePath = path.join(UPLOAD_DIR, path.basename(file));
  if (!fs.existsSync(filePath)) return res.status(400).json({ error: 'El archivo no existe en la carpeta: ' + file });

  // Fase BD con reintento si el tunel se cae a mitad
  async function dbPhase() {
    const c = await getDbClient();
    try {
      const plInfo = await c.query('SELECT lqprclqid::text AS id, lqprcdsc AS descripcion, lqprcperliq AS periodo FROM lqprc WHERE lqprclqid = $1', [pl]);

      let table = 'lqlegcon';
      let query = '';
      if (plInfo.rows.length) {
        const inLeg = await c.query('SELECT count(*) AS n FROM lqlegcon WHERE lqprclqid = $1', [pl]);
        const inHis = await c.query('SELECT count(*) AS n FROM lqhislegcon WHERE lqhisprclqid = $1', [pl]);
        if (parseInt(inLeg.rows[0].n) === 0 && parseInt(inHis.rows[0].n) > 0) table = 'lqhislegcon';
      }

      if (table === 'lqlegcon') {
        query = `SELECT c.lqlegnroleg::text AS legajo, c.lqconliqcod::text AS concepto,
                        c.lqlegconimp::float8 AS importe, c.lqlegconrutid AS rutina, c.lqlegconrutconsid AS consulta,
                        c.lqlegcondias AS dias, c.lqlegconcnt::float8 AS cnt,
                        c.lqlegconofijudid::text AS oficio,
                        l.lqlegapenom AS apenom, l.lqlegnrodoc AS documento, l.lqlegpuereg AS regcod
                 FROM lqlegcon c
                 LEFT JOIN lqlegajo l ON l.lqprclqid = c.lqprclqid AND l.lqlegnroleg = c.lqlegnroleg
                 WHERE c.lqprclqid = $1`;
      } else {
        query = `SELECT c.lqhislegnroleg::text AS legajo, c.lqhisconliqcod::text AS concepto,
                        c.lqhislegconimp::float8 AS importe, c.lqhislegconrutid AS rutina, c.lqhislegconrutconsid AS consulta,
                        c.lqhislegcondias AS dias, c.lqhislegconcnt::float8 AS cnt,
                        c.lqhislegconofijudid::text AS oficio,
                        l.lqhislegapenom AS apenom, l.lqhislegnrodoc AS documento, l.lqhislegpuereg AS regcod
                 FROM lqhislegcon c
                 LEFT JOIN lqhislegajo l ON l.lqhisprclqid = c.lqhisprclqid AND l.lqhislegnroleg = c.lqhislegnroleg
                 WHERE c.lqhisprclqid = $1`;
      }

      const dbRes = await c.query(query, [pl]);
      // Los conceptos 9000+ son internos de calculo (totales) que el archivo TERCOP nunca trae
      const internosBd = dbRes.rows.filter(r => Number(r.concepto) >= 9000).length;
      const usedRows = dbRes.rows.filter(r => Number(r.concepto) < 9000);
      const mapping = await getConceptMapping(c, plInfo.rows.length ? plInfo.rows[0].periodo : null);

      // Info de conceptos (nombre, modalidad de carga, tipo) desde conliq
      const conliqRes = await c.query('SELECT conliqcod::text AS cod, conliqnom AS nom, conliqmodcar AS modcar, conliqtip::int AS tip FROM conliq');
      const conliqMap = new Map();
      for (const r of conliqRes.rows) conliqMap.set(r.cod, r);

      // Mapa de regimenes: regsalcod -> regsalnom
      const regsalRes = await c.query('SELECT regsalcod::text AS cod, regsalnom AS nom FROM regsal');
      const regsalMap = new Map();
      for (const r of regsalRes.rows) regsalMap.set(r.cod, str(r.nom));

      return { plInfo, table, usedRows, internosBd, mapping, conliqMap, regsalMap };
    } finally {
      try { await c.closeTunnel(); } catch (_) {}
    }
  }

  try {
    const tercop = parseTercop(filePath);

    let db = null;
    try {
      db = await dbPhase();
    } catch (e) {
      // el tunel pudo caerse: reintentar con un tunel nuevo
      db = await dbPhase();
    }
    const { plInfo, table, usedRows, internosBd, mapping, conliqMap, regsalMap } = db;

    // Agregar filas de TERCOP: pasar el concepto externo a interno (conliqext) usando el regimen (empresa -> regsal)
    const fileByKey = new Map();
    const sinMapeo = [];
    for (const t of tercop.rows) {
      let concepto = '';
      if (tercop.sinConcepto) {
        // el archivo no trae conceptos: se compara solo por legajo
      } else {
        const regsalcod = mapping.regByNom.get(normHeader(t.empresa));
        concepto = regsalcod !== undefined ? mapping.resolve(t.concepto_ext, regsalcod) : null;
        if (concepto === null) {
          sinMapeo.push({ ...t, regsalcod: regsalcod !== undefined ? regsalcod : '' });
          continue;
        }
      }
      const key = t.legajo + '|' + concepto;
      let g = fileByKey.get(key);
      if (!g) {
        g = { legajo: t.legajo, concepto, concepto_ext: t.concepto_ext, documento: '', apellido: '', empresa: '', estado: '', dependencia: '', cargo: '', descripcion_concepto: '', unidades: 0, importe: 0, cantidad_horas: 0, n: 0 };
        fileByKey.set(key, g);
      }
      if (!g.concepto_ext) g.concepto_ext = t.concepto_ext;
      if (!g.documento) g.documento = t.documento;
      if (!g.apellido) g.apellido = t.apellido;
      if (!g.empresa) g.empresa = t.empresa;
      if (!g.estado) g.estado = t.estado;
      if (!g.dependencia) g.dependencia = t.dependencia;
      if (!g.cargo) g.cargo = t.cargo;
      if (!g.descripcion_concepto) g.descripcion_concepto = t.descripcion_concepto;
      g.unidades += t.unidades;
      g.importe += t.importe;
      g.cantidad_horas += t.cantidad_horas;
      g.n++;
    }
    const fileRows = [...fileByKey.values()];

    // Agregar filas de BD por legajo+concepto
    const dbMap = new Map();
    for (const r of usedRows) {
      const legajo = str(r.legajo).replace(/\D/g, '');
      const concepto = tercop.sinConcepto ? '' : str(r.concepto);
      const key = legajo + '|' + concepto;
      if (!dbMap.has(key)) dbMap.set(key, { legajo, concepto, importe: 0, oficio: '', apenom: '', documento: '', regimen: '', items: [] });
      const g = dbMap.get(key);
      g.importe += num(r.importe);
      if (!g.oficio) g.oficio = r.oficio !== undefined ? str(r.oficio) : '';
      if (!g.apenom) g.apenom = str(r.apenom);
      if (!g.documento) g.documento = str(r.documento);
      if (!g.regimen) g.regimen = regsalMap.get(str(r.regcod)) || '';
      g.items.push({ importe: r.importe, rutina: r.rutina, consulta: r.consulta, dias: r.dias, cnt: r.cnt, oficio: r.oficio !== undefined ? str(r.oficio) : '' });
    }

    // Clave presente en BD y no en TERCOP
    const tercopKeys = new Set(fileRows.map(t => t.legajo + '|' + t.concepto));
    const soloBD = [];
    for (const [key, g] of dbMap) {
      if (!tercopKeys.has(key)) {
        soloBD.push({ legajo: g.legajo, concepto: g.concepto, importe: g.importe, oficio: g.oficio, items: g.items });
      }
    }

    // Comparar cada fila de TERCOP contra BD
    const diferencias = [];
    const okCount = [];
    const soloTercop = [];
    for (const t of fileRows) {
      const db = dbMap.get(t.legajo + '|' + t.concepto);
      if (!db) {
        soloTercop.push({ ...t, db_importe: 0 });
        continue;
      }
      const diff = round2(t.importe) - round2(db.importe);
      const ci = conliqMap.get(t.concepto);
      const row = {
        ...t,
        db_importe: round2(db.importe),
        diferencia: diff,
        oficio: db.oficio,
        genrrhh_apenom: db.apenom,
        genrrhh_documento: db.documento,
        genrrhh_regimen: db.regimen,
        concepto_nombre: ci ? str(ci.nom) : '',
        concepto_modcar: ci ? str(ci.modcar) : '',
        concepto_tip: ci ? tipoConcepto(ci.tip) : '',
        rutina: db.items.map(i => i.rutina).join('+'),
        consulta: db.items.map(i => i.consulta).join('+'),
        items: db.items
      };
      if (Math.abs(diff) > 0.009) diferencias.push(row);
      else okCount.push(row);
    }

    // Resumen comparativo por concepto
    const porCon = new Map();
    function aggCon(concepto, ci, campo, impTercop, impBD) {
      if (!concepto) return;
      let c = porCon.get(concepto);
      if (!c) {
        c = {
          concepto,
          nombre: ci ? str(ci.nom) : '',
          modcar: ci ? str(ci.modcar) : '',
          tip: ci ? tipoConcepto(ci.tip) : '',
          ok: 0, con_diferencia: 0, solo_tercop: 0, solo_bd: 0,
          imp_tercop: 0, imp_bd: 0
        };
        porCon.set(concepto, c);
      }
      c[campo]++;
      if (impTercop !== undefined) c.imp_tercop += num(impTercop);
      if (impBD !== undefined) c.imp_bd += num(impBD);
    }
    for (const r of okCount) aggCon(r.concepto, conliqMap.get(r.concepto), 'ok', r.importe, r.db_importe);
    for (const r of diferencias) aggCon(r.concepto, conliqMap.get(r.concepto), 'con_diferencia', r.importe, r.db_importe);
    for (const r of soloTercop) aggCon(r.concepto, conliqMap.get(r.concepto), 'solo_tercop', r.importe, 0);
    for (const r of soloBD) aggCon(r.concepto, conliqMap.get(r.concepto), 'solo_bd', 0, r.importe);
    const porConcepto = [...porCon.values()].map(c => ({
      ...c,
      imp_tercop: round2(c.imp_tercop),
      imp_bd: round2(c.imp_bd),
      diferencia: round2(c.imp_tercop - c.imp_bd)
    })).sort((a, b) => a.concepto.localeCompare(b.concepto, undefined, { numeric: true }));

    const payload = {
      pl,
      pl_desc: plInfo.rows.length ? plInfo.rows[0].descripcion : '(PL no encontrado en lqprc)',
      pl_periodo: plInfo.rows.length ? plInfo.rows[0].periodo : '',
      tabla: table,
      archivo: file,
      columnas: tercop.cols,
      usa_mapeo_conliqext: !tercop.sinConcepto,
      resumen: {
        filas_tercop: fileRows.length,
        filas_bd: usedRows.length,
        internos_bd: internosBd,
        ok: okCount.length,
        con_diferencia: diferencias.length,
        solo_tercop: soloTercop.length,
        solo_bd: soloBD.length,
        sin_mapeo: sinMapeo.length
      },
      diferencias,
      soloTercop,
      soloBD,
      sinMapeo,
      porConcepto
    };
    cacheSet(pl, file, payload);
    res.json(payload);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Exportar resultado a xlsx
app.post('/api/export', async (req, res) => {
  const { pl, file } = req.body || {};
  const result = req.body && req.body.result ? req.body.result : (pl !== undefined ? cacheGet(pl, file) : null);
  if (!result) return res.status(400).json({ error: 'Sin datos. Ejecute primero la comparacion.' });

  const wb = XLSX.utils.book_new();

  const diffRows = result.diferencias.map(r => ({
    'Legajo TERCOP': r.legajo, 'Documento TERCOP': r.documento, 'Apellido TERCOP': r.apellido,
    'Regimen TERCOP (Empresa)': r.empresa, 'Concepto TERCOP': r.concepto_ext, 'Nombre concepto TERCOP': r.descripcion_concepto,
    'Unidades TERCOP': r.unidades, 'Cantidad de Horas TERCOP': r.cantidad_horas,
    'Dependencia TERCOP': r.dependencia, 'Cargo TERCOP': r.cargo, 'Importe TERCOP': r.importe,
    'Legajo GENRRHH': r.legajo, 'Nombre y Apellido GENRRHH': r.genrrhh_apenom, 'Regimen GENRRHH': r.genrrhh_regimen,
    'Documento GENRRHH': r.genrrhh_documento,
    'Concepto GENRRHH': r.concepto, 'Nombre concepto GENRRHH': r.concepto_nombre,
    'Modalidad de carga GENRRHH': r.concepto_modcar, 'Tipo de concepto GENRRHH': r.concepto_tip,
    'Oficio GENRRHH': r.oficio, 'Rutina GENRRHH': r.rutina, 'Consulta GENRRHH': r.consulta,
    'Diferencia': r.diferencia
  }));
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(diffRows), 'Diferencias');

  const wsE2 = XLSX.utils.json_to_sheet(
    result.soloTercop.map(r => ({
      legajo: r.legajo, concepto: r.concepto, descripcion_concepto: r.descripcion_concepto,
      documento: r.documento, apellido: r.apellido, empresa: r.empresa, estado: r.estado,
      dependencia: r.dependencia, cargo: r.cargo, unidades: r.unidades, cantidad_horas: r.cantidad_horas,
      'Importe TERCOP': r.importe, 'Importe BD': r.db_importe
    })));
  XLSX.utils.book_append_sheet(wb, wsE2, 'Solo TERCOP');

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
    result.soloBD.map(r => ({
      legajo: r.legajo, concepto: r.concepto, 'Importe BD': r.importe,
      oficio: r.oficio, rutina: r.items.map(i => i.rutina).join('+'), consulta: r.items.map(i => i.consulta).join('+')
    }))), 'Solo BD');

  if (result.sinMapeo && result.sinMapeo.length) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
      result.sinMapeo.map(r => ({
        legajo: r.legajo, concepto_externo: r.concepto_ext, regimen: r.empresa,
        regsalcod: r.regsalcod, apellido: r.apellido, documento: r.documento,
        dependencia: r.dependencia, cargo: r.cargo, 'Importe TERCOP': r.importe, unidades: r.unidades, cantidad_horas: r.cantidad_horas
      }))), 'Sin mapeo');
  }

  if (result.porConcepto && result.porConcepto.length) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(
      result.porConcepto.map(c => ({
        'Concepto GENRRHH': c.concepto, 'Nombre concepto': c.nombre,
        'Modalidad de carga': c.modcar, 'Tipo de concepto': c.tip,
        'Cuadran (OK)': c.ok, 'Con diferencia': c.con_diferencia,
        'Solo TERCOP': c.solo_tercop, 'Solo BD': c.solo_bd,
        'Importe TERCOP': c.imp_tercop, 'Importe BD': c.imp_bd, 'Diferencia': c.diferencia
      }))), 'Comparativa por concepto');
  }

  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([
    { PL: result.pl, Descripcion: result.pl_desc, Periodo: result.pl_periodo, Tabla: result.tabla, Archivo: result.archivo },
    { 'Filas TERCOP': result.resumen.filas_tercop, 'Filas BD': result.resumen.filas_bd,
      'Internos BD excluidos': result.resumen.internos_bd, OK: result.resumen.ok,
      'Con diferencia': result.resumen.con_diferencia, 'Solo TERCOP': result.resumen.solo_tercop, 'Solo BD': result.resumen.solo_bd,
      'Sin mapeo': result.resumen.sin_mapeo }
  ]), 'Resumen');

  const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

  // Guardar copia en la carpeta RESULTADO (si el archivo esta abierto/locked, no tirar abajo el server)
  const RESULT_DIR = path.join(__dirname, 'RESULTADO');
  if (!fs.existsSync(RESULT_DIR)) fs.mkdirSync(RESULT_DIR, { recursive: true });
  const baseArchivo = String(result.archivo || '').replace(/\.(xlsx|xls|csv)$/i, '');
  const fileName = `Comparativa - ${baseArchivo} vs PL ${result.pl}.xlsx`;
  const outFile = path.join(RESULT_DIR, fileName);
  try {
    fs.writeFileSync(outFile, buf);
  } catch (e) {
    console.error('[export] no se pudo guardar en RESULTADO (' + e.code + '): ' + outFile);
  }

  res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
  res.send(buf);
});

function round2(n) { return Math.round(n * 100) / 100; }

app.get('/', (req, res) => serveIndex(res));

app.get('/api/health', async (req, res) => {
  const useTunnel = !!process.env.SSH_HOST;
  const cfg = {
    modo: useTunnel ? 'tunel_ssh' : 'directo',
    sshHost: process.env.SSH_HOST, sshPort: process.env.SSH_PORT, sshUser: process.env.SSH_USERNAME,
    dbHost: process.env.DB_HOST, dbPort: process.env.DB_PORT, dbName: process.env.DB_NAME, dbUser: process.env.DB_USER
  };
  try {
    const c = await getDbClient();
    const r = await c.query('SELECT 1 AS ok');
    await c.closeTunnel();
    res.json({ tunnel: useTunnel, db: true, ok: r.rows[0].ok, cfg });
  } catch (e) {
    res.status(500).json({ tunnel: useTunnel, db: false, error: e.message, cfg });
  }
});

app.listen(PORT, () => {
  console.log(`CUADRATURA corriendo en http://localhost:${PORT}`);
});
