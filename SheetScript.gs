// ============================================================
// SMG15 — BigQuery → Google Sheets → Dashboard Online
// Cole isso em: Extensões > Apps Script da sua planilha
// ============================================================

const SHEET_CONFIG = {
  PROJECT_ID: 'meli-bi-data',
  FACILITY:   'SMG15',
  SITE:       'MLB',
  SHEET_DOCAS:    'DOCAS_LIVE',   // aba com dados ao vivo
  SHEET_HISTORICO:'HISTORICO',    // aba com histórico por turno
};

// ── Roda automaticamente (configure trigger a cada 1 min) ───
function atualizarDados() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoje = Utilities.formatDate(new Date(), 'America/Recife', 'yyyy-MM-dd');

  try {
    const rows = queryYMS(hoje, 1);
    gravarDocas(ss, rows);
    gravarHistorico(ss, rows);
    gravarMetadata(ss);
    Logger.log('✅ Atualizado: ' + rows.length + ' registros');
  } catch(e) {
    Logger.log('❌ Erro: ' + e.message);
  }
}

// ── Query YMS ───────────────────────────────────────────────
function queryYMS(date, diasAtras) {
  const sql = `
  SELECT
    COALESCE(dck.DOCK.NAME, 'SEM_DOCA')                                            AS DOCK_NAME,
    COALESCE(yms.REQUEST.VEHICLE.LICENSE_PLATE, yms.OCCASIONAL.PLATE)              AS PLACA,
    yms.REQUEST.CARRIER.NAME                                                        AS TRANSPORTADORA,
    yms_purpose.OPERATION                                                           AS OPERACAO,
    yms_purpose.MILE                                                                AS MILE,
    yms_purpose.ROUTE_ID                                                            AS ROTA,
    yms_purpose.QUANTITY_SHIPMENT                                                   AS QTD_SHIPMENTS,
    COALESCE(yms.REQUEST.VEHICLE.DESCRIPTION, yms.OCCASIONAL.VEHICLE_TYPE)         AS TIPO_VEICULO,
    DATETIME(TIMESTAMP(DATETIME_ADD(yms.GATE_CHECKIN_DATE,   INTERVAL 4 HOUR)), 'America/Buenos_Aires') AS GATE_CHECKIN,
    DATETIME(TIMESTAMP(DATETIME_ADD(yms.GATE.CHECKOUT.DATE,  INTERVAL 4 HOUR)), 'America/Buenos_Aires') AS GATE_CHECKOUT,
    DATETIME(TIMESTAMP(DATETIME_ADD(dck.CHECKIN_DATE,        INTERVAL 4 HOUR)), 'America/Buenos_Aires') AS DOCA_CHECKIN,
    DATETIME(TIMESTAMP(DATETIME_ADD(dck.CHECKOUT_DATE,       INTERVAL 4 HOUR)), 'America/Buenos_Aires') AS DOCA_CHECKOUT,
    yms.GATE.FACILITY_STAY_TIME_MINUTES                                             AS PERMANENCIA_MIN,
    CAST(DATETIME(TIMESTAMP(DATETIME_ADD(yms.GATE.REFERENCE.DATE, INTERVAL 4 HOUR)), 'America/Buenos_Aires') AS DATE) AS DATA_REF
  FROM \`meli-bi-data.WHOWNER.BT_SHP_MT_YMS_GATE\` yms
  LEFT JOIN UNNEST(yms.PURPOSE) AS yms_purpose
  LEFT JOIN \`meli-bi-data.WHOWNER.BT_SHP_MT_YMS_DOCK\` dck
    ON yms_purpose.ID = dck.PURPOSE.ID AND dck.SITE_ID = 'MLB'
  WHERE
    CAST(DATETIME(TIMESTAMP(DATETIME_ADD(yms.GATE.REFERENCE.DATE, INTERVAL 4 HOUR)), 'America/Buenos_Aires') AS DATE)
    >= DATE_SUB(DATE('${date}'), INTERVAL ${diasAtras} DAY)
    AND yms.SITE_ID = 'MLB'
    AND yms.REQUEST_FACILITY_ID = 'SMG15'
  ORDER BY GATE_CHECKIN DESC
  LIMIT 2000`;

  const resp = BigQuery.Jobs.query(
    { query: sql, useLegacySql: false, timeoutMs: 30000, location: 'US' },
    SHEET_CONFIG.PROJECT_ID
  );

  if (!resp.rows) return [];
  const fields = resp.schema.fields.map(f => f.name);
  return resp.rows.map(r => {
    const obj = {};
    r.f.forEach((c,i) => { obj[fields[i]] = c.v; });
    return obj;
  });
}

// ── Grava aba DOCAS_LIVE ─────────────────────────────────────
function gravarDocas(ss, rows) {
  let sheet = ss.getSheetByName(SHEET_CONFIG.SHEET_DOCAS);
  if (!sheet) sheet = ss.insertSheet(SHEET_CONFIG.SHEET_DOCAS);
  sheet.clearContents();

  const headers = ['DOCA_NUM','DOCK_NAME','PLACA','TRANSPORTADORA','OPERACAO',
    'MILE','ROTA','QTD_SHIPMENTS','TIPO_VEICULO',
    'GATE_CHECKIN','GATE_CHECKOUT','DOCA_CHECKIN','DOCA_CHECKOUT',
    'PERMANENCIA_MIN','STATUS','INICIO_STATUS','DATA_REF'];

  const hoje = Utilities.formatDate(new Date(), 'America/Recife', 'yyyy-MM-dd');

  // filtra só registros de hoje
  const hoje_rows = rows.filter(r => (r.DATA_REF||'').slice(0,10) === hoje);

  // deduplica por doca (pega o mais recente)
  const docaMap = {};
  hoje_rows.forEach(r => {
    const match = (r.DOCK_NAME||'').match(/\d+/);
    const num = match ? parseInt(match[0]) : 0;
    if (num < 1 || num > 75) return;
    if (!docaMap[num] || r.DOCA_CHECKIN > (docaMap[num].DOCA_CHECKIN||'')) {
      docaMap[num] = {...r, DOCA_NUM: num};
    }
  });

  // determina status
  const data = [headers];
  for (let n = 1; n <= 75; n++) {
    const r = docaMap[n];
    if (!r) {
      data.push([n,'','','','','','',0,'','','','','','','livre','','']);
      continue;
    }
    let status = 'livre';
    let inicioStatus = '';
    const op = (r.OPERACAO||'').toUpperCase();
    if (r.DOCA_CHECKIN && !r.DOCA_CHECKOUT) {
      inicioStatus = r.DOCA_CHECKIN;
      if (op.includes('ADUANA') || op.includes('CUSTOM'))       status = 'aduana';
      else if (op.includes('CONTAGEM') || op.includes('COUNT')) status = 'contagem';
      else if (op.includes('GAIOLA') || op.includes('CAGE'))    status = 'gaiola';
      else if (op.includes('FINAL'))                             status = 'finalizado';
      else                                                        status = 'aduana';
    } else if (r.GATE_CHECKOUT || r.DOCA_CHECKOUT) {
      status = 'expedida';
    }
    data.push([n, r.DOCK_NAME||'', r.PLACA||'', r.TRANSPORTADORA||'',
      r.OPERACAO||'', r.MILE||'', r.ROTA||'', r.QTD_SHIPMENTS||0,
      r.TIPO_VEICULO||'', r.GATE_CHECKIN||'', r.GATE_CHECKOUT||'',
      r.DOCA_CHECKIN||'', r.DOCA_CHECKOUT||'', r.PERMANENCIA_MIN||'',
      status, inicioStatus, r.DATA_REF||'']);
  }

  sheet.getRange(1, 1, data.length, headers.length).setValues(data);
  Logger.log('DOCAS_LIVE: ' + (data.length-1) + ' linhas gravadas');
}

// ── Grava aba HISTORICO ──────────────────────────────────────
function gravarHistorico(ss, rows) {
  let sheet = ss.getSheetByName(SHEET_CONFIG.SHEET_HISTORICO);
  if (!sheet) sheet = ss.insertSheet(SHEET_CONFIG.SHEET_HISTORICO);
  sheet.clearContents();

  const headers = ['DATA','TURNO','TOTAL','COM_CHECKOUT',
    'STAY_AVG_MIN','DENTRO_30','ACIMA_30','PCT_ONTIME'];

  const turnos = {};
  rows.forEach(r => {
    const ref = (r.DATA_REF||'').slice(0,10);
    if (!ref) return;
    const checkinStr = r.GATE_CHECKIN || r.DOCA_CHECKIN || '';
    const h = checkinStr ? new Date(checkinStr).getHours() : 0;
    const turno = h < 6 ? 'MN' : h < 12 ? 'AM1' : h < 18 ? 'PM1' : 'PM2';
    const key = `${ref}_${turno}`;
    if (!turnos[key]) turnos[key] = {data:ref, turno, total:0, comCheckout:0, staySum:0, stayN:0, dentro30:0, acima30:0};
    const t = turnos[key];
    t.total++;
    if (r.GATE_CHECKOUT || r.DOCA_CHECKOUT) t.comCheckout++;
    const stay = parseFloat(r.PERMANENCIA_MIN||0);
    if (stay > 0) { t.staySum += stay; t.stayN++; if(stay<=30) t.dentro30++; else t.acima30++; }
  });

  const data = [headers];
  Object.values(turnos)
    .sort((a,b) => b.data.localeCompare(a.data) || a.turno.localeCompare(b.turno))
    .forEach(t => {
      const avg = t.stayN > 0 ? Math.round(t.staySum/t.stayN) : 0;
      const pct = t.stayN > 0 ? Math.round((t.dentro30/t.stayN)*100) : 0;
      data.push([t.data, t.turno, t.total, t.comCheckout, avg, t.dentro30, t.acima30, pct]);
    });

  sheet.getRange(1, 1, data.length, headers.length).setValues(data);
  Logger.log('HISTORICO: ' + (data.length-1) + ' linhas gravadas');
}

// ── Grava metadados (timestamp de atualização) ───────────────
function gravarMetadata(ss) {
  let sheet = ss.getSheetByName('META');
  if (!sheet) sheet = ss.insertSheet('META');
  sheet.clearContents();
  const now = new Date();
  sheet.getRange('A1').setValue('UPDATED_AT');
  sheet.getRange('B1').setValue(now.toISOString());
  sheet.getRange('A2').setValue('FACILITY');
  sheet.getRange('B2').setValue('SMG15');
}

// ── Configura trigger automático de 1 minuto ─────────────────
function configurarTrigger() {
  // Apaga triggers existentes
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  // Cria novo trigger a cada 1 minuto
  ScriptApp.newTrigger('atualizarDados')
    .timeBased()
    .everyMinutes(1)
    .create();
  Logger.log('✅ Trigger configurado: atualizarDados a cada 1 minuto');
}
