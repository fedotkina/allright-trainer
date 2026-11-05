/**
 * Тренажёр звонков AllRight — связка между ElevenLabs, Google Drive и страницей админов.
 *
 * Что делает:
 *  1. doPost — принимает post-call webhook от ElevenLabs: пишет звонок строкой в таблицу
 *     и сохраняет его файлом в папку «звонки (JSON)», откуда прототип забирает звонки на оценку.
 *  2. doGet  — без параметров отдаёт странице админов свежий файл сценариев;
 *     с ?action=results&name=... отдаёт оценки звонков этого админа.
 *  3. syncScores — раз в час переносит баллы из оценок в таблицу.
 *
 *  5. scorePending — каждые 5 минут оценивает новые звонки, если в Script Properties есть ANTHROPIC_API_KEY.
 *  4. buildDashboard — раз в час собирает данные для дашборда (звонки и баллы без расшифровок).
 *
 * Разворачивается один раз: Deploy → New deployment → Web app,
 * Execute as: Me, Who has access: Anyone. После правок кода: Manage deployments → New version.
 */

// ID таблицы «Тренажёр звонков — звонки админов»
const SHEET_ID = '1Mb09C2mI4ONGvxRJUBAU9OYfI65BClnQ_jidJ_DR7SY';
// ID папки «Тренажёр звонков — сценарии»
const FOLDER_ID = '1psY55TrBuVCc1jLb76et5wHVHaDhl0zN';
// ID папки «Тренажёр звонков — звонки (JSON)»
const CALLS_FOLDER_ID = '1CWk3dJ3hQmaDKpsHFZzmo_lH3R0lh2Cv';
// ID папки «Тренажёр звонков — оценки»
const RESULTS_FOLDER_ID = '1wWNhBDPtaaFjOqpOzGczs06Fuh3fEcdj';
// Файлы данных для дашборда (папка «Тренажёр звонков — дашборд (данные)»)
const DASH_CALLS_FILE_ID = '11KIVcCAHfp0kim5myX_Sjbp6dPjlDGwj';
const DASH_ITEMS_FILE_ID = '1v4m_6_kxJNFxi5do0dk48mGCtNvH1Qll';
// Имя файла сценариев, который выгружает прототип
const SCENARIOS_NAME = 'scenarios.dat';

const HEADER = ['received_at', 'conversation_id', 'admin_name', 'scenario_id', 'scenario_lang',
  'scenario_dept', 'started_at', 'duration_sec', 'status', 'transcript_json', 'evaluated_at', 'score'];
const CELL_LIMIT = 49000; // в ячейке Google Таблиц максимум 50 000 символов

/* ---------- 1. Webhook от ElevenLabs ---------- */
function doPost(e) {
  const secret = PropertiesService.getScriptProperties().getProperty('WEBHOOK_TOKEN');
  if (!secret || !e || !e.parameter || e.parameter.token !== secret) {
    return json_({ ok: false, error: 'unauthorized' });
  }
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }

  // Нам нужен только итог звонка с расшифровкой
  if (body.type && body.type !== 'post_call_transcription') return json_({ ok: true, skipped: body.type });

  const d = body.data || body;
  const init = d.conversation_initiation_client_data || {};
  const vars = init.dynamic_variables || {};
  const meta = d.metadata || {};
  const transcript = (d.transcript || []).map(function (t) {
    let text = t.message || '';
    // Убираем подсказки интонации для голоса вроде [neutral], [thoughtful]: в звонке их не слышно
    if (t.role === 'agent') text = text.replace(/\[[A-Za-z][A-Za-z _-]{1,24}\]\s*/g, '').trim();
    return { role: t.role === 'agent' ? 'mom' : 'admin', text: text, t: Math.round(t.time_in_call_secs || 0) };
  }).filter(function (t) { return t.text; });

  const call = {
    conversation_id: d.conversation_id || '',
    admin_name: vars.admin_name || init.user_id || d.user_id || '',
    admin_id: vars.admin_id || '',
    scenario_id: vars.scenario_id || '',
    scenario_lang: vars.scenario_lang || '',
    scenario_dept: vars.scenario_dept || '',
    started_at: meta.start_time_unix_secs ? meta.start_time_unix_secs * 1000 : Date.now(),
    duration_sec: meta.call_duration_secs || 0,
    status: d.status || '',
    transcript: transcript
  };

  let trJson = JSON.stringify(transcript);
  if (trJson.length > CELL_LIMIT) trJson = trJson.slice(0, CELL_LIMIT);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    // Защита от дублей: ElevenLabs может повторить доставку
    if (call.conversation_id && sh.getLastRow() > 1) {
      const ids = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; });
      if (ids.indexOf(call.conversation_id) !== -1) return json_({ ok: true, duplicate: true });
    }
    sh.appendRow([
      new Date(),
      call.conversation_id,
      call.admin_name,
      call.scenario_id,
      call.scenario_lang,
      call.scenario_dept,
      meta.start_time_unix_secs ? new Date(meta.start_time_unix_secs * 1000) : '',
      call.duration_sec || '',
      call.status,
      trJson,
      '',
      ''
    ]);
    // Файл звонка для прототипа. В base64, чтобы текст не искажался при чтении через Google Drive.
    if (call.conversation_id && transcript.length) {
      const b64 = Utilities.base64Encode(JSON.stringify(call), Utilities.Charset.UTF_8);
      DriveApp.getFolderById(CALLS_FOLDER_ID).createFile('call-' + call.conversation_id + '.txt', b64, MimeType.PLAIN_TEXT);
    }
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

/* ---------- 2. Страница админов ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'results') return json_(resultsFor_(p.name || '', p.id || ''));
  const files = scenarioFiles_();
  if (!files.length) return text_('');
  return text_(files[0].getBlob().getDataAsString());
}

function normName_(s) {
  return String(s || '').toLowerCase().replace(/[’ʼ`´]/g, "'").replace(/\s+/g, ' ').trim();
}

function resultsFor_(name, id) {
  const key = normName_(name);
  if (!key && !id) return { ok: false, error: 'no_name', results: [], pending: [] };
  const mine = function (o) { return (id && o.admin_id === id) || (key && normName_(o.admin_name) === key); };
  const byConv = {};
  const it = DriveApp.getFolderById(RESULTS_FOLDER_ID).getFiles();
  while (it.hasNext()) {
    const f = it.next();
    let r;
    try { r = JSON.parse(f.getBlob().getDataAsString()); } catch (err) { continue; }
    if (!r || !mine(r)) continue;
    const created = f.getDateCreated().getTime();
    // Если звонок оценивали повторно, берём самую свежую оценку
    if (!byConv[r.conversation_id] || byConv[r.conversation_id]._created < created) {
      r._created = created;
      byConv[r.conversation_id] = r;
    }
  }
  const results = Object.keys(byConv).map(function (k) { return byConv[k]; })
    .sort(function (a, b) { return (b.started_at || 0) - (a.started_at || 0); })
    .slice(0, 50);
  // Звонки, которые ещё не оценены
  const pending = [];
  const cit = DriveApp.getFolderById(CALLS_FOLDER_ID).getFiles();
  while (cit.hasNext()) {
    const f = cit.next();
    try {
      const c = JSON.parse(Utilities.newBlob(Utilities.base64Decode(f.getBlob().getDataAsString().trim())).getDataAsString('UTF-8'));
      if (!c || byConv[c.conversation_id] || !mine(c)) continue;
      if (!(c.transcript || []).some(function (t) { return t.role === 'admin'; })) continue;
      pending.push({ conversation_id: c.conversation_id, scenario_id: c.scenario_id, lang: c.scenario_lang, started_at: c.started_at, duration_sec: c.duration_sec });
    } catch (err) { /* пропускаем */ }
  }
  return { ok: true, results: results, pending: pending };
}

/* ---------- 3. Баллы в таблицу ---------- */
function syncScores() {
  const sh = sheet_();
  if (sh.getLastRow() < 2) return;
  const ids = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; });
  const it = DriveApp.getFolderById(RESULTS_FOLDER_ID).getFiles();
  const latest = {};
  while (it.hasNext()) {
    const f = it.next();
    let r;
    try { r = JSON.parse(f.getBlob().getDataAsString()); } catch (err) { continue; }
    const created = f.getDateCreated().getTime();
    if (!latest[r.conversation_id] || latest[r.conversation_id].created < created) latest[r.conversation_id] = { r: r, created: created };
  }
  Object.keys(latest).forEach(function (conv) {
    const row = ids.indexOf(conv);
    if (row === -1) return;
    const r = latest[conv].r;
    sh.getRange(row + 2, 11, 1, 2).setValues([[new Date(r.evaluated_at || latest[conv].created), r.score]]);
  });
  buildDashboard();
}

/* ---------- 4. Данные для дашборда ---------- */
function buildDashboard() {
  const tz = Session.getScriptTimeZone();
  // Последняя оценка по каждому звонку
  const results = {};
  const rit = DriveApp.getFolderById(RESULTS_FOLDER_ID).getFiles();
  while (rit.hasNext()) {
    const f = rit.next();
    let r;
    try { r = JSON.parse(f.getBlob().getDataAsString()); } catch (err) { continue; }
    if (!r || !r.conversation_id) continue;
    const created = f.getDateCreated().getTime();
    if (!results[r.conversation_id] || results[r.conversation_id].created < created) results[r.conversation_id] = { r: r, created: created };
  }
  // Все звонки, в том числе ещё не оценённые
  const calls = {};
  const cit = DriveApp.getFolderById(CALLS_FOLDER_ID).getFiles();
  while (cit.hasNext()) {
    const f = cit.next();
    try {
      const c = JSON.parse(Utilities.newBlob(Utilities.base64Decode(f.getBlob().getDataAsString().trim())).getDataAsString('UTF-8'));
      if (c && c.conversation_id) calls[c.conversation_id] = c;
    } catch (err) { /* пропускаем повреждённый файл */ }
  }
  const names = {};
  Object.keys(results).forEach(function (k) { const r = results[k].r; if (r.scenario_id && r.scenario_name) names[r.scenario_id] = r.scenario_name; });

  const ids = {};
  Object.keys(calls).forEach(function (k) { ids[k] = 1; });
  Object.keys(results).forEach(function (k) { ids[k] = 1; });

  const rows = [];
  const items = {};
  Object.keys(ids).forEach(function (id) {
    const c = calls[id] || {};
    const r = results[id] ? results[id].r : null;
    const scenarioId = (r && r.scenario_id) || c.scenario_id || '';
    if (!scenarioId) return; // тесты из ElevenLabs без сценария
    const hasAdmin = r || (c.transcript || []).some(function (t) { return t.role === 'admin'; });
    if (!hasAdmin) return; // звонок без единой реплики админа
    const started = (r && r.started_at) || c.started_at;
    if (!started) return;
    const d = new Date(started);
    const dow = (Number(Utilities.formatDate(d, tz, 'u')) + 6) % 7; // 0 = понедельник
    const monday = new Date(d.getTime() - dow * 86400000);
    const score = r && typeof r.score === 'number' ? r.score : null;
    const scenario = (r && r.scenario_name) || names[scenarioId] || scenarioId;
    rows.push({
      conversation_id: id,
      date: Utilities.formatDate(d, tz, 'yyyy-MM-dd'),
      week: Utilities.formatDate(monday, tz, 'yyyy-MM-dd'),
      admin: String((r && r.admin_name) || c.admin_name || '').trim(),
      scenario_id: scenarioId,
      scenario: scenario,
      lang: (r && r.lang) || c.scenario_lang || '',
      dept: c.scenario_dept || 'sales',
      duration_min: Math.round(((r && r.duration_sec) || c.duration_sec || 0) / 6) / 10,
      score: score,
      result: score == null ? 'не оценён' : score >= 80 ? 'сдан' : score >= 60 ? 'повторить' : 'пересдать',
      real_reason: r ? !!r.real_reason_found : null
    });
    if (r) (r.metrics || []).forEach(function (m) {
      if (m.passed === null || m.passed === undefined) return;
      const key = scenario + '|' + m.group + '|' + m.item;
      const a = items[key] || (items[key] = { scenario: scenario, group: m.group, item: m.item, n: 0, passed: 0 });
      a.n++;
      if (m.passed === true) a.passed++;
    });
  });
  rows.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  DriveApp.getFileById(DASH_CALLS_FILE_ID).setContent(JSON.stringify(rows));
  DriveApp.getFileById(DASH_ITEMS_FILE_ID).setContent(JSON.stringify(Object.keys(items).map(function (k) { return items[k]; })));
  Logger.log('Дашборд: звонков ' + rows.length + ', пунктов чек-листа ' + Object.keys(items).length);
}


/* ---------- 5. Автоматическая оценка (когда есть ключ Anthropic) ---------- */
// Ключ хранится в Script Properties под именем ANTHROPIC_API_KEY. Без ключа функция ничего не делает,
// и звонки оценивает супервайзер кнопкой в прототипе.
const BUNDLE_KEY = 'allright-trainer-v1';
const SCORING_MODEL = 'claude-sonnet-5-5';

function loadBundle_() {
  const files = scenarioFiles_();
  if (!files.length) return null;
  const raw = Utilities.base64Decode(files[0].getBlob().getDataAsString().trim());
  const key = Utilities.newBlob(BUNDLE_KEY).getBytes();
  const out = raw.map(function (b, i) { return (b ^ key[i % key.length]); });
  return JSON.parse(Utilities.newBlob(out).getDataAsString('UTF-8'));
}

function scorePending() {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) return;
  const bundle = loadBundle_();
  if (!bundle) return;
  const done = {};
  const rit = DriveApp.getFolderById(RESULTS_FOLDER_ID).getFiles();
  while (rit.hasNext()) { const m = /^result-(.+)\.json$/.exec(rit.next().getName()); if (m) done[m[1]] = 1; }
  const props = PropertiesService.getScriptProperties();
  const tries = JSON.parse(props.getProperty('SCORE_TRIES') || '{}');
  const week = Date.now() - 7 * 86400000;
  const cit = DriveApp.getFolderById(CALLS_FOLDER_ID).getFiles();
  let n = 0;
  while (cit.hasNext() && n < 5) {
    const f = cit.next();
    const m = /^call-(.+)\.txt$/.exec(f.getName());
    if (!m || done[m[1]] || (tries[m[1]] || 0) >= 3 || f.getDateCreated().getTime() < week) continue;
    let c;
    try { c = JSON.parse(Utilities.newBlob(Utilities.base64Decode(f.getBlob().getDataAsString().trim())).getDataAsString('UTF-8')); } catch (err) { continue; }
    if (!(c.transcript || []).some(function (t) { return t.role === 'admin'; })) continue;
    const scn = (bundle.scenarios || []).filter(function (x) { return x.id === c.scenario_id; })[0];
    const L = scn && scn.langs[c.scenario_lang];
    if (!L || !L.check) continue;
    tries[m[1]] = (tries[m[1]] || 0) + 1;
    props.setProperty('SCORE_TRIES', JSON.stringify(tries));
    n++;
    try {
      const tr = c.transcript.map(function (t) { return (t.role === 'mom' ? 'МАМА: ' : 'АДМИНИСТРАТОР: ') + t.text; }).join('\n');
      const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify({ model: SCORING_MODEL, max_tokens: 4000, messages: [{ role: 'user', content: L.check.scoringPrompt.replace('{{TRANSCRIPT}}', tr) }] })
      });
      if (resp.getResponseCode() !== 200) { Logger.log('Anthropic ' + resp.getResponseCode() + ': ' + resp.getContentText().slice(0, 300)); continue; }
      const text = (JSON.parse(resp.getContentText()).content || []).map(function (b) { return b.text || ''; }).join('');
      const r = JSON.parse(text.replace(/```json|```/g, '').trim());
      const byId = {};
      (r.metrics || []).forEach(function (x) { byId[x.id] = x; });
      const metrics = L.check.items.map(function (it) {
        const x = byId[it.id] || {};
        return { id: it.id, group: it.group, item: it.item, passed: x.passed === undefined ? null : x.passed, evidence: x.evidence || '', comment: x.comment || '' };
      });
      const score = metrics.length ? Math.round(metrics.filter(function (x) { return x.passed === true; }).length / metrics.length * 100) : 0;
      const out = {
        auto: true, conversation_id: c.conversation_id, admin_name: c.admin_name, admin_id: c.admin_id || '',
        scenario_id: c.scenario_id, scenario_name: L.name || scn.id, lang: c.scenario_lang, started_at: c.started_at, duration_sec: c.duration_sec,
        evaluated_at: Date.now(), score: score, ai_score: score, real_reason_found: !!r.real_reason_found,
        summary: r.summary || '', strengths: r.strengths || [], improvements: r.improvements || [], metrics: metrics,
        steps: (r.steps || []).map(function (x) { return { text: L.check.steps[x.n - 1] || ('Крок ' + x.n), result: x.result, comment: x.comment || '' }; }),
        objections: (r.objections || []).map(function (x) { return { title: L.check.objections[x.n - 1] || ('Заперечення ' + x.n), result: x.result, comment: x.comment || '' }; }),
        transcript: c.transcript
      };
      DriveApp.getFolderById(RESULTS_FOLDER_ID).createFile('result-' + c.conversation_id + '.json', JSON.stringify(out), MimeType.PLAIN_TEXT);
    } catch (err) { Logger.log('Оценка не удалась: ' + err); }
  }
}

/* ---------- Служебное ---------- */
function scenarioFiles_() {
  const it = DriveApp.getFolderById(FOLDER_ID).getFilesByName(SCENARIOS_NAME);
  const list = [];
  while (it.hasNext()) list.push(it.next());
  list.sort(function (a, b) { return b.getDateCreated() - a.getDateCreated(); });
  return list;
}

function sheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sh = ss.getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(HEADER);
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function text_(s) {
  return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT);
}

/**
 * Запустить вручную (кнопка ▶ Выполнить) после обновления кода:
 *  — создаёт секретный ключ для webhook (если его ещё нет),
 *  — включает ежечасный перенос баллов в таблицу и обновление данных дашборда,
 *  — проверяет доступ к таблице и папкам,
 *  — выводит в журнал ключ для адреса webhook.
 */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('WEBHOOK_TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('WEBHOOK_TOKEN', token);
  }
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncScores') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncScores').timeBased().everyHours(1).create();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'scorePending') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('scorePending').timeBased().everyMinutes(5).create();
  sheet_();
  DriveApp.getFolderById(FOLDER_ID);
  DriveApp.getFolderById(CALLS_FOLDER_ID);
  DriveApp.getFolderById(RESULTS_FOLDER_ID);
  DriveApp.getFileById(DASH_CALLS_FILE_ID);
  DriveApp.getFileById(DASH_ITEMS_FILE_ID);
  Logger.log('Ключ webhook: ' + token);
  Logger.log('Адрес для webhook: <URL веб-приложения>?token=' + token);
}

/**
 * Один раз: перенести в папку «звонки (JSON)» звонки, которые пришли до этой версии скрипта.
 * Запустить вручную после обновления кода.
 */
function backfillCalls() {
  const sh = sheet_();
  if (sh.getLastRow() < 2) return;
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 10).getValues();
  const folder = DriveApp.getFolderById(CALLS_FOLDER_ID);
  let n = 0;
  rows.forEach(function (r) {
    const conv = r[1];
    if (!conv || folder.getFilesByName('call-' + conv + '.txt').hasNext()) return;
    let tr = [];
    try { tr = JSON.parse(r[9]); } catch (err) { return; }
    tr = tr.map(function (t) {
      return { role: t.role, text: t.role === 'mom' ? String(t.text).replace(/\[[A-Za-z][A-Za-z _-]{1,24}\]\s*/g, '').trim() : t.text, t: t.t };
    }).filter(function (t) { return t.text; });
    if (!tr.length) return;
    const call = {
      conversation_id: conv, admin_name: r[2], scenario_id: r[3], scenario_lang: r[4], scenario_dept: r[5],
      started_at: r[6] instanceof Date ? r[6].getTime() : Date.now(), duration_sec: r[7] || 0, status: r[8], transcript: tr
    };
    folder.createFile('call-' + conv + '.txt', Utilities.base64Encode(JSON.stringify(call), Utilities.Charset.UTF_8), MimeType.PLAIN_TEXT);
    n++;
  });
  Logger.log('Перенесено звонков: ' + n);
}
