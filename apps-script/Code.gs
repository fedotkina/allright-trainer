/**
 * Тренажёр звонков AllRight — связка между ElevenLabs, Google Drive и страницей админов.
 *
 * Что делает:
 *  1. doPost — принимает post-call webhook от ElevenLabs и пишет звонок строкой в таблицу.
 *  2. doGet  — отдаёт странице админов самый свежий файл сценариев из папки на Drive.
 *
 * Разворачивается один раз: Deploy → New deployment → Web app,
 * Execute as: Me, Who has access: Anyone. Подробности в README.md.
 */

// ID таблицы «Тренажёр звонков — звонки админов»
const SHEET_ID = '1Mb09C2mI4ONGvxRJUBAU9OYfI65BClnQ_jidJ_DR7SY';
// ID папки «Тренажёр звонков — сценарии»
const FOLDER_ID = '1psY55TrBuVCc1jLb76et5wHVHaDhl0zN';
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
    return { role: t.role === 'agent' ? 'mom' : 'admin', text: t.message || '', t: Math.round(t.time_in_call_secs || 0) };
  }).filter(function (t) { return t.text; });

  let trJson = JSON.stringify(transcript);
  if (trJson.length > CELL_LIMIT) trJson = trJson.slice(0, CELL_LIMIT);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_();
    // Защита от дублей: ElevenLabs может повторить доставку
    const convId = d.conversation_id || '';
    if (convId && sh.getLastRow() > 1) {
      const ids = sh.getRange(2, 2, sh.getLastRow() - 1, 1).getValues().map(function (r) { return r[0]; });
      if (ids.indexOf(convId) !== -1) return json_({ ok: true, duplicate: true });
    }
    sh.appendRow([
      new Date(),
      convId,
      vars.admin_name || init.user_id || d.user_id || '',
      vars.scenario_id || '',
      vars.scenario_lang || '',
      vars.scenario_dept || '',
      meta.start_time_unix_secs ? new Date(meta.start_time_unix_secs * 1000) : '',
      meta.call_duration_secs || '',
      d.status || '',
      trJson,
      '',
      ''
    ]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

/* ---------- 2. Сценарии для страницы админов ---------- */
function doGet(e) {
  const files = scenarioFiles_();
  if (!files.length) return text_('');
  return text_(files[0].getBlob().getDataAsString());
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
 * Запустить ОДИН раз вручную (кнопка ▶ Run):
 *  — создаёт секретный ключ для webhook,
 *  — проверяет доступ к таблице и папке,
 *  — выводит в журнал ключ для адреса webhook.
 */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let token = props.getProperty('WEBHOOK_TOKEN');
  if (!token) {
    token = Utilities.getUuid().replace(/-/g, '');
    props.setProperty('WEBHOOK_TOKEN', token);
  }
  sheet_();
  DriveApp.getFolderById(FOLDER_ID);
  Logger.log('Ключ webhook: ' + token);
  Logger.log('Адрес для Назара: <URL веб-приложения>?token=' + token);
}
