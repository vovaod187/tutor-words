/**
 * «Слова для уроков» — сервер (Google Apps Script, привязан к таблице «Репетитор — слова»).
 *
 * Секреты и настройки — в свойствах скрипта (Настройки проекта → Свойства скрипта):
 *   BOT_TOKEN       — токен бота из @BotFather
 *   ALLOWED_IDS     — Telegram ID через запятую (кто может пользоваться)
 *   GEMINI_API_KEY  — ключ из Google AI Studio
 *   MODELS          — (необязательно) модели Gemini через запятую, по порядку
 *   GROQ_API_KEY    — (необязательно) ключ Groq: станет последней запасной моделью
 *   GROQ_MODEL      — (необязательно) имя модели Groq
 *
 * Ручные функции: setup, testModels, testSheet.
 */

const VERSION = 1;
const TZ = 'Europe/Minsk';
const LEVELS = ['A1', 'A2', 'B1', 'B2', 'C1'];
const MAX_BATCH = 20;
const DAILY_MODEL_LIMIT = 200;
const DEFAULT_MODELS = 'gemini-3.5-flash,gemini-3.6-flash-lite,gemini-3.8-flash';
const DEFAULT_GROQ_MODEL = 'llama-3.3-70b-versatile';
const SKIP_MINUTES = 15;

const SH = { students: 'Ученики', words: 'Слова', lists: 'Листы', log: 'Журнал', settings: 'Настройки' };
const HEAD = {
  'Ученики': ['id', 'имя', 'уровень', 'создан', 'активен'],
  'Слова': ['id', 'ученик', 'ввод', 'слово', 'транскрипция', 'перевод', 'пример', 'выделить', 'уровень', 'лист', 'добавлено', 'модель'],
  'Листы': ['id', 'ученик', 'номер', 'создан', 'слов'],
  'Журнал': ['время', 'действие', 'модель', 'слов', 'мс', 'итог'],
  'Настройки': ['ключ', 'значение', 'пояснение']
};
const S = { id: 0, name: 1, level: 2, created: 3, active: 4 };
const W = { id: 0, student: 1, input: 2, word: 3, ipa: 4, tr: 5, ex: 6, hit: 7, level: 8, list: 9, added: 10, model: 11 };
const L = { id: 0, student: 1, no: 2, created: 3, count: 4 };
const DEFAULT_SETTINGS = [
  ['fontSize', '12', 'размер шрифта в Word, pt (от 10 до 14)'],
  ['maxWords', '30', 'не больше стольких слов на лист; сколько влезает на А4, бот считает сам']
];

const ACTIONS = {
  state: actState_,
  addStudent: actAddStudent_,
  updateStudent: actUpdateStudent_,
  words: actWords_,
  generate: actGenerate_,
  regenerate: actRegenerate_,
  updateWord: actUpdateWord_,
  deleteWord: actDeleteWord_,
  buildSheet: actBuildSheet_,
  reopenSheet: actReopenSheet_
};

// ───────────────────────── вход ─────────────────────────

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad data' });
  }
  try {
    const user = verify_(req.initData);
    if (!user || !user.id) return json_({ ok: false, error: 'denied' });
    if (req.action === 'whoami') {
      return json_({ ok: true, version: VERSION, allowed: isAllowed_(user.id), id: String(user.id) });
    }
    if (!isAllowed_(user.id)) return json_({ ok: false, error: 'not_allowed', id: String(user.id) });
    const fn = Object.prototype.hasOwnProperty.call(ACTIONS, req.action) ? ACTIONS[req.action] : null;
    if (!fn) return json_({ ok: false, error: 'bad action' });
    const res = fn(req, user) || {};
    res.ok = true;
    return json_(res);
  } catch (err) {
    if (err instanceof Fail) return json_(Object.assign({ ok: false, error: err.code }, err.extra));
    console.error(err && err.stack ? err.stack : err);
    return json_({ ok: false, error: 'server error' });
  }
}

function doGet() {
  return ContentService.createTextOutput('Слова для уроков: сервер работает, версия ' + VERSION);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function Fail(code, extra) {
  this.code = code;
  this.extra = extra || {};
}

function fail_(code, extra) {
  throw new Fail(code, extra);
}

// ───────────────────────── доступ ─────────────────────────

/** Свойство скрипта без лишних пробелов и кавычек по краям (частая ошибка при вставке). */
function prop_(key) {
  const v = PropertiesService.getScriptProperties().getProperty(key);
  return v == null ? null : String(v).trim().replace(/^["'«»“”]+|["'«»“”]+$/g, '').trim();
}

/** Токен или ключ: дополнительно убираем пробелы и переносы строк внутри. */
function secret_(key) {
  return String(prop_(key) || '').replace(/\s+/g, '');
}

function allowedIds_() {
  return String(prop_('ALLOWED_IDS') || '').split(/[\s,;]+/).filter(Boolean);
}

function isAllowed_(id) {
  return allowedIds_().indexOf(String(id)) >= 0;
}

/** Проверка подписи Telegram initData (HMAC-SHA256, не старше 24 часов). */
function verify_(initData) {
  if (!initData || typeof initData !== 'string') return null;
  const token = secret_('BOT_TOKEN');
  if (!token) fail_('no_token');
  const data = {};
  let hash = '';
  initData.split('&').forEach(function (pair) {
    const i = pair.indexOf('=');
    if (i < 0) return;
    let k, v;
    try {
      k = decodeURIComponent(pair.slice(0, i).replace(/\+/g, ' '));
      v = decodeURIComponent(pair.slice(i + 1).replace(/\+/g, ' '));
    } catch (err) {
      return;
    }
    if (k === 'hash') hash = v;
    else data[k] = v;
  });
  if (!hash) return null;
  const check = Object.keys(data).sort().map(function (k) { return k + '=' + data[k]; }).join('\n');
  const secret = Utilities.computeHmacSha256Signature(token, 'WebAppData');
  const sig = Utilities.computeHmacSha256Signature(Utilities.newBlob(check).getBytes(), secret);
  if (hex_(sig) !== String(hash).toLowerCase()) return null;
  const age = Date.now() / 1000 - Number(data.auth_date || 0);
  if (!(age >= -300 && age < 86400)) return null;
  try {
    return JSON.parse(data.user);
  } catch (err) {
    return null;
  }
}

function hex_(bytes) {
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

// ───────────────────────── таблица ─────────────────────────

function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const ss = ss_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    const n = HEAD[name].length;
    sh.getRange(1, 1, sh.getMaxRows(), n).setNumberFormat('@');
    sh.getRange(1, 1, 1, n).setValues([HEAD[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Все строки листа без заголовка; у каждой строки есть свойство row (номер в таблице). */
function rows_(name) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEAD[name].length).getValues().map(function (r, i) {
    r.row = i + 2;
    return r;
  });
}

function appendRows_(name, rows) {
  if (!rows.length) return;
  const sh = sheet_(name);
  const n = HEAD[name].length;
  const start = sh.getLastRow() + 1;
  const need = start + rows.length - 1;
  if (sh.getMaxRows() < need) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows() + 50);
  const range = sh.getRange(start, 1, rows.length, n);
  range.setNumberFormat('@');
  range.setValues(rows.map(function (r) { return r.map(safe_); }));
}

/** Защита от формул: текст, начинающийся с = + - @, пишем с апострофом. */
function safe_(v) {
  const s = v == null ? '' : String(v);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/** Чтение ячейки: строка без защитного апострофа. */
function str_(v) {
  const s = v == null ? '' : String(v);
  return /^'[=+\-@]/.test(s) ? s.slice(1) : s;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function stamp_() {
  return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
}

function dateRu_(stamp) {
  const m = String(stamp || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return m ? m[3] + '.' + m[2] + '.' + m[1] : Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy');
}

function newId_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 12);
}

function settings_() {
  const out = {};
  DEFAULT_SETTINGS.forEach(function (s) { out[s[0]] = s[1]; });
  rows_(SH.settings).forEach(function (r) {
    const k = str_(r[0]).trim();
    if (k) out[k] = str_(r[1]).trim();
  });
  return out;
}

function maxWords_(set) {
  const n = parseInt((set || settings_()).maxWords, 10);
  return n >= 1 && n <= 40 ? n : 30;
}

function fontSize_(set) {
  const n = parseFloat(String((set || settings_()).fontSize).replace(',', '.'));
  return n >= 10 && n <= 14 ? n : 12;
}

/** Слова черновика ученика и расчёт листа: сколько влезет, на сколько % заполнен, сколько ещё примерно войдёт. */
function sheetInfo_(words, set) {
  const s = set || settings_();
  const f = fitSheet_(words, fontSize_(s), maxWords_(s));
  return { fit: f.fit, fill: Math.round(f.fill * 100), left: f.left, more: f.more };
}

function draft_(sid, set) {
  const words = rows_(SH.words)
    .filter(function (r) { return String(r[W.student]) === String(sid) && !String(r[W.list]); })
    .map(wordObj_);
  return { words: words, sheet: sheetInfo_(words, set) };
}

function log_(action, model, n, ms, result) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    appendRows_(SH.log, [[stamp_(), action, model || '', n === '' || n == null ? '' : String(n), ms === '' || ms == null ? '' : String(ms), String(result || '').slice(0, 300)]]);
    const sh = sheet_(SH.log);
    if (sh.getLastRow() > 3000) sh.deleteRows(2, 1000);
  } catch (err) {
    console.error('log: ' + err);
  } finally {
    lock.releaseLock();
  }
}

// ───────────────────────── ученики ─────────────────────────

function studentObj_(r) {
  const level = String(r[S.level]);
  return { id: String(r[S.id]), name: str_(r[S.name]), level: LEVELS.indexOf(level) >= 0 ? level : 'A2' };
}

function student_(id) {
  const r = rows_(SH.students).filter(function (x) { return String(x[S.id]) === String(id); })[0];
  if (!r) fail_('no_student');
  return studentObj_(r);
}

function listsOf_(sid) {
  return rows_(SH.lists)
    .filter(function (r) { return String(r[L.student]) === String(sid); })
    .map(function (r) { return { no: Number(r[L.no]) || 0, created: String(r[L.created]), date: dateRu_(r[L.created]), count: Number(r[L.count]) || 0 }; })
    .sort(function (a, b) { return b.no - a.no; });
}

function actState_() {
  const set = settings_();
  const draft = {};
  const lists = {};
  rows_(SH.words).forEach(function (r) {
    if (!String(r[W.list])) {
      const sid = String(r[W.student]);
      (draft[sid] = draft[sid] || []).push(wordObj_(r));
    }
  });
  rows_(SH.lists).forEach(function (r) {
    const sid = String(r[L.student]);
    lists[sid] = Math.max(lists[sid] || 0, Number(r[L.no]) || 0);
  });
  const students = rows_(SH.students)
    .filter(function (r) { return String(r[S.active]).trim().toLowerCase() !== 'нет'; })
    .map(function (r) {
      const s = studentObj_(r);
      const d = draft[s.id] || [];
      s.draft = d.length;
      s.sheet = sheetInfo_(d, set);
      s.lists = lists[s.id] || 0;
      return s;
    });
  return { students: students, levels: LEVELS };
}

function cleanName_(name) {
  const s = clean_(name);
  if (!s || s.length > 40) fail_('bad_name');
  return s;
}

function cleanLevel_(level) {
  const s = String(level || '').toUpperCase();
  if (LEVELS.indexOf(s) < 0) fail_('bad_level');
  return s;
}

function actAddStudent_(req) {
  const name = cleanName_(req.name);
  const level = cleanLevel_(req.level || 'A2');
  const id = newId_();
  withLock_(function () {
    appendRows_(SH.students, [[id, name, level, stamp_(), 'да']]);
  });
  return { student: { id: id, name: name, level: level, draft: 0, lists: 0, sheet: sheetInfo_([]) } };
}

function actUpdateStudent_(req) {
  const name = req.name != null ? cleanName_(req.name) : null;
  const level = req.level != null ? cleanLevel_(req.level) : null;
  return withLock_(function () {
    const r = rows_(SH.students).filter(function (x) { return String(x[S.id]) === String(req.id); })[0];
    if (!r) fail_('no_student');
    const sh = sheet_(SH.students);
    if (name != null) sh.getRange(r.row, S.name + 1).setValue(safe_(name));
    if (level != null) sh.getRange(r.row, S.level + 1).setValue(level);
    if (name != null) r[S.name] = name;
    if (level != null) r[S.level] = level;
    return { student: studentObj_(r) };
  });
}

// ───────────────────────── слова ─────────────────────────

function wordObj_(r) {
  return {
    id: String(r[W.id]),
    input: str_(r[W.input]),
    word: str_(r[W.word]),
    ipa: str_(r[W.ipa]),
    tr: str_(r[W.tr]),
    ex: str_(r[W.ex]),
    hit: str_(r[W.hit]),
    level: String(r[W.level]),
    list: String(r[W.list]),
    added: String(r[W.added])
  };
}

function findWord_(id) {
  const r = rows_(SH.words).filter(function (x) { return String(x[W.id]) === String(id); })[0];
  if (!r) fail_('no_word');
  return r;
}

function actWords_(req) {
  const st = student_(req.studentId);
  const listNo = req.listNo ? String(req.listNo) : '';
  const words = rows_(SH.words)
    .filter(function (r) { return String(r[W.student]) === st.id && String(r[W.list]) === listNo; })
    .map(wordObj_);
  const res = { student: st, words: words, lists: listsOf_(st.id) };
  if (!listNo) res.sheet = sheetInfo_(words);
  return res;
}

/** Разбор ввода: запятая, точка с запятой, новая строка; запятые в скобках не делят. */
function parseItems_(text) {
  const parts = [];
  let cur = '';
  let depth = 0;
  String(text || '').split('').forEach(function (ch) {
    if (ch === '(') depth++;
    if (ch === ')') depth = Math.max(0, depth - 1);
    if (depth === 0 && (ch === ',' || ch === ';' || ch === '\n' || ch === '\r')) {
      parts.push(cur);
      cur = '';
    } else {
      cur += ch;
    }
  });
  parts.push(cur);
  const out = [];
  const seen = {};
  parts.forEach(function (p) {
    const input = p.replace(/\s+/g, ' ').replace(/^\s*\d+\s*[.)]\s*/, '').trim().slice(0, 80);
    if (!input) return;
    const m = input.match(/^(.*?)\s*\((.*)\)\s*$/);
    const base = (m ? m[1] : input).trim();
    const hint = m ? m[2].trim() : '';
    if (!base) return;
    const key = key_(base);
    if (!key || seen[key]) return;
    seen[key] = 1;
    out.push({ input: input, base: base, hint: hint, key: key });
  });
  return out;
}

/** Ключ для поиска повторов: без подсказки в скобках, регистра и лишних знаков. */
function key_(s) {
  return String(s || '')
    .toLowerCase()
    .replace(/\(.*?\)/g, ' ')
    .replace(/[’`´]/g, "'")
    .replace(/[^a-z0-9' -]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function actGenerate_(req) {
  const st = student_(req.studentId);
  const items = parseItems_(req.text);
  if (!items.length) fail_('empty_input');
  if (items.length > MAX_BATCH) fail_('too_many', { max: MAX_BATCH });

  let fresh = items;
  const duplicates = [];
  if (!req.force) {
    const seen = {};
    rows_(SH.words).forEach(function (r) {
      if (String(r[W.student]) !== st.id) return;
      const info = { list: String(r[W.list]), date: dateRu_(r[W.added]) };
      [key_(str_(r[W.word])), key_(str_(r[W.input]))].forEach(function (k) {
        if (k && !seen[k]) seen[k] = info;
      });
    });
    fresh = items.filter(function (it) { return !seen[it.key]; });
    items.forEach(function (it) {
      if (seen[it.key]) duplicates.push({ input: it.input, list: seen[it.key].list, date: seen[it.key].date });
    });
  }

  let added = [];
  let model = '';
  if (fresh.length) {
    const res = callModel_(fresh, st.level, '');
    model = res.model;
    const now = stamp_();
    const rows = res.items.map(function (it) {
      return [newId_(), st.id, it.input, it.word, it.ipa, it.tr, it.ex, it.hit, st.level, '', now, res.model];
    });
    withLock_(function () { appendRows_(SH.words, rows); });
    added = rows.map(wordObj_);
  }
  const d = draft_(st.id);
  return { added: added, duplicates: duplicates, model: model, words: d.words, sheet: d.sheet };
}

function actRegenerate_(req) {
  const r = findWord_(req.id);
  const st = student_(String(r[W.student]));
  const word = str_(r[W.word]);
  const items = [{ input: word, base: word, hint: '', key: key_(word) }];
  const extra = 'Write a NEW example sentence, clearly different from this one: "' + str_(r[W.ex]) + '"\n';
  const res = callModel_(items, st.level, extra);
  const it = res.items[0];
  return withLock_(function () {
    const cur = findWord_(req.id);
    const sh = sheet_(SH.words);
    sh.getRange(cur.row, W.ex + 1, 1, 3).setValues([[safe_(it.ex), safe_(it.hit), st.level]]);
    sh.getRange(cur.row, W.model + 1).setValue(res.model);
    cur[W.ex] = it.ex;
    cur[W.hit] = it.hit;
    cur[W.level] = st.level;
    const out = { word: wordObj_(cur) };
    if (!String(cur[W.list])) out.sheet = draft_(st.id).sheet;
    return out;
  });
}

function actUpdateWord_(req) {
  const word = clean_(req.word);
  if (!word) fail_('bad_word');
  const ipa = clean_(req.ipa) ? ipa_(req.ipa) : '';
  const tr = clean_(req.tr);
  const ex = clean_(req.ex);
  return withLock_(function () {
    const cur = findWord_(req.id);
    const hit = findHit_(ex, str_(cur[W.hit]), word);
    sheet_(SH.words).getRange(cur.row, W.word + 1, 1, 5).setValues([[word, ipa, tr, ex, hit].map(safe_)]);
    cur[W.word] = word;
    cur[W.ipa] = ipa;
    cur[W.tr] = tr;
    cur[W.ex] = ex;
    cur[W.hit] = hit;
    const out = { word: wordObj_(cur) };
    if (!String(cur[W.list])) out.sheet = draft_(String(cur[W.student])).sheet;
    return out;
  });
}

function actDeleteWord_(req) {
  return withLock_(function () {
    const cur = findWord_(req.id);
    sheet_(SH.words).deleteRow(cur.row);
    return String(cur[W.list]) ? {} : { sheet: draft_(String(cur[W.student])).sheet };
  });
}

// ───────────────────────── листы ─────────────────────────

function actBuildSheet_(req, user) {
  const st = student_(req.studentId);
  const set = settings_();
  const all = rows_(SH.words).filter(function (r) { return String(r[W.student]) === st.id; });
  let no;
  let words;
  let date;
  const isNew = !req.listNo;
  if (!isNew) {
    no = Number(req.listNo);
    words = all.filter(function (r) { return String(r[W.list]) === String(no); }).map(wordObj_);
    if (!words.length) fail_('no_list');
    const meta = listsOf_(st.id).filter(function (l) { return l.no === no; })[0];
    date = meta ? meta.date : dateRu_('');
  } else {
    const draftWords = all.filter(function (r) { return !String(r[W.list]); }).map(wordObj_);
    if (!draftWords.length) fail_('empty_draft');
    words = draftWords.slice(0, sheetInfo_(draftWords, set).fit);
    no = listsOf_(st.id).reduce(function (m, l) { return Math.max(m, l.no); }, 0) + 1;
    date = dateRu_(stamp_());
  }

  sendSheet_(st.name, no, date, words, user.id, fontSize_(set));

  let left = null;
  if (isNew) {
    withLock_(function () {
      const ids = {};
      words.forEach(function (w) { ids[w.id] = 1; });
      const sh = sheet_(SH.words);
      rows_(SH.words).forEach(function (r) {
        if (ids[String(r[W.id])]) sh.getRange(r.row, W.list + 1).setValue(String(no));
      });
      appendRows_(SH.lists, [[newId_(), st.id, String(no), stamp_(), String(words.length)]]);
    });
    left = all.filter(function (r) { return !String(r[W.list]); }).length - words.length;
  }
  log_('лист', '', words.length, '', st.name + ' · лист ' + no + (isNew ? '' : ' (повтор)'));
  return { no: no, count: words.length, left: left };
}

function actReopenSheet_(req) {
  const st = student_(req.studentId);
  const no = Number(req.listNo);
  return withLock_(function () {
    const lists = rows_(SH.lists).filter(function (r) { return String(r[L.student]) === st.id; });
    const max = lists.reduce(function (m, r) { return Math.max(m, Number(r[L.no]) || 0); }, 0);
    if (!no || no !== max) fail_('not_last');
    const shW = sheet_(SH.words);
    rows_(SH.words).forEach(function (r) {
      if (String(r[W.student]) === st.id && String(r[W.list]) === String(no)) shW.getRange(r.row, W.list + 1).setValue('');
    });
    const lr = lists.filter(function (r) { return Number(r[L.no]) === no; })[0];
    if (lr) sheet_(SH.lists).deleteRow(lr.row);
    return {};
  });
}

function plural_(n, forms) {
  const a = Math.abs(n) % 100;
  const b = a % 10;
  if (a > 10 && a < 20) return forms[2];
  if (b > 1 && b < 5) return forms[1];
  if (b === 1) return forms[0];
  return forms[2];
}

function fileName_(name, no, date) {
  const safe = String(name).replace(/[\\\/:*?"<>|]+/g, ' ').replace(/\s+/g, ' ').trim() || 'Ученик';
  return safe + ' - лист ' + no + ' - ' + date + '.docx';
}

function sendSheet_(name, no, date, words, chatId, fontSize) {
  const blob = docx_(docxParts_(name, date, words, fontSize), fileName_(name, no, date));
  const caption = name + ' · лист ' + no + ' · ' + words.length + ' ' + plural_(words.length, ['слово', 'слова', 'слов']);
  return tg_('sendDocument', { chat_id: String(chatId), caption: caption, document: blob });
}

function tg_(method, payload) {
  const token = secret_('BOT_TOKEN');
  const resp = UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/' + method, {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  });
  let data = {};
  try {
    data = JSON.parse(resp.getContentText());
  } catch (err) {
    data = {};
  }
  if (data.ok) return data.result;
  const desc = String(data.description || ('HTTP ' + resp.getResponseCode()));
  log_('telegram', method, '', '', desc);
  if (data.error_code === 403 || /chat not found|initiate/i.test(desc)) fail_('start_bot');
  fail_('telegram', { message: desc.slice(0, 200) });
}

// ───────────────────────── модель ─────────────────────────

const LEVEL_GUIDE = {
  A1: 'beginner: present simple, the most common everyday words, 5–8 words',
  A2: 'elementary: present and past simple, going to, common words, 6–10 words',
  B1: 'intermediate: common tenses, one linking word (because, when, if), 8–13 words',
  B2: 'upper-intermediate: varied tenses and structures, natural phrasing, 10–16 words',
  C1: 'advanced: rich vocabulary and complex structures, 12–20 words'
};

const SYSTEM_PROMPT = [
  'You prepare vocabulary for printed worksheets. The teacher is a private English tutor; the students are Russian-speaking teenagers aged 10–18.',
  'For every numbered item return one object, in the same order:',
  '- "input": the item exactly as given (without the number).',
  '- "word": the English word or phrase in dictionary form (lowercase unless it is a proper noun); fix an obvious typo.',
  '- "ipa": British English (RP) pronunciation in IPA between slashes, e.g. "/ˈbɒrəʊ/". For a phrase, transcribe the whole phrase, e.g. "/ɡɪv ˈʌp/".',
  '- "translation": the 1–3 most common Russian meanings, short; separate synonyms with ", " and different meanings or parts of speech with "; ". No explanations, no English.',
  '- "example": ONE natural English sentence at the requested CEFR level about something a teenager relates to (school, hobbies, friends, sport, family, travel, technology). It must contain the word or phrase; an inflected form such as "looking forward to" is fine. Avoid alcohol, romance, violence and politics.',
  '- "hit": the exact part of "example" that is the word or phrase as used, copied character for character (e.g. "looking forward to").',
  'If an item has a meaning hint, use that meaning for the translation and the example, and do not put the hint into "word".',
  'Reply with JSON only: {"items": [ ... ]}.'
].join('\n');

const SCHEMA_GEMINI = {
  type: 'OBJECT',
  properties: {
    items: {
      type: 'ARRAY',
      items: {
        type: 'OBJECT',
        properties: {
          input: { type: 'STRING' },
          word: { type: 'STRING' },
          ipa: { type: 'STRING' },
          translation: { type: 'STRING' },
          example: { type: 'STRING' },
          hit: { type: 'STRING' }
        },
        required: ['input', 'word', 'ipa', 'translation', 'example', 'hit']
      }
    }
  },
  required: ['items']
};

function buildPrompt_(items, level, extra) {
  const lv = LEVELS.indexOf(level) >= 0 ? level : 'A2';
  const user = 'CEFR level: ' + lv + ' — ' + LEVEL_GUIDE[lv] + '.\n' + (extra || '') + 'Items:\n' +
    items.map(function (it, i) { return (i + 1) + '. ' + it.base + (it.hint ? ' — meaning: ' + it.hint : ''); }).join('\n');
  return { system: SYSTEM_PROMPT, user: user };
}

function modelChain_() {
  const list = String(prop_('MODELS') || DEFAULT_MODELS).split(',')
    .map(function (s) { return s.trim(); })
    .filter(Boolean)
    .map(function (m) { return { provider: 'gemini', model: m }; });
  if (prop_('GROQ_API_KEY')) list.push({ provider: 'groq', model: prop_('GROQ_MODEL') || DEFAULT_GROQ_MODEL });
  return list;
}

function countCall_() {
  withLock_(function () {
    const props = PropertiesService.getScriptProperties();
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    let c = {};
    try {
      c = JSON.parse(props.getProperty('MODEL_CALLS') || '{}');
    } catch (err) {
      c = {};
    }
    if (c.d !== today) c = { d: today, n: 0 };
    if (c.n >= DAILY_MODEL_LIMIT) fail_('daily_limit');
    c.n++;
    props.setProperty('MODEL_CALLS', JSON.stringify(c));
  });
}

/** Цепочка моделей: первая, что вернула правильный JSON. */
function callModel_(items, level, extra) {
  countCall_();
  const prompt = buildPrompt_(items, level, extra);
  const cache = CacheService.getScriptCache();
  const chain = modelChain_();
  let badKey = false;
  for (let i = 0; i < chain.length; i++) {
    const m = chain[i];
    if (badKey && m.provider === 'gemini') continue;
    const skipKey = 'skip:' + m.model;
    if (cache.get(skipKey)) continue;
    const r = tryModel_(m, prompt, items);
    if (r.items) return r;
    if (r.badKey && m.provider === 'gemini') badKey = true;
    if (r.skip) cache.put(skipKey, '1', SKIP_MINUTES * 60);
  }
  fail_(badKey ? 'bad_key' : 'model_unavailable');
}

/** Одна модель, до двух попыток (вторая — без настроек размышлений). */
function tryModel_(m, prompt, items) {
  let last = '';
  for (let attempt = 0; attempt < 2; attempt++) {
    const t0 = Date.now();
    let r;
    try {
      r = m.provider === 'groq' ? groq_(m.model, prompt) : gemini_(m.model, prompt, attempt);
    } catch (err) {
      r = { error: 'сеть: ' + (err && err.message ? err.message : err), skip: true };
    }
    const ms = Date.now() - t0;
    if (r.text != null) {
      const parsed = validate_(r.text, items);
      log_('модель', m.model, items.length, ms, parsed ? 'ok' : 'кривой JSON');
      if (parsed) return { items: parsed, model: m.model, ms: ms };
      last = 'кривой JSON';
      continue;
    }
    log_('модель', m.model, items.length, ms, r.error);
    last = r.error;
    if (!r.retry) return { error: last, skip: !!r.skip, badKey: !!r.badKey };
  }
  return { error: last };
}

function gemini_(model, p, attempt) {
  const key = secret_('GEMINI_API_KEY');
  if (!key) return { error: 'нет GEMINI_API_KEY', badKey: true };
  const gen = { responseMimeType: 'application/json', responseSchema: SCHEMA_GEMINI };
  if (attempt === 0) gen.thinkingConfig = { thinkingLevel: 'low' };
  const body = {
    systemInstruction: { parts: [{ text: p.system }] },
    contents: [{ role: 'user', parts: [{ text: p.user }] }],
    generationConfig: gen
  };
  const resp = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': key },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const txt = resp.getContentText();
  if (code === 200) {
    let data;
    try {
      data = JSON.parse(txt);
    } catch (err) {
      return { error: 'ответ не JSON', retry: true };
    }
    const cand = (data.candidates || [])[0] || {};
    const parts = (cand.content || {}).parts || [];
    const out = parts.filter(function (x) { return !x.thought && typeof x.text === 'string'; }).map(function (x) { return x.text; }).join('');
    if (!out) return { error: 'пустой ответ ' + (cand.finishReason || ''), retry: true };
    return { text: out };
  }
  const msg = 'HTTP ' + code + ': ' + shortErr_(txt);
  if (code === 400 && /api key not valid|API_KEY_INVALID/i.test(txt)) return { error: 'Google не принял ключ GEMINI_API_KEY (' + msg + ')', badKey: true };
  if (code === 400) return { error: msg, retry: attempt === 0 };
  if (code === 404 || code === 429 || code >= 500) return { error: msg, skip: true };
  return { error: msg, skip: code === 403 };
}

function groq_(model, p) {
  const key = secret_('GROQ_API_KEY');
  if (!key) return { error: 'нет GROQ_API_KEY' };
  const body = {
    model: model,
    messages: [{ role: 'system', content: p.system }, { role: 'user', content: p.user }],
    response_format: { type: 'json_object' },
    temperature: 0.7
  };
  const resp = UrlFetchApp.fetch('https://api.groq.com/openai/v1/chat/completions', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  const code = resp.getResponseCode();
  const txt = resp.getContentText();
  if (code === 200) {
    try {
      const data = JSON.parse(txt);
      return { text: String(data.choices[0].message.content || '') };
    } catch (err) {
      return { error: 'ответ не JSON', retry: true };
    }
  }
  const msg = 'HTTP ' + code + ': ' + shortErr_(txt);
  if (code === 404 || code === 429 || code >= 500) return { error: msg, skip: true };
  return { error: msg, skip: code === 403 };
}

function shortErr_(txt) {
  try {
    const d = JSON.parse(txt);
    return String((d.error && (d.error.message || d.error)) || txt).slice(0, 200);
  } catch (err) {
    return String(txt).slice(0, 200);
  }
}

function clean_(s) {
  return String(s == null ? '' : s).replace(/\s+/g, ' ').trim().slice(0, 300);
}

function ipa_(s) {
  const t = clean_(s).replace(/^[\/\[\]]+|[\/\[\]]+$/g, '').trim();
  return t ? '/' + t + '/' : '';
}

function findHit_(ex, hit, word) {
  const low = String(ex || '').toLowerCase();
  const tries = [hit, word];
  for (let i = 0; i < tries.length; i++) {
    const t = clean_(tries[i]);
    if (!t) continue;
    const at = low.indexOf(t.toLowerCase());
    if (at >= 0) return ex.substr(at, t.length);
  }
  return '';
}

function validate_(text, items) {
  let data;
  try {
    data = JSON.parse(String(text).replace(/^\s*```(?:json)?/i, '').replace(/```\s*$/, ''));
  } catch (err) {
    return null;
  }
  const arr = Array.isArray(data) ? data : data && data.items;
  if (!Array.isArray(arr) || arr.length !== items.length) return null;
  const out = [];
  for (let i = 0; i < arr.length; i++) {
    const a = arr[i] || {};
    const word = clean_(a.word) || items[i].base;
    const ipa = ipa_(a.ipa);
    const tr = clean_(a.translation);
    const ex = clean_(a.example);
    if (!ipa || !tr || !ex) return null;
    out.push({ input: items[i].input, word: word, ipa: ipa, tr: tr, ex: ex, hit: findHit_(ex, a.hit, word) });
  }
  return out;
}

// ───────────────────────── Word (.docx) ─────────────────────────

const SHEET_COLS = [600, 1813, 1989, 2193, 3611]; // ширины колонок, twips (сумма = ширина текста A4)
const CELL_PAD_X = 200; // поля ячейки слева+справа, twips
// Высота страницы под строки таблицы, pt: 757 (текст А4) − шапка − строка заголовков − запас.
// Коэффициенты подобраны по вёрстке LibreOffice со шрифтами, совместимыми с Times New Roman и Arial.
const FIT = { usable: 757, header: 38, tableHead: 20, reserve: 16, lineK: 1.207, rowPad: 7.6 };

/** Примерная ширина символа в долях кегля (Times New Roman / Arial). */
function charEm_(ch, sans) {
  if (/[\s]/.test(ch)) return 0.25;
  if (/[iljtf.,;:'!|ɪ`ˈˌ]/.test(ch)) return sans ? 0.26 : 0.28;
  if (/[rsI()\[\]\/\-ʃ]/.test(ch)) return sans ? 0.36 : 0.36;
  if (/[mwMWшщжмыюШЩЖМЫЮ]/.test(ch)) return sans ? 0.84 : 0.78;
  if (/[A-ZА-ЯЁ]/.test(ch)) return 0.7;
  return sans ? 0.56 : 0.5;
}

function textEm_(s, bold, sans) {
  let w = 0;
  for (const ch of String(s)) w += charEm_(ch, sans);
  return w * (bold ? 1.08 : 1);
}

/** Сколько строк займёт текст в колонке шириной widthPt при кегле sizePt. */
function wrapLines_(text, widthPt, sizePt, bold, sans) {
  const words = String(text || '').split(/\s+/).filter(Boolean);
  if (!words.length) return 1;
  const space = 0.25 * sizePt;
  let lines = 1;
  let cur = 0;
  words.forEach(function (tok) {
    let w = textEm_(tok, bold, sans) * sizePt;
    if (cur > 0 && cur + space + w <= widthPt) {
      cur += space + w;
      return;
    }
    if (cur > 0) lines++;
    while (w > widthPt) {
      lines++;
      w -= widthPt;
    }
    cur = w;
  });
  return lines;
}

function rowHeight_(w, fs) {
  const inner = SHEET_COLS.map(function (c) { return (c - CELL_PAD_X) / 20; });
  const n = Math.max(
    wrapLines_(w.word, inner[1], fs, true, false),
    wrapLines_(w.ipa, inner[2], fs - 1, false, true),
    wrapLines_(w.tr, inner[3], fs, false, false),
    wrapLines_(w.ex, inner[4], fs, false, false)
  );
  return n * fs * FIT.lineK + FIT.rowPad;
}

/** Сколько слов (по порядку) помещается на один лист и насколько он заполнен. */
function fitSheet_(words, fs, maxWords) {
  const cap = FIT.usable - FIT.header - FIT.tableHead - FIT.reserve;
  let used = 0;
  let fit = 0;
  for (let i = 0; i < words.length && fit < maxWords; i++) {
    const h = rowHeight_(words[i], fs);
    if (fit > 0 && used + h > cap) break;
    used += h;
    fit++;
  }
  const avg = fit ? used / fit : fs * FIT.lineK * 2 + FIT.rowPad;
  const left = fit >= maxWords ? 0 : Math.max(0, Math.floor((cap - used) / avg));
  return { fit: fit, fill: Math.min(1, used / cap), left: left, more: words.length - fit };
}

function xmlEsc_(s) {
  return String(s == null ? '' : s)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function splitHit_(ex, hit) {
  const e = String(ex || '');
  const h = String(hit || '');
  const at = h ? e.toLowerCase().indexOf(h.toLowerCase()) : -1;
  if (at < 0) return [e, '', ''];
  return [e.slice(0, at), e.slice(at, at + h.length), e.slice(at + h.length)];
}

/** Части файла .docx (XML), лист «Б»: без вертикальных линий, полосы через строку. */
function docxParts_(name, date, words, fontSize) {
  const fs = Math.round((fontSize || 12) * 2);
  const TEXT_W = 10206; // ширина текста A4 при полях 1,5 см, в twips
  const cols = SHEET_COLS;
  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

  function run(text, o) {
    o = o || {};
    if (text === '' || text == null) return '';
    let rpr = '';
    if (o.font) rpr += '<w:rFonts w:ascii="' + o.font + '" w:hAnsi="' + o.font + '" w:cs="' + o.font + '"/>';
    if (o.bold) rpr += '<w:b/><w:bCs/>';
    if (o.caps) rpr += '<w:caps/>';
    if (o.color) rpr += '<w:color w:val="' + o.color + '"/>';
    if (o.spacing) rpr += '<w:spacing w:val="' + o.spacing + '"/>';
    if (o.size) rpr += '<w:sz w:val="' + o.size + '"/><w:szCs w:val="' + o.size + '"/>';
    if (o.lang) rpr += '<w:lang w:val="' + o.lang + '"/>';
    return '<w:r>' + (rpr ? '<w:rPr>' + rpr + '</w:rPr>' : '') + '<w:t xml:space="preserve">' + xmlEsc_(text) + '</w:t></w:r>';
  }

  function para(runs, o) {
    o = o || {};
    let ppr = '';
    if (o.border) ppr += '<w:pBdr><w:bottom w:val="single" w:sz="8" w:space="4" w:color="111418"/></w:pBdr>';
    if (o.tab) ppr += '<w:tabs><w:tab w:val="right" w:pos="' + o.tab + '"/></w:tabs>';
    ppr += '<w:spacing w:before="0" w:after="' + (o.after || 0) + '" w:line="' + (o.line || 252) + '" w:lineRule="auto"/>';
    if (o.jc) ppr += '<w:jc w:val="' + o.jc + '"/>';
    return '<w:p><w:pPr>' + ppr + '</w:pPr>' + runs + '</w:p>';
  }

  function cell(width, content, o) {
    o = o || {};
    let tcpr = '<w:tcW w:w="' + width + '" w:type="dxa"/>';
    if (o.bottom) tcpr += '<w:tcBorders><w:bottom w:val="single" w:sz="' + o.bottom + '" w:space="0" w:color="111418"/></w:tcBorders>';
    if (o.fill) tcpr += '<w:shd w:val="clear" w:color="auto" w:fill="' + o.fill + '"/>';
    tcpr += '<w:vAlign w:val="top"/>';
    return '<w:tc><w:tcPr>' + tcpr + '</w:tcPr>' + content + '</w:tc>';
  }

  const label = { font: 'Arial', bold: true, caps: true, color: '374151', spacing: 10, size: 17 };
  let rows = '<w:tr><w:trPr><w:cantSplit/><w:tblHeader/></w:trPr>' +
    cell(cols[0], para(run('№', label), { jc: 'right' }), { bottom: 12 }) +
    cell(cols[1], para(run('Слово', label)), { bottom: 12 }) +
    cell(cols[2], para(run('Транскрипция', label)), { bottom: 12 }) +
    cell(cols[3], para(run('Перевод', label)), { bottom: 12 }) +
    cell(cols[4], para(run('Пример', label)), { bottom: 12 }) +
    '</w:tr>';

  words.forEach(function (w, i) {
    const fill = i % 2 === 1 ? 'F4F6F9' : '';
    const part = splitHit_(w.ex, w.hit);
    rows += '<w:tr><w:trPr><w:cantSplit/></w:trPr>' +
      cell(cols[0], para(run(String(i + 1), { color: '4B5563' }), { jc: 'right' }), { fill: fill }) +
      cell(cols[1], para(run(w.word, { bold: true, lang: 'en-GB' })), { fill: fill }) +
      cell(cols[2], para(run(w.ipa, { font: 'Arial', size: fs - 2, color: '1F2937', lang: 'en-GB' })), { fill: fill }) +
      cell(cols[3], para(run(w.tr, { lang: 'ru-RU' })), { fill: fill }) +
      cell(cols[4], para(run(part[0], { lang: 'en-GB' }) + run(part[1], { bold: true, lang: 'en-GB' }) + run(part[2], { lang: 'en-GB' })), { fill: fill }) +
      '</w:tr>';
  });

  const none = 'w:val="nil"';
  const table = '<w:tbl><w:tblPr>' +
    '<w:tblW w:w="' + TEXT_W + '" w:type="dxa"/>' +
    '<w:tblBorders><w:top ' + none + '/><w:left ' + none + '/>' +
    '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="C9CED8"/><w:right ' + none + '/>' +
    '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="C9CED8"/><w:insideV ' + none + '/></w:tblBorders>' +
    '<w:tblLayout w:type="fixed"/>' +
    '<w:tblCellMar><w:top w:w="70" w:type="dxa"/><w:left w:w="100" w:type="dxa"/><w:bottom w:w="70" w:type="dxa"/><w:right w:w="100" w:type="dxa"/></w:tblCellMar>' +
    '</w:tblPr><w:tblGrid>' + cols.map(function (c) { return '<w:gridCol w:w="' + c + '"/>'; }).join('') + '</w:tblGrid>' +
    rows + '</w:tbl>';

  const header = para(run(name, { bold: true, size: 36 }) + '<w:r><w:tab/></w:r>' + run(date, {}), { border: true, tab: TEXT_W, after: 200, line: 240 });
  const tail = '<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="20" w:lineRule="exact"/></w:pPr></w:p>';
  const sect = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="850" w:right="850" w:bottom="850" w:left="850" w:header="425" w:footer="425" w:gutter="0"/></w:sectPr>';

  const decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  return {
    '[Content_Types].xml': decl + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '</Types>',
    '_rels/.rels': decl + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
    'word/_rels/document.xml.rels': decl + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>',
    'word/styles.xml': decl + '<w:styles ' + NS + '><w:docDefaults>' +
      '<w:rPrDefault><w:rPr><w:rFonts w:ascii="Times New Roman" w:eastAsia="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>' +
      '<w:sz w:val="' + fs + '"/><w:szCs w:val="' + fs + '"/><w:lang w:val="ru-RU" w:eastAsia="en-US" w:bidi="ar-SA"/></w:rPr></w:rPrDefault>' +
      '<w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault>' +
      '</w:docDefaults>' +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
      '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:unhideWhenUsed/>' +
      '<w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>' +
      '</w:styles>',
    'word/document.xml': decl + '<w:document ' + NS + '><w:body>' + header + table + tail + sect + '</w:body></w:document>'
  };
}

function docx_(parts, fileName) {
  const blobs = Object.keys(parts).map(function (path) {
    return Utilities.newBlob('', 'application/xml', path).setDataFromString(parts[path], 'UTF-8');
  });
  return Utilities.zip(blobs, fileName).setContentType('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
}

// ───────────────────────── ручные функции ─────────────────────────

/** Запусти один раз: создаёт листы таблицы и настройки по умолчанию. */
function setup() {
  Object.keys(HEAD).forEach(function (name) { sheet_(name); });
  const have = {};
  rows_(SH.settings).forEach(function (r) { have[str_(r[0])] = 1; });
  appendRows_(SH.settings, DEFAULT_SETTINGS.filter(function (s) { return !have[s[0]]; }));
  const missing = ['BOT_TOKEN', 'ALLOWED_IDS', 'GEMINI_API_KEY'].filter(function (k) { return !prop_(k); });
  console.log('Листы готовы: ' + Object.keys(HEAD).join(', ') + '.');
  console.log(missing.length ? 'Не хватает свойств скрипта: ' + missing.join(', ') : 'Все нужные свойства скрипта на месте.');
  console.log('Модели по порядку: ' + modelChain_().map(function (m) { return m.model; }).join(' → '));
}

/** Проверка каждой модели по отдельности на трёх словах. Пришли журнал выполнения. */
function testModels() {
  const raw = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  const key = secret_('GEMINI_API_KEY');
  if (!key) {
    console.log('✗ Свойство GEMINI_API_KEY не найдено. Проверь имя свойства: GEMINI_API_KEY, заглавными, без пробелов.');
  } else {
    console.log('Ключ Gemini: ' + key.length + ' ' + plural_(key.length, ['символ', 'символа', 'символов']) + ', начинается с «' + key.slice(0, 4) + '».');
    if (raw !== key) console.log('В значении были пробелы, переносы или кавычки — код их убрал.');
  }
  const items = parseItems_('borrow, give up, bank (берег)');
  const prompt = buildPrompt_(items, 'B1', '');
  let badKey = false;
  modelChain_().forEach(function (m) {
    if (badKey && m.provider === 'gemini') {
      console.log('— ' + m.model + ': не проверяю, у всех моделей Gemini один ключ');
      return;
    }
    const t0 = Date.now();
    const r = tryModel_(m, prompt, items);
    const ms = Date.now() - t0;
    if (!r.items) {
      console.log('✗ ' + m.model + ' — ' + r.error + ' (' + ms + ' мс)');
      if (r.badKey && m.provider === 'gemini') {
        badKey = true;
        console.log('  → Ключ не принят. Скопируй ключ заново в AI Studio (Get API key → нажми на ключ → Copy) и вставь в свойство GEMINI_API_KEY целиком, без кавычек. Если не помогло — создай новый ключ (Create API key).');
      }
      return;
    }
    console.log('✓ ' + m.model + ' — ' + ms + ' мс\n' + r.items.map(function (x) {
      return '  ' + x.word + '  ' + x.ipa + '  — ' + x.tr + '\n    ' + x.ex + '   [жирным: ' + (x.hit || '—') + ']';
    }).join('\n'));
  });
}

const SAMPLE_WORDS = [
  ['ambitious', '/æmˈbɪʃəs/', 'амбициозный, целеустремлённый', 'Kate is very ambitious: she wants to become a doctor.', 'ambitious'],
  ['give up', '/ɡɪv ˈʌp/', 'сдаваться, бросать', "Don't give up if the test seems difficult.", 'give up'],
  ['borrow', '/ˈbɒrəʊ/', 'брать взаймы, одалживать', 'Can I borrow your pen for a minute?', 'borrow'],
  ['reliable', '/rɪˈlaɪəbl/', 'надёжный', 'My best friend is really reliable.', 'reliable'],
  ['look forward to', '/lʊk ˈfɔːwəd tuː/', 'с нетерпением ждать', "I'm looking forward to the summer holidays.", 'looking forward to'],
  ['tidy', '/ˈtaɪdi/', 'убирать; опрятный', 'Please tidy your room before dinner.', 'tidy'],
  ['advice', '/ədˈvaɪs/', 'совет', 'My teacher gave me good advice about the exam.', 'advice'],
  ['crowded', '/ˈkraʊdɪd/', 'переполненный, людный', 'The bus was so crowded that I had to stand.', 'crowded'],
  ['support', '/səˈpɔːt/', 'поддерживать; поддержка', 'My parents always support me in sports.', 'support'],
  ['disappointed', '/ˌdɪsəˈpɔɪntɪd/', 'разочарованный', 'I was disappointed when the match was cancelled.', 'disappointed'],
  ['take part in', '/teɪk ˈpɑːt ɪn/', 'участвовать в', 'Our class will take part in the science fair.', 'take part in'],
  ['confident', '/ˈkɒnfɪdənt/', 'уверенный', 'Tom feels confident before his English test.', 'confident'],
  ['waste', '/weɪst/', 'тратить впустую; отходы', "Don't waste water when you brush your teeth.", 'waste'],
  ['rarely', '/ˈreəli/', 'редко', 'I rarely watch TV on school days.', 'rarely'],
  ['set up', '/set ˈʌp/', 'создавать; настраивать', 'We set up a chess club at school.', 'set up']
];

/** Присылает пробный лист из 15 слов первому ID из ALLOWED_IDS. Таблицу не меняет. */
function testSheet() {
  const chat = allowedIds_()[0];
  if (!chat) throw new Error('Сначала впиши свой Telegram ID в свойство ALLOWED_IDS');
  const set = settings_();
  const pool = [];
  for (let i = 0; i < 40; i++) {
    const w = SAMPLE_WORDS[i % SAMPLE_WORDS.length];
    pool.push({ word: w[0], ipa: w[1], tr: w[2], ex: w[3], hit: w[4] });
  }
  const info = sheetInfo_(pool, set);
  const words = pool.slice(0, info.fit);
  sendSheet_('Пробный лист', 1, dateRu_(stamp_()), words, chat, fontSize_(set));
  console.log('Готово: полный пробный лист (' + words.length + ' слов, шрифт ' + fontSize_(set) + ' pt, заполнен на ~' + info.fill + '%) отправлен в чат с ботом. Проверь, что в Word он на ОДНОЙ странице.');
}

/**
 * Сравнение моделей на «трудных» словах: омографы (wind, lead, bow, read), сложная транскрипция,
 * фразовые глаголы. Каждая модель из MODELS получает один и тот же список. Пришли журнал.
 */
function compareModels() {
  const list = [
    ['wind', 'ветер'], ['wind', 'заводить (часы), наматывать'], ['lead', 'свинец'], ['bow', 'поклон'],
    ['read', 'прошедшее время'], ['thorough', ''], ['conscientious', ''], ['entrepreneur', ''],
    ['put up with', ''], ['get over', '']
  ];
  const items = list.map(function (x) {
    return { input: x[0] + (x[1] ? ' (' + x[1] + ')' : ''), base: x[0], hint: x[1], key: key_(x[0]) };
  });
  const prompt = buildPrompt_(items, 'B2', '');
  modelChain_().forEach(function (m) {
    const t0 = Date.now();
    const r = tryModel_(m, prompt, items);
    const ms = Date.now() - t0;
    if (!r.items) {
      console.log('✗ ' + m.model + ' — ' + r.error + ' (' + ms + ' мс)');
      return;
    }
    console.log('■ ' + m.model + ' — ' + ms + ' мс\n' + r.items.map(function (x, i) {
      return (i + 1) + '. ' + x.input + ' → ' + x.ipa + ' — ' + x.tr + '\n   ' + x.ex + (x.hit ? '' : '   [слово не найдено в примере]');
    }).join('\n'));
  });
}
