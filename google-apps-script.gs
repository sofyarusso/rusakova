/**
 * Приём заявок с сайта → Google Таблица + почта + Telegram.
 * Вставь этот код в Расширения → Apps Script своей таблицы (см. README.md).
 *
 * Настройки (Project Settings → Script Properties):
 *   EMAIL       — куда слать письма (если пусто — на почту владельца скрипта)
 *   TG_TOKEN    — токен бота от @BotFather
 *   TG_CHAT_ID  — твой chat_id (узнать: напиши боту, открой https://api.telegram.org/bot<TOKEN>/getUpdates)
 */
const FIELDS = { name: 80, contact: 120, service: 200, budget: 60, message: 8000, page: 300 }; // message длинный — туда приходит весь бриф

function doPost(e) {
  try {
    return json_(handle_(e && e.parameter ? e.parameter : {}));
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: 'server' });
  }
}

function handle_(p) {
  if (p.website) return { ok: true }; // honeypot: бот заполнил скрытое поле — молча игнорируем

  const d = {};
  Object.keys(FIELDS).forEach((k) => { d[k] = String(p[k] || '').trim().slice(0, FIELDS[k]); });
  if (!d.name || !d.contact) return { ok: false, error: 'required' };
  if (p.consent !== 'да') return { ok: false, error: 'consent' };

  // антиспам: не больше 5 заявок в минуту на весь сайт
  const cache = CacheService.getScriptCache();
  const n = Number(cache.get('rate') || 0);
  if (n >= 5) return { ok: false, error: 'rate' };
  cache.put('rate', String(n + 1), 60);

  const when = Utilities.formatDate(new Date(), 'Europe/Moscow', 'dd.MM.yyyy HH:mm');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) sheet.appendRow(['Дата', 'Имя', 'Контакт', 'Услуга', 'Бюджет', 'Сообщение', 'Страница']);
    // апостроф — чтобы таблица не приняла текст за формулу
    sheet.appendRow([when, d.name, d.contact, d.service, d.budget, d.message, d.page].map((v) => (/^[=+\-@]/.test(v) ? "'" + v : v)));
  } finally {
    lock.releaseLock();
  }

  const text = [
    '🖤 Новая заявка с сайта',
    'Имя: ' + d.name,
    'Контакт: ' + d.contact,
    'Услуга: ' + (d.service || '—'),
    'Бюджет: ' + (d.budget || '—'),
    'Сообщение: ' + (d.message || '—'),
    when,
  ].join('\n');

  const props = PropertiesService.getScriptProperties();
  // почта и Telegram не должны ронять заявку: она уже сохранена в таблице
  try {
    MailApp.sendEmail(props.getProperty('EMAIL') || Session.getEffectiveUser().getEmail(), 'Заявка с сайта: ' + d.name, text);
  } catch (err) { console.error('mail', err); }
  const token = props.getProperty('TG_TOKEN'), chat = props.getProperty('TG_CHAT_ID');
  if (token && chat) {
    try {
      UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/sendMessage', {
        method: 'post', contentType: 'application/json', muteHttpExceptions: true,
        payload: JSON.stringify({ chat_id: chat, text: text.slice(0, 4000) }), // без parse_mode; Telegram режет сообщения > 4096 символов — полный бриф есть в почте и таблице
      });
    } catch (err) { console.error('tg', err); }
  }
  return { ok: true };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** Запусти вручную из редактора: проверяет валидацию и шлёт тестовую заявку. */
function testSubmit() {
  console.assert(handle_({ website: 'spam' }).ok === true, 'honeypot');
  console.assert(handle_({ name: '', contact: 'x', consent: 'да' }).error === 'required', 'required');
  console.assert(handle_({ name: 'a', contact: 'x' }).error === 'consent', 'consent');
  const r = handle_({ name: 'Тест', contact: '@test', service: 'Логотип', message: '=1+1 проверка', consent: 'да', page: 'test' });
  console.log('test result', r);
}
