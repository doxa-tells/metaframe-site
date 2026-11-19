/**
 * METAFRAME: заявки МК → Google Таблица + Telegram + WhatsApp (Wazzapo API v2)
 * Заменяет ЦЕЛИКОМ код в Apps Script проекта таблицы заявок.
 *
 * ПОДКЛЮЧЕНИЕ (после вставки кода):
 * 1. Настройки проекта (шестерёнка) → Свойства скрипта → добавить:
 *      TG_TOKEN     — токен телеграм-бота (перенести из старого кода)
 *      TG_CHAT      — chat_id группы (перенести из старого кода)
 *      WZ_API_KEY   — ключ Wazzapo (кабинет → API → «Создать ключ», wz_live_...)
 *      WZ_PHONE_ID  — phone_number_id (функция wzListNumbers ниже покажет его в логе)
 *      WZ_TEMPLATE  — имя одобренного шаблона (после createTemplate: mk_zayavka_22_nov)
 *      POSTER_URL   — https://metaframe.ltd/assets/poster_22_nov.png
 * 2. Выполнить wzListNumbers() → взять phone_number_id из лога → вписать в WZ_PHONE_ID.
 * 3. Выполнить createTemplate() ОДИН РАЗ → шаблон уйдёт на проверку Meta.
 *    Статус: templateStatus() — ждать APPROVED (обычно минуты-часы).
 * 4. После APPROVED: тест на своём номере — testWhatsApp('7705XXXXXXX').
 * 5. Развернуть → Управление развёртываниями → карандаш → Новая версия → Развернуть.
 *    До шага 5 боевой поток работает по-старому (таблица + Telegram), WhatsApp не шлётся.
 *
 * В таблице добавить заголовки колонок H, I, J: whatsapp_status | whatsapp_time | whatsapp_error
 */

var P = PropertiesService.getScriptProperties();
var WZ_BASE = 'https://wazzapo.kz/api/v2/external';
// Telegram: берётся из Свойств скрипта, пока не заданы — используются текущие значения
var TG_TOKEN = P.getProperty('TG_TOKEN') || '8673079878:AAHpWbR11bdoPqGwr0R-aj39sThTQln3jIM';
var TG_CHAT  = P.getProperty('TG_CHAT')  || '-5380084978';

/* ===================== ПРИЁМ ЗАЯВКИ ===================== */
function doPost(e) {
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var p = e.parameter;
  var row = [new Date(), p.name||'', p.phone||'', p.work||'', p.exp||'', p.goal||'', p.deposit||'', 'pending', '', ''];
  s.appendRow(row);
  var rowIdx = s.getLastRow();
  sendTg(p);                      // Telegram-поток не меняется
  tryWhatsApp(p, s, rowIdx);      // WhatsApp; его ошибка не ломает заявку
  return ContentService.createTextOutput('ok');
}

/* ===================== TELEGRAM (без изменений) ===================== */
function sendTg(p) {
  try {
    var pay = /Да/.test(p.deposit || '');
    var msg = (pay ? '🟢 ЛИД — ГОТОВ ПЛАТИТЬ 9900' : '🔴 Лид — НЕ готов платить') + '\n\n' +
      'Имя: ' + (p.name||'—') + '\nТелефон: ' + (p.phone||'—') + '\nЗанятие: ' + (p.work||'—') +
      '\nОпыт: ' + (p.exp||'—') + '\nЦель: ' + (p.goal||'—');
    UrlFetchApp.fetch('https://api.telegram.org/bot' + TG_TOKEN + '/sendMessage',
      {method:'post', payload:{chat_id: TG_CHAT, text: msg}, muteHttpExceptions:true});
  } catch(err) {}
}

/* ===================== WHATSAPP ===================== */
function normPhone(raw) {
  var d = String(raw||'').replace(/\D/g,'');
  if (d.charAt(0) === '8') d = '7' + d.slice(1);
  if (d.length === 10) d = '7' + d;
  return d;
}

function setWa(s, rowIdx, status, error) {
  s.getRange(rowIdx, 8).setValue(status);
  s.getRange(rowIdx, 9).setValue(new Date());
  s.getRange(rowIdx, 10).setValue(error || '');
}

function tryWhatsApp(p, s, rowIdx) {
  try {
    var key = P.getProperty('WZ_API_KEY');
    if (!key || !P.getProperty('WZ_TEMPLATE')) { setWa(s,rowIdx,'skipped_not_configured'); return; }
    var phone = normPhone(p.phone);
    if (phone.length < 11) { setWa(s,rowIdx,'failed','bad phone'); return; }

    // 1) дубль заявки: этому номеру уже отправляли из формы
    var vals = s.getDataRange().getValues();
    for (var i = 1; i < vals.length; i++) {
      if (i+1 === rowIdx) continue;
      if (normPhone(vals[i][2]) === phone && vals[i][7] === 'sent') { setWa(s,rowIdx,'skipped_duplicate'); return; }
    }

    // 2) человек уже писал нам в WhatsApp сам (его встретил сценарий Wazzapo)
    if (hasIncomingChat(phone)) { setWa(s,rowIdx,'skipped_existing_chat'); return; }

    // 3) отправка шаблона; Idempotency-Key по номеру — повтор не создаст дубль
    var body = {
      phone_number_id: P.getProperty('WZ_PHONE_ID'),
      to: phone,
      type: 'template',
      template: {
        name: P.getProperty('WZ_TEMPLATE'),
        language: { code: 'ru' },
        components: [
          { type: 'header', parameters: [{ type: 'image', image: { link: P.getProperty('POSTER_URL') } }] },
          { type: 'body',   parameters: [{ type: 'text', text: (p.name||'друг').slice(0,40) }] }
        ]
      }
    };
    var res = UrlFetchApp.fetch(WZ_BASE + '/messages/send', {
      method: 'post', contentType: 'application/json',
      headers: { 'X-API-Key': key, 'Idempotency-Key': 'mk22-' + phone },
      payload: JSON.stringify(body), muteHttpExceptions: true
    });
    var code = res.getResponseCode();
    if (code >= 200 && code < 300) { setWa(s,rowIdx,'sent'); }
    else {
      var errCode = '';
      try { errCode = JSON.parse(res.getContentText()).error.code; } catch(e2) {}
      setWa(s,rowIdx,'failed', 'HTTP ' + code + (errCode ? ' ' + errCode : ''));  // без токенов и ПД
    }
  } catch(err) { try { setWa(s,rowIdx,'failed', String(err).slice(0,120)); } catch(e3) {} }
}

/**
 * Писал ли номер нам сам: контакт в Wazzapo с source == 'whatsapp'
 * (создаётся, когда человек сам написал и его встретил сценарий).
 * Ограничение API: прямой выборки диалога по номеру нет; контакт, созданный
 * рассылкой или вручную, имеет другой source и фильтр его пропустит — это ок:
 * в худшем случае человек получит афишу повторно, но не из-за формы-дубля.
 */
function hasIncomingChat(phone) {
  try {
    var res = UrlFetchApp.fetch(WZ_BASE + '/contacts?search=' + encodeURIComponent(phone), {
      headers: { 'X-API-Key': P.getProperty('WZ_API_KEY') }, muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) return false;
    var data = JSON.parse(res.getContentText()).data || [];
    for (var i = 0; i < data.length; i++) {
      if (normPhone(data[i].phone) === phone && data[i].source === 'whatsapp') return true;
    }
  } catch(e) {}
  return false;
}

/* ===================== ОДНОРАЗОВЫЕ УТИЛИТЫ ===================== */
// Шаг 2: показать подключённые номера → phone_number_id в лог
function wzListNumbers() {
  var res = UrlFetchApp.fetch(WZ_BASE + '/phone-numbers', { headers: { 'X-API-Key': P.getProperty('WZ_API_KEY') } });
  Logger.log(res.getContentText());
}

// Шаг 3: создать шаблон (ОДИН РАЗ). После — ждать одобрения Meta.
function createTemplate() {
  var body = {
    phone_number_id: P.getProperty('WZ_PHONE_ID'),
    name: 'mk_zayavka_22_nov',
    language: 'ru',
    category: 'MARKETING',
    components: [
      { type: 'HEADER', format: 'IMAGE', example: { media_url: P.getProperty('POSTER_URL') } },
      { type: 'BODY',
        text: 'Привет, {{1}}! Это METAFRAME — получили твою заявку на последний офлайн мастер-класс по AI-контенту.\n\n' +
              '📅 22 ноября 2026, воскресенье, 14:00–16:00\n📍 Nazarbayev University, Астана\n🎫 Участие — 9 900 ₸\n' +
              '🎤 Ведущий — Дархан Телманов, ex. Generative Director Higgsfield, победитель Higgsfield World Championship 2025\n\n' +
              'Оплата переводом на Kaspi: 8 (705) 992-64-40, получатель Азамат У. После оплаты пришли сюда чек — внесём тебя в список участников.\n\n' +
              'Instagram: instagram.com/metaframe_ai',
        example: { body_text: [['Айдар']] } },
      { type: 'BUTTONS', buttons: [
        { type: 'URL', text: 'Страница МК', url: 'https://metaframe.ltd/mk.html' },
        { type: 'URL', text: 'Сайт METAFRAME', url: 'https://metaframe.ltd' }
      ]}
    ]
  };
  var res = UrlFetchApp.fetch(WZ_BASE + '/templates', {
    method: 'post', contentType: 'application/json',
    headers: { 'X-API-Key': P.getProperty('WZ_API_KEY') },
    payload: JSON.stringify(body), muteHttpExceptions: true
  });
  Logger.log(res.getContentText()); // сохраните template_id из ответа
}

// Статус проверки шаблона Meta
function templateStatus() {
  var res = UrlFetchApp.fetch(WZ_BASE + '/templates?phone_number_id=' + P.getProperty('WZ_PHONE_ID') + '&name=mk_zayavka', {
    headers: { 'X-API-Key': P.getProperty('WZ_API_KEY') }, muteHttpExceptions: true
  });
  Logger.log(res.getContentText());
}

// Шаг 4: безопасный тест на СВОЁМ номере (одна отправка, без формы)
function testWhatsApp(myPhone) {
  myPhone = myPhone || '77XXXXXXXXX'; // ← свой номер
  var s = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  s.appendRow([new Date(), 'ТЕСТ WA', myPhone, 'тест', 'тест', 'тест', 'тест', 'pending', '', '']);
  tryWhatsApp({ name: 'Тест', phone: myPhone }, s, s.getLastRow());
}
