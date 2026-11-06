/**
 * Приём заявок для приложения «ОпенМир · Предприниматель».
 * Сюда приходят заявки из форм Tilda и из ChatPlace (Instagram, Telegram).
 * Приложение забирает новые строки и создаёт из них клиентов в CRM.
 *
 * КАК ПОДКЛЮЧИТЬ (один раз, около 10 минут):
 * 1. Создайте новую Google-таблицу (название любое, например «Заявки»). Таблицу не открывайте для всех: доступ только вам.
 * 2. В таблице: Расширения → Apps Script. Удалите всё, что там есть, и вставьте этот файл целиком. Сохраните (значок дискеты).
 * 3. Если листов «Заявки» и «Настройки» в таблице ещё нет: выберите вверху функцию «настроить» и нажмите «Выполнить»,
 *    разрешите доступ (Google спросит один раз). В листе «Настройки» ячейка B1 — ваш ключ доступа.
 * 4. Развернуть → Новое развёртывание → тип «Веб-приложение».
 *    «Выполнять от имени»: я. «У кого есть доступ»: все. Нажмите «Развернуть», скопируйте «URL веб-приложения»
 *    (заканчивается на /exec).
 * 5. В приложении: CRM → Ещё → Настройки → «Заявки из Tilda и мессенджеров». Вставьте URL и ключ, нажмите «Сохранить» и «Проверить».
 * 6. Tilda: Настройки сайта → Формы → Webhook (или в самой форме: Контент → Куда отправлять данные → Webhook).
 *    Адрес: ваш URL + ?k=КЛЮЧ&src=tilda  (пример: https://script.google.com/macros/s/XXXX/exec?k=abc123&src=tilda).
 *    В полях формы переменные оставьте как есть: Name, Phone, Email, Textarea (или Comment).
 * 7. ChatPlace: в сценарии добавьте действие «HTTP-запрос» (доступно на тарифе Pro), метод POST, формат JSON.
 *    Адрес: ваш URL + ?k=КЛЮЧ. Для Instagram в теле укажите "src":"instagram", для Telegram "src":"telegram".
 *    Готовое тело запроса: см. комментарий внизу файла.
 *
 * Если вы измените этот код, нужно заново: Развернуть → Управление развёртываниями → карандаш → Версия «Новая» → Развернуть.
 * Адрес при этом не меняется.
 *
 * Столбец «Все поля» заполняется для каждой новой заявки: туда попадают все поля формы Tilda (имя поля: значение).
 * Названия полей берутся из переменных формы в Tilda, поэтому давайте полям понятные названия.
 *
 * Если в столбце «Телефон» видно #ERROR!: после обновления кода новые заявки будут записываться правильно.
 * Старые строки поправьте вручную: выделите столбец E, Формат → Числа → Обычный текст, впишите номер заново (он есть в столбце «Исходные данные»).
 */

var ЛИСТ = 'Заявки';
var ШАПКА = ['id', 'Получено', 'Источник', 'Имя', 'Телефон', 'Почта', 'Instagram', 'Telegram', 'Сообщение', 'Страница или форма', 'UTM', 'Исходные данные', 'Все поля'];

function настроить() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(ЛИСТ) || ss.insertSheet(ЛИСТ);
  if (sh.getLastRow() === 0) {
    sh.appendRow(ШАПКА);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, ШАПКА.length).setFontWeight('bold');
  }
  var st = ss.getSheetByName('Настройки') || ss.insertSheet('Настройки');
  var key = String(st.getRange('B1').getValue() || '').trim() || Utilities.getUuid().replace(/-/g, '');
  st.clear();
  st.getRange('A1:B1').setValues([['Ключ доступа (вставьте в приложение)', key]]);
  st.getRange('A3').setValue('Адрес веб-приложения (/exec) появится после развёртывания: Развернуть → Новое развёртывание → Веб-приложение.');
  st.getRange('A4').setValue('Не показывайте ключ и таблицу посторонним: в ней персональные данные людей.');
  st.setColumnWidth(1, 330);
  st.setColumnWidth(2, 330);
}

/* в листе должно быть столько столбцов, сколько в шапке (для таблиц, созданных старой версией) */
function столбцы_(sh) {
  if (sh.getMaxColumns() < ШАПКА.length) sh.insertColumnsAfter(sh.getMaxColumns(), ШАПКА.length - sh.getMaxColumns());
  if (!sh.getRange(1, ШАПКА.length).getValue()) sh.getRange(1, ШАПКА.length).setValue(ШАПКА[ШАПКА.length - 1]).setFontWeight('bold');
}

function ответ_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ключ_() {
  var st = SpreadsheetApp.getActive().getSheetByName('Настройки');
  var k = st ? String(st.getRange('B1').getValue() || '').trim() : '';
  return k || PropertiesService.getScriptProperties().getProperty('KEY');
}

function чисто_(v) {
  if (v === undefined || v === null) return '';
  var s = String(v).trim();
  if (/^\{\{.*\}\}$/.test(s)) return '';
  return s;
}

function первое_(d, шаблоны) {
  var ключи = Object.keys(d);
  for (var i = 0; i < шаблоны.length; i++) {
    for (var j = 0; j < ключи.length; j++) {
      if (шаблоны[i].test(ключи[j])) {
        var v = чисто_(d[ключи[j]]);
        if (v) return v;
      }
    }
  }
  return '';
}

function doPost(e) {
  var p = (e && e.parameter) || {};
  if (!ключ_() || p.k !== ключ_()) return ответ_({ ok: false, error: 'key' });
  var d = {};
  try {
    if (e.postData && /json/i.test(e.postData.type || '')) d = JSON.parse(e.postData.contents);
  } catch (x) {}
  Object.keys(p).forEach(function (k) { if (d[k] === undefined) d[k] = p[k]; });
  if (чисто_(d.test) === 'test') return ContentService.createTextOutput('ok');

  var src = чисто_(d.src) || ((d.tranid || d.formid || d.formname) ? 'tilda' : 'other');
  var имя = первое_(d, [/^(fullname|name|имя|fio|фио)$/i, /name|имя|фио/i]);
  var тел = первое_(d, [/^(phone|телефон|tel)$/i, /client_phone/i, /phone|тел/i]);
  var почта = первое_(d, [/^(email|e-mail|почта)$/i, /mail|почт/i]);
  var инст = src === 'instagram' ? первое_(d, [/^(username|ig|instagram)$/i]) : первое_(d, [/^(ig|instagram)$/i]);
  var тг = src === 'telegram' ? первое_(d, [/^(username|tg|telegram)$/i]) : первое_(d, [/^(tg|telegram)$/i]);
  var сообщение = [первое_(d, [/^(msg|message|textarea|comment|text|сообщение|вопрос)/i, /вопрос_клиента/i, /textarea|comment|message|сообщ/i]),
    первое_(d, [/^problem_description$/i]), первое_(d, [/^interest$/i]), первое_(d, [/^service_type$/i]), первое_(d, [/^tax_regime$/i])]
    .filter(function (x) { return x; }).join('. ');
  var страница = первое_(d, [/^(page|pageurl|formname|formid|chatlink)$/i]);
  var utm = ['utm_source', 'utm_medium', 'utm_campaign', 'utm_content', 'utm_term'].map(function (k) { return d[k] ? k.replace('utm_', '') + '=' + d[k] : ''; }).filter(function (x) { return x; }).join('; ');
  var служебные = /^(k|src|test|tranid|formid|id|cookies|formname)$/i;
  var все = Object.keys(d).filter(function (k) { return !служебные.test(k) && чисто_(d[k]); })
    .map(function (k) { return k + ': ' + чисто_(d[k]).slice(0, 1000); }).join('\n').slice(0, 8000);
  var ид = чисто_(d.tranid) ? 'tilda:' + чисто_(d.tranid) : (чисто_(d.id) ? src + ':' + чисто_(d.id) : src + ':' + Utilities.getUuid());

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(ЛИСТ);
    if (!sh) { настроить(); sh = SpreadsheetApp.getActive().getSheetByName(ЛИСТ); }
    столбцы_(sh);
    var n = sh.getLastRow();
    if (n > 1) {
      var from = Math.max(2, n - 300);
      var ids = sh.getRange(from, 1, n - from + 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (ids[i][0] === ид && чисто_(d.id) && src !== 'tilda') {
          /* повторное обращение того же человека из мессенджера: пишем как новую строку с новым id */
          ид = ид + '#' + Utilities.formatDate(new Date(), 'Europe/Minsk', 'yyMMddHHmm');
          break;
        }
        if (ids[i][0] === ид) return ответ_({ ok: true, dup: true });
      }
    }
    var raw = JSON.stringify(d);
    /* всё пишем как текст, иначе телефон вида +375... таблица считает формулой и показывает #ERROR! */
    var строка = sh.getLastRow() + 1;
    var диап = sh.getRange(строка, 1, 1, ШАПКА.length);
    диап.setNumberFormat('@');
    диап.setValues([[ид, Utilities.formatDate(new Date(), 'Europe/Minsk', 'yyyy-MM-dd HH:mm:ss'), src, имя, тел, почта, инст, тг, сообщение, страница, utm, raw.slice(0, 1500), все]]);
  } finally {
    lock.releaseLock();
  }
  return ответ_({ ok: true });
}

/* если в ячейке телефона ошибка (#ERROR!), берём номер из столбца «Исходные данные» */
function тел_(v) {
  var t = String(v[4]);
  if (/^[#=]/.test(t)) {
    var m = String(v[11]).match(/"(?:phone|Phone|tel|Tel|client_phone|Телефон)"\s*:\s*"([^"]+)"/);
    if (m) return m[1];
  }
  return t;
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!ключ_() || p.k !== ключ_()) return ответ_({ ok: false, error: 'key' });
  var sh = SpreadsheetApp.getActive().getSheetByName(ЛИСТ);
  if (!sh) return ответ_({ ok: false, error: 'sheet' });
  var last = sh.getLastRow();
  if (p.action === 'ping') return ответ_({ ok: true, last: last });
  столбцы_(sh);
  var since = Math.max(1, parseInt(p.since, 10) || 1);
  if (since > last) since = last;
  var rows = [];
  if (last > since) {
    var cnt = Math.min(200, last - since);
    var vals = sh.getRange(since + 1, 1, cnt, ШАПКА.length).getValues();
    for (var i = 0; i < vals.length; i++) {
      var v = vals[i];
      rows.push({ row: since + 1 + i, id: String(v[0]), ts: String(v[1]), src: String(v[2]), name: String(v[3]), phone: тел_(v), email: String(v[5]), ig: String(v[6]), tg: String(v[7]), msg: String(v[8]), page: String(v[9]), utm: String(v[10]), all: String(v[12] || '') });
    }
    return ответ_({ ok: true, rows: rows, last: since + cnt, total: last });
  }
  return ответ_({ ok: true, rows: [], last: last, total: last });
}

/*
 ТЕЛО HTTP-ЗАПРОСА ДЛЯ CHATPLACE (метод POST, формат JSON). Для Instagram:

 {"src":"instagram","id":"{{clientId}}","name":"{{fullName}}","username":"{{username}}","phone":"{{client_phone}}","email":"{{email}}","msg":"{{вопрос_клиента}}","problem_description":"{{problem_description}}","interest":"{{interest}}","service_type":"{{service_type}}","tax_regime":"{{tax_regime}}","chatLink":"{{chatLink}}"}

 Для Telegram замените "src":"instagram" на "src":"telegram".
 Если какой-то переменной у человека ещё нет, поле придёт пустым, это нормально.
 Запрос лучше ставить в сценарий после того, как человек оставил телефон или ответил на вопросы.
*/
