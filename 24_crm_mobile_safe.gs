/********************************
 * SAFE MOBILE API
 ********************************/

function CRM_mobile_getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Поранені');
  if (!sh) throw new Error('Аркуш "Поранені" не знайдено');
  return sh;
}

function CRM_mobile_norm_(v) {
  return String(v || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function CRM_mobile_headers_(sheet) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(2, 1, 1, lastCol).getDisplayValues()[0];
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var k = CRM_mobile_norm_(headers[i]);
    if (k) map[k] = i + 1; // 1-based
  }
  return { headers: headers, map: map, lastCol: lastCol };
}

function CRM_mobile_findCol_(headerMap, names) {
  for (var i = 0; i < names.length; i++) {
    var key = CRM_mobile_norm_(names[i]);
    if (headerMap[key]) return headerMap[key];
  }
  return null;
}

/** Пошук */
function CRM_mobile_searchRows(query) {
  var text = String(query || '').trim().toLowerCase();
  if (!text) return [];

  var sheet = CRM_mobile_getSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 3) return [];

  var h = CRM_mobile_headers_(sheet);
  var vals = sheet.getRange(3, 1, lastRow - 2, h.lastCol).getDisplayValues();

  var colPib = CRM_mobile_findCol_(h.map, ['ПІБ В/с', 'ПІБ в/с', 'ПІБ']);
  var colPhone = CRM_mobile_findCol_(h.map, ['Номер телефону', 'Телефон', 'контактний телефон']);
  var colStatus = CRM_mobile_findCol_(h.map, ['Статус супроводження', 'Статус', 'Стан']);
  var colCity = CRM_mobile_findCol_(h.map, ['Населений пункт', 'Місто']);
  var colHospital = CRM_mobile_findCol_(h.map, ['Назва медичного закладу', 'Лікарня']);

  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var rowVals = vals[i];
    if (rowVals.join(' | ').toLowerCase().indexOf(text) === -1) continue;

    out.push({
      row: i + 3,
      pib: colPib ? (rowVals[colPib - 1] || 'Без ПІБ') : 'Без ПІБ',
      phone: colPhone ? (rowVals[colPhone - 1] || '') : '',
      status: colStatus ? (rowVals[colStatus - 1] || '') : '',
      city: colCity ? (rowVals[colCity - 1] || '') : '',
      hospital: colHospital ? (rowVals[colHospital - 1] || '') : ''
    });

    if (out.length >= 50) break;
  }

  return out;
}

/** Опції документів */
function CRM_mobile_getDocOptions() {
  var sheet = CRM_mobile_getSheet_();
  var h = CRM_mobile_headers_(sheet);
  var lastRow = sheet.getLastRow();

  function uniqFrom(names) {
    var c = CRM_mobile_findCol_(h.map, names);
    if (!c || lastRow < 3) return [];
    var arr = sheet.getRange(3, c, lastRow - 2, 1).getDisplayValues().flat()
      .map(function(v) { return String(v || '').trim(); })
      .filter(function(v) { return !!v; });
    return Array.from(new Set(arr)).sort();
  }

  return {
    f100: uniqFrom(['Первинна медична карта (ф. 100) (Ф001)', 'Первинна медична карта (ф. 100)', 'Ф100']),
    f5: uniqFrom(['Довідка 5 про обставини поранення (травми) Ф5', 'Ф5']),
    f6: uniqFrom(['Довідка 6 про безпосередню участь у бойових діях (ф. 6)', 'Ф6']),
    ubd: uniqFrom(['Посвідчення учасника бойових дій', 'УБД'])
  };
}

/** Картка */
function CRM_mobile_getCard(row) {
  return CRM_getInitialSidebarData24('Поранені', Number(row));
}

/** Швидке збереження */
function CRM_mobile_saveQuick(payload) {
  payload = payload || {};
  payload.sheetName = 'Поранені';
  payload.row = Number(payload.row || 0);
  return CRM_saveCardData24(payload);
}

/** Звіт 112 */
function CRM_mobile_getReport112() {
  var sh = CRM_mobile_getSheet_();
  var startRow = 4;
  var lastRow = sh.getLastRow();
  if (lastRow < startRow) return {};

  var lastCol = Math.max(sh.getLastColumn(), 43); // AQ
  var rows = sh.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getDisplayValues();

  function col(letters) {
    var n = 0;
    for (var i = 0; i < letters.length; i++) n = n * 26 + (letters.charCodeAt(i) - 64);
    return n - 1;
  }

  var B = col('B'), C = col('C'), I = col('I'), J = col('J'),
      K = col('K'), O = col('O'), R = col('R'), U = col('U'),
      Y = col('Y'), AA = col('AA'), AJ = col('AJ'), AK = col('AK'), AQ = col('AQ');

  function norm(s) { return String(s || '').trim().toLowerCase(); }
  function eq(v, x) { return norm(v) === norm(x); }
  function inArr(v, arr) {
    var nv = norm(v);
    for (var i = 0; i < arr.length; i++) if (norm(arr[i]) === nv) return true;
    return false;
  }
  function notEmpty(v) { return String(v || '').trim() !== ''; }

  function parseDate(v) {
    var s = String(v || '').trim();
    if (!s) return null;
    var m = s.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    var d = new Date(s);
    return isNaN(d.getTime()) ? null : d;
  }
  function dateGte(v, d) {
    var x = parseDate(v);
    if (!x) return false;
    x.setHours(0, 0, 0, 0);
    return x >= d;
  }

  var today = new Date();
  today.setHours(0, 0, 0, 0);
  var d7 = new Date(today); d7.setDate(d7.getDate() - 7);
  var d30 = new Date(today); d30.setDate(d30.getDate() - 30);

  function count(fn) {
    var n = 0;
    for (var i = 0; i < rows.length; i++) if (fn(rows[i])) n++;
    return n;
  }

  var totalOnRecord = count(function(r){ return notEmpty(r[C]); });
  var totalLight = count(function(r){ return eq(r[I], 'легке'); });
  var totalHeavy = count(function(r){ return eq(r[I], 'тяжке'); });

  var weekLight = count(function(r){
    return inArr(r[B], ['В роботі','Новий','Не вдалось звʼязатися']) &&
      eq(r[I], 'легке') && dateGte(r[J], d7);
  });
  var weekHeavy = count(function(r){ return eq(r[B], 'В роботі') && eq(r[I], 'тяжке') && dateGte(r[J], d7); });
  var weekUnknown = count(function(r){ return eq(r[B], 'В роботі') && eq(r[I], 'інформація уточнюється') && dateGte(r[J], d7); });

  var comm30 = count(function(r){ return eq(r[B], 'В роботі') && dateGte(r[AK], d30); });
  var visits30 = count(function(r){ return eq(r[B], 'В роботі') && dateGte(r[AJ], d30); });
  var new30Light = count(function(r){ return eq(r[B], 'В роботі') && eq(r[I], 'легке') && dateGte(r[J], d30); });
  var new30Heavy = count(function(r){ return eq(r[B], 'В роботі') && eq(r[I], 'тяжке') && dateGte(r[J], d30); });
  var new30Unknown = count(function(r){ return eq(r[B], 'В роботі') && eq(r[I], 'інформація уточнюється') && dateGte(r[J], d30); });

  var statusDone = count(function(r){ return eq(r[B], 'Завершено'); });
  var statusWork = count(function(r){ return inArr(r[B], ['В роботі','Новий']); });
  var statusNoContact = count(function(r){ return eq(r[B], 'Не вдалось звʼязатися'); });
  var statusNew = count(function(r){ return eq(r[B], 'Новий'); });

  var inWorkLight = count(function(r){ return eq(r[B], 'В роботі') && eq(r[I], 'легке'); });
  var inWorkHeavy = count(function(r){ return eq(r[B], 'В роботі') && eq(r[I], 'тяжке'); });
  var inWorkUnknown = count(function(r){ return eq(r[B], 'В роботі') && eq(r[I], 'інформація уточнюється'); });
  var inWorkTotal = inWorkLight + inWorkHeavy + inWorkUnknown;
  var inWorkNot = totalOnRecord - inWorkTotal;

  var kyiv = count(function(r){ return eq(r[K], 'на лікуванні в межах України') && eq(r[R], 'Україна') && eq(r[O], 'Київ'); });

  var f5Yes = count(function(r){ return eq(r[AA], 'наявна'); });
  var f5No = count(function(r){ return eq(r[AA], 'відсутня'); });
  var amputation = count(function(r){ return eq(r[U], 'так'); });
  var reward = count(function(r){ return eq(r[Y], 'отримує виплати'); });

  var uaLight = count(function(r){ return eq(r[K], 'на лікуванні в межах України') && eq(r[I], 'легке'); });
  var uaHeavy = count(function(r){ return eq(r[K], 'на лікуванні в межах України') && eq(r[I], 'тяжке'); });
  var uaUnknown = count(function(r){ return eq(r[K], 'на лікуванні в межах України') && eq(r[I], 'інформація уточнюється'); });

  var treatmentUa = count(function(r){ return eq(r[K], 'на лікуванні в межах України'); });
  var rehabUa = count(function(r){ return eq(r[K], 'на реабілітації в межах України'); });
  var treatmentAbroad = count(function(r){ return eq(r[K], 'на лікуванні за кордоном'); });
  var rehabAbroad = count(function(r){ return eq(r[K], 'на реабілітації за кордоном'); });
  var dead = count(function(r){ return eq(r[K], 'помер'); });
  var finished = count(function(r){ return eq(r[K], 'закінчив лікування та вибув до частини'); });
  var inMedicalTotal = treatmentUa + rehabUa + treatmentAbroad + rehabAbroad;
  var healthLeave = count(function(r){ return eq(r[AQ], 'Відпустка за станом здоров’я'); });

  return {
    total: { onRecord: totalOnRecord, light: totalLight, heavy: totalHeavy },
    week: { light: weekLight, heavy: weekHeavy, unknown: weekUnknown },
    m30: { comm: comm30, visit: visits30, newLight: new30Light, newHeavy: new30Heavy, newUnknown: new30Unknown },
    status: { done: statusDone, work: statusWork, noContact: statusNoContact, fresh: statusNew },
    inWork: { light: inWorkLight, heavy: inWorkHeavy, unknown: inWorkUnknown, total: inWorkTotal, notWork: inWorkNot },
    location: { kyiv: kyiv },
    docs: { f5Yes: f5Yes, f5No: f5No, amputation: amputation, reward: reward },
    rehab: { uaLight: uaLight, uaHeavy: uaHeavy, uaUnknown: uaUnknown, treatmentUa: treatmentUa, rehabUa: rehabUa, treatmentAbroad: treatmentAbroad, rehabAbroad: rehabAbroad, dead: dead, finished: finished, inMedicalTotal: inMedicalTotal, healthLeave: healthLeave }
  };
}
