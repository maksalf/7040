/********************************
 * 24_crm_api.gs
 ********************************/

function CRM_jsonResponse24_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  var ping = e && e.parameter ? e.parameter.ping : '';
  if (String(ping) === '1') {
    return CRM_jsonResponse24_({ ok: true, message: 'pong', ts: new Date().toISOString() });
  }

  return HtmlService
    .createHtmlOutputFromFile('mobile_crm')
    .setTitle('CRM Mobile');
}

function doPost(e) {
  try {
    var raw = e && e.postData ? e.postData.contents : '{}';
    var body = JSON.parse(raw);
    var action = String(body.action || '');
    var data = body.data || {};

    if (action === 'search') {
      return CRM_jsonResponse24_({ ok: true, data: { items: CRM_mobile_searchRows(data.query || '') } });
    }

    if (action === 'getCard') {
      return CRM_jsonResponse24_({ ok: true, data: CRM_mobile_getCard(Number(data.row || 0)) });
    }

    if (action === 'saveQuick') {
      return CRM_jsonResponse24_({ ok: true, data: { message: CRM_mobile_saveQuick(data || {}) } });
    }

    if (action === 'report112') {
      return CRM_jsonResponse24_({ ok: true, data: CRM_mobile_getReport112() });
    }

    return CRM_jsonResponse24_({ ok: false, error: 'Unknown action' });
  } catch (err) {
    return CRM_jsonResponse24_({ ok: false, error: (err && err.message) ? err.message : String(err) });
  }
}
