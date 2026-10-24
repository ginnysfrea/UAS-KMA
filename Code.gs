function doPost(e) {
  var d = JSON.parse(e.postData.contents);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  var ss = SpreadsheetApp.getActiveSpreadsheet(), t = new Date();
  var r = ss.getSheetByName('Ringkasan') || ss.insertSheet('Ringkasan');
  var j = ss.getSheetByName('Jawaban') || ss.insertSheet('Jawaban');
  if (!r.getLastRow()) r.appendRow(['Waktu','Nama','NIM','Skor PG (maks 50)','Skor Essay (maks 50)','Total (maks 100)']);
  if (!j.getLastRow()) {
    var h = ['Waktu','Nama','NIM'], i;
    for (i = 1; i <= 20; i++) h.push('PG ' + i);
    for (i = 1; i <= 10; i++) h.push('Essay ' + i);
    j.appendRow(h);
  }
  r.appendRow([t, d.nama, "'" + d.nim, d.pg, d.essay, d.total]);
  j.appendRow([t, d.nama, "'" + d.nim].concat(d.p, d.e));
  lock.releaseLock();
  return ContentService.createTextOutput('OK');
}
