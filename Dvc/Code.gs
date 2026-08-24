function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Office Status Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getDashboardData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var rows = sheet.getDataRange().getValues();
  var result = [];

  // Row 4 (Index 3) se lekar saara data read karein
  for (var i = 3; i < rows.length; i++) {
    var id = String(rows[i][0] || '').trim(); // Column A (B1, B2, MR-1 etc.)
    if (!id) continue;

    var booking = String(rows[i][1] || '').trim().toLowerCase(); // Column B
    var isBusy = (booking === 'yes' || booking === 'true');

    result.push({
      id: id,
      type: id.toUpperCase().startsWith('MR') ? 'Meeting Room' : '3D Printer',
      status: isBusy ? 'busy' : 'free',
      time: String(rows[i][2] || '—'),    // Column C (Booking time)
      user: String(rows[i][3] || '—'),    // Column D (User)
      after: String(rows[i][4] || '—')   // Column E (After booking)
    });
  }
  return result;
}