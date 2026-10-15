const SPREADSHEET_ID = '10xMlixR0oK0qcjzmQBUNLHCQbFecPLF5SqOkg1LqCc0';
const SETTING_SHEET_NAME = '⚙️ システム設定';

function doGet(e) {
  // パラメータが存在する場合、または外部（GitHub等）からのデータ要求（Fetch）の場合はJSONを返却する
  if (e && (e.parameter.dataType === 'json' || Object.keys(e.parameter).length > 0)) {
    const settings = getSystemSettings();
    return ContentService.createTextOutput(JSON.stringify(settings))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // 通常通りブラウザで直接開かれた場合はHTML画面を表示する
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('塾ポータル（入口）')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSystemSettings() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SETTING_SHEET_NAME);
    if (!sheet) throw new Error('設定シートが見つかりません');

    return {
      isMaintenance: sheet.getRange('A2').getValue() === 'ON',
      maintenanceEndTime: sheet.getRange('B2').getValue() ? sheet.getRange('B2').getValue().toString() : '未定',
      mainAppUrl: sheet.getRange('C2').getValue()
    };
  } catch (error) {
    return { isMaintenance: true, maintenanceEndTime: '未定 (システムエラー)', mainAppUrl: '' };
  }
}
