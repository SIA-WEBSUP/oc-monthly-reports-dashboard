function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Monthly Reports')
    .addItem('1. Compile Monthly Report', 'compileMonthlyReport')
    .addItem('2. Create Index Doc', 'createIndexGoogleDoc')
    .addSeparator()
    .addItem('3. Build Dashboard Tabs (CAREFUL!!!)', 'setupOCMonthlyReportsDashboard')
    .addItem('4. Create Source Docs (CAREFUL!!!)', 'createSourceDocsFromTemplate')
    .addSeparator()
    .addItem('Test Log Entry', 'testDebugLog')
    .addToUi();
}

function makeRunId() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
}

function getSetting(key, fallbackValue = '') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_NAMES.settings);
  if (!sheet) throw new Error(`Missing tab: ${TAB_NAMES.settings}`);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return fallbackValue;

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  for (const [k, v] of values) {
    if (k === key) return v;
  }
  return fallbackValue;
}

function setSetting(key, value, notes = '') {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_NAMES.settings);
  if (!sheet) throw new Error(`Missing tab: ${TAB_NAMES.settings}`);

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === key) {
        sheet.getRange(i + 2, 2).setValue(value);
        if (notes !== '') sheet.getRange(i + 2, 3).setValue(notes);
        return;
      }
    }
  }
  sheet.appendRow([key, value, notes]);
}

function extractGoogleDocId(url) {
  if (!url || typeof url !== 'string') return '';
  const match = url.match(/\/document\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : '';
}

function debugLog({
  runId = '',
  monthLabel = '',
  level = 'INFO',
  action = '',
  sourceLabel = 'System',
  sourceUrl = '',
  comment = '',
} = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(TAB_NAMES.log);
  if (!logSheet) throw new Error(`Missing required tab: ${TAB_NAMES.log}`);

  const nextRow = logSheet.getLastRow() + 1;
  logSheet.getRange(nextRow, 1, 1, 7).setValues([[
    new Date(),
    runId,
    monthLabel,
    level,
    action,
    sourceLabel,
    comment,
  ]]);
  logSheet.getRange(nextRow, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');

  if (sourceUrl) {
    const richText = SpreadsheetApp.newRichTextValue()
      .setText(sourceLabel)
      .setLinkUrl(sourceUrl)
      .build();
    logSheet.getRange(nextRow, 6).setRichTextValue(richText);
  }
}

function testDebugLog() {
  debugLog({
    runId: makeRunId(),
    level: 'INFO',
    action: 'TEST_LOG',
    sourceLabel: 'System',
    comment: 'This is a test log entry.',
  });
  SpreadsheetApp.getUi().alert('Test log entry written.');
}