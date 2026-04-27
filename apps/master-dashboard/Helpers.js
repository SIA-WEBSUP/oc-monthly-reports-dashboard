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

function appendItemsTable_(body, items, options) {
  const opts = options || {};
  const columns = Number(opts.columns) > 0 ? Number(opts.columns) : 4;
  const hasFontSize = Number(opts.fontSize) > 0;
  const fontSize = hasFontSize ? Number(opts.fontSize) : null;
  const borderWidth = Number.isFinite(opts.borderWidth) ? Number(opts.borderWidth) : 0;
  const glyph = String(opts.glyph ?? '');
  const cellPadding = opts.cellPadding;
  const cellPaddingTop = opts.cellPaddingTop;
  const cellPaddingRight = opts.cellPaddingRight;
  const cellPaddingBottom = opts.cellPaddingBottom;
  const cellPaddingLeft = opts.cellPaddingLeft;

  const normalizedItems = items
    .map(raw => {
      if (raw && typeof raw === 'object') {
        const text = String(raw.text ?? raw.label ?? raw.position ?? '').trim();
        const url = String(raw.url ?? '').trim();
        return { text, url };
      }

      return {
        text: String(raw || '').trim(),
        url: '',
      };
    })
    .filter(item => item.text);

  if (!normalizedItems.length) {
    return null;
  }

  const rowCount = Math.ceil(normalizedItems.length / columns);
  const tableData = Array.from({ length: rowCount }, () => Array(columns).fill(''));
  const table = body.appendTable(tableData);
  table.setBorderWidth(borderWidth);

  normalizedItems.forEach((item, index) => {
    const rowIndex = Math.floor(index / columns);
    const colIndex = index % columns;
    const cell = table.getCell(rowIndex, colIndex);

    cell.clear();
    applyCellPadding_(cell, {
      all: cellPadding,
      top: cellPaddingTop,
      right: cellPaddingRight,
      bottom: cellPaddingBottom,
      left: cellPaddingLeft,
    });

    const displayText = glyph ? `${glyph} ${item.text}` : item.text;
    const paragraph = cell.appendParagraph(displayText);

    paragraph
      .setSpacingBefore(0)
      .setSpacingAfter(0)
      .setLineSpacing(1);

    if (hasFontSize) {
      paragraph.setFontSize(fontSize);
    }

    if (item.url) {
      paragraph.editAsText().setLinkUrl(0, displayText.length - 1, item.url);
    }
  });

  return table;
}

function appendStringArrayTable_(body, items, options) {
  return appendItemsTable_(body, items, options);
}

function applyCellPadding_(cell, paddingOptions) {
  const opts = paddingOptions || {};
  const all = Number.isFinite(opts.all) ? Number(opts.all) : null;
  const top = Number.isFinite(opts.top) ? Number(opts.top) : all;
  const right = Number.isFinite(opts.right) ? Number(opts.right) : all;
  const bottom = Number.isFinite(opts.bottom) ? Number(opts.bottom) : all;
  const left = Number.isFinite(opts.left) ? Number(opts.left) : all;

  if (Number.isFinite(top) && typeof cell.setPaddingTop === 'function') {
    cell.setPaddingTop(top);
  }

  if (Number.isFinite(right) && typeof cell.setPaddingRight === 'function') {
    cell.setPaddingRight(right);
  }

  if (Number.isFinite(bottom) && typeof cell.setPaddingBottom === 'function') {
    cell.setPaddingBottom(bottom);
  }

  if (Number.isFinite(left) && typeof cell.setPaddingLeft === 'function') {
    cell.setPaddingLeft(left);
  }
}