/**
 * O&C Monthly Reports Dashboard
 * Updated to include Email Address and Web Page fields.
 */

const TAB_NAMES = {
  directory: 'Directory',
  compiled: 'Compiled',
  settings: 'Settings',
  log: 'Log',
};

const DIRECTORY_HEADERS = [
  'Active',
  'Title',
  'Position',
  'Name',
  'Email Address',
  'Web Page',
  'Source Doc Link',
  'Source Doc ID',
  'Notes',
];

const COMPILED_HEADERS = [
  'Month Label',
  'Compiled Doc Link',
  'Compiled Doc ID',
  'Created On',
  'Missing Reports',
  'Notes',
];

const SETTINGS_HEADERS = [
  'Key',
  'Value',
  'Notes',
];

const LOG_HEADERS = [
  'Timestamp',
  'Run ID',
  'Month Label',
  'Level',
  'Action',
  'Source',
  'Comment',
];

const DIRECTORY_ROWS = [
  ['yes', 'Officer', 'Chairperson', 'Tom B.', 'chairperson@suffolkny-aa.org', 'https://suffolkny-aa.org', '', '', ''],
  ['yes', 'Officer', 'Alt-Chairperson', 'David V.', 'alt-chairperson@suffolkny-aa.org', 'https://suffolkny-aa.org', '', '', ''],
  ['yes', 'Officer', 'Treasurer', 'Meredith F.', 'treasurer@suffolkny-aa.org', 'https://suffolkny-aa.org/7th-tradition', '', '', ''],
  ['yes', 'Officer', 'Recording Secretary', 'TBD', 'recsec@suffolkny-aa.org', 'https://suffolkny-aa.org/sia-business-meeting', '', '', ''],
  ['yes', 'Officer', 'Corresponding Secretary', 'Nancy S.', 'recsec@suffolkny-aa.org', 'https://suffolkny-aa.org/siaoffice', '', '', ''],

  ['yes', 'Special Worker', 'Office Manager', 'Charlie D.', 'siaoffice@suffolkny-aa.org', 'https://suffolkny-aa.org/siaoffice', '', '', ''],

  ['yes', 'Chairperson', 'Archives', 'Evan H.', 'archives@suffolkny-aa.org', 'https://suffolkny-aa.org/archives', '', '', ''],
  ['yes', 'Chairperson', 'Big Meeting', 'TBD', 'bigmeeting@suffolkny-aa.org', 'https://suffolkny-aa.org', '', '', ''],
  ['yes', 'Editor', 'Bulletin', 'Tim D.', 'bulletins@suffolkny-aa.org', 'https://suffolkny-aa.org/bulletin', '', '', ''],
  ['yes', 'Chairperson', 'Corrections', 'Karen C.', 'correct@suffolkny-aa.org', 'https://suffolkny-aa.org/corrections', '', '', ''],
  ['yes', 'Chairperson', 'Grapevine', 'Laurie A.', 'grapevine@suffolkny-aa.org', 'https://suffolkny-aa.org/grapevine', '', '', ''],
  ['yes', 'Chairperson', 'Hot Line', 'Howie L.', 'phones@suffolkny-aa.org', 'https://suffolkny-aa.org/hotline', '', '', ''],
  ['yes', 'Chairperson', 'Literature', 'Mike R.', 'literature@suffolkny-aa.org', 'https://suffolkny-aa.org/literature', '', '', ''],
  ['yes', 'Chairperson', 'Meeting List', 'Adam B.', 'meetings@suffolkny-aa.org', 'https://suffolkny-aa.org/meeting-list', '', '', ''],
  ['yes', 'Chairperson', 'Public Information', 'Ed A.', 'pubinfo@suffolkny-aa.org', 'https://suffolkny-aa.org/pubinfo', '', '', ''],
  ['yes', 'Chairperson', 'Schools', 'TBD', 'schools@suffolkny-aa.org', 'https://suffolkny-aa.org/pubinfo', '', '', ''],
  ['yes', 'Chairperson', 'Share A Thon', 'TBD', 'share@suffolkny-aa.org', 'https://suffolkny-aa.org', '', '', ''],
  ['yes', 'Chairperson', 'Special Events', 'Pete K.', 'specialevents@suffolkny-aa.org', 'https://suffolkny-aa.org', '', '', ''],
  ['yes', 'Chairperson', 'Third Legacy', 'Natalie S.', 'thirdleg@suffolkny-aa.org', 'https://suffolkny-aa.org/thirdleg', '', '', ''],
  ['yes', 'Chairperson', 'Treatment Facilities', 'Rob M.', 'treatment@suffolkny-aa.org', 'https://suffolkny-aa.org/treatment', '', '', ''],
  ['yes', 'Chairperson', 'Webmaster', 'David P.', 'webmaster@suffolkny-aa.org', 'https://suffolkny-aa.org', '', '', ''],

  ['yes', 'Liaison', 'Hispanic Intergroup Liaison', 'Jose R.', 'hispanicrep@suffolkny-aa.org', '', '', '', ''],
  ['yes', 'Liaison', 'General Service Liaison', 'Zoie S.', 'sgsliaison@suffolkny-aa.org', 'https://aasuffolkgs-ny.org/', '', '', ''],
  ['yes', 'Liaison', 'Al-Anon Liaison', 'Matt C.', 'AAliaison@al-anon-suffolk-ny.org', 'https://al-anon-suffolk-ny.org', '', '', '']
];

const DEFAULT_SETTINGS = [
  ['compiled_doc_title_prefix', 'O&C Compiled Report', 'Used when naming compiled monthly report docs'],
  ['compiled_report_template_doc_id', '1GdDtJs0DB2KMhiAj4W0XkBN0Rgc4V3nOnx4h5oT6RPo', 'Google Doc ID of the master compiled report template'],
  ['compiled_report_folder_id', '16lHNWsRWjEL_NKKVzNtEpkm38wlPRf4V', 'Google Drive folder ID where compiled reports will be created'],
  ['template_doc_id', '1JP9t7ybcHneAGaD_SKuPQRaZ09am77GCOGwQAt3_Dso', 'Google Doc ID of the master monthly report template'],
  ['source_docs_folder_id', '1yoL_GRTgmDuBQeVy0I4aITEPENZHb8OX', 'Google Drive folder ID where individual source docs should be created'],
  ['index_doc_template_doc_id', '10tQrUcuBOvpun_h-EmzYy2uPPtYm9cvLqCAporJrzIA', 'Google Doc ID of the master index doc template'],
  ['index_doc_folder_id', '1XF5PeDsigQIC5Jv5oL3Whm6J8HAvF9xC', 'Optional folder ID where the index/directory doc should be created'],
  ['include_missing_placeholders', 'yes', 'If yes, compiled reports include missing placeholders'],
  ['reuse_existing_compiled_doc', 'no', 'Future compile scripts may use this'],
];

function setupOCMonthlyReportsDashboard() {
  
  // Guards to make sure you want to destroy the existing tabs
  const ui = SpreadsheetApp.getUi();
  const firstResponse = ui.alert(
    'Overwrite existing tabs?',
    'This will destroy the existing dashboard tabs and destroy work you added manually.\n\nDo you want to continue?',
    ui.ButtonSet.YES_NO
  );

  if (firstResponse !== ui.Button.YES) {
    ui.alert('Cancelled. No tabs were overwritten.');
    return;
  }

  const secondResponse = ui.alert(
    'Final confirmation',
    'Are you absolutely sure?\n\nThis action cannot be undone easily.',
    ui.ButtonSet.YES_NO
  );

  if (secondResponse !== ui.Button.YES) {
    ui.alert('Cancelled. No tabs were overwritten.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  [TAB_NAMES.directory, TAB_NAMES.compiled, TAB_NAMES.settings, TAB_NAMES.log].forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) ss.deleteSheet(sh);
  });

  const directorySheet = ss.insertSheet(TAB_NAMES.directory);
  const compiledSheet = ss.insertSheet(TAB_NAMES.compiled);
  const settingsSheet = ss.insertSheet(TAB_NAMES.settings);
  const logSheet = ss.insertSheet(TAB_NAMES.log);

  setupDirectorySheet_(directorySheet);
  setupCompiledSheet_(compiledSheet);
  setupSettingsSheet_(settingsSheet);
  setupLogSheet_(logSheet);

  ss.setActiveSheet(directorySheet);

  debugLog({
    runId: makeRunId(),
    level: 'INFO',
    action: 'SETUP_COMPLETE',
    sourceLabel: 'System',
    comment: 'Dashboard created with updated Directory columns including Email Address and Web Page.',
  });
}

function setupDirectorySheet_(sheet) {
  sheet.clear();
  sheet.getRange(1, 1, 1, DIRECTORY_HEADERS.length).setValues([DIRECTORY_HEADERS]);
  sheet.getRange(2, 1, DIRECTORY_ROWS.length, DIRECTORY_HEADERS.length).setValues(DIRECTORY_ROWS);

  formatHeaderRow_(sheet, DIRECTORY_HEADERS.length);
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 80);   // Active
  sheet.setColumnWidth(2, 140);  // Title
  sheet.setColumnWidth(3, 240);  // Position
  sheet.setColumnWidth(4, 180);  // Name
  sheet.setColumnWidth(5, 240);  // Email Address
  sheet.setColumnWidth(6, 240);  // Web Page
  sheet.setColumnWidth(7, 360);  // Source Doc Link
  sheet.setColumnWidth(8, 220);  // Source Doc ID
  sheet.setColumnWidth(9, 220);  // Notes

  sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).createFilter();

  const activeValidation = SpreadsheetApp.newDataValidation()
    .requireValueInList(['yes','no'], true)
    .setAllowInvalid(true)
    .build();
  sheet.getRange(2, 1, DIRECTORY_ROWS.length, 1).setDataValidation(activeValidation);

  sheet.getRange('A1').setNote('Leave blank to skip this row in future scripts.');
  sheet.getRange('B1').setNote('Officer, Chairperson, Liaison, or Special Worker.');
  sheet.getRange('C1').setNote('Row order determines output order in future compiled reports.');
  sheet.getRange('E1').setNote('Optional email address to place into the source doc.');
  sheet.getRange('F1').setNote('Optional web page URL to place into the source doc.');
  sheet.getRange('G1').setNote('Full Google Doc URL.');
  sheet.getRange('H1').setNote('Google Doc ID.');

  applyDirectoryBanding_(sheet);
  applyInactiveRowFormatting_(sheet);

}

function applyDirectoryBanding_(sheet) {
  const range = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn());

  range.applyRowBanding(
    SpreadsheetApp.BandingTheme.CYAN,
    true,
    false
  );
}


function applyInactiveRowFormatting_(sheet) {
  const range = sheet.getRange(2, 1, sheet.getMaxRows() - 1, DIRECTORY_HEADERS.length);

  const rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=$A2<>"yes"')
    .setFontColor('#999999')
    .setRanges([range])
    .build();

  const existingRules = sheet.getConditionalFormatRules();
  existingRules.push(rule);
  sheet.setConditionalFormatRules(existingRules);
}


function setupCompiledSheet_(sheet) {
  sheet.clear();
  sheet.getRange(1, 1, 1, COMPILED_HEADERS.length).setValues([COMPILED_HEADERS]);
  formatHeaderRow_(sheet, COMPILED_HEADERS.length);
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 320);
  sheet.setColumnWidth(3, 220);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 120);
  sheet.setColumnWidth(6, 260);
}

function setupSettingsSheet_(sheet) {
  sheet.clear();
  sheet.getRange(1, 1, 1, SETTINGS_HEADERS.length).setValues([SETTINGS_HEADERS]);
  sheet.getRange(2, 1, DEFAULT_SETTINGS.length, SETTINGS_HEADERS.length).setValues(DEFAULT_SETTINGS);
  formatHeaderRow_(sheet, SETTINGS_HEADERS.length);
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 420);
}

function setupLogSheet_(sheet) {
  sheet.clear();
  sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]);
  formatHeaderRow_(sheet, LOG_HEADERS.length);
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 140);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 180);
  sheet.setColumnWidth(6, 220);
  sheet.setColumnWidth(7, 520);
}

function formatHeaderRow_(sheet, numColumns) {
  sheet.getRange(1, 1, 1, numColumns).setFontWeight('bold').setBackground('#d9ead3');
}