function createIndexGoogleDoc() {
  const runId = makeRunId();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_NAMES.directory);

  if (!sheet) throw new Error(`Missing tab: ${TAB_NAMES.directory}`);

  const folderId = getSetting('index_doc_folder_id') || getSetting('source_docs_folder_id');
  if (!folderId) throw new Error('Settings is missing index_doc_folder_id and source_docs_folder_id.');

  const targetFolder = DriveApp.getFolderById(folderId);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No Directory rows found.');
    return;
  }

  // Active, title, Position, Name, Email Address, Web Page, Source Doc Link, Source Doc ID, Notes
  const values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  const groups = {
    'Officers': [],
    'Special Workers': [],
    'Chairpersons': [],
    'Liaisons': [],
  };

  values.forEach(row => {
    const [
      active,
      title,
      position,
      name,
      emailAddress,
      webPage,
      sourceDocLink,
      sourceDocId,
    ] = row;

    if (String(active).trim().toLowerCase() !== 'yes') return;

    const item = {
      title: String(title).trim(),
      position: String(position).trim(),
      url: String(sourceDocLink || '').trim(),
      id: String(sourceDocId || '').trim(),
    };

    if (!item.position) return;

    if (item.title === 'Officer') {
      groups['Officers'].push(item);
    } else if (item.title === 'Special Worker') {
      groups['Special Workers'].push(item);
    } else if (item.title === 'Chairperson' || item.title === 'Editor') {
      groups['Chairpersons'].push(item);
    } else if (item.title === 'Liaison') {
      groups['Liaisons'].push(item);
    }
  });

  const docTitle = 'O&C Monthly Reports Directory';
  const doc = DocumentApp.create(docTitle);
  const file = DriveApp.getFileById(doc.getId());

  targetFolder.addFile(file);

  try {
    DriveApp.getRootFolder().removeFile(file);
  } catch (e) {}

  const body = doc.getBody();
  body.clear();

  body.appendParagraph('O&C Monthly Reports Directory')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  body.appendParagraph('Suffolk Intergroup Association')
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  body.appendParagraph('');

  appendThreeColumnIndexSection_(body, 'Officers', groups['Officers']);
  appendThreeColumnIndexSection_(body, 'Special Workers', groups['Special Workers']);
  appendThreeColumnIndexSection_(body, 'Chairpersons', groups['Chairpersons']);
  appendThreeColumnIndexSection_(body, 'Liaisons', groups['Liaisons']);

  doc.saveAndClose();

  const compiledSheet = ss.getSheetByName(TAB_NAMES.compiled);
  if (!compiledSheet) throw new Error(`Missing tab: ${TAB_NAMES.compiled}`);

  compiledSheet.appendRow([
    'Index',
    doc.getUrl(),
    doc.getId(),
    new Date(),
    '',
    'Created O&C Monthly Reports Directory doc.',
  ]);

  debugLog({
    runId,
    level: 'INFO',
    action: 'INDEX_DOC_CREATED',
    sourceLabel: 'Compiled',
    sourceUrl: doc.getUrl(),
    comment: 'Created O&C Monthly Reports Directory doc.',
  });

  SpreadsheetApp.getUi().alert(`Index doc created:\n${doc.getUrl()}`);
}


function appendThreeColumnIndexSection_(body, sectionName, items) {
  body.appendParagraph(sectionName)
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  if (!items.length) {
    body.appendParagraph('No entries.');
    body.appendParagraph('');
    return;
  }

  const tableData = [];

  for (let i = 0; i < items.length; i += 3) {
    tableData.push(['', '', '']);
  }

  const table = body.appendTable(tableData);
  table.setBorderWidth(0);

  items.forEach((item, index) => {
    const rowIndex = Math.floor(index / 3);
    const colIndex = index % 3;
    const cell = table.getCell(rowIndex, colIndex);

    cell.clear();

    const label = item.position;
    const p = cell.appendParagraph(label);

    if (item.url) {
      p.editAsText().setLinkUrl(0, label.length - 1, item.url);
    }
  });

  body.appendParagraph('');
}