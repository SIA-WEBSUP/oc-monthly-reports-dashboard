function createIndexGoogleDoc() {
  const runId = makeRunId();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_NAMES.directory);

  if (!sheet) throw new Error(`Missing tab: ${TAB_NAMES.directory}`);

  const templateDocId = getSetting('index_doc_template_doc_id');
  const folderId = getSetting('index_doc_folder_id') || getSetting('source_docs_folder_id');

  if (!templateDocId) throw new Error('Settings is missing index_doc_template_doc_id.');
  if (!folderId) throw new Error('Settings is missing index_doc_folder_id and source_docs_folder_id.');

  const templateFile = DriveApp.getFileById(templateDocId);
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
    'Service Committees': [],
    'Special Workers': [],
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
    } else if (item.title === 'Chairperson' || item.title === 'Editor') {
      groups['Service Committees'].push(item);
    } else if (item.title === 'Special Worker') {
      groups['Special Workers'].push(item);
    } else if (item.title === 'Liaison') {
      groups['Liaisons'].push(item);
    }
  });

  const docTitle = 'O&C Monthly Reports Directory';
  const indexFile = templateFile.makeCopy(docTitle, targetFolder);
  const doc = DocumentApp.openById(indexFile.getId());

  const body = doc.getBody();

  const sectionOrder = [
    'Officers',
    'Service Committees',
    'Special Workers',
    'Liaisons',
  ];

  const sectionTableOptions = {
    columns: 3,
    glyph: '',
    borderWidth: 0,
    cellPaddingTop: 0,
    cellPaddingBottom: 0,
  };

  sectionOrder.forEach(sectionName => {
    body.appendParagraph(sectionName)
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);

    appendThreeColumnIndexSection_(body, groups[sectionName], sectionTableOptions);
  });

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

  showIndexCreatedDialog_({
    docName: indexFile.getName(),
    docUrl: doc.getUrl(),
  });
}


function appendThreeColumnIndexSection_(body, items, tableOptions) {
  if (!items.length) {
    body.appendParagraph('No entries.');
    body.appendParagraph('');
    return;
  }

  appendItemsTable_(body, items.map(item => ({
    text: item.position,
    url: item.url,
  })), tableOptions || {});

  body.appendParagraph('');
}


function showIndexCreatedDialog_(data) {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 16px; line-height: 1.45;">
      <h2 style="margin-top: 0;">Index doc created</h2>

      <p>
        Created <strong>${escapeHtml_(data.docName)}</strong>.
      </p>

      <p>
        <a href="${escapeHtml_(data.docUrl)}" target="_blank" rel="noopener noreferrer"
           style="font-size: 16px; font-weight: bold;">
          Open newly created index
        </a>
      </p>
    </div>
  `)
    .setWidth(420)
    .setHeight(220);

  SpreadsheetApp.getUi().showModalDialog(html, 'Index Doc Created');
}