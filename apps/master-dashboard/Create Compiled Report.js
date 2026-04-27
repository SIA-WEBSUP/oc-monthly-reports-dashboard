function compileMonthlyReport() {
  const runId = makeRunId();
  const ui = SpreadsheetApp.getUi();

  try {
    debugLog({
      runId,
      level: 'INFO',
      action: 'COMPILE_STARTED',
      sourceLabel: 'System',
      comment: 'Started monthly report compile.',
    });

    const defaultMonth = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'MMMM yyyy'
    );

    const response = ui.prompt(
      'Compile Monthly Report',
      `Enter month to compile. Examples: ${defaultMonth}, Apr 26, 4/26, 04/2026`,
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) {
      debugLog({
        runId,
        level: 'INFO',
        action: 'COMPILE_CANCELLED',
        sourceLabel: 'System',
        comment: 'User cancelled monthly report compile.',
      });
      return;
    }

    const input = response.getResponseText().trim() || defaultMonth;
    const targetDate = parseFlexibleMonthInput_(input);

    const monthLabel = Utilities.formatDate(
      targetDate,
      Session.getScriptTimeZone(),
      'MMMM yyyy'
    );

    debugLog({
      runId,
      level: 'INFO',
      action: 'TARGET_MONTH_SET',
      sourceLabel: 'System',
      comment: `Target month set to ${monthLabel}.`,
    });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(TAB_NAMES.directory);

    if (!sheet) {
      throw new Error(`Missing tab: ${TAB_NAMES.directory}`);
    }

    const newFilePrefix = getSetting('compiled_doc_title_prefix');
    const templateDocId = getSetting('compiled_report_template_doc_id');
    const folderId = getSetting('compiled_report_folder_id');

    if (!newFilePrefix) {
      throw new Error('Missing setting: compiled_doc_title_prefix');
    }

    if (!templateDocId) {
      throw new Error('Missing setting: compiled_report_template_doc_id');
    }

    if (!folderId) {
      throw new Error('Missing setting: compiled_report_folder_id');
    }

    debugLog({
      runId,
      level: 'INFO',
      action: 'SETTINGS_LOADED',
      sourceLabel: 'System',
      comment: 'Compiled report settings loaded.',
    });

    const templateFile = DriveApp.getFileById(templateDocId);
    const folder = DriveApp.getFolderById(folderId);

    const newFile = templateFile.makeCopy(`${newFilePrefix} - ${monthLabel}`, folder);
    const compiledDoc = DocumentApp.openById(newFile.getId());
    const compiledBody = compiledDoc.getBody();

    debugLog({
      runId,
      level: 'INFO',
      action: 'COMPILED_DOC_CREATED',
      sourceLabel: 'System',
      comment: `Created compiled doc: ${newFile.getName()} | ${newFile.getUrl()}`,
    });

    compiledBody.replaceText('\\{Month Year\\}', monthLabel);
    replaceLinkToThisDoc_(compiledBody, newFile.getUrl());

    const values = sheet.getDataRange().getValues();
    const headers = values.shift();

    const sourceDocCol = findColumnIndex_(headers, [
      'Source Doc',
      'Source Doc URL',
      'Source Doc ID',
      'Doc',
      'Doc URL',
    ]);

    const positionCol = findColumnIndex_(headers, ['Position']);
    const nameCol = findColumnIndex_(headers, ['Name']);
    const activeCol = findColumnIndex_(headers, ['Active']);

    if (sourceDocCol === -1) {
      throw new Error(
        'Missing source doc column. Expected one of: Source Doc, Source Doc URL, Source Doc ID, Doc, Doc URL'
      );
    }

    let insertedCount = 0;
    let missingCount = 0;
    let skippedCount = 0;
    let errorCount = 0;
    const notSubmitted = [];
    const skipped = [];

    values.forEach((row, index) => {
      const rowNumber = index + 2;

      const position = positionCol >= 0 ? String(row[positionCol]).trim() : '';
      const name = nameCol >= 0 ? String(row[nameCol]).trim() : '';
      const sourceLabel = position || name || `Row ${rowNumber}`;

      try {
        if (activeCol >= 0) {
          const activeValue = String(row[activeCol]).trim().toLowerCase();

          if (activeValue && activeValue !== 'yes' && activeValue !== 'true') {
            skippedCount++;
            skipped.push(sourceLabel);

            debugLog({
              runId,
              level: 'INFO',
              action: 'ROW_SKIPPED_INACTIVE',
              sourceLabel,
              comment: `Row ${rowNumber} skipped because Active is "${activeValue}".`,
            });

            return;
          }
        }

        const sourceDocValue = row[sourceDocCol];

        if (!sourceDocValue) {
          skippedCount++;
          skipped.push(sourceLabel);

          debugLog({
            runId,
            level: 'WARN',
            action: 'ROW_SKIPPED_NO_SOURCE_DOC',
            sourceLabel,
            comment: `Row ${rowNumber} skipped because no source doc is listed.`,
          });

          return;
        }

        const sourceDocId = extractDocId_(sourceDocValue);

        if (!sourceDocId) {
          skippedCount++;
          skipped.push(sourceLabel);

          debugLog({
            runId,
            level: 'WARN',
            action: 'ROW_SKIPPED_BAD_SOURCE_DOC',
            sourceLabel,
            comment: `Row ${rowNumber} skipped because source doc ID could not be extracted.`,
          });

          return;
        }

        debugLog({
          runId,
          level: 'INFO',
          action: 'SOURCE_DOC_CHECK_STARTED',
          sourceLabel,
          comment: `Checking source doc for ${monthLabel}.`,
        });

        const sourceDoc = DocumentApp.openById(sourceDocId);
        const sectionElements = getMonthlySectionElements_(sourceDoc, monthLabel);

        if (!sectionElements.length) {
          missingCount++;
          notSubmitted.push(sourceLabel);

          debugLog({
            runId,
            level: 'WARN',
            action: 'REPORT_NOT_FOUND',
            sourceLabel,
            comment: `No ${monthLabel} report found in source doc.`,
          });

          return;
        }

        appendReportSection_(compiledBody, position, name, sectionElements);
        insertedCount++;

        debugLog({
          runId,
          level: 'INFO',
          action: 'REPORT_INSERTED',
          sourceLabel,
          comment: `Inserted ${monthLabel} report.`,
        });
      } catch (rowErr) {
        errorCount++;

        debugLog({
          runId,
          level: 'ERROR',
          action: 'ROW_ERROR',
          sourceLabel,
          comment: `Row ${rowNumber}: ${rowErr.message}`,
        });
      }
    });

    removeUnusedReportPlaceholders_(compiledBody);
    appendEndOfReportsSection_(compiledBody, notSubmitted, skipped);
    compiledDoc.saveAndClose();

    addCompiledReportToSheet_(monthLabel, newFile, missingCount);

    debugLog({
      runId,
      level: 'INFO',
      action: 'COMPILE_COMPLETE',
      sourceLabel: 'System',
      comment: `Finished. Inserted: ${insertedCount}; Missing: ${missingCount}; Skipped: ${skippedCount}; Errors: ${errorCount}; URL: ${newFile.getUrl()}`,
    });

    showCompileCompleteDialog_({
      monthLabel,
      insertedCount,
      missingCount,
      skippedCount,
      errorCount,
      docName: newFile.getName(),
      docUrl: newFile.getUrl(),
    });

  } catch (err) {
    debugLog({
      runId,
      level: 'ERROR',
      action: 'COMPILE_FAILED',
      sourceLabel: 'System',
      comment: err.message,
    });

    ui.alert('Compile failed', err.message, ui.ButtonSet.OK);

    throw err;
  }
}


function appendEndOfReportsSection_(body, notSubmitted, skipped) {

  body.appendParagraph('END of Reports')
    .setHeading(DocumentApp.ParagraphHeading.HEADING1);

  const COLS = 4;

  if (notSubmitted.length) {
    body.appendParagraph('Reports Not Submitted')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    appendStringArrayTable_(body, notSubmitted, {
      columns: COLS,
      glyph: '-',
      fontSize: 10,
    });
  }

  if (skipped.length) {
    body.appendParagraph('Reports Skipped')
      .setHeading(DocumentApp.ParagraphHeading.HEADING2);

    appendStringArrayTable_(body, skipped, {
      columns: COLS,
      glyph: '-',
      fontSize: 9,
    });
  }

  const generatedAt = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss z'
  );

  body.appendParagraph(`Report generated at: ${generatedAt}`)
    .setSpacingBefore(12)
    .setFontSize(9)
    .setItalic(true);
}


function appendReportSection_(body, position, name, sectionElements) {
  const placeholder = body.findText('\\{REPORT\\}');

  if (!placeholder) {
    if (position) {
      body.appendParagraph(position)
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);
    }

    sectionElements.forEach(element => {
      appendElementToBody_(body, element);
    });

    return;
  }

  const reportText = placeholder.getElement().asText();
  const start = placeholder.getStartOffset();
  const end = placeholder.getEndOffsetInclusive();

  reportText.deleteText(start, end);

  const placeholderParagraph = reportText.getParent().asParagraph();
  const insertIndex = body.getChildIndex(placeholderParagraph);

  let currentIndex = insertIndex + 1;

  if (name) {
    body.insertParagraph(currentIndex, name).setItalic(true);
    currentIndex++;
  }

  sectionElements.forEach(element => {
    insertElementAt_(body, currentIndex, element);
    currentIndex++;
  });
}


function insertElementAt_(body, index, element) {
  const type = element.getType();

  if (type === DocumentApp.ElementType.PARAGRAPH) {
    body.insertParagraph(index, element.asParagraph());
    return;
  }

  if (type === DocumentApp.ElementType.LIST_ITEM) {
    const inserted = body.insertListItem(index, element.asListItem());
    normalizeListItemGlyph_(inserted);
    return;
  }

  if (type === DocumentApp.ElementType.TABLE) {
    body.insertTable(index, element.asTable());
    return;
  }

  body.insertParagraph(index, element.getText ? element.getText() : '');
}


function copyElementToBody_(element, body) {
  appendElementToBody_(body, element);
}


function appendElementToBody_(body, element) {
  const type = element.getType();

  if (type === DocumentApp.ElementType.PARAGRAPH) {
    body.appendParagraph(element.asParagraph());
    return;
  }

  if (type === DocumentApp.ElementType.LIST_ITEM) {
    const appended = body.appendListItem(element.asListItem());
    normalizeListItemGlyph_(appended);
    return;
  }

  if (type === DocumentApp.ElementType.TABLE) {
    body.appendTable(element.asTable());
    return;
  }

  body.appendParagraph(element.getText ? element.getText() : '');
}


function normalizeListItemGlyph_(listItem) {
  const level = listItem.getNestingLevel();

  const glyphCycle = [
    DocumentApp.GlyphType.BULLET,          // solid circle
    DocumentApp.GlyphType.HOLLOW_BULLET,   // hollow circle
    DocumentApp.GlyphType.SQUARE_BULLET,   // square
    DocumentApp.GlyphType.HOLLOW_SQUARE,   // hollow square
    DocumentApp.GlyphType.ARROW,           // arrow
  ];

  const glyph = glyphCycle[level % glyphCycle.length];

  listItem.setGlyphType(glyph);
}


function getMonthlySectionElements_(doc, monthLabel) {
  const body = doc.getBody();
  const elements = [];
  let inTargetSection = false;

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);

    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const paragraph = child.asParagraph();
      const text = paragraph.getText().trim();

      if (isMonthHeading_(text)) {
        if (normalizeMonthLabel_(text) === monthLabel) {
          inTargetSection = true;
          continue;
        }

        if (inTargetSection) break;
      }
    }

    if (inTargetSection) {
      elements.push(child.copy());
    }
  }

  return elements;
}


function replaceLinkToThisDoc_(body, docUrl) {
  const found = body.findText('\\{linkToThisDoc\\}');
  if (!found) return;

  const textElement = found.getElement().asText();
  const start = found.getStartOffset();
  const end = found.getEndOffsetInclusive();

  textElement.deleteText(start, end);
  textElement.insertText(start, '(click here)');
  textElement.setLinkUrl(start, start + '(click here)'.length - 1, docUrl);
}


function removeUnusedReportPlaceholders_(body) {
  let found = body.findText('\\{REPORT\\}');

  while (found) {
    const textElement = found.getElement().asText();

    textElement.deleteText(
      found.getStartOffset(),
      found.getEndOffsetInclusive()
    );

    found = body.findText('\\{REPORT\\}');
  }
}


function parseFlexibleMonthInput_(input) {
  input = String(input).trim();

  let date = new Date(input);

  if (!isNaN(date.getTime())) {
    return new Date(date.getFullYear(), date.getMonth(), 1);
  }

  const numericMatch = input.match(/^(\d{1,2})[\/\-](\d{2}|\d{4})$/);

  if (numericMatch) {
    const month = Number(numericMatch[1]);
    let year = Number(numericMatch[2]);

    if (year < 100) year += 2000;

    if (month >= 1 && month <= 12) {
      return new Date(year, month - 1, 1);
    }
  }

  throw new Error(
    `Could not understand month: "${input}". Try April 2026, Apr 26, 4/26, or 04/2026.`
  );
}


function isMonthHeading_(text) {
  try {
    normalizeMonthLabel_(text);
    return true;
  } catch (e) {
    return false;
  }
}


function normalizeMonthLabel_(text) {
  const date = parseFlexibleMonthInput_(text);

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'MMMM yyyy'
  );
}


function extractDocId_(value) {
  const text = String(value).trim();

  if (/^[a-zA-Z0-9_-]{25,}$/.test(text)) {
    return text;
  }

  const match = text.match(/\/document\/d\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}


function findColumnIndex_(headers, possibleNames) {
  const normalizedHeaders = headers.map(header =>
    String(header).trim().toLowerCase()
  );

  for (const name of possibleNames) {
    const index = normalizedHeaders.indexOf(name.toLowerCase());
    if (index !== -1) return index;
  }

  return -1;
}

function addCompiledReportToSheet_(monthLabel, file, missingCount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_NAMES.compiled);

  if (!sheet) {
    throw new Error(`Missing tab: ${TAB_NAMES.compiled}`);
  }

  const nextRow = sheet.getLastRow() + 1;

  sheet.getRange(nextRow, 1, 1, 6).setValues([[
    monthLabel,
    file.getUrl(),
    file.getId(),
    new Date(),
    missingCount,
    '',
  ]]);

  sheet.getRange(nextRow, 2).setFormula(
    `=HYPERLINK("${file.getUrl()}","${file.getName().replace(/"/g, '""')}")`
  );
}


function showCompileCompleteDialog_(data) {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 16px; line-height: 1.45;">
      <h2 style="margin-top: 0;">Monthly report compiled</h2>

      <p>
        Created report for <strong>${escapeHtml_(data.monthLabel)}</strong>.
      </p>

      <p>
        <a href="${escapeHtml_(data.docUrl)}" target="_blank" rel="noopener noreferrer"
           style="font-size: 16px; font-weight: bold;">
          Open newly created report
        </a>
      </p>

      <hr>

      <p>
        Reports inserted: <strong>${data.insertedCount}</strong><br>
        Reports missing: <strong>${data.missingCount}</strong><br>
        Rows skipped: <strong>${data.skippedCount}</strong><br>
        Errors: <strong>${data.errorCount}</strong>
      </p>
    </div>
  `)
    .setWidth(420)
    .setHeight(300);

  SpreadsheetApp.getUi().showModalDialog(html, 'Monthly Report Compiled');
}


function escapeHtml_(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}