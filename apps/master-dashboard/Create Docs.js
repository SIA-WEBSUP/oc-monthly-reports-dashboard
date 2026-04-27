function createSourceDocsFromTemplate() {
  const runId = makeRunId();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(TAB_NAMES.directory);

  if (!sheet) throw new Error(`Missing tab: ${TAB_NAMES.directory}`);

  const templateDocId = getSetting('template_doc_id');
  const folderId = getSetting('source_docs_folder_id');

  if (!templateDocId) throw new Error('Settings is missing template_doc_id.');
  if (!folderId) throw new Error('Settings is missing source_docs_folder_id.');

  const templateFile = DriveApp.getFileById(templateDocId);
  const targetFolder = DriveApp.getFolderById(folderId);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No Directory rows found.');
    return;
  }

  // Active, Title, Position, Name, Email Address, Web Page, Source Doc Link, Source Doc ID, Notes
  const values = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  debugLog({
    runId,
    level: 'INFO',
    action: 'CREATE_SOURCE_DOCS_START',
    sourceLabel: 'System',
    comment: 'Started source doc creation from template.',
  });

  let createdCount = 0;
  let skippedCount = 0;
  let failedCount = 0;
  let sharedCount = 0;
  let shareFailedCount = 0;
  let loggerInstalledCount = 0;
  let loggerFailedCount = 0;

  values.forEach((row, idx) => {
    const rowNum = idx + 2;
    const [
      active,
      title,
      position,
      name,
      emailAddress,
      webPage,
      sourceDocLink,
      sourceDocId
    ] = row;

    if (String(active).trim().toLowerCase() !== 'yes') {
      skippedCount++;
      debugLog({
        runId,
        level: 'INFO',
        action: 'SOURCE_DOC_SKIPPED_INACTIVE',
        sourceLabel: position || `Row ${rowNum}`,
        comment: `Skipped row ${rowNum} because Active is not yes.`,
      });
      return;
    }

    if (!position) {
      failedCount++;
      debugLog({
        runId,
        level: 'ERROR',
        action: 'SOURCE_DOC_SKIPPED_NO_POSITION',
        sourceLabel: `Row ${rowNum}`,
        comment: `Skipped row ${rowNum} because Position is blank.`,
      });
      return;
    }

    if (sourceDocId) {
      skippedCount++;
      debugLog({
        runId,
        level: 'INFO',
        action: 'SOURCE_DOC_SKIPPED_EXISTS',
        sourceLabel: position,
        sourceUrl: sourceDocLink || '',
        comment: 'Skipped because Source Doc ID already exists.',
      });
      return;
    }

    try {
      const newName = `${position} Report`;
      const copiedFile = templateFile.makeCopy(newName, targetFolder);
      const newDocId = copiedFile.getId();
      const newDocUrl = copiedFile.getUrl();

      debugLog({
        runId,
        level: 'INFO',
        action: 'SOURCE_DOC_COPY_CREATED',
        sourceLabel: position,
        sourceUrl: newDocUrl,
        comment: `Created Drive copy "${newName}" with doc ID ${newDocId}.`,
      });

      const doc = DocumentApp.openById(newDocId);
      const body = doc.getBody();

      body.replaceText('\\{Position\\}', position || '');
      body.replaceText('\\{Title\\}', title || '');
      body.replaceText('\\{Name\\}', name || '');
      body.replaceText('\\{EmailAddress\\}', emailAddress || '');
      body.replaceText('\\{WebPage\\}', webPage || '');

      if (emailAddress) {
        linkTextInBody_(body, emailAddress, `mailto:${emailAddress}`);
      }

      if (webPage) {
        linkTextInBody_(body, webPage, webPage);
      }
      doc.saveAndClose();

      try {
        installAccessLoggerForDoc_(newDocId);

        loggerInstalledCount++;

        debugLog({
          runId,
          level: 'INFO',
          action: 'SOURCE_DOC_ACCESS_LOGGER_INSTALLED',
          sourceLabel: position,
          sourceUrl: newDocUrl,
          comment: `Installed access logger for "${newName}".`,
        });

      } catch (loggerErr) {
        loggerFailedCount++;

        debugLog({
          runId,
          level: 'ERROR',
          action: 'SOURCE_DOC_ACCESS_LOGGER_INSTALL_FAILED',
          sourceLabel: position,
          sourceUrl: newDocUrl,
          comment: `Doc was created, but access logger install failed: ${loggerErr.message}`,
        });
      }

      // G = Source Doc Link, H = Source Doc ID
      sheet.getRange(rowNum, 7).setValue(newDocUrl);
      sheet.getRange(rowNum, 8).setValue(newDocId);

      debugLog({
        runId,
        level: 'INFO',
        action: 'SOURCE_DOC_CREATED',
        sourceLabel: position,
        sourceUrl: newDocUrl,
        comment: `Created source doc "${newName}" and wrote link/ID back to Directory.`,
      });

      createdCount++;

      // Grant editor access using Email Address column
      const email = String(emailAddress || '').trim();
      if (email) {
        try {
          debugLog({
            runId,
            level: 'INFO',
            action: 'SOURCE_DOC_EDITOR_GRANT_ATTEMPT',
            sourceLabel: position,
            sourceUrl: newDocUrl,
            comment: `Attempting to grant editor access to ${email}.`,
          });

          // Add without generating an email notification
          Drive.Permissions.insert(
            {
              role: 'writer',
              type: 'user',
              value: email
            },
            newDocId,
            {
              sendNotificationEmails: false,
              supportsAllDrives: true
            }
          );


          // Try to verify whether the account now appears among editors/viewers.
          // Note: domain/group/shared-drive policies may affect what is visible here.
          let editorEmails = [];
          let viewerEmails = [];

          try {
            editorEmails = copiedFile.getEditors().map(u => u.getEmail()).filter(Boolean);
            viewerEmails = copiedFile.getViewers().map(u => u.getEmail()).filter(Boolean);
          } catch (verifyErr) {
            debugLog({
              runId,
              level: 'WARN',
              action: 'SOURCE_DOC_EDITOR_VERIFY_FAILED',
              sourceLabel: position,
              sourceUrl: newDocUrl,
              comment: `Editor grant was attempted for ${email}, but verification lookup failed: ${verifyErr.message}`,
            });
          }

          const foundAsEditor = editorEmails.includes(email);
          const foundAsViewer = viewerEmails.includes(email);

          if (foundAsEditor) {
            sharedCount++;
            debugLog({
              runId,
              level: 'INFO',
              action: 'SOURCE_DOC_EDITOR_GRANTED',
              sourceLabel: position,
              sourceUrl: newDocUrl,
              comment: `Granted editor access to ${email}. Verified in editor list.`,
            });
          } else if (foundAsViewer) {
            shareFailedCount++;
            debugLog({
              runId,
              level: 'WARN',
              action: 'SOURCE_DOC_EDITOR_GRANTED_AS_VIEWER_ONLY',
              sourceLabel: position,
              sourceUrl: newDocUrl,
              comment: `Sharing attempt completed for ${email}, but account appears only in viewer list, not editor list.`,
            });
          } else {
            // Sometimes group/domain/shared drive permission behavior may not reflect cleanly here.
            sharedCount++;
            debugLog({
              runId,
              level: 'WARN',
              action: 'SOURCE_DOC_EDITOR_GRANT_UNVERIFIED',
              sourceLabel: position,
              sourceUrl: newDocUrl,
              comment: `Called Drive.Permissions.insert(${email}) successfully, but could not verify the address in editor/viewer lists. Check sharing settings, group aliases, and Shared Drive inheritance.`,
            });
          }

        } catch (shareErr) {
          shareFailedCount++;
          debugLog({
            runId,
            level: 'ERROR',
            action: 'SOURCE_DOC_EDITOR_GRANT_FAILED',
            sourceLabel: position,
            sourceUrl: newDocUrl,
            comment: `Doc was created, but Drive.Permissions.insert(${email}) failed: ${shareErr.message}`,
          });
        }
      } else {
        debugLog({
          runId,
          level: 'INFO',
          action: 'SOURCE_DOC_EDITOR_SKIPPED',
          sourceLabel: position,
          sourceUrl: newDocUrl,
          comment: 'No email address provided, so editor sharing was skipped.',
        });
      }

    } catch (err) {
      failedCount++;
      debugLog({
        runId,
        level: 'ERROR',
        action: 'SOURCE_DOC_CREATE_FAILED',
        sourceLabel: position,
        comment: `Failed to create source doc: ${err.message}`,
      });
    }
  });

  debugLog({
    runId,
    level: 'INFO',
    action: 'CREATE_SOURCE_DOCS_COMPLETE',
    sourceLabel: 'System',
    comment: `Finished source doc creation pass. Created: ${createdCount}. Skipped: ${skippedCount}. Failed: ${failedCount}. Shared: ${sharedCount}. Share failed: ${shareFailedCount}. Logger installed: ${loggerInstalledCount}. Logger failed: ${loggerFailedCount}.`,
  });

  SpreadsheetApp.getUi().alert(
    `Source doc creation pass complete.\n\nCreated: ${createdCount}\nSkipped: ${skippedCount}\nFailed: ${failedCount}\nShared: ${sharedCount}\nShare failed: ${shareFailedCount}`
  );
}

function installAccessLoggerForDoc_(docId) {
   ReportDocLib.installAccessLoggerForDoc(docId);
}

function linkTextInBody_(body, textToFind, url) {
  const escapedText = textToFind.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const found = body.findText(escapedText);

  if (!found) return;

  const element = found.getElement().asText();
  const start = found.getStartOffset();
  const end = found.getEndOffsetInclusive();

  element.setLinkUrl(start, end, url);
}
