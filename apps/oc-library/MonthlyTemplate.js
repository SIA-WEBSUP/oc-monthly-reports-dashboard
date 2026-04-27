/**
 * Adds a new monthly report entry at the top of the Report tab
 */
function addBlankTemplateForCurrentMonth() {
  const doc = DocumentApp.getActiveDocument();

  const reportTab = findTabByName_(doc, 'Report');
  const templateTab = findTabByName_(doc, 'Monthly Template');

  if (!reportTab) {
    throw new Error('Report tab not found.');
  }

  if (!templateTab) {
    throw new Error('Monthly Template tab not found.');
  }

  const reportBody = reportTab.asDocumentTab().getBody();
  const templateBody = templateTab.asDocumentTab().getBody();

  const firstMonthlyHeading = findFirstMonthlyHeading_(reportBody);

  if (!firstMonthlyHeading) {
    throw new Error(
      'Could not find first monthly heading in Report tab.'
    );
  }

  const insertIndex = firstMonthlyHeading.index;

  const metadata = extractTitleAndName_(
    reportBody,
    insertIndex
  );

  const templateStart =
    findFirstHeading1InTemplate_(
      templateBody
    );

  if (!templateStart) {
    throw new Error(
      'Could not find first HEADING1 in Monthly Template tab.'
    );
  }

  const insertedElements =
    copyTemplateSection_(
      templateBody,
      templateStart.index,
      reportBody,
      insertIndex
    );

  replaceTemplatePlaceholders_(
    insertedElements,
    metadata
  );

  DocumentApp.getUi().alert(
    'New monthly template added successfully.'
  );
}


/**
 * Find tab by exact name
 */
function findTabByName_(doc, tabName) {
  const tabs = doc.getTabs();

  for (const tab of tabs) {
    if (tab.getTitle() === tabName) {
      return tab;
    }
  }

  return null;
}


/**
 * Finds first HEADING1 in Monthly Template tab
 */
function findFirstHeading1InTemplate_(body) {
  for (
    let i = 0;
    i < body.getNumChildren();
    i++
  ) {
    const child = body.getChild(i);

    if (
      child.getType() ===
      DocumentApp.ElementType.PARAGRAPH
    ) {
      const paragraph =
        child.asParagraph();

      if (
        paragraph.getHeading() ===
        DocumentApp
          .ParagraphHeading
          .HEADING1
      ) {
        return {
          index: i,
          element: paragraph
        };
      }
    }
  }

  return null;
}


/**
 * Finds first monthly heading
 *
 * Accepts:
 * October 2026
 * Oct 2026
 * Oct. 2026
 * Sept 2026
 */
function findFirstMonthlyHeading_(body) {
  const validMonths = [
    'jan',
    'feb',
    'mar',
    'apr',
    'may',
    'jun',
    'jul',
    'aug',
    'sep',
    'oct',
    'nov',
    'dec'
  ];

  for (
    let i = 0;
    i < body.getNumChildren();
    i++
  ) {
    const child = body.getChild(i);

    if (
      child.getType() !==
      DocumentApp.ElementType.PARAGRAPH
    ) {
      continue;
    }

    const paragraph =
      child.asParagraph();

    if (
      paragraph.getHeading() !==
      DocumentApp
        .ParagraphHeading
        .HEADING1
    ) {
      continue;
    }

    const text =
      paragraph.getText().trim();

    const normalized = text
      .replace(/\./g, '')
      .trim();

    const parts =
      normalized.split(/\s+/);

    if (parts.length < 2) {
      continue;
    }

    const monthPart = parts[0]
      .substring(0, 3)
      .toLowerCase();

    const yearPart = parts[1];

    const validMonth =
      validMonths.includes(
        monthPart
      );

    const validYear =
      /^\d{4}$/.test(yearPart);

    if (validMonth && validYear) {
      return {
        index: i,
        element: paragraph,
        text: text
      };
    }
  }

  return null;
}


/**
 * Reads line after monthly heading
 *
 * Expected:
 * Title: Name
 */
function extractTitleAndName_(
  body,
  headingIndex
) {
  const nextIndex =
    headingIndex + 1;

  if (
    nextIndex >=
    body.getNumChildren()
  ) {
    return {};
  }

  const nextChild =
    body.getChild(nextIndex);

  if (
    nextChild.getType() !==
    DocumentApp.ElementType.PARAGRAPH
  ) {
    return {};
  }

  const text = nextChild
    .asParagraph()
    .getText()
    .trim();

  const match = text.match(
    /^(.+?):\s*(.+)$/
  );

  if (!match) {
    return {};
  }

  return {
    title: match[1].trim(),
    name: match[2].trim()
  };
}


function copyTemplateSection_(
  sourceBody,
  startIndex,
  targetBody,
  insertIndex
) {
  const numChildren = sourceBody.getNumChildren();
  const insertedElements = [];

  for (let i = numChildren - 1; i >= startIndex; i--) {
    const child = sourceBody.getChild(i).copy();
    const type = child.getType();
    let insertedElement = null;

    switch (type) {
      case DocumentApp.ElementType.PARAGRAPH:
        insertedElement = targetBody.insertParagraph(insertIndex, child.asParagraph());
        break;

      case DocumentApp.ElementType.LIST_ITEM:
        const sourceItem = child.asListItem();
        insertedElement = targetBody.insertListItem(insertIndex, sourceItem);
        
        // FORCE THE GLYPH HERE
        // 1. Link it to a list (using itself as a reference creates/joins a list)
        insertedElement.setListId(insertedElement);
        // 2. Set the Nesting Level (critical for sub-bullets)
        insertedElement.setNestingLevel(sourceItem.getNestingLevel());
        // 3. Force the specific bullet style (e.g., SOLID_SQUARE or BULLET)
        insertedElement.setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET);
        break;

      case DocumentApp.ElementType.TABLE:
        insertedElement = targetBody.insertTable(insertIndex, child.asTable());
        break;

      case DocumentApp.ElementType.HORIZONTAL_RULE:
        insertedElement = targetBody.insertHorizontalRule(insertIndex);
        break;

      case DocumentApp.ElementType.PAGE_BREAK:
        insertedElement = targetBody.insertPageBreak(insertIndex);
        break;

      default:
        continue;
    }

    if (insertedElement) {
      insertedElements.unshift(insertedElement);
    }
  }

  return insertedElements;
}


/**
 * Replace placeholders
 */
function replaceTemplatePlaceholders_(
  insertedElements,
  metadata
) {
  const now = new Date();
  const tz =
    Session.getScriptTimeZone();

  const replacements = {
    '{Month}':
      Utilities.formatDate(
        now,
        tz,
        'MMMM'
      ),

    '{Year}':
      Utilities.formatDate(
        now,
        tz,
        'yyyy'
      )
  };

  if (metadata.title) {
    replacements[
      '{Title}'
    ] = metadata.title;
  }

  if (metadata.name) {
    replacements[
      '{Name}'
    ] = metadata.name;
  }

  insertedElements.forEach(
    el => {
      const type =
        el.getType();

      if (
        type ===
          DocumentApp
            .ElementType
            .PARAGRAPH ||
        type ===
          DocumentApp
            .ElementType
            .LIST_ITEM
      ) {
        const text =
          el.editAsText();

        Object.entries(
          replacements
        ).forEach(
          ([
            findText,
            replaceText
          ]) => {
            text.replaceText(
              escapeRegex_(
                findText
              ),
              replaceText
            );
          }
        );
      }

      if (
        type ===
        DocumentApp
          .ElementType
          .TABLE
      ) {
        replaceTextInTable_(
          el.asTable(),
          replacements
        );
      }
    }
  );
}


/**
 * Replace text inside tables
 */
function replaceTextInTable_(
  table,
  replacements
) {
  for (
    let r = 0;
    r < table.getNumRows();
    r++
  ) {
    const row =
      table.getRow(r);

    for (
      let c = 0;
      c < row.getNumCells();
      c++
    ) {
      const cell =
        row.getCell(c);

      for (
        let i = 0;
        i <
        cell.getNumChildren();
        i++
      ) {
        const child =
          cell.getChild(i);

        if (
          child.getType() ===
            DocumentApp
              .ElementType
              .PARAGRAPH ||
          child.getType() ===
            DocumentApp
              .ElementType
              .LIST_ITEM
        ) {
          const text =
            child.editAsText();

          Object.entries(
            replacements
          ).forEach(
            ([
              findText,
              replaceText
            ]) => {
              text.replaceText(
                escapeRegex_(
                  findText
                ),
                replaceText
              );
            }
          );
        }
      }
    }
  }
}


/**
 * Escape regex characters
 */
function escapeRegex_(
  text
) {
  return text.replace(
    /[.*+?^${}()|[\]\\]/g,
    '\\$&'
  );
}