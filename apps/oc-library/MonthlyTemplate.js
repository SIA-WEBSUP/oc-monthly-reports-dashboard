/**
 * Adds a new monthly report entry at the top of the Report tab
 */
const TEMPLATE_BEGIN_MARKER = 'TEMPLATE BEGIN';

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
    findTemplateContentStart_(
      templateBody
    );

  if (!templateStart) {
    throw new Error(
      `Could not find required marker heading: "${TEMPLATE_BEGIN_MARKER}" in Monthly Template tab.`
    );
  }

  // Template content starts below the marker line.
  const templateContentStart = templateStart.index + 1;

  // Create and insert Month Year heading
  const monthYearHeading = reportBody.insertParagraph(insertIndex, '');
  monthYearHeading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const monthYear = Utilities.formatDate(now, tz, 'MMMM yyyy');
  monthYearHeading.setText(monthYear);

  // Create and insert Title: Name paragraph
  const titleNamePara = reportBody.insertParagraph(insertIndex + 1, '');
  titleNamePara.setText(metadata.title && metadata.name ? `${metadata.title}: ${metadata.name}` : '');
  titleNamePara.setIndentStart(9.36);       // left indent 0.13 inches
  titleNamePara.setIndentFirstLine(9.36);  // first line hangs back to margin
  const titleText = titleNamePara.editAsText();
  titleText.setItalic(true);
  titleText.setBold(true);
  titleText.setForegroundColor('#434343');

  // Copy template content from the detected template start.
  copyTemplateSection_(
    templateBody,
    templateContentStart,
    reportBody,
    insertIndex + 2
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
 * Finds required template marker heading in Monthly Template tab.
 * The marker must be a HEADING1 line with this text:
 * TEMPLATE BEGIN --- DO NOT DELETE THIS LINE
 */
function findTemplateContentStart_(body) {
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

    const p = child.asParagraph();
    const heading = p.getHeading();
    const text = p
      .getText()
      .trim()
      .replace(/\s+/g, ' ')
      .toUpperCase();

    if (
      heading === DocumentApp.ParagraphHeading.HEADING1 &&
      text.startsWith(TEMPLATE_BEGIN_MARKER)
    ) {
      return {
        index: i,
        element: child
      };
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
 * Debug helper: list all elements in Monthly Template tab.
 */
function debugListMonthlyTemplateElements() {
  const doc = DocumentApp.getActiveDocument();
  const templateTab = findTabByName_(doc, 'Monthly Template');

  if (!templateTab) {
    DocumentApp.getUi().alert('Monthly Template tab not found.');
    return;
  }

  const body = templateTab.asDocumentTab().getBody();
  const lines = [];

  lines.push('Monthly Template element scan');
  lines.push(`Total children: ${body.getNumChildren()}`);
  lines.push('');

  for (let i = 0; i < body.getNumChildren(); i++) {
    const child = body.getChild(i);
    const type = child.getType();
    lines.push(formatTemplateElementLine_(i, child, type));
  }

  const output = lines.join('\n');
  Logger.log(output);

  // Alerts can be truncated; show the beginning and log the full output.
  const preview = output.length > 3500
    ? `${output.substring(0, 3500)}\n\n...truncated in alert. Check execution logs for full output.`
    : output;

  DocumentApp.getUi().alert(preview);
}


function formatTemplateElementLine_(index, child, type) {
  if (type === DocumentApp.ElementType.HORIZONTAL_RULE) {
    return `[${index}] HORIZONTAL_RULE`;
  }

  if (type === DocumentApp.ElementType.PARAGRAPH) {
    const p = child.asParagraph();
    const text = p.getText().trim();
    const heading = String(p.getHeading()).replace('ParagraphHeading.', '');
    const preview = text.length > 80
      ? `${text.substring(0, 80)}...`
      : text;
    return `[${index}] PARAGRAPH heading=${heading} text="${preview}"`;
  }

  if (type === DocumentApp.ElementType.LIST_ITEM) {
    const li = child.asListItem();
    const text = li.getText().trim();
    const preview = text.length > 80
      ? `${text.substring(0, 80)}...`
      : text;
    return `[${index}] LIST_ITEM level=${li.getNestingLevel()} text="${preview}"`;
  }

  if (type === DocumentApp.ElementType.TABLE) {
    const t = child.asTable();
    return `[${index}] TABLE rows=${t.getNumRows()}`;
  }

  return `[${index}] ${String(type).replace('ElementType.', '')}`;
}

