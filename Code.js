/**
 * Smart Unified Code Highlighter & Professional Table Formatter for Google Docs
 *
 * @license MIT
 * @repository https://github.com/
 */

function onOpen() {
  DocumentApp.getUi()
    .createMenu('⚡ Code, Table & Typography Suite')
    .addItem('🚀 Smart Auto-Format Entire Document', 'quickSmartAutoFormatDocument')
    .addSeparator()
    .addItem('Open Sidebar (Styles & Colors)', 'showSidebar')
    .addSeparator()
    .addItem('✍️ Format All Document Typography', 'quickFormatDocumentTypography')
    .addItem('✍️ Format Selected Text Auto', 'quickFormatSelectedTypography')
    .addItem('📖 Format Selection as Title', 'quickFormatSelectedAsTitle')
    .addItem('📌 Format Selection as Heading 1', 'quickFormatSelectedAsH1')
    .addItem('📑 Format Selection as Sub-Heading', 'quickFormatSelectedAsSub')
    .addItem('↩ Undo Document Text Formatting', 'quickUndoDocumentTypography')
    .addSeparator()
    .addItem('📊 Format & Center All Tables', 'quickFormatAllTables')
    .addItem('📊 Format Selected Table', 'quickFormatSelectedTable')
    .addItem('↩ Undo All Tables Formatting', 'quickUndoAllTables')
    .addSeparator()
    .addItem('⚡ Highlight All Code (Quick Run)', 'quickHighlightAll')
    .addItem('⚡ Highlight Selected Code', 'formatSelectedCodeBlock')
    .addItem('📐 Auto-Indent All Code Blocks', 'quickIndentAllCode')
    .addItem('↩ Undo All Code Blocks Formatting', 'quickUndoAllCodeBlocks')
    .addSeparator()
    .addItem('🔖 Remove All Bookmarks', 'quickRemoveAllBookmarks')
    .addSeparator()
    .addSubMenu(DocumentApp.getUi().createMenu('✨ Text Cleanup & Spacing')
      .addItem('🧹 Cleanup Text (Spaces, Quotes, Dashes)', 'quickCleanupDocumentText')
      .addItem('📐 Normalize Paragraph Spacing', 'quickNormalizeDocumentSpacing')
      .addItem('✦ Apply Inline Markdown (bold/italic/~~strikethrough~~)', 'quickApplyInlineMarkdown')
      .addItem('🔢 Auto-Number Headings (1. / 1.1 / 1.1.1)', 'quickApplyHeadingNumbers')
      .addItem('📋 Convert Markdown Lists (- item, • item)', 'quickConvertMarkdownLists')
      .addItem('🔡 Heading → UPPERCASE', 'quickHeadingsToUpperCase')
      .addItem('🔡 Heading → Title Case', 'quickHeadingsToTitleCase')
    )
    .addToUi();
}

/**
 * Quick action to remove all bookmarks from the document menu
 */
function quickRemoveAllBookmarks() {
  const result = removeAllBookmarks();
  const msg = result.count > 0
    ? 'Removed ' + result.count + ' bookmark(s) from the document.'
    : 'No bookmarks found in the document.';
  DocumentApp.getUi().alert('🔖 Remove Bookmarks', msg, DocumentApp.getUi().ButtonSet.OK);
}


/**
 * Opens the styling sidebar in Google Docs
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutput(getSidebarHtml())
    .setTitle('⚡ Code & Table Formatter')
    .setWidth(340);
  DocumentApp.getUi().showSidebar(html);
}

/* ==========================================================================
   TABLE ALIGNMENT & PROFESSIONAL FORMATTING ENGINE
   ========================================================================== */

/**
 * Quick action to format and center all tables from the menu
 */
function quickFormatAllTables() {
  const prefs = getTablePreferences();
  const result = formatAllDocumentTables(prefs);
  DocumentApp.getUi().alert('📊 Table Formatter', result.message, DocumentApp.getUi().ButtonSet.OK);
}

/**
 * Quick action to format the currently selected table from the menu
 */
function quickFormatSelectedTable() {
  const prefs = getTablePreferences();
  const result = formatSelectedTable(prefs);
  DocumentApp.getUi().alert('📊 Table Formatter', result.message, DocumentApp.getUi().ButtonSet.OK);
}

/**
 * Quick action to undo table formatting from the menu
 */
function quickUndoAllTables() {
  const result = undoAllTableFormatting();
  DocumentApp.getUi().alert('📊 Table Formatter', result.message, DocumentApp.getUi().ButtonSet.OK);
}

/**
 * Formats and centers all data tables in the document.
 * Code block tables (1x1 monospace) are centered without altering code syntax colors.
 */
function formatAllDocumentTables(options) {
  options = options || getTablePreferences();
  saveTablePreferences(options);

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  // Snapshot table list before any DOM mutations
  const tables = body.getTables();

  if (!tables || tables.length === 0) {
    return { success: true, count: 0, message: 'No tables found in this document.' };
  }

  let formattedCount = 0;
  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];

    // Check if this table is a 1x1 code block or heading banner
    if (isCodeBlockTable(table)) {
      // Center code block table to page width without overriding syntax styles
      centerTableOnPage(table, body);
      ensureDoubleNewlineAfterElement(body, table);
      continue;
    }
    if (isHeadingBannerTable(table)) {
      // Center heading banner table to page width without overriding typography styles
      centerTableOnPage(table, body);
      ensureDoubleNewlineAfterElement(body, table);
      continue;
    }

    formatSingleTable(table, options, body);
    // Note: formatSingleTable already calls ensureDoubleNewlineAfterElement internally
    formattedCount++;
  }

  return {
    success: true,
    count: formattedCount,
    message: formattedCount > 0
      ? 'Aligned and formatted ' + formattedCount + ' table(s) successfully!'
      : 'All existing tables are code blocks and were centered.'
  };
}

/**
 * Formats only the table currently containing cursor or selection
 */
function formatSelectedTable(options) {
  options = options || getTablePreferences();
  saveTablePreferences(options);

  const table = getSelectedTable();
  if (!table) {
    return {
      success: false,
      message: 'Please place your cursor inside a table or select it first.'
    };
  }

  const doc = DocumentApp.getActiveDocument();
  formatSingleTable(table, options, doc.getBody());
  return { success: true, count: 1, message: 'Selected table aligned and formatted!' };
}

/**
 * Locates the table currently enclosing the user cursor or selection
 */
function getSelectedTable() {
  const doc = DocumentApp.getActiveDocument();
  
  // 1. Try cursor position
  const cursor = doc.getCursor();
  if (cursor) {
    let el = cursor.getElement();
    while (el) {
      if (el.getType() === DocumentApp.ElementType.TABLE) {
        return el.asTable();
      }
      el = el.getParent();
    }
  }

  // 2. Try selection
  const selection = doc.getSelection();
  if (selection) {
    const elements = selection.getSelectedElements();
    for (let i = 0; i < elements.length; i++) {
      let el = elements[i].getElement();
      while (el) {
        if (el.getType() === DocumentApp.ElementType.TABLE) {
          return el.asTable();
        }
        el = el.getParent();
      }
    }
  }

  return null;
}

/**
 * Undoes custom formatting on all data tables in the document,
 * reverting borders, backgrounds, padding, and text styling back to defaults.
 * Skips code block tables.
 */
function undoAllTableFormatting() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const tables = body.getTables();

  if (!tables || tables.length === 0) {
    return { success: true, count: 0, message: 'No tables found in this document.' };
  }

  let count = 0;
  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    if (isCodeBlockTable(table) || isHeadingBannerTable(table)) continue;
    undoSingleTableFormatting(table);
    count++;
  }

  return {
    success: true,
    count: count,
    message: count > 0
      ? 'Reset formatting for ' + count + ' table(s) to document defaults.'
      : 'No data tables found to reset.'
  };
}

/**
 * Undoes formatting on the currently selected table.
 */
function undoSelectedTableFormatting() {
  const table = getSelectedTable();
  if (!table) {
    return { success: false, message: 'Please place your cursor inside a table or select it first.' };
  }
  if (isCodeBlockTable(table)) {
    return { success: false, message: 'The selected table is a code block. Use the Code Blocks tab to undo it.' };
  }
  if (isHeadingBannerTable(table)) {
    return { success: false, message: 'The selected table is a heading banner. Use the Document Text & Headings tab to undo it.' };
  }

  undoSingleTableFormatting(table);
  return { success: true, count: 1, message: 'Selected table formatting reset to defaults!' };
}

/**
 * Resets a single table to standard clean document styling:
 * - Borders: 1pt black
 * - Background: white / clear
 * - Padding: standard 5pt
 * - Alignment: left
 * - Font: Arial 10pt, black, regular (non-bold), no background highlights
 */
function undoSingleTableFormatting(table) {
  const numRows = table.getNumRows();
  if (numRows === 0) return;

  table.setBorderWidth(1);
  table.setBorderColor('#000000');

  for (let r = 0; r < numRows; r++) {
    const row = table.getRow(r);
    const numCells = row.getNumCells();
    for (let c = 0; c < numCells; c++) {
      const cell = row.getCell(c);
      cell.setBackgroundColor('#FFFFFF');
      cell.setPaddingTop(5);
      cell.setPaddingBottom(5);
      cell.setPaddingLeft(5);
      cell.setPaddingRight(5);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);

      const numChildren = cell.getNumChildren();
      for (let pIdx = 0; pIdx < numChildren; pIdx++) {
        const child = cell.getChild(pIdx);
        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
          const para = child.asParagraph();
          para.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
          try { para.setFontSize(10); } catch(e) {}
          try { para.setFontFamily('Arial'); } catch(e) {}
          para.setLineSpacing(1.15);

          const textObj = para.editAsText();
          if (textObj.getText().length > 0) {
            try { textObj.setFontFamily('Arial'); } catch(e) {}
            try { textObj.setFontSize(10); } catch(e) {}
            try { textObj.setBold(false); } catch(e) {}
            try { textObj.setForegroundColor('#000000'); } catch(e) {}
            try { textObj.setBackgroundColor(null); } catch(e) {}
          }
        }
      }
    }
  }
}

/**
 * Detects whether a table is a 1x1 code block container
 */
function isCodeBlockTable(table) {
  if (table.getNumRows() !== 1) return false;
  const row = table.getRow(0);
  if (row.getNumCells() !== 1) return false;

  const cell = row.getCell(0);
  const numChildren = cell.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const child = cell.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const text = child.asParagraph().editAsText();
      if (text.getText().length > 0) {
        const font = text.getFontFamily(0);
        if (font && (
          font === 'Consolas' ||
          font === 'JetBrains Mono' ||
          font === 'Courier New' ||
          font === 'Roboto Mono' ||
          font === 'Inconsolata' ||
          font === 'Source Code Pro' ||
          font === 'Space Mono' ||
          font === 'PT Mono' ||
          font === 'Ubuntu Mono' ||
          font.toLowerCase().includes('mono')
        )) {
          return true;
        }
      }
    }
  }
  return false;
}

/**
 * Detects whether a table is a 1x1 heading banner container
 */
function isHeadingBannerTable(table) {
  if (table.getNumRows() !== 1) return false;
  const row = table.getRow(0);
  if (row.getNumCells() !== 1) return false;
  if (isCodeBlockTable(table)) return false;

  const cell = row.getCell(0);
  const numChildren = cell.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const child = cell.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const heading = child.asParagraph().getHeading();
      if (heading && heading !== DocumentApp.ParagraphHeading.NORMAL) {
        return true;
      }
    }
  }
  return false;
}

/**
 * Aligns a table to the exact center of the page by sizing its columns
 * to match the printable width between left and right document margins.
 */
/**
 * Ensures exactly two empty (NORMAL) paragraphs exist immediately after a given
 * table or block element in the body. Adds missing empty paragraphs; trims
 * extras beyond two. This guarantees visual breathing room after code blocks
 * and data tables without creating runaway blank lines.
 */
function ensureDoubleNewlineAfterElement(body, element) {
  try {
    const idx = body.getChildIndex(element);
    if (idx < 0) return;

    const total = body.getNumChildren();
    let blankCount = 0;
    let scanIdx = idx + 1;

    // Count consecutive empty paragraphs right after the element
    while (scanIdx < total) {
      const sibling = body.getChild(scanIdx);
      if (sibling.getType() === DocumentApp.ElementType.PARAGRAPH &&
          sibling.asParagraph().getText().trim() === '') {
        blankCount++;
        scanIdx++;
      } else {
        break;
      }
    }

    // Add missing empty paragraphs
    for (let add = blankCount; add < 2; add++) {
      const newPara = body.insertParagraph(idx + 1 + blankCount + (add - blankCount), '');
      newPara.setFontFamily('Arial');
      newPara.setFontSize(11);
      try { safeSetHeading(newPara, DocumentApp.ParagraphHeading.NORMAL); } catch(e) {}
    }

    // Remove surplus empty paragraphs beyond two
    if (blankCount > 2) {
      for (let rem = blankCount; rem > 2; rem--) {
        try { body.getChild(idx + 1 + rem - 1).removeFromParent(); } catch(e) { break; }
      }
    }
  } catch(e) {}
}

function centerTableOnPage(table, body) {
  const numRows = table.getNumRows();
  if (numRows === 0) return;
  const numCols = table.getRow(0).getNumCells();
  if (numCols === 0) return;

  const pageWidth = body.getPageWidth();
  const marginLeft = body.getMarginLeft();
  const marginRight = body.getMarginRight();
  const printableWidth = Math.max(100, pageWidth - marginLeft - marginRight);

  const optimalWidths = calculateOptimalColumnWidths(table, numCols, numRows, printableWidth);
  for (let c = 0; c < numCols; c++) {
    try {
      table.setColumnWidth(c, optimalWidths[c]);
    } catch(e) {}
  }
}

/**
 * Calculates balanced column widths to span printableWidth symmetrically.
 * Preserves user manual column adjustments proportionally, or allocates
 * based on text length with guaranteed minimum bounds.
 */
function calculateOptimalColumnWidths(table, numCols, numRows, printableWidth) {
  let existingTotal = 0;
  const existingWidths = [];
  let allHaveWidths = true;

  for (let c = 0; c < numCols; c++) {
    const w = table.getColumnWidth(c);
    if (!w || w <= 0) {
      allHaveWidths = false;
      break;
    }
    existingWidths.push(w);
    existingTotal += w;
  }

  // Preserve existing column proportions if all columns had explicit widths
  if (allHaveWidths && existingTotal > 0) {
    const scale = printableWidth / existingTotal;
    return existingWidths.map(w => Math.round(w * scale * 10) / 10);
  }

  // Calculate text weights based on content length
  const colCharCounts = new Array(numCols).fill(0);
  for (let r = 0; r < numRows; r++) {
    const row = table.getRow(r);
    const cellsCount = row.getNumCells();
    for (let c = 0; c < Math.min(numCols, cellsCount); c++) {
      const textLen = row.getCell(c).getText().trim().length;
      colCharCounts[c] = Math.max(colCharCounts[c], textLen);
    }
  }

  // Base weight guarantees minimum readable width; capped variance prevents extreme skew
  const weights = colCharCounts.map(count => 16 + Math.min(count, 48));
  const totalWeight = weights.reduce((sum, w) => sum + w, 0);

  return weights.map(w => Math.round((w / totalWeight) * printableWidth * 10) / 10);
}

/**
 * Core table formatting:
 * - Table aligned in the middle of the page (centered).
 * - Table Header text in the middle (center-aligned) with LARGER bold font.
 * - Table Body text LEFT-ALIGNED (not in middle or right) with standard font.
 * - Font Family applied across all cells.
 * - Colorful header background, alternating row zebra striping, and clean borders.
 */
function formatSingleTable(table, options, body) {
  const numRows = table.getNumRows();
  if (numRows === 0) return;

  const firstRow = table.getRow(0);
  const numCols = firstRow.getNumCells();
  if (numCols === 0) return;

  // 1. Center table across printable margins
  centerTableOnPage(table, body);

  // 2. Borders
  const borderWidth = options.borderWidth !== undefined ? Number(options.borderWidth) : 1;
  const borderColor = options.borderColor || '#93C5FD';
  table.setBorderWidth(borderWidth);
  table.setBorderColor(borderColor);

  // 3. Spacing & Padding
  let padTop = 7, padBottom = 7, padLeft = 10, padRight = 10;
  if (options.padding === 'compact') {
    padTop = 4; padBottom = 4; padLeft = 6; padRight = 6;
  } else if (options.padding === 'relaxed') {
    padTop = 10; padBottom = 10; padLeft = 14; padRight = 14;
  }

  // 4. Color Palette & Typography settings
  const headerBg = options.headerBg || '#1E3A8A';
  const headerText = options.headerText || '#FFFFFF';
  const altRowBg = options.altRowBg || '#F0F7FF';
  const normalRowBg = options.normalRowBg || '#FFFFFF';
  const textColor = options.textColor || '#1E293B';
  const fontFamily = options.fontFamily || 'Roboto';
  const headerFontSize = Number(options.headerFontSize) || 11;
  const bodyFontSize = Number(options.bodyFontSize) || 9.5;

  // 5. Apply row and cell styles
  for (let r = 0; r < numRows; r++) {
    const row = table.getRow(r);
    const isHeader = (r === 0);
    // Header gets headerBg; subsequent rows alternate for clear zebra striping
    const rowBg = isHeader ? headerBg : (r % 2 === 0 ? altRowBg : normalRowBg);

    const cellsCount = row.getNumCells();
    for (let c = 0; c < cellsCount; c++) {
      const cell = row.getCell(c);

      // Cell background, padding & vertical alignment (Always Top-aligned)
      cell.setBackgroundColor(rowBg);
      cell.setPaddingTop(isHeader ? padTop + 2 : padTop);
      cell.setPaddingBottom(isHeader ? padBottom + 2 : padBottom);
      cell.setPaddingLeft(padLeft);
      cell.setPaddingRight(padRight);
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);

      // Cell Paragraph Content & Typography
      const numChildren = cell.getNumChildren();
      for (let pIdx = 0; pIdx < numChildren; pIdx++) {
        const child = cell.getChild(pIdx);
        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
          const para = child.asParagraph();

          // Alignment Rule:
          // - Header text: MUST be in the middle (center-aligned)
          // - Table text (data rows): MUST be left-aligned (not in middle or right)
          if (isHeader) {
            para.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
            try { para.setFontSize(headerFontSize); } catch(e) {}
            try { para.setFontFamily(fontFamily); } catch(e) {}
          } else {
            para.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
            try { para.setFontSize(bodyFontSize); } catch(e) {}
            try { para.setFontFamily(fontFamily); } catch(e) {}
          }

          para.setLineSpacing(1.15);

          // Font color, size, family & weight: Header is BIGGER and BOLD
          const textObj = para.editAsText();
          if (textObj.getText().length > 0) {
            try { textObj.setFontFamily(fontFamily); } catch(e) {}
            if (isHeader) {
              textObj.setBold(true);
              try { textObj.setBold(0, textObj.getText().length - 1, true); } catch(e) {}
              textObj.setForegroundColor(headerText);
              try { textObj.setFontSize(headerFontSize); } catch(e) {}
            } else {
              textObj.setBold(false);
              textObj.setForegroundColor(textColor);
              try { textObj.setFontSize(bodyFontSize); } catch(e) {}
              // Highlight any inline code text found in data row cells
              if (options.inlineCodeHighlight !== false) {
                applyInlineTableCodeHighlight(textObj, options);
              }
            }
          }
        }
      }
    }
  }

  // Ensure two empty paragraphs follow this table for readability
  ensureDoubleNewlineAfterElement(body, table);
}

/**
 * Scans a table cell's Text object for code-like tokens and applies
 * monospace font + a subtle highlight background to those inline spans.
 *
 * Detects:
 *   - Backtick-wrapped text:  `code`
 *   - Programming keywords:   function, const, let, var, return, import, etc.
 *   - Method/property calls:  object.method(), array.length
 *   - Identifiers with parens: myFunc(), render()
 *   - File paths / extensions: file.js, index.html, style.css
 *   - CLI commands:           npm install, git push, python run.py
 *   - Type annotations:       string, boolean, number, null, undefined, true, false
 */
function applyInlineTableCodeHighlight(textObj, options) {
  const fullText = textObj.getText();
  if (!fullText || fullText.length === 0) return;

  const codeFont = 'Consolas';
  const codeBg = options.inlineCodeBg || '#EFF1F3';
  const codeColor = options.inlineCodeColor || '#B45309';

  // Pattern groups (ordered: most-specific first)
  const codePatterns = [
    // 1. Backtick-wrapped: `any code here` — highest priority
    /`[^`]+`/g,
    // 2. File names with known code extensions
    /\b[\w./\-]+\.(?:js|ts|py|html|css|json|jsx|tsx|sh|bash|md|sql|java|cpp|cs|php|rb|go|rs|kt|swift|yaml|yml|xml|env)\b/g,
    // 3. Method/property chains: obj.method() or object.prop
    /\b[A-Za-z_$][\w$]*(?:\.[A-Za-z_$][\w$]*)+(?:\(\))?/g,
    // 4. Function calls: myFunction()
    /\b[A-Za-z_$][\w$]*\(\)/g,
    // 5. CLI commands at start or after space
    /(?:^|(?<=[\s,]))(?:npm|yarn|pnpm|pip|pip3|python|python3|node|npx|git|docker|curl|wget|bash|sh|cd|ls|mkdir|rm|cp|mv|cat|echo|export|source)\s+[\w./\-@=<>"'\s]+/gm,
    // 6. Programming keywords (word-boundary)
    /\b(?:function|const|let|var|return|import|export|class|extends|async|await|if|else|for|while|switch|try|catch|finally|new|this|typeof|instanceof|void|delete|in|of|from|default|static|public|private|protected|abstract|interface|enum|type|def|lambda|print|pass|raise|yield|with|as|nil|None|True|False|null|undefined|true|false)\b/g,
    // 7. Type-like tokens: string, number, boolean, any, unknown, never, void
    /\b(?:string|number|boolean|any|unknown|never|void|object|array|tuple|Record|Promise|Map|Set|List|Dict|int|float|double|char|long|short|byte|uint)\b/g,
    // 8. Common UPPERCASE constants / env vars
    /\b[A-Z][A-Z0-9_]{2,}\b/g
  ];

  // We collect all matches as {start, end} spans, then apply styles
  // Using a coverage map to avoid double-applying overlapping ranges
  const covered = [];

  codePatterns.forEach(pattern => {
    pattern.lastIndex = 0;
    let match;
    while ((match = pattern.exec(fullText)) !== null) {
      const start = match.index;
      const end = start + match[0].length - 1;
      if (end < start) continue;

      // Check for significant overlap with already-covered spans
      const overlaps = covered.some(s =>
        !(end < s.start || start > s.end)
      );
      if (!overlaps) {
        covered.push({ start, end });
        try {
          textObj.setFontFamily(start, end, codeFont);
          textObj.setForegroundColor(start, end, codeColor);
          textObj.setBackgroundColor(start, end, codeBg);
        } catch(e) {}
      }
    }
  });
}

/**
 * Preferences for Table Formatter
 */
function saveTablePreferences(prefs) {
  PropertiesService.getUserProperties().setProperty('table_formatter_prefs', JSON.stringify(prefs));
}

function getTablePreferences() {
  const saved = PropertiesService.getUserProperties().getProperty('table_formatter_prefs');
  if (saved) {
    try { return JSON.parse(saved); } catch(e) {}
  }
  return {
    theme: 'corporate-navy',
    headerBg: '#1E3A8A',
    headerText: '#FFFFFF',
    altRowBg: '#F0F7FF',
    normalRowBg: '#FFFFFF',
    borderColor: '#93C5FD',
    textColor: '#1E293B',
    fontFamily: 'Roboto',
    headerFontSize: 11,
    bodyFontSize: 9.5,
    padding: 'normal',
    borderWidth: 1,
    inlineCodeHighlight: true,
    inlineCodeBg: '#EFF1F3',
    inlineCodeColor: '#B45309'
  };
}

/* ==========================================================================
   CODE HIGHLIGHTER ENGINE
   ========================================================================== */

function quickHighlightAll() {
  const prefs = getUserPreferences();
  const result = highlightAllCodeBlocks(prefs);
  DocumentApp.getUi().alert('⚡ Code Highlighter', result.message, DocumentApp.getUi().ButtonSet.OK);
}

function quickIndentAllCode() {
  const prefs = getUserPreferences();
  const result = indentAllCodeBlocks(prefs);
  DocumentApp.getUi().alert('⚡ Code Indentation', result.message, DocumentApp.getUi().ButtonSet.OK);
}

function quickUndoAllCodeBlocks() {
  const result = undoAllCodeBlocksFormatting();
  DocumentApp.getUi().alert('⚡ Code Highlighter', result.message, DocumentApp.getUi().ButtonSet.OK);
}

/**
 * Core scanning & formatting function for code blocks
 */
function highlightAllCodeBlocks(options) {
  options = options || getUserPreferences();
  saveUserPreferences(options);

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();
  
  let codeGroups = [];
  let inCode = false;
  let inFencedCode = false;
  let braceDepth = 0;
  let classContext = false; // true when inside a class/struct/pseudocode class body
  let currentGroup = [];
  let blankBuffer = [];

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    
    // Skip items already inside a table cell
    if (p.getParent().getType() === DocumentApp.ElementType.TABLE_CELL) {
      if (inCode) finishGroup();
      continue;
    }

    const text = p.getText();
    const trimmed = text.trim();

    // 1. Fenced Code Block Handler: ``` or ```lang
    if (/^```/.test(trimmed)) {
      if (!inFencedCode) {
        if (inCode) finishGroup();
        inFencedCode = true;
        inCode = true;
        currentGroup.push(p);
        continue;
      } else {
        currentGroup.push(p);
        inFencedCode = false;
        finishGroup();
        continue;
      }
    }

    if (inFencedCode) {
      currentGroup.push(p);
      continue;
    }

      // 2. Syntax-based / Indented Code Block Handler
    if (!inCode) {
      if (isCodeStart(text)) {
        inCode = true;
        currentGroup.push(p);
        blankBuffer = [];
        braceDepth = Math.max(0, getNetBraceCount(text));
        // Detect class/pseudocode context (e.g. "class Invoice:", "class Foo:")
        if (/^(class|interface|struct|enum)\s+\w.*:?\s*$/.test(trimmed)) {
          classContext = true;
        }
      }
    } else {
      // 1. LOCKED BY CURLY BRACES
      if (braceDepth > 0) {
        if (blankBuffer.length > 0) {
          currentGroup.push(blankBuffer[0]); // at most 1 blank line carried forward
          blankBuffer = [];
        }
        currentGroup.push(p);
        braceDepth = Math.max(0, braceDepth + getNetBraceCount(text));
        continue;
      }

      // 2. OUTSIDE BRACES (braceDepth === 0)
      if (trimmed.length === 0) {
        blankBuffer.push(p);
        // Close if more than 3 consecutive blank lines
        if (blankBuffer.length > 3) finishGroup();
      } else if (isEllipsis(trimmed)) {
        // Ellipsis / continuation marker (... or …) is always part of pseudocode.
        // Only carry forward 1 blank line max to avoid huge gaps.
        if (blankBuffer.length > 0) {
          currentGroup.push(blankBuffer[0]); // at most 1 blank line per gap
          blankBuffer = [];
        }
        currentGroup.push(p);
      } else if (isStrongProse(text)) {
        // Strong prose ALWAYS closes the block — even inside a class context.
        // This prevents body-text paragraphs from being absorbed into code.
        finishGroup();
      } else if (isCodeLine(text)) {
        // Carry forward at most 1 blank line per gap to avoid excess spacing
        if (blankBuffer.length > 0) {
          currentGroup.push(blankBuffer[0]);
          blankBuffer = [];
        }
        currentGroup.push(p);
        braceDepth = Math.max(0, braceDepth + getNetBraceCount(text));
        // Extend class context on nested function/method declarations
        if (/^(function|def|method|procedure|constructor)\s+\w.*[(:)]/.test(trimmed)) {
          classContext = true;
        }
      } else if (classContext && /^[A-Za-z_$][\w$]*$/.test(trimmed)) {
        // Bare single-word identifier inside a class body (e.g. field names 'items', 'customer').
        // Only absorb when classContext is active and line is a simple identifier.
        if (blankBuffer.length > 0) {
          currentGroup.push(blankBuffer[0]);
          blankBuffer = [];
        }
        currentGroup.push(p);
      } else {
        finishGroup();
      }
    }
  }

  if (inCode) finishGroup();

  function finishGroup() {
    while (currentGroup.length > 0 && currentGroup[currentGroup.length - 1].getText().trim().length === 0) {
      currentGroup.pop();
    }
    if (currentGroup.length > 0) {
      codeGroups.push(currentGroup);
    }
    currentGroup = [];
    blankBuffer = [];
    inCode = false;
    inFencedCode = false;
    braceDepth = 0;
    classContext = false; // reset class/pseudocode context on group close
  }

  if (codeGroups.length === 0) {
    return { success: true, count: 0, message: 'No unformatted code blocks found.' };
  }

  if (body.getParagraphs().length <= paragraphs.length) {
    body.appendParagraph('');
  }

  // Reverse loop (bottom-to-top) to preserve paragraph indices
  for (let g = codeGroups.length - 1; g >= 0; g--) {
    convertParagraphsToCodeBlock(body, codeGroups[g], options);
  }

  return { success: true, count: codeGroups.length, message: 'Formatted ' + codeGroups.length + ' code block(s)!' };
}

function getNetBraceCount(text) {
  if (!text) return 0;
  let clean = text.replace(/\/\*[\s\S]*?\*\/|\/\/.*$|#.*$/gm, '');
  clean = clean.replace(/(["'`])(?:(?=(\\?))\2[\s\S])*?\1/g, '');
  const openCount = (clean.match(/\{/g) || []).length;
  const closeCount = (clean.match(/\}/g) || []).length;
  return openCount - closeCount;
}

function isEllipsis(trimmed) {
  // Matches literal '...', the Unicode ellipsis '…', or '--' standalone continuation markers
  return trimmed === '...' || trimmed === '\u2026' || trimmed === '--' || /^\.{2,6}$/.test(trimmed);
}

function isCodeStart(text) {
  const trimmed = text.trim();
  if (!trimmed) return false;
  if (/^```/.test(trimmed)) return true;
  if (isStrongProse(trimmed)) return false;

  // Never match markdown headings as code comments
  if (/^#{1,6}\s+/.test(trimmed)) return false;

  // Pseudocode / class declarations with colon terminator (e.g. "class Invoice:", "function foo():")
  if (/^(class|interface|struct|enum|record)\s+\w[\w\s<>,]*:?\s*$/.test(trimmed)) return true;
  if (/^(function|def|method|procedure|constructor|fn)\s+\w[\w\s(),<>]*:?\s*$/.test(trimmed)) return true;

  const startPatterns = [
    /^(def|class|function|const|let|var|import|export|public|private|protected|static|package|namespace)\b/,
    /^(if|for|while|switch|catch|with|elif)\s*(\(|:)/,
    /^(try|except|finally|else):?$/,
    /^(SELECT|INSERT|UPDATE|DELETE|CREATE|ALTER|DROP|FROM|WHERE)\b/i,
    /^\s*(\/\/|\/\*|\*\/|<!--)/,
    /^\s*#(?!\s+[A-Za-z0-9])|^\s*#\s*$/, // Code comment, but NOT # Heading
    /[A-Za-z0-9_$]+\.[A-Za-z0-9_$]+\(/,
    /(=>|===|!==|\+=|-=|\+\+|--|&&|\|\|)/,
    /^[A-Za-z0-9_$]+\s*=\s*.+/,
    /[{};]$/,
    /^\s{2,}|\t/,
    /^(npm|yarn|pnpm|pip|pip3|python|python3|node|docker|git|curl)\s+/
  ];
  return startPatterns.some(rx => rx.test(trimmed));
}

function isCodeLine(text) {
  const trimmed = text.trim();
  if (!trimmed) return true;
  if (/^```/.test(trimmed)) return true;
  if (isEllipsis(trimmed)) return true;  // ... and … are always code continuations
  if (isStrongProse(trimmed)) return false;
  if (/^\s*[}\])];?$/.test(trimmed)) return true;
  if (/^\s*\.[A-Za-z0-9_$]+/.test(trimmed)) return true;
  if (/^\s{2,}|\t/.test(text)) return true;
  if (/^(return|break|continue|pass|throw|export|default)\b/.test(trimmed)) return true;
  if (/[{};]$/.test(trimmed)) return true;
  // Colon-terminated identifiers (pseudocode method/property declarations)
  if (/^[A-Za-z_$][\w$]*\s*[:(]/.test(trimmed)) return true;
  // Single bare identifier (field name in a class body like 'items', 'customer')
  if (/^[A-Za-z_$][\w$]*$/.test(trimmed)) return true;
  return isCodeStart(text);
}

function isStrongProse(text) {
  const trimmed = text.trim();
  if (!trimmed) return false;
  if (/^```/.test(trimmed)) return false;
  if (isEllipsis(trimmed)) return false;  // never treat ellipsis as prose
  if (/^(def|class|function|const|let|var|import|export|return|public|private|static|if|for|while)\b/.test(trimmed)) return false;
  if (/[{};]$/.test(trimmed) || /(=>|===|!==)/.test(trimmed)) return false;
  if (/^\s{2,}|\t/.test(text)) return false;
  // Single-word lines are never strong prose (could be variable/field names)
  if (!/\s/.test(trimmed)) return false;

  const words = trimmed.toLowerCase().split(/\s+/);
  const stopWords = ['the', 'this', 'that', 'with', 'and', 'are', 'can', 'you', 'we', 'our', 'will', 'should', 'about', 'there', 'from', 'which', 'because', 'also', 'here', 'when', 'then', 'into', 'have', 'been', 'would', 'could'];
  const proseCount = words.filter(w => stopWords.includes(w.replace(/[^a-z]/g, ''))).length;

  if (proseCount >= 2 && /[.!?:]$/.test(trimmed)) return true;
  if (proseCount >= 4) return true;
  return false;
}

/**
 * Normalizes code and pseudocode line spacing to professional standards:
 * - Trims leading and trailing empty lines from the code block.
 * - Collapses consecutive empty lines into at most one empty line.
 * - Removes unnecessary blank lines right after opening statements (class, function, def, opening braces/colons).
 * - Removes unnecessary blank lines right before closing braces/brackets.
 * - Normalizes spacing around pseudocode continuation markers (... or …) so they don't produce excessive vertical gaps.
 */
function normalizeCodeSpacing(lines) {
  if (!lines || lines.length === 0) return [''];

  // 1. Strip leading and trailing empty lines
  let start = 0;
  while (start < lines.length && lines[start].trim().length === 0) {
    start++;
  }
  let end = lines.length - 1;
  while (end >= start && lines[end].trim().length === 0) {
    end--;
  }
  if (start > end) return [''];
  const trimmedLines = lines.slice(start, end + 1);

  // 2. Collapse consecutive blank lines and strip redundant blanks
  const cleaned = [];
  for (let i = 0; i < trimmedLines.length; i++) {
    const raw = trimmedLines[i];
    const trimmed = raw.trim();
    const isBlank = trimmed.length === 0;

    if (isBlank) {
      // Never allow multiple consecutive blank lines
      if (cleaned.length > 0 && cleaned[cleaned.length - 1].trim().length === 0) {
        continue;
      }
      // Never allow a blank line immediately after class / function declaration or opening bracket/colon
      if (cleaned.length > 0) {
        const prev = cleaned[cleaned.length - 1].trim();
        if (/^(class|interface|struct|enum|record|def|function|method|procedure)\s+\w.*:?\s*$/.test(prev) ||
            /[:{[(]$/.test(prev)) {
          continue;
        }
      }
      cleaned.push('');
    } else {
      // Never allow a blank line immediately before a closing bracket
      if (/^[}\])]$/.test(trimmed)) {
        if (cleaned.length > 0 && cleaned[cleaned.length - 1].trim().length === 0) {
          cleaned.pop();
        }
      }
      cleaned.push(raw);
    }
  }

  // 3. Normalize spacing around ellipsis / continuation markers (... or …)
  const result = [];
  for (let i = 0; i < cleaned.length; i++) {
    const raw = cleaned[i];
    const trimmed = raw.trim();
    if (isEllipsis(trimmed)) {
      // If preceded by a blank line, remove that blank line
      if (result.length > 0 && result[result.length - 1].trim().length === 0) {
        result.pop();
      }
      result.push(raw);
      // If followed by a blank line, skip that blank line
      if (i + 1 < cleaned.length && cleaned[i + 1].trim().length === 0) {
        i++;
      }
      continue;
    }
    result.push(raw);
  }

  return result.length > 0 ? result : [''];
}

/**
 * Creates and styles the code table container centered in the document
 */
function convertParagraphsToCodeBlock(body, paragraphGroup, options) {
  const firstParagraph = paragraphGroup[0];
  const insertIndex = body.getChildIndex(firstParagraph);

  const table = body.insertTable(insertIndex);
  table.setBorderWidth(1);
  table.setBorderColor(options.borderColor || '#D0D7DE');
  
  // Center code container across document margins
  try {
    const pageWidth = body.getPageWidth();
    const marginLeft = body.getMarginLeft();
    const marginRight = body.getMarginRight();
    table.setColumnWidth(0, Math.max(100, pageWidth - marginLeft - marginRight));
  } catch(e) {}

  const cell = table.appendTableRow().appendTableCell();
  cell.setBackgroundColor(options.bgColor || '#F6F8FA');
  cell.setPaddingTop(8);
  cell.setPaddingBottom(8);
  cell.setPaddingLeft(12);
  cell.setPaddingRight(12);

  const fontSize = Number(options.fontSize) || 9.5;
  const fontFamily = options.fontFamily || 'Consolas';
  const textColor = options.textColor || '#24292F';

  // Apply indentation handling
  let rawLines = paragraphGroup.map(p => p.getText());
  if (rawLines.length >= 2 && /^```/.test(rawLines[0].trim()) && /^```/.test(rawLines[rawLines.length - 1].trim())) {
    rawLines = rawLines.slice(1, -1);
  } else if (rawLines.length >= 1 && /^```/.test(rawLines[0].trim())) {
    rawLines = rawLines.slice(1);
  } else if (rawLines.length >= 1 && /^```/.test(rawLines[rawLines.length - 1].trim())) {
    rawLines = rawLines.slice(0, -1);
  }
  if (rawLines.length === 0) rawLines = [''];

  // Normalize code spacing: remove extra newlines, collapse blanks, clean up pseudocode
  rawLines = normalizeCodeSpacing(rawLines);

  const indentStyle = options.indentStyle || 'auto-2';
  const formattedLines = formatCodeIndentation(rawLines, indentStyle);

  formattedLines.forEach((textLine, idx) => {
    let line;
    if (idx === 0) {
      line = cell.getChild(0).asParagraph();
      line.setText(textLine);
    } else {
      line = cell.appendParagraph(textLine);
    }
    line.setFontFamily(fontFamily);
    line.setFontSize(fontSize);
    line.setLineSpacing(1.15);
    line.setSpacingBefore(0);
    line.setSpacingAfter(0);
    line.setForegroundColor(textColor);
    
    const textObj = line.editAsText();
    if (textObj.getText().length > 0) {
      textObj.setBold(false);
      try { textObj.setBold(0, textObj.getText().length - 1, false); } catch(e) {}
    }
    
    applySyntaxHighlight(textObj, options);
  });

  paragraphGroup.forEach(p => {
    p.removeFromParent();
  });

  // Ensure two empty paragraphs follow the code block for readability
  ensureDoubleNewlineAfterElement(body, table);
}

function formatSelectedCodeBlock(options) {
  options = options || getUserPreferences();
  const selection = DocumentApp.getActiveDocument().getSelection();
  if (!selection) {
    return { success: false, message: 'Please highlight/select the text first.' };
  }

  const elements = selection.getSelectedElements();
  const paragraphs = [];
  elements.forEach(el => {
    let element = el.getElement();
    if (element.getType() === DocumentApp.ElementType.TEXT) element = element.getParent();
    if (element.getType() === DocumentApp.ElementType.PARAGRAPH) paragraphs.push(element);
  });

  if (paragraphs.length > 0) {
    convertParagraphsToCodeBlock(DocumentApp.getActiveDocument().getBody(), paragraphs, options);
    return { success: true, count: 1, message: 'Selected code formatted!' };
  }
  return { success: false, message: 'No valid paragraphs selected.' };
}

/**
 * Undoes all code blocks formatting across the document,
 * converting 1x1 code block tables back into regular document paragraphs.
 */
function undoAllCodeBlocksFormatting() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const tables = body.getTables();

  if (!tables || tables.length === 0) {
    return { success: true, count: 0, message: 'No code blocks found in this document.' };
  }

  let count = 0;
  // Iterate backwards to preserve child index positions when removing tables
  for (let i = tables.length - 1; i >= 0; i--) {
    const table = tables[i];
    if (isCodeBlockTable(table)) {
      undoSingleCodeBlock(body, table);
      count++;
    }
  }

  return {
    success: true,
    count: count,
    message: count > 0
      ? 'Reverted ' + count + ' code block(s) back to standard document paragraphs.'
      : 'No formatted code blocks found.'
  };
}

/**
 * Undoes formatting for the selected code block.
 */
function undoSelectedCodeBlockFormatting() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const table = getSelectedTable();

  if (!table || !isCodeBlockTable(table)) {
    return { success: false, message: 'Please place your cursor inside a formatted code block first.' };
  }

  undoSingleCodeBlock(body, table);
  return { success: true, count: 1, message: 'Selected code block reverted to standard text!' };
}

/**
 * Converts a 1x1 code block table back into normal document paragraphs
 */
function undoSingleCodeBlock(body, table) {
  const insertIndex = body.getChildIndex(table);
  const cell = table.getRow(0).getCell(0);
  const numChildren = cell.getNumChildren();

  const lines = [];
  for (let i = 0; i < numChildren; i++) {
    const child = cell.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const text = child.asParagraph().getText();
      // Skip empty first paragraph if dummy placeholder
      if (i === 0 && text === '' && numChildren > 1) {
        continue;
      }
      lines.push(text);
    }
  }

  if (lines.length === 0) {
    lines.push('');
  }

  // Insert standard paragraphs at the table's index
  for (let i = 0; i < lines.length; i++) {
    const p = body.insertParagraph(insertIndex + i, lines[i]);
    p.setFontFamily('Arial');
    p.setFontSize(11);
    p.setLineSpacing(1.15);
    p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    const textObj = p.editAsText();
    if (textObj.getText().length > 0) {
      try { textObj.setFontFamily('Arial'); } catch(e) {}
      try { textObj.setFontSize(11); } catch(e) {}
      try { textObj.setBold(false); } catch(e) {}
      try { textObj.setForegroundColor('#000000'); } catch(e) {}
      try { textObj.setBackgroundColor(null); } catch(e) {}
    }
  }

  table.removeFromParent();
}

/**
 * Auto-indents all formatted code blocks in the document according to options.indentStyle
 */
function indentAllCodeBlocks(options) {
  options = options || getUserPreferences();
  const indentStyle = options.indentStyle || 'auto-2';
  saveUserPreferences(options);

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const tables = body.getTables();

  if (!tables || tables.length === 0) {
    return { success: true, count: 0, message: 'No code blocks found in this document.' };
  }

  let count = 0;
  for (let i = 0; i < tables.length; i++) {
    const table = tables[i];
    if (isCodeBlockTable(table)) {
      indentSingleCodeBlock(table, indentStyle, options);
      count++;
    }
  }

  return {
    success: true,
    count: count,
    message: count > 0
      ? 'Indented ' + count + ' code block(s) with ' + getIndentLabel(indentStyle) + '!'
      : 'No code blocks found to indent.'
  };
}

/**
 * Auto-indents the selected code block
 */
function indentSelectedCodeBlock(options) {
  options = options || getUserPreferences();
  const indentStyle = options.indentStyle || 'auto-2';
  saveUserPreferences(options);

  const table = getSelectedTable();
  if (!table || !isCodeBlockTable(table)) {
    return { success: false, message: 'Please place your cursor inside a formatted code block first.' };
  }

  indentSingleCodeBlock(table, indentStyle, options);
  return { success: true, count: 1, message: 'Selected code block indented with ' + getIndentLabel(indentStyle) + '!' };
}

/**
 * Indents a single code block in place
 */
function indentSingleCodeBlock(table, indentStyle, options) {
  const cell = table.getRow(0).getCell(0);
  const numChildren = cell.getNumChildren();
  const paragraphs = [];
  const rawLines = [];

  for (let i = 0; i < numChildren; i++) {
    const child = cell.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = child.asParagraph();
      // Skip empty first paragraph if dummy placeholder
      if (i === 0 && p.getText() === '' && numChildren > 1) {
        continue;
      }
      paragraphs.push(p);
      rawLines.push(p.getText());
    }
  }

  if (rawLines.length === 0) return;

  const indentedLines = formatCodeIndentation(rawLines, indentStyle);
  const fontSize = Number(options.fontSize) || 9.5;
  const fontFamily = options.fontFamily || 'Consolas';
  const textColor = options.textColor || '#24292F';

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    p.setText(indentedLines[i]);
    p.setFontFamily(fontFamily);
    p.setFontSize(fontSize);
    p.setLineSpacing(1.15);
    p.setSpacingBefore(0);
    p.setSpacingAfter(0);
    p.setForegroundColor(textColor);
    const textObj = p.editAsText();
    if (textObj.getText().length > 0) {
      textObj.setBold(false);
      try { textObj.setBold(0, textObj.getText().length - 1, false); } catch(e) {}
    }
    applySyntaxHighlight(textObj, options);
  }
}

/**
 * Formats code indentation across an array of text lines
 */
function formatCodeIndentation(textLines, indentStyle) {
  if (!textLines || textLines.length === 0) return [];
  if (!indentStyle || indentStyle === 'keep') return textLines;

  if (indentStyle === 'tab-2') {
    return textLines.map(line => line.replace(/\t/g, '  '));
  }
  if (indentStyle === 'tab-4') {
    return textLines.map(line => line.replace(/\t/g, '    '));
  }

  const unit = (indentStyle === 'auto-4' || indentStyle === '4-spaces') ? '    ' : '  ';

  let hasExistingIndent = false;
  for (let i = 0; i < textLines.length; i++) {
    if (/^\s+/.test(textLines[i]) && textLines[i].trim().length > 0) {
      hasExistingIndent = true;
      break;
    }
  }

  if (hasExistingIndent) {
    return normalizeExistingIndent(textLines, unit);
  } else {
    return smartIndentFlatLines(textLines, unit);
  }
}

/**
 * Normalizes existing irregular leading whitespace and tabs to clean unit levels
 */
function normalizeExistingIndent(textLines, unit) {
  const unTabbed = textLines.map(line => line.replace(/\t/g, '  '));
  let minIndent = 0;
  for (let i = 0; i < unTabbed.length; i++) {
    const match = unTabbed[i].match(/^ +/);
    if (match) {
      const len = match[0].length;
      if (minIndent === 0 || (len < minIndent && len > 0)) {
        minIndent = len;
      }
    }
  }
  if (minIndent === 0) minIndent = 2;

  return unTabbed.map(line => {
    if (line.trim().length === 0) return '';
    const match = line.match(/^( +)/);
    if (!match) return line.trim();
    const spaces = match[1].length;
    const level = Math.round(spaces / minIndent);
    return unit.repeat(level) + line.trim();
  });
}

/**
 * Applies syntax and bracket-based smart auto-indentation to flat unindented code
 */
function smartIndentFlatLines(textLines, unit) {
  let level = 0;
  const result = [];

  for (let i = 0; i < textLines.length; i++) {
    const raw = textLines[i];
    const trimmed = raw.trim();
    if (trimmed.length === 0) {
      result.push('');
      continue;
    }

    const clean = trimmed
      .replace(/\/\*[\s\S]*?\*\/|\/\/.*$|#.*$/gm, '')
      .replace(/(["'`])(?:(?=(\\?))\2[\s\S])*?\1/g, '');

    const startsClosing = /^[}\])]|^(<\/[a-zA-Z0-9_-]+>)/.test(clean) ||
                          /^(else|elif|catch|finally|except)\b/.test(clean);

    if (startsClosing) {
      level = Math.max(0, level - 1);
    }

    // Step back to member level for sibling function/method declarations in pseudocode or Python
    if (/^(function|def|method|procedure)\s+\w.*[(:]/.test(trimmed)) {
      if (level >= 2) level = 1;
    }

    result.push(unit.repeat(level) + trimmed);

    const opens = (clean.match(/[{[(]|<[a-zA-Z0-9_-]+(?:\s+[^>]*?)?(?<!\/)>/g) || []).length;
    const closes = (clean.match(/[}\])]|<\/[a-zA-Z0-9_-]+>/g) || []).length;
    let net = opens - closes;

    if (/:\s*$/.test(clean) && !startsClosing && opens === closes) {
      net = 1;
    }

    if (startsClosing) {
      level = Math.max(0, level + net + 1);
    } else {
      level = Math.max(0, level + net);
    }
  }

  return result;
}

function getIndentLabel(style) {
  switch (style) {
    case 'auto-2': return 'Smart 2-Space Indent';
    case 'auto-4': return 'Smart 4-Space Indent';
    case 'tab-2': return '2 Spaces (Tabs Converted)';
    case 'tab-4': return '4 Spaces (Tabs Converted)';
    case 'keep': return 'Original Indentation';
    default: return '2 Spaces';
  }
}

/**
 * Universal theme palettes for syntax highlighting
 */
const CODE_THEMES = {
  'github-light': {
    bg: '#F6F8FA', text: '#24292F', border: '#D0D7DE',
    kw: '#CF222E', str: '#0A3069', com: '#6E7781', num: '#0550AE',
    fn: '#8250DF', type: '#953800', bool: '#CF222E', special: '#0550AE', op: '#0550AE'
  },
  'one-dark': {
    bg: '#21252B', text: '#ABB2BF', border: '#3B4048',
    kw: '#C678DD', str: '#98C379', com: '#5C6370', num: '#D19A66',
    fn: '#61AFEF', type: '#E5C07B', bool: '#D19A66', special: '#E06C75', op: '#56B6C2'
  },
  'dracula': {
    bg: '#282A36', text: '#F8F8F2', border: '#44475A',
    kw: '#FF79C6', str: '#F1FA8C', com: '#6272A4', num: '#BD93F9',
    fn: '#50FA7B', type: '#8BE9FD', bool: '#BD93F9', special: '#FFB86C', op: '#FF79C6'
  },
  'monokai': {
    bg: '#272822', text: '#F8F8F2', border: '#3E3D32',
    kw: '#F92672', str: '#E6DB74', com: '#75715E', num: '#AE81FF',
    fn: '#A6E22E', type: '#66D9EF', bool: '#AE81FF', special: '#FD971F', op: '#F92672'
  },
  'solarized-light': {
    bg: '#FDF6E3', text: '#657B83', border: '#EEE8D5',
    kw: '#859900', str: '#2AA198', com: '#93A1A1', num: '#D33682',
    fn: '#268BD2', type: '#B58900', bool: '#CB4B16', special: '#268BD2', op: '#859900'
  }
};

/**
 * Master Syntax Highlighter: applies rich, multi-token syntax highlighting
 * with non-overlapping range tracking to ensure clean colors without bleed.
 * Supports:
 * - Comments: single-line //, #, and block /* ... * /
 * - Strings: double, single, template literals
 * - Keywords: control flow, declarations, storage modifiers, SQL
 * - Special keywords: this, self, super
 * - Types & Classes: primitives (int, bool, string, void) and PascalCase classes (Invoice, Customer)
 * - Functions & Methods: calls foo(), declarations, pseudocode method headers
 * - Booleans, Null & Constants: true, false, null, undefined, ALL_CAPS
 * - Numbers: hex, binary, float, integer
 * - Operators: =>, ===, !==, +, -, *, /, &&, ||, etc.
 */
function applySyntaxHighlight(textObj, options) {
  const text = textObj.getText();
  if (!text || text.length === 0) return;

  // Code block text must never be bold
  try {
    textObj.setBold(false);
    textObj.setBold(0, text.length - 1, false);
  } catch(e) {}

  options = options || {};
  const themeName = options.theme || 'github-light';
  const themeDef = (typeof CODE_THEMES !== 'undefined' && CODE_THEMES[themeName])
    ? CODE_THEMES[themeName]
    : CODE_THEMES['github-light'];

  const kwColor   = options.keywordColor  || themeDef.kw;
  const strColor  = options.stringColor   || themeDef.str;
  const comColor  = options.commentColor  || themeDef.com;
  const numColor  = options.numberColor   || themeDef.num;
  const fnColor   = options.functionColor || themeDef.fn || '#8250DF';
  const typeColor = options.typeColor     || themeDef.type || '#953800';
  const boolColor = options.boolColor     || themeDef.bool || kwColor;
  const specColor = options.specialColor  || themeDef.special || kwColor;
  const opColor   = options.operatorColor || themeDef.op || kwColor;

  const occupied = new Array(text.length).fill(false);

  function apply(regex, color, useGroup1) {
    if (!color) return;
    let m;
    regex.lastIndex = 0;
    while ((m = regex.exec(text)) !== null) {
      const idx = useGroup1 && m[1] !== undefined ? m.index + m[0].indexOf(m[1]) : m.index;
      const val = useGroup1 && m[1] !== undefined ? m[1] : m[0];
      const end = idx + val.length - 1;
      let conflict = false;
      for (let i = idx; i <= end; i++) {
        if (occupied[i]) { conflict = true; break; }
      }
      if (!conflict) {
        for (let i = idx; i <= end; i++) occupied[i] = true;
        try {
          textObj.setForegroundColor(idx, end, color);
        } catch(e) {}
      }
    }
  }

  // 1. Comments (claim entire comment span first to prevent keywords matching inside comments)
  apply(/(\/\/.*$|#.*$|\/\*[\s\S]*?\*\/)/gm, comColor, false);

  // 2. Strings (claim entire string span next)
  apply(/(["'`])(?:(?=(\\?))\2[\s\S])*?\1/g, strColor, false);

  // 3. Keywords & Declarations
  apply(/\b(def|class|function|const|let|var|val|fn|sub|procedure|constructor|method|interface|struct|enum|record|type|alias|namespace|package|import|export|from|as|using|public|private|protected|static|final|abstract|override|readonly|mut|volatile|return|if|else|elif|for|while|do|switch|case|default|break|continue|throw|try|catch|finally|yield|await|async|new|delete|typeof|instanceof|lambda|pass|with|in|is|not|and|or|SELECT|FROM|WHERE|INSERT|UPDATE|DELETE|JOIN|INNER|LEFT|RIGHT|ON|GROUP|BY|ORDER|HAVING|LIMIT|CREATE|TABLE|ALTER|DROP|SET|VALUES|INTO)\b/g, kwColor, false);

  // 4. Special Keywords (this, self, super)
  apply(/\b(this|self|super)\b/g, specColor, false);

  // 5. Types & Classes (primitives + PascalCase identifier names)
  apply(/\b(int|float|double|char|bool|boolean|void|string|number|any|unknown|never|byte|short|long|unsigned|object|symbol|bigint)\b/g, typeColor, false);
  apply(/\b([A-Z][a-zA-Z0-9_$]*)\b/g, typeColor, false);

  // 6. Function / Method calls & headers
  apply(/\b([a-zA-Z_$][\w$]*)\s*(?=\()/g, fnColor, true);
  apply(/\b([a-zA-Z_$][\w$]*)\s*(?=:\s*$)/g, fnColor, true);

  // 7. Booleans & Constants
  apply(/\b(true|false|null|undefined|nil|None|True|False|NaN|Infinity)\b/g, boolColor, false);
  apply(/\b([A-Z_][A-Z0-9_]{2,})\b/g, boolColor, false);

  // 8. Numbers (hex, binary, float, integer)
  apply(/\b(0x[0-9a-fA-F]+|0b[01]+|\d+(\.\d+)?([eE][+-]?\d+)?)\b/g, numColor, false);

  // 9. Operators
  apply(/(=>|===|!==|==|!=|<=|>=|&&|\|\||\+\+|--|\+=|-=|\*=|\/=|%=|\+|-|\*|\/|%|=|->)/g, opColor, false);
}

function saveUserPreferences(prefs) {
  PropertiesService.getUserProperties().setProperty('code_highlighter_prefs', JSON.stringify(prefs));
}

function getUserPreferences() {
  const saved = PropertiesService.getUserProperties().getProperty('code_highlighter_prefs');
  let prefs = null;
  if (saved) {
    try { prefs = JSON.parse(saved); } catch(e) {}
  }
  return Object.assign({
    theme: 'github-light',
    fontSize: '9.5',
    fontFamily: 'Consolas',
    indentStyle: 'auto-2',
    bgColor: '#F6F8FA',
    textColor: '#24292F',
    borderColor: '#D0D7DE',
    keywordColor: '#CF222E',
    stringColor: '#0A3069',
    commentColor: '#6E7781',
    numberColor: '#0550AE',
    functionColor: '#8250DF',
    typeColor: '#953800'
  }, prefs || {});
}

/* ==========================================================================
   DOCUMENT TYPOGRAPHY & HEADING FORMATTING ENGINE
   ========================================================================== */

/**
 * Quick action to format document typography from the menu
 */
function quickFormatDocumentTypography() {
  const prefs = getTypographyPreferences();
  const result = formatDocumentTypography(prefs);
  DocumentApp.getUi().alert('✍️ Typography Formatter', result.message, DocumentApp.getUi().ButtonSet.OK);
}

/**
 * Quick action to format selected text/heading typography from the menu
 */
function quickFormatSelectedTypography() {
  const prefs = getTypographyPreferences();
  const result = formatSelectedTypography(prefs);
  DocumentApp.getUi().alert('✍️ Typography Formatter', result.message, DocumentApp.getUi().ButtonSet.OK);
}

/**
 * Quick action to format selected text as Title from the menu
 */
function quickFormatSelectedAsTitle() {
  const prefs = getTypographyPreferences();
  const result = formatSelectedAs('title', prefs);
  DocumentApp.getUi().alert('✍️ Typography Formatter', result.message, DocumentApp.getUi().ButtonSet.OK);
}

/**
 * Quick action to format selected text as Heading 1 from the menu
 */
function quickFormatSelectedAsH1() {
  const prefs = getTypographyPreferences();
  const result = formatSelectedAs('heading1', prefs);
  DocumentApp.getUi().alert('✍️ Typography Formatter', result.message, DocumentApp.getUi().ButtonSet.OK);
}

/**
 * Quick action to format selected text as Sub-Heading from the menu
 */
function quickFormatSelectedAsSub() {
  const prefs = getTypographyPreferences();
  const result = formatSelectedAs('sub', prefs);
  DocumentApp.getUi().alert('✍️ Typography Formatter', result.message, DocumentApp.getUi().ButtonSet.OK);
}

/**
 * Quick action to undo typography formatting from the menu
 */
function quickUndoDocumentTypography() {
  const result = undoDocumentTypography();
  DocumentApp.getUi().alert('✍️ Typography Formatter', result.message, DocumentApp.getUi().ButtonSet.OK);
}

/**
 * Master 1-Click Auto-Formatter:
 * Intelligently scans the entire document to auto-detect what formatting
 * to apply to every element:
 * 1. Code blocks (fenced ``` and programming syntax) -> formatted & indented into code containers
 * 2. Data tables -> centered on page and formatted with professional theme & zebra striping
 * 3. Typography -> Title, Heading 1, Sub-headings (H2, H3), and Body text with inline code highlights
 */
function quickSmartAutoFormatDocument() {
  try {
    const result = smartAutoFormatEntireDocument();
    DocumentApp.getUi().alert('🚀 Smart Auto-Formatter', result.message, DocumentApp.getUi().ButtonSet.OK);
  } catch (err) {
    Logger.log('quickSmartAutoFormatDocument error: ' + err);
    DocumentApp.getUi().alert('⚠️ Smart Auto-Formatter', 'An error occurred during formatting: ' + (err.message || err.toString()), DocumentApp.getUi().ButtonSet.OK);
  }
}

function smartAutoFormatEntireDocument(codeOptions, tableOptions, typoOptions) {
  codeOptions = codeOptions || getUserPreferences();
  tableOptions = tableOptions || getTablePreferences();
  typoOptions = typoOptions || getTypographyPreferences();

  // 1. Remove all bookmarks from the document
  try {
    removeAllBookmarks();
  } catch(e) {
    Logger.log('removeAllBookmarks error: ' + e);
  }

  // 2. Format code blocks first (extracts them into table containers so typography ignores them)
  let codeResult = null;
  try {
    codeResult = highlightAllCodeBlocks(codeOptions);
  } catch(e) {
    Logger.log('highlightAllCodeBlocks error: ' + e);
  }

  // 3. Format and center all data tables (code block tables are centered without losing syntax styles)
  let tableResult = null;
  try {
    tableResult = formatAllDocumentTables(tableOptions);
  } catch(e) {
    Logger.log('formatAllDocumentTables error: ' + e);
  }

  // 4. Auto-detect and format all typography (Title, H1, Subheadings, Body, inline code)
  let typoResult = null;
  try {
    typoResult = formatDocumentTypography(typoOptions);
  } catch(e) {
    Logger.log('formatDocumentTypography error: ' + e);
  }

  const summary = [];
  if (codeResult && codeResult.count > 0) summary.push(codeResult.count + ' code block(s) formatted');
  if (tableResult && tableResult.count > 0) summary.push(tableResult.count + ' table(s) aligned & styled');
  if (typoResult && typoResult.count > 0) summary.push(typoResult.message || 'Typography formatted');

  const message = summary.length > 0
    ? '🚀 Smart Auto-Format Complete!\n\n• ' + summary.join('\n• ')
    : 'Document scanned. All elements are formatted!';

  return {
    success: true,
    codeResult: codeResult,
    tableResult: tableResult,
    typoResult: typoResult,
    message: message
  };
}

/**
 * Removes all bookmark annotations from the active document.
 * Bookmarks in Google Docs are named anchors shown as blue flags in the margin.
 */
function removeAllBookmarks() {
  try {
    const doc = DocumentApp.getActiveDocument();
    const bookmarks = doc.getBookmarks();
    for (let i = 0; i < bookmarks.length; i++) {
      try { bookmarks[i].remove(); } catch(e) {}
    }
    return { success: true, count: bookmarks.length };
  } catch(e) {
    return { success: false, count: 0, error: e.toString() };
  }
}

/* ==========================================================================
   TEXT CLEANUP, SPACING & INLINE FORMATTING ENGINE
   ========================================================================== */

/**
 * Quick menu wrappers
 */
function quickCleanupDocumentText() {
  const result = cleanupDocumentText({ smartTypography: true, stripDoubleSpaces: true, normalizeBlankLines: true });
  DocumentApp.getUi().alert('🧹 Text Cleanup', result.message, DocumentApp.getUi().ButtonSet.OK);
}
function quickNormalizeDocumentSpacing() {
  const result = normalizeDocumentSpacing({ headingSpacingBefore: 14, headingSpacingAfter: 4, bodySpacingAfter: 6, keepHeadingsWithNext: true });
  DocumentApp.getUi().alert('📐 Spacing Normalizer', result.message, DocumentApp.getUi().ButtonSet.OK);
}
function quickApplyInlineMarkdown() {
  const result = applyInlineMarkdown();
  DocumentApp.getUi().alert('✦ Inline Markdown', result.message, DocumentApp.getUi().ButtonSet.OK);
}
function quickApplyHeadingNumbers() {
  const result = applyHeadingNumbers();
  DocumentApp.getUi().alert('🔢 Heading Numbers', result.message, DocumentApp.getUi().ButtonSet.OK);
}
function quickConvertMarkdownLists() {
  const result = convertMarkdownLists();
  DocumentApp.getUi().alert('📋 List Converter', result.message, DocumentApp.getUi().ButtonSet.OK);
}
function quickHeadingsToUpperCase() {
  const result = transformHeadingCase('upper');
  DocumentApp.getUi().alert('🔡 Heading Case', result.message, DocumentApp.getUi().ButtonSet.OK);
}
function quickHeadingsToTitleCase() {
  const result = transformHeadingCase('title');
  DocumentApp.getUi().alert('🔡 Heading Case', result.message, DocumentApp.getUi().ButtonSet.OK);
}

/**
 * Cleans up document text:
 *  - Collapses multiple spaces within paragraphs
 *  - Strips trailing whitespace from each line
 *  - Normalizes consecutive blank paragraphs to at most 2
 *  - Smart typography: straight quotes → curly, -- → —, ... → …
 */
function cleanupDocumentText(options) {
  options = options || {};
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  let fixedParas = 0;
  let fixedChars = 0;

  const numChildren = body.getNumChildren();
  let prevWasBlank = false;
  let blankStreak = 0;
  const toRemove = [];

  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) {
      prevWasBlank = false;
      blankStreak = 0;
      continue;
    }
    const p = child.asParagraph();
    // Skip paragraphs inside table cells
    if (p.getParent() && p.getParent().getType() === DocumentApp.ElementType.TABLE_CELL) continue;

    let text = p.getText();
    const trimmed = text.trim();
    const isEmpty = trimmed.length === 0;

    // Track blank line streaks for normalization
    if (isEmpty) {
      blankStreak++;
      if (options.normalizeBlankLines && blankStreak > 2) {
        toRemove.push(p);
      }
      continue;
    } else {
      blankStreak = 0;
    }

    let newText = text;

    // Strip double (or more) spaces within a line
    if (options.stripDoubleSpaces) {
      newText = newText.replace(/  +/g, ' ');
    }

    // Strip trailing whitespace
    newText = newText.replace(/\s+$/, '');

    // Smart typography substitutions
    if (options.smartTypography) {
      // Em dash: ' -- ' or '--' → ' — '
      newText = newText.replace(/ -- /g, ' \u2014 ').replace(/--/g, '\u2014');
      // Ellipsis: ... → …
      newText = newText.replace(/\.\.\./g, '\u2026');
      // Curly double quotes
      newText = newText.replace(/(^|[\s([{])"(\S)/g, '$1\u201C$2');
      newText = newText.replace(/(\S)"([\s)\]}.!?,;:]|$)/g, '$1\u201D$2');
      // Curly single quotes / apostrophes
      newText = newText.replace(/(\w)'(\w)/g, '$1\u2019$2'); // contractions
      newText = newText.replace(/(^|[\s([{])'(\S)/g, '$1\u2018$2');
      newText = newText.replace(/(\S)'([\s)\]}.!?,;:]|$)/g, '$1\u2019$2');
    }

    if (newText !== text) {
      p.setText(newText);
      fixedParas++;
      fixedChars += Math.abs(newText.length - text.length);
    }
  }

  // Remove excess blank paragraphs (bottom-up to preserve indices)
  for (let i = toRemove.length - 1; i >= 0; i--) {
    try { toRemove[i].removeFromParent(); } catch(e) {}
  }

  const msg = [];
  if (fixedParas > 0) msg.push(fixedParas + ' paragraph(s) cleaned');
  if (toRemove.length > 0) msg.push(toRemove.length + ' excess blank line(s) removed');

  return {
    success: true,
    fixedParas: fixedParas,
    removedBlanks: toRemove.length,
    message: msg.length > 0 ? '🧹 ' + msg.join(', ') + '.' : '✓ Document text is already clean.'
  };
}

/**
 * Normalizes paragraph spacing across the entire document:
 *  - Sets spacingBefore/After on headings
 *  - Sets spacingAfter on body paragraphs
 *  - Sets keepWithNext=true on all headings so they don't orphan at page breaks
 *  - Sets keepTogether=true on headings
 */
function normalizeDocumentSpacing(options) {
  options = options || {};
  const hBefore = options.headingSpacingBefore !== undefined ? options.headingSpacingBefore : 14;
  const hAfter  = options.headingSpacingAfter  !== undefined ? options.headingSpacingAfter  : 4;
  const bAfter  = options.bodySpacingAfter     !== undefined ? options.bodySpacingAfter     : 6;
  const keepWithNext = options.keepHeadingsWithNext !== false;

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const PH = DocumentApp.ParagraphHeading;
  let count = 0;

  const numChildren = body.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const p = child.asParagraph();
    if (p.getParent() && p.getParent().getType() === DocumentApp.ElementType.TABLE_CELL) continue;

    const heading = p.getHeading();
    const isHeading = heading === PH.TITLE || heading === PH.SUBTITLE ||
                      heading === PH.HEADING1 || heading === PH.HEADING2 ||
                      heading === PH.HEADING3 || heading === PH.HEADING4 ||
                      heading === PH.HEADING5 || heading === PH.HEADING6;
    try {
      if (isHeading) {
        p.setSpacingBefore(hBefore);
        p.setSpacingAfter(hAfter);
        if (keepWithNext) {
          try { p.setKeepWithNext(true); } catch(e) {}
        }
        try { p.setKeepTogether(true); } catch(e) {}
      } else if (heading === PH.NORMAL) {
        if (p.getText().trim().length > 0) {
          p.setSpacingAfter(bAfter);
        }
      }
      count++;
    } catch(e) {}
  }

  return {
    success: true,
    count: count,
    message: '📐 Spacing normalized for ' + count + ' paragraph(s). Headings pinned to next paragraph.'
  };
}

/**
 * Parses inline Markdown syntax in all body paragraphs and applies
 * the corresponding text formatting using Apps Script text ranges:
 *   **bold**          → setBold
 *   *italic*          → setItalic (also _italic_)
 *   ~~strikethrough~~ → setStrikethrough
 *   ==highlight==     → setBackgroundColor('#FFFF00')
 * Strips the markdown markers from the text after applying formatting.
 */
function applyInlineMarkdown(options) {
  options = options || {};
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  let count = 0;

  const numChildren = body.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const p = child.asParagraph();
    if (p.getParent() && p.getParent().getType() === DocumentApp.ElementType.TABLE_CELL) continue;

    if (applyMarkdownToParagraph(p)) count++;
  }

  return {
    success: true,
    count: count,
    message: count > 0
      ? '✦ Inline markdown applied to ' + count + ' paragraph(s).'
      : '✓ No inline markdown markers found.'
  };
}

/**
 * Applies markdown formatting to a single paragraph.
 * Returns true if any changes were made.
 */
function applyMarkdownToParagraph(p) {
  const raw = p.getText();
  if (!raw || raw.length === 0) return false;

  // Ordered list of patterns: [ regex, formatFn, markerLength ]
  // We repeatedly scan because applying one pattern shifts character indices.
  const PATTERNS = [
    // ~~strikethrough~~ (must come before *italic*)
    { re: /~~([^~]+)~~/, apply: function(t, s, e) { t.setStrikethrough(s, e, true); }, marker: 2 },
    // ==highlight==
    { re: /==([^=]+)==/, apply: function(t, s, e) { t.setBackgroundColor(s, e, '#FFFF00'); }, marker: 2 },
    // **bold**
    { re: /\*\*([^*]+)\*\*/, apply: function(t, s, e) { t.setBold(s, e, true); }, marker: 2 },
    // *italic* or _italic_
    { re: /\*([^*]+)\*|_([^_]+)_/, apply: function(t, s, e) { t.setItalic(s, e, true); }, marker: 1 }
  ];

  let changed = false;
  for (let pi = 0; pi < PATTERNS.length; pi++) {
    const pat = PATTERNS[pi];
    let text = p.getText();
    let match;
    // Use a loop to handle multiple occurrences in this paragraph
    while ((match = pat.re.exec(text)) !== null) {
      const fullMatch = match[0];
      const inner = match[1] || match[2];
      const startFull = match.index;
      const mk = pat.marker;
      // Content spans from startFull+mk to startFull+mk+inner.length-1
      const contentStart = startFull + mk;
      const contentEnd   = startFull + mk + inner.length - 1;

      // Build new text without the markers
      const before = text.substring(0, startFull);
      const after  = text.substring(startFull + fullMatch.length);
      const newText = before + inner + after;
      p.setText(newText);

      // Apply formatting to the content range in the new text
      try {
        const textObj = p.editAsText();
        pat.apply(textObj, contentStart, contentStart + inner.length - 1);
      } catch(e) {}

      changed = true;
      // Re-read text for next iteration
      text = p.getText();
    }
  }
  return changed;
}

/**
 * Auto-numbers all headings in the document using outline numbering:
 *   Heading1 → 1.  2.  3.
 *   Heading2 → 1.1  1.2  2.1
 *   Heading3 → 1.1.1  1.1.2  1.2.1
 * Strips any existing leading numbering before re-applying.
 */
function applyHeadingNumbers(options) {
  options = options || {};
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const PH = DocumentApp.ParagraphHeading;

  let counters = [0, 0, 0]; // H1, H2, H3
  let count = 0;

  const numChildren = body.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const p = child.asParagraph();
    if (p.getParent() && p.getParent().getType() === DocumentApp.ElementType.TABLE_CELL) continue;

    const heading = p.getHeading();
    let level = -1;
    if (heading === PH.HEADING1) level = 0;
    else if (heading === PH.HEADING2) level = 1;
    else if (heading === PH.HEADING3) level = 2;
    if (level < 0) continue;

    counters[level]++;
    // Reset deeper levels
    for (let c = level + 1; c < counters.length; c++) counters[c] = 0;

    const prefix = level === 0
      ? counters[0] + '.'
      : level === 1
        ? counters[0] + '.' + counters[1]
        : counters[0] + '.' + counters[1] + '.' + counters[2];

    // Strip existing leading number pattern
    let text = p.getText().replace(/^\d+(\.\d+)*\.?\s+/, '').trim();
    p.setText(prefix + '  ' + text);
    count++;
  }

  return {
    success: true,
    count: count,
    message: count > 0
      ? '🔢 Numbered ' + count + ' heading(s) with outline numbering.'
      : '✓ No headings found to number.'
  };
}

/**
 * Transforms heading text case across the document.
 * @param {string} caseType - 'upper', 'lower', 'title', or 'sentence'
 */
function transformHeadingCase(caseType) {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const PH = DocumentApp.ParagraphHeading;
  let count = 0;

  const numChildren = body.getNumChildren();
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    const p = child.asParagraph();
    if (p.getParent() && p.getParent().getType() === DocumentApp.ElementType.TABLE_CELL) continue;

    const heading = p.getHeading();
    const isHeading = heading === PH.TITLE || heading === PH.HEADING1 ||
                      heading === PH.HEADING2 || heading === PH.HEADING3 ||
                      heading === PH.HEADING4 || heading === PH.HEADING5 ||
                      heading === PH.HEADING6;
    if (!isHeading) continue;

    const text = p.getText();
    if (!text || text.trim().length === 0) continue;

    let newText;
    switch (caseType) {
      case 'upper':
        newText = text.toUpperCase();
        break;
      case 'lower':
        newText = text.toLowerCase();
        break;
      case 'title':
        newText = text.replace(/\w\S*/g, function(w) {
          // Lowercase small words unless first word
          const small = /^(a|an|the|and|but|or|for|nor|on|at|to|by|in|of|up|as|is|it)$/i;
          return small.test(w) ? w.toLowerCase() : w.charAt(0).toUpperCase() + w.substr(1).toLowerCase();
        });
        // Always capitalize first word
        newText = newText.charAt(0).toUpperCase() + newText.slice(1);
        break;
      case 'sentence':
        newText = text.charAt(0).toUpperCase() + text.slice(1).toLowerCase();
        break;
      default:
        newText = text;
    }

    if (newText !== text) {
      p.setText(newText);
      count++;
    }
  }

  return {
    success: true,
    count: count,
    message: count > 0
      ? '🔡 Applied ' + caseType + ' case to ' + count + ' heading(s).'
      : '✓ No heading text changed.'
  };
}

/**
 * Converts plain-text markdown-style list items into proper Google Docs
 * ListItem elements with bullet styling.
 * Detects lines starting with: - , • , * , or 1. 2. 3. (ordered)
 */
function convertMarkdownLists() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  let converted = 0;

  // We iterate via index because convertToListItem may change the DOM
  let i = 0;
  while (i < body.getNumChildren()) {
    const child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) { i++; continue; }
    const p = child.asParagraph();
    if (p.getParent() && p.getParent().getType() === DocumentApp.ElementType.TABLE_CELL) { i++; continue; }

    const text = p.getText();
    const unordered = /^[-•*]\s+(.+)$/.exec(text);
    const ordered   = /^\d+\.\s+(.+)$/.exec(text);

    if (unordered || ordered) {
      const inner = (unordered || ordered)[1];
      try {
        // Replace text then convert to list item
        p.setText(inner);
        p.setGlyphType(unordered
          ? DocumentApp.GlyphType.BULLET
          : DocumentApp.GlyphType.NUMBER);
        converted++;
      } catch(e) {
        // setGlyphType not available; fall back to re-inserting as ListItem
        try {
          const idx = body.getChildIndex(p);
          const li = body.insertListItem(idx, inner);
          li.setGlyphType(unordered
            ? DocumentApp.GlyphType.BULLET
            : DocumentApp.GlyphType.NUMBER);
          li.setNestingLevel(0);
          p.removeFromParent();
          converted++;
        } catch(e2) {}
      }
    }
    i++;
  }

  return {
    success: true,
    count: converted,
    message: converted > 0
      ? '📋 Converted ' + converted + ' plain-text list item(s) to proper document lists.'
      : '✓ No markdown list items found.'
  };
}

/**
 * Safely sets the heading style of a paragraph without throwing an exception.
 * In Google Apps Script, the ParagraphHeading constants do not have underscores:
 * HEADING1, HEADING2, HEADING3, HEADING4, HEADING5, HEADING6, TITLE, SUBTITLE, NORMAL.
 */
function safeSetHeading(p, headingEnum) {
  if (!p || headingEnum === undefined || headingEnum === null) return;
  try {
    // Only attempt on genuine standalone paragraphs
    if (typeof p.getType === 'function' && p.getType() !== DocumentApp.ElementType.PARAGRAPH) {
      return;
    }
    const parent = p.getParent();
    if (parent && typeof parent.getType === 'function') {
      const pType = parent.getType();
      // Google Docs throws service exceptions if setHeading is called inside tables or list items
      if (pType === DocumentApp.ElementType.TABLE_CELL ||
          pType === DocumentApp.ElementType.TABLE_ROW ||
          pType === DocumentApp.ElementType.TABLE ||
          pType === DocumentApp.ElementType.LIST_ITEM) {
        return;
      }
    }
    if (typeof p.setHeading === 'function') {
      p.setHeading(headingEnum);
    }
  } catch(e) {
    // Gracefully catch Google Docs internal heading assignment exceptions
  }
}

function getHeadingConstant(type) {
  const PH = DocumentApp.ParagraphHeading;
  if (!PH) return null;
  switch (type) {
    case 'TITLE': return PH.TITLE;
    case 'SUBTITLE': return PH.SUBTITLE;
    case 'HEADING1': case 'HEADING_1': case 1: return PH.HEADING1 || PH['HEADING_1'];
    case 'HEADING2': case 'HEADING_2': case 2: return PH.HEADING2 || PH['HEADING_2'];
    case 'HEADING3': case 'HEADING_3': case 3: return PH.HEADING3 || PH['HEADING_3'];
    case 'HEADING4': case 'HEADING_4': case 4: return PH.HEADING4 || PH['HEADING_4'];
    case 'HEADING5': case 'HEADING_5': case 5: return PH.HEADING5 || PH['HEADING_5'];
    case 'HEADING6': case 'HEADING_6': case 6: return PH.HEADING6 || PH['HEADING_6'];
    case 'NORMAL': default: return PH.NORMAL;
  }
}

/**
 * Intelligent classifier for document paragraphs:
 * Analyzes structure, markdown markers, numbering patterns, typography hints,
 * sentence punctuation, and context to automatically detect whether a paragraph
 * is a Title, Heading 1, Sub-Heading (H2/H3), or Body Text.
 */
function detectTextRole(p, text, isFirstNonEmpty, state) {
  state = state || {};
  if (!text || text.length === 0) {
    return { role: 'body', headingLevel: 0, cleanText: '' };
  }

  // 1. Markdown syntax check (applies regardless of prior headings)
  if (/^#\s+(.+)$/.test(text)) {
    const clean = text.replace(/^#\s+/, '').trim();
    if (!state.hasTitle && isFirstNonEmpty) {
      state.hasTitle = true;
      return { role: 'title', headingLevel: 1, cleanText: clean };
    }
    state.lastHeadingLevel = 1;
    return { role: 'heading1', headingLevel: 1, cleanText: clean };
  }
  if (/^##\s+(.+)$/.test(text)) {
    const clean = text.replace(/^##\s+/, '').trim();
    state.lastHeadingLevel = 2;
    return { role: 'sub', headingLevel: 2, cleanText: clean };
  }
  if (/^###\s+(.+)$/.test(text)) {
    const clean = text.replace(/^###\s+/, '').trim();
    state.lastHeadingLevel = 3;
    return { role: 'sub', headingLevel: 3, cleanText: clean };
  }
  if (/^#{4,6}\s+(.+)$/.test(text)) {
    const clean = text.replace(/^#{4,6}\s+/, '').trim();
    state.lastHeadingLevel = 3;
    return { role: 'sub', headingLevel: 3, cleanText: clean };
  }

  // 2. Explicit Title prefixes: "Title: ...", "Document Title: ..."
  const titlePrefixMatch = text.match(/^(?:Document\s+Title|Title)\s*[:–—]\s*(.+)$/i);
  if (titlePrefixMatch) {
    state.hasTitle = true;
    return { role: 'title', headingLevel: 1, cleanText: titlePrefixMatch[1].trim() };
  }

  // 3. Pre-existing Google Docs formal headings (preserve unless overridden)
  const existingHeading = p.getHeading();
  if (existingHeading && existingHeading !== DocumentApp.ParagraphHeading.NORMAL) {
    if (existingHeading === DocumentApp.ParagraphHeading.TITLE) {
      state.hasTitle = true;
      return { role: 'title', headingLevel: 1, cleanText: text };
    }
    if (existingHeading === (DocumentApp.ParagraphHeading.HEADING1 || DocumentApp.ParagraphHeading.HEADING_1)) {
      state.lastHeadingLevel = 1;
      return { role: 'heading1', headingLevel: 1, cleanText: text };
    }
    if (
      existingHeading === (DocumentApp.ParagraphHeading.HEADING2 || DocumentApp.ParagraphHeading.HEADING_2) ||
      existingHeading === DocumentApp.ParagraphHeading.SUBTITLE
    ) {
      state.lastHeadingLevel = 2;
      return { role: 'sub', headingLevel: 2, cleanText: text };
    }
    state.lastHeadingLevel = 3;
    return { role: 'sub', headingLevel: 3, cleanText: text };
  }

  // 4. Exclude obvious non-headings (code blocks, quotes, bullet lists)
  if (/^```/.test(text) || /^[-*•]\s+/.test(text) || /^>\s+/.test(text)) {
    return { role: 'body', headingLevel: 0, cleanText: text };
  }

  // Check if text ends in sentence punctuation (. ! ?)
  const endsInPunctuation = /[.!?]$/.test(text) && !/^(?:etc|vs|vol|no|dr|mr|mrs|prof)\.$/i.test(text);

  // 5. First Non-Empty line Document Title heuristic
  if (isFirstNonEmpty && !state.hasTitle) {
    if (text.length <= 110 && !endsInPunctuation) {
      const words = text.split(/\s+/);
      if (words.length <= 14) {
        state.hasTitle = true;
        return { role: 'title', headingLevel: 1, cleanText: text };
      }
    }
  }

  // 6. Numbered Sub-sub-sections (H3): e.g. "1.1.1 Overview", "(a) Item", "(1) Item"
  if (text.length <= 90 && !endsInPunctuation) {
    if (/^\d+\.\d+\.\d+\.?\s+[A-Za-z]/.test(text) || /^\([a-z0-9]+\)\s+[A-Za-z]/i.test(text)) {
      state.lastHeadingLevel = 3;
      return { role: 'sub', headingLevel: 3, cleanText: text };
    }
  }

  // 7. Numbered Sub-sections (H2): e.g. "1.1 Background", "2.3 Implementation", "a) Setup", "Step 1:"
  if (text.length <= 90 && !endsInPunctuation) {
    if (
      /^\d+\.\d+\.?\s+[A-Za-z]/.test(text) ||
      /^[a-z]\)\s+[A-Za-z]/.test(text) ||
      /^(?:Sub-?section|Sub-?topic|Step)\s+(?:\d+|[A-Za-z0-9.]+)/i.test(text)
    ) {
      state.lastHeadingLevel = 2;
      return { role: 'sub', headingLevel: 2, cleanText: text };
    }
  }

  // 8. Numbered Main Sections (H1): e.g. "1. Introduction", "2. Architecture", "I. Executive Summary", "A. Background"
  if (text.length <= 90 && !endsInPunctuation) {
    if (
      /^(?:\d+|[A-Z]|[IVXLCDM]+)\.\s+[A-Z]/.test(text) ||
      /^(?:Section|Chapter|Part|Module|Unit|Phase)\s+(?:\d+|[A-Z]|[IVXLCDM]+)[:.\s–—]/i.test(text)
    ) {
      state.lastHeadingLevel = 1;
      return { role: 'heading1', headingLevel: 1, cleanText: text };
    }
  }

  // 9. Typographic & Font styling cues from Google Docs:
  let isEntirelyBold = false;
  let maxFontSize = 0;
  try {
    const textObj = p.editAsText();
    if (textObj.getText().length > 0) {
      isEntirelyBold = textObj.isBold() === true;
      maxFontSize = Number(textObj.getFontSize(0)) || 0;
    }
  } catch(e) {}

  if (maxFontSize >= 22 && text.length <= 110 && !endsInPunctuation) {
    state.hasTitle = true;
    return { role: 'title', headingLevel: 1, cleanText: text };
  }

  if (isEntirelyBold && text.length <= 80 && !endsInPunctuation) {
    if (state.lastHeadingLevel === 1 && text.length < 50) {
      state.lastHeadingLevel = 2;
      return { role: 'sub', headingLevel: 2, cleanText: text };
    }
    state.lastHeadingLevel = 1;
    return { role: 'heading1', headingLevel: 1, cleanText: text };
  }

  if (maxFontSize >= 16 && text.length <= 80 && !endsInPunctuation) {
    state.lastHeadingLevel = 1;
    return { role: 'heading1', headingLevel: 1, cleanText: text };
  }
  if (maxFontSize >= 13 && text.length <= 80 && !endsInPunctuation) {
    state.lastHeadingLevel = 2;
    return { role: 'sub', headingLevel: 2, cleanText: text };
  }

  // 10. ALL CAPS Standalone Headings (H1): e.g. "EXECUTIVE SUMMARY", "METHODOLOGY", "RESULTS"
  if (text.length >= 3 && text.length <= 55 && !endsInPunctuation) {
    if (/^[A-Z0-9\s&,/:–—\-]+$/.test(text) && /[A-Z]{2,}/.test(text)) {
      state.lastHeadingLevel = 1;
      return { role: 'heading1', headingLevel: 1, cleanText: text };
    }
  }

  // 11. Title Case Short Headers without punctuation (e.g. "Key Architecture Overview")
  if (text.length >= 4 && text.length <= 65 && !endsInPunctuation) {
    const words = text.split(/\s+/);
    if (words.length >= 2 && words.length <= 8) {
      const majorWords = words.filter(w => !['and','or','the','in','on','at','to','for','with','of','a','an'].includes(w.toLowerCase()));
      const capCount = majorWords.filter(w => /^[A-Z]/.test(w)).length;
      if (majorWords.length > 0 && capCount / majorWords.length >= 0.8) {
        const hasCommonVerb = words.some(w => ['is','are','was','were','has','have','can','could','should','will','would'].includes(w.toLowerCase()));
        if (!hasCommonVerb) {
          const role = (state.lastHeadingLevel === 1) ? 'sub' : 'heading1';
          state.lastHeadingLevel = (role === 'sub') ? 2 : 1;
          return { role: role, headingLevel: (role === 'sub') ? 2 : 1, cleanText: text };
        }
      }
    }
  }

  return { role: 'body', headingLevel: 0, cleanText: text };
}

/**
 * Formats all headings and optionally body text according to typography options,
 * with intelligent full-document structure and role auto-detection.
 */
function formatDocumentTypography(options) {
  options = options || getTypographyPreferences();
  saveTypographyPreferences(options);

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  // 1. Unroll any existing heading banner tables so headings are standard body paragraphs
  try {
    unrollHeadingBannerTables(body);
  } catch(e) {}

  let titleCount = 0;
  let h1Count = 0;
  let subCount = 0;
  let bodyCount = 0;

  // 2. Scan and classify all paragraphs
  const numChildren = body.getNumChildren();
  const paragraphs = [];
  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      paragraphs.push(child.asParagraph());
    }
  }

  const state = { hasTitle: false, lastHeadingLevel: 0 };
  let firstNonEmptyFound = false;

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    if (!p.getParent()) continue;

    // Skip items already inside a table cell (e.g. data tables, code blocks)
    if (p.getParent().getType() === DocumentApp.ElementType.TABLE_CELL) {
      continue;
    }

    const rawText = p.getText();
    const trimmed = rawText.trim();
    if (trimmed.length === 0) continue;

    const isFirstNonEmpty = !firstNonEmptyFound;
    if (isFirstNonEmpty) firstNonEmptyFound = true;

    // Detect role for this paragraph
    const detection = detectTextRole(p, trimmed, isFirstNonEmpty, state);

    // Apply cleaned text if markdown markers were stripped
    if (detection.cleanText && detection.cleanText !== trimmed) {
      p.setText(detection.cleanText);
    }

    if (detection.role === 'title') {
      safeSetHeading(p, getHeadingConstant('TITLE'));
      formatSingleHeading(body, p, options.title);
      titleCount++;
    } else if (detection.role === 'heading1') {
      safeSetHeading(p, getHeadingConstant('HEADING1'));
      formatSingleHeading(body, p, options.heading1);
      h1Count++;
    } else if (detection.role === 'sub') {
      const headingLevelEnum = detection.headingLevel === 3
        ? getHeadingConstant('HEADING3')
        : getHeadingConstant('HEADING2');
      safeSetHeading(p, headingLevelEnum);
      formatSingleHeading(body, p, options.subHeading);
      subCount++;
    } else if (detection.role === 'body') {
      safeSetHeading(p, getHeadingConstant('NORMAL'));
      if (options.body && options.body.applyToBody) {
        formatSingleBodyParagraph(p, options.body);
        bodyCount++;
      }
    }
  }

  const parts = [];
  if (titleCount > 0) parts.push(titleCount + ' title');
  if (h1Count > 0) parts.push(h1Count + ' main heading(s)');
  if (subCount > 0) parts.push(subCount + ' sub-heading(s)');
  if (bodyCount > 0) parts.push(bodyCount + ' body paragraph(s)');

  return {
    success: true,
    count: titleCount + h1Count + subCount + bodyCount,
    message: parts.length > 0
      ? 'Formatted ' + parts.join(', ') + ' successfully!'
      : 'No text was formatted. Ensure "Apply to General Body Text" is checked if formatting normal text.'
  };
}

/**
 * Formats the selected text or paragraph with intelligent auto-detection of role
 */
function formatSelectedTypography(options) {
  options = options || getTypographyPreferences();
  saveTypographyPreferences(options);

  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  if (!selection) {
    return {
      success: false,
      message: 'Please highlight or select text in your document first.'
    };
  }

  const elements = selection.getSelectedElements();
  const body = doc.getBody();
  let titleCount = 0;
  let h1Count = 0;
  let subCount = 0;
  let bodyCount = 0;

  const state = { hasTitle: false, lastHeadingLevel: 0 };
  let firstNonEmptyFound = false;

  for (let i = 0; i < elements.length; i++) {
    const el = elements[i];
    let p = null;
    if (el.getElement().getType() === DocumentApp.ElementType.PARAGRAPH) {
      p = el.getElement().asParagraph();
    } else if (el.getElement().getType() === DocumentApp.ElementType.TEXT) {
      const parent = el.getElement().getParent();
      if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
        p = parent.asParagraph();
      }
    }

    if (!p) continue;
    const text = p.getText().trim();
    if (text.length === 0) continue;

    const isFirstNonEmpty = !firstNonEmptyFound;
    if (isFirstNonEmpty) firstNonEmptyFound = true;

    // Run auto-detection on this selected element
    const detection = detectTextRole(p, text, isFirstNonEmpty, state);

    if (detection.cleanText && detection.cleanText !== text) {
      p.setText(detection.cleanText);
    }

    if (detection.role === 'title') {
      safeSetHeading(p, getHeadingConstant('TITLE'));
      applyHeadingStyles(p, options.title, options.title.bgEnabled && options.title.bgStyle === 'inline' ? options.title.bgColor : null);
      titleCount++;
    } else if (detection.role === 'heading1') {
      safeSetHeading(p, getHeadingConstant('HEADING1'));
      applyHeadingStyles(p, options.heading1, options.heading1.bgEnabled && options.heading1.bgStyle === 'inline' ? options.heading1.bgColor : null);
      h1Count++;
    } else if (detection.role === 'sub') {
      const hLevelEnum = detection.headingLevel === 3
        ? getHeadingConstant('HEADING3')
        : getHeadingConstant('HEADING2');
      safeSetHeading(p, hLevelEnum);
      applyHeadingStyles(p, options.subHeading, options.subHeading.bgEnabled && options.subHeading.bgStyle === 'inline' ? options.subHeading.bgColor : null);
      subCount++;
    } else {
      safeSetHeading(p, getHeadingConstant('NORMAL'));
      formatSingleBodyParagraph(p, options.body);
      bodyCount++;
    }
  }

  const parts = [];
  if (titleCount > 0) parts.push(titleCount + ' title');
  if (h1Count > 0) parts.push(h1Count + ' main heading(s)');
  if (subCount > 0) parts.push(subCount + ' sub-heading(s)');
  if (bodyCount > 0) parts.push(bodyCount + ' body paragraph(s)');

  return {
    success: true,
    count: titleCount + h1Count + subCount + bodyCount,
    message: parts.length > 0
      ? 'Auto-formatted selection: ' + parts.join(', ') + '!'
      : 'No text elements could be formatted from selection.'
  };
}

/**
 * Formats currently selected text as a specific typography element (title, heading1, sub, or body)
 */
function formatSelectedAs(targetType, options) {
  options = options || getTypographyPreferences();
  saveTypographyPreferences(options);

  const doc = DocumentApp.getActiveDocument();
  const selection = doc.getSelection();
  if (!selection) {
    return {
      success: false,
      message: 'Please highlight or select text in your document first.'
    };
  }

  const elements = selection.getSelectedElements();
  const body = doc.getBody();
  let count = 0;

  for (let i = 0; i < elements.length; i++) {
    const el = elements[i];
    let p = null;
    if (el.getElement().getType() === DocumentApp.ElementType.PARAGRAPH) {
      p = el.getElement().asParagraph();
    } else if (el.getElement().getType() === DocumentApp.ElementType.TEXT) {
      const parent = el.getElement().getParent();
      if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
        p = parent.asParagraph();
      }
    }

    if (!p) continue;

    if (targetType === 'title') {
      safeSetHeading(p, getHeadingConstant('TITLE'));
      formatSingleHeading(body, p, options.title);
      count++;
    } else if (targetType === 'heading1') {
      safeSetHeading(p, getHeadingConstant('HEADING1'));
      formatSingleHeading(body, p, options.heading1);
      count++;
    } else if (targetType === 'sub') {
      safeSetHeading(p, getHeadingConstant('HEADING2'));
      formatSingleHeading(body, p, options.subHeading);
      count++;
    } else if (targetType === 'body') {
      safeSetHeading(p, getHeadingConstant('NORMAL'));
      formatSingleBodyParagraph(p, options.body);
      count++;
    }
  }

  const labelMap = {
    title: 'Document Title',
    heading1: 'Heading 1',
    sub: 'Sub-Heading',
    body: 'Body Text'
  };

  return {
    success: count > 0,
    count: count,
    message: count > 0
      ? 'Applied ' + (labelMap[targetType] || targetType) + ' formatting to ' + count + ' selected element(s)!'
      : 'No text element could be formatted from selection.'
  };
}

/**
 * Formats a single heading paragraph, handling background banner vs inline highlight
 */
function formatSingleHeading(body, p, config) {
  if (!config) return;

  const bgEnabled = !!config.bgEnabled;
  const bgStyle = config.bgStyle || 'inline';
  const bgColor = config.bgColor || '#EFF6FF';

  let targetPara = p;

  if (bgEnabled && bgStyle === 'banner') {
    targetPara = wrapParagraphInBanner(body, p, bgColor);
  }

  applyHeadingStyles(targetPara, config, bgEnabled && bgStyle === 'inline' ? bgColor : null);
}

/**
 * Wraps a paragraph into a full-width borderless 1x1 table banner
 */
function wrapParagraphInBanner(body, p, bgColor) {
  try {
    const parent = p.getParent();
    if (parent.getType() !== DocumentApp.ElementType.BODY_SECTION) {
      return p;
    }

    const childIndex = body.getChildIndex(p);
    const headingType = p.getHeading();
    const text = p.getText();

    // In Google Apps Script, body.insertTable(childIndex) creates a table at childIndex
    const table = body.insertTable(childIndex);
    table.setBorderWidth(0);
    try { table.setBorderColor(bgColor); } catch(e) {}

    const row = table.appendTableRow();
    const cell = row.appendTableCell();
    cell.setBackgroundColor(bgColor);
    cell.setPaddingTop(8);
    cell.setPaddingBottom(8);
    cell.setPaddingLeft(12);
    cell.setPaddingRight(12);
    cell.setVerticalAlignment(DocumentApp.VerticalAlignment.TOP);

    const cellPara = cell.getChild(0).asParagraph();
    cellPara.setText(text);

    try {
      const pageWidth = body.getPageWidth();
      const marginLeft = body.getMarginLeft();
      const marginRight = body.getMarginRight();
      table.setColumnWidth(0, Math.max(100, pageWidth - marginLeft - marginRight));
    } catch(e) {}

    p.removeFromParent();
    return cellPara;
  } catch(err) {
    return p;
  }
}

/**
 * Applies font, size, bold, color, and alignment styles to a heading paragraph
 */
function applyHeadingStyles(para, config, inlineBgColor) {
  if (!para) return;

  let align = DocumentApp.HorizontalAlignment.LEFT;
  if (config.alignment === 'CENTER') align = DocumentApp.HorizontalAlignment.CENTER;
  if (config.alignment === 'RIGHT') align = DocumentApp.HorizontalAlignment.RIGHT;
  if (config.alignment === 'JUSTIFY') align = DocumentApp.HorizontalAlignment.JUSTIFY;
  para.setAlignment(align);

  const fontSize = Number(config.fontSize) || 16;
  const fontFamily = config.fontFamily || 'Arial';
  const textColor = config.textColor || '#000000';
  const bold = config.bold !== undefined ? config.bold : true;

  try { para.setFontFamily(fontFamily); } catch(e) {}
  try { para.setFontSize(fontSize); } catch(e) {}
  para.setLineSpacing(1.15);

  const textObj = para.editAsText();
  const textLen = textObj.getText().length;
  if (textLen > 0) {
    try { textObj.setFontFamily(0, textLen - 1, fontFamily); } catch(e) {
      try { textObj.setFontFamily(fontFamily); } catch(e2) {}
    }
    try { textObj.setFontSize(0, textLen - 1, fontSize); } catch(e) {
      try { textObj.setFontSize(fontSize); } catch(e2) {}
    }
    try { textObj.setBold(0, textLen - 1, bold); } catch(e) {
      try { textObj.setBold(bold); } catch(e2) {}
    }
    try { textObj.setForegroundColor(0, textLen - 1, textColor); } catch(e) {
      try { textObj.setForegroundColor(textColor); } catch(e2) {}
    }
    try { textObj.setBackgroundColor(0, textLen - 1, inlineBgColor || null); } catch(e) {
      try { textObj.setBackgroundColor(inlineBgColor || null); } catch(e2) {}
    }
  }
}

/**
 * Applies typography options to a normal body paragraph
 */
function formatSingleBodyParagraph(para, config) {
  if (!para) return;

  let align = DocumentApp.HorizontalAlignment.LEFT;
  if (config.alignment === 'CENTER') align = DocumentApp.HorizontalAlignment.CENTER;
  if (config.alignment === 'RIGHT') align = DocumentApp.HorizontalAlignment.RIGHT;
  if (config.alignment === 'JUSTIFY') align = DocumentApp.HorizontalAlignment.JUSTIFY;
  para.setAlignment(align);

  const fontSize = Number(config.fontSize) || 11;
  const fontFamily = config.fontFamily || 'Arial';
  const textColor = config.textColor || '#1F2937';

  try { para.setFontFamily(fontFamily); } catch(e) {}
  try { para.setFontSize(fontSize); } catch(e) {}
  para.setLineSpacing(1.15);

  const textObj = para.editAsText();
  const textLen = textObj.getText().length;
  if (textLen > 0) {
    try { textObj.setFontFamily(0, textLen - 1, fontFamily); } catch(e) {
      try { textObj.setFontFamily(fontFamily); } catch(e2) {}
    }
    try { textObj.setFontSize(0, textLen - 1, fontSize); } catch(e) {
      try { textObj.setFontSize(fontSize); } catch(e2) {}
    }
    try { textObj.setBold(0, textLen - 1, false); } catch(e) {
      try { textObj.setBold(false); } catch(e2) {}
    }
    try { textObj.setForegroundColor(0, textLen - 1, textColor); } catch(e) {
      try { textObj.setForegroundColor(textColor); } catch(e2) {}
    }
    try { textObj.setBackgroundColor(0, textLen - 1, null); } catch(e) {
      try { textObj.setBackgroundColor(null); } catch(e2) {}
    }

    // Automatically detect and highlight inline code tokens (`code`)
    applyInlineBodyCodeHighlight(textObj);
  }
}

/**
 * Detects backtick-wrapped tokens (e.g. `npm install` or `myVar`) in body text
 * and applies clean monospace font (Consolas) with subtle background shading.
 */
function applyInlineBodyCodeHighlight(textObj) {
  const fullText = textObj.getText();
  if (!fullText || fullText.length === 0) return;

  const codeFont = 'Consolas';
  const codeColor = '#B45309';
  const codeBg = '#EFF1F3';

  const backtickRegex = /`([^`]+)`/g;
  let match;
  while ((match = backtickRegex.exec(fullText)) !== null) {
    const start = match.index;
    const end = start + match[0].length - 1;
    try {
      textObj.setFontFamily(start, end, codeFont);
      textObj.setForegroundColor(start, end, codeColor);
      textObj.setBackgroundColor(start, end, codeBg);
    } catch(e) {}
  }
}

/**
 * Unrolls any 1x1 heading banner tables back into standard body paragraphs
 */
function unrollHeadingBannerTables(body) {
  const tables = body.getTables();
  if (!tables || tables.length === 0) return 0;

  let unrolled = 0;
  for (let i = tables.length - 1; i >= 0; i--) {
    const table = tables[i];
    if (isHeadingBannerTable(table)) {
      unrollSingleHeadingBanner(body, table);
      unrolled++;
    }
  }
  return unrolled;
}

/**
 * Extracts paragraphs from a heading banner table into body and removes the table
 */
function unrollSingleHeadingBanner(body, table) {
  const insertIndex = body.getChildIndex(table);
  const cell = table.getRow(0).getCell(0);
  const numChildren = cell.getNumChildren();
  const createdParas = [];

  for (let i = 0; i < numChildren; i++) {
    const child = cell.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = child.asParagraph();
      const text = p.getText();
      const heading = p.getHeading();
      if (i === 0 && text === '' && numChildren > 1) continue;
      const newP = body.insertParagraph(insertIndex + createdParas.length, text);
      safeSetHeading(newP, heading);
      createdParas.push(newP);
    }
  }

  if (createdParas.length === 0) {
    createdParas.push(body.insertParagraph(insertIndex, ''));
  }

  table.removeFromParent();
  return createdParas;
}

/**
 * Resets all document headings and body text back to standard Google Docs defaults
 */
function undoDocumentTypography() {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  unrollHeadingBannerTables(body);

  const numChildren = body.getNumChildren();
  let count = 0;

  for (let i = 0; i < numChildren; i++) {
    const child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const p = child.asParagraph();
      const heading = p.getHeading();

      let defaultFont = 'Arial';
      let defaultSize = 11;
      let defaultBold = false;
      let defaultColor = '#000000';

      if (heading === DocumentApp.ParagraphHeading.TITLE) {
        defaultSize = 26;
        defaultBold = true;
      } else if (heading === (DocumentApp.ParagraphHeading.HEADING1 || DocumentApp.ParagraphHeading.HEADING_1)) {
        defaultSize = 20;
        defaultBold = true;
      } else if (heading === (DocumentApp.ParagraphHeading.HEADING2 || DocumentApp.ParagraphHeading.HEADING_2)) {
        defaultSize = 16;
        defaultBold = true;
      } else if (heading === (DocumentApp.ParagraphHeading.HEADING3 || DocumentApp.ParagraphHeading.HEADING_3)) {
        defaultSize = 14;
        defaultBold = true;
        defaultColor = '#434343';
      } else if (heading === DocumentApp.ParagraphHeading.SUBTITLE) {
        defaultSize = 15;
        defaultColor = '#666666';
      } else if (heading !== DocumentApp.ParagraphHeading.NORMAL) {
        defaultSize = 13;
        defaultBold = true;
      }

      p.setAlignment(DocumentApp.HorizontalAlignment.LEFT);
      try { p.setFontFamily(defaultFont); } catch(e) {}
      try { p.setFontSize(defaultSize); } catch(e) {}
      p.setLineSpacing(1.15);

      const textObj = p.editAsText();
      if (textObj.getText().length > 0) {
        try { textObj.setFontFamily(defaultFont); } catch(e) {}
        try { textObj.setFontSize(defaultSize); } catch(e) {}
        try { textObj.setBold(defaultBold); } catch(e) {}
        try { textObj.setForegroundColor(defaultColor); } catch(e) {}
        try { textObj.setBackgroundColor(null); } catch(e) {}
      }
      count++;
    }
  }

  return {
    success: true,
    count: count,
    message: 'Reset typography for ' + count + ' paragraph(s) and headings back to document defaults.'
  };
}

/**
 * Retrieves typography preferences with defaults
 */
function getTypographyPreferences() {
  const defaults = {
    preset: 'executive-navy',
    title: {
      fontFamily: 'Montserrat',
      fontSize: 26,
      textColor: '#1E3A8A',
      alignment: 'CENTER',
      bold: true,
      bgEnabled: false,
      bgColor: '#EFF6FF',
      bgStyle: 'inline'
    },
    heading1: {
      fontFamily: 'Montserrat',
      fontSize: 18,
      textColor: '#1E3A8A',
      alignment: 'LEFT',
      bold: true,
      bgEnabled: true,
      bgColor: '#EFF6FF',
      bgStyle: 'inline'
    },
    subHeading: {
      fontFamily: 'Montserrat',
      fontSize: 14,
      textColor: '#2563EB',
      alignment: 'LEFT',
      bold: true,
      bgEnabled: false,
      bgColor: '#F1F5F9',
      bgStyle: 'inline'
    },
    body: {
      applyToBody: true,
      fontFamily: 'Roboto',
      fontSize: 11,
      textColor: '#1F2937',
      alignment: 'LEFT'
    }
  };

  try {
    const raw = PropertiesService.getUserProperties().getProperty('DOCUMENT_TYPOGRAPHY_PREFS');
    if (raw) {
      const parsed = JSON.parse(raw);
      return Object.assign({}, defaults, parsed);
    }
  } catch (e) {}

  return defaults;
}

/**
 * Persists typography preferences to user properties
 */
function saveTypographyPreferences(options) {
  try {
    PropertiesService.getUserProperties().setProperty('DOCUMENT_TYPOGRAPHY_PREFS', JSON.stringify(options));
  } catch (e) {}
}

/* ==========================================================================
   SIDEBAR UI (TABBED: 📊 TABLES & ⚡ CODE)
   ========================================================================== */

function getSidebarHtml() {
  const codePrefs = getUserPreferences();
  const tablePrefs = getTablePreferences();
  const typoPrefs = getTypographyPreferences();

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <link rel="preconnect" href="https://fonts.googleapis.com">
        <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
        <link href="https://fonts.googleapis.com/css2?family=Inconsolata:wght@400;700&family=Inter:wght@400;600;700&family=JetBrains+Mono:wght@400;700&family=Lato:wght@400;700&family=Merriweather:wght@400;700&family=Montserrat:wght@400;600;700&family=Open+Sans:wght@400;600;700&family=Roboto+Mono:wght@400;700&family=Roboto:wght@400;500;700&family=Source+Code+Pro:wght@400;700&family=Space+Mono:wght@400;700&display=swap" rel="stylesheet">
        <style>
          * { box-sizing: border-box; }
          body {
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
            font-size: 13px;
            color: #1f2937;
            padding: 12px;
            margin: 0;
            background: #ffffff;
          }
          
          /* Navigation Tabs */
          .tab-header {
            display: flex;
            background: #f3f4f6;
            border-radius: 8px;
            padding: 3px;
            margin-bottom: 14px;
            gap: 3px;
          }
          .tab-btn {
            flex: 1;
            border: none;
            background: transparent;
            padding: 7px 2px;
            font-size: 11px;
            font-weight: 600;
            color: #6b7280;
            border-radius: 6px;
            cursor: pointer;
            transition: all 0.2s ease;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 4px;
            white-space: nowrap;
          }
          .tab-btn.active {
            background: #ffffff;
            color: #1d4ed8;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
          }
          
          .tab-content { display: none; }
          .tab-content.active { display: block; }

          /* Typography Card Component */
          .typo-card {
            background: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            padding: 9px 10px;
            margin-bottom: 10px;
          }
          .typo-card-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 6px;
          }
          .typo-card-title {
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.4px;
            color: #334155;
          }
          .row-2col {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 6px;
            margin-bottom: 6px;
          }
          .checkbox-label {
            font-size: 11px;
            font-weight: 500;
            color: #475569;
            display: flex;
            align-items: center;
            gap: 5px;
            cursor: pointer;
          }

          /* Form Controls */
          .section-title {
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 0.5px;
            color: #6b7280;
            margin: 12px 0 6px 0;
          }
          .control-group { margin-bottom: 12px; }
          label { display: block; font-weight: 600; margin-bottom: 4px; font-size: 12px; color: #374151; }
          select, input[type="text"] {
            width: 100%;
            padding: 7px 10px;
            border: 1px solid #d1d5db;
            border-radius: 6px;
            font-size: 12px;
            background-color: #ffffff;
            color: #1f2937;
          }
          select:focus, input[type="text"]:focus {
            outline: none;
            border-color: #2563eb;
            box-shadow: 0 0 0 2px rgba(37,99,235,0.15);
          }

          .info-badge {
            font-size: 11px;
            color: #374151;
            margin-bottom: 12px;
            background: #f0fdf4;
            border: 1px solid #bbf7d0;
            padding: 7px 9px;
            border-radius: 6px;
            line-height: 1.4;
          }
          
          /* Color Pickers */
          .color-grid {
            display: grid;
            grid-template-columns: 1fr;
            gap: 6px;
            background: #f9fafb;
            padding: 10px;
            border-radius: 8px;
            border: 1px solid #e5e7eb;
            margin-bottom: 12px;
          }
          .color-row {
            display: flex;
            align-items: center;
            justify-content: space-between;
          }
          .color-row label {
            margin-bottom: 0;
            font-size: 11px;
            font-weight: 500;
            color: #4b5563;
          }
          .color-row input[type="color"] {
            width: 32px;
            height: 24px;
            border: 1px solid #cbd5e1;
            border-radius: 4px;
            cursor: pointer;
            padding: 0;
            background: none;
          }

          /* Live Previews */
          .preview-container {
            margin: 12px 0;
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            overflow: hidden;
            background: #ffffff;
            box-shadow: 0 1px 2px rgba(0,0,0,0.05);
          }
          .preview-header {
            background: #f9fafb;
            padding: 5px 10px;
            font-size: 11px;
            font-weight: 600;
            color: #6b7280;
            border-bottom: 1px solid #e5e7eb;
            display: flex;
            justify-content: space-between;
          }
          .preview-body { padding: 10px; }

          /* Table Live Preview */
          .preview-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 11px;
            transition: all 0.2s;
          }
          .preview-table th {
            padding: 7px 8px;
            border-width: 1px;
            border-style: solid;
            text-align: center;
            vertical-align: top;
            font-weight: 700;
            font-size: 12px;
          }
          .preview-table td {
            padding: 6px 8px;
            border-width: 1px;
            border-style: solid;
            text-align: left;
            vertical-align: top;
            font-size: 10px;
          }

          /* Code Live Preview */
          .preview-code-box {
            border-radius: 4px;
            padding: 8px 10px;
            font-size: 10px;
            font-weight: normal;
            line-height: 1.35;
            overflow-x: auto;
            border-width: 1px;
            border-style: solid;
          }

          /* Buttons */
          .btn-primary {
            width: 100%;
            background: #2563eb;
            color: #ffffff;
            border: none;
            padding: 10px;
            border-radius: 6px;
            font-weight: 600;
            cursor: pointer;
            font-size: 13px;
            margin-top: 6px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            transition: background 0.15s;
          }
          .btn-primary:hover { background: #1d4ed8; }
          .btn-primary:disabled { background: #93c5fd; cursor: not-allowed; }

          .btn-secondary {
            width: 100%;
            background: #ffffff;
            color: #2563eb;
            border: 1px solid #bfdbfe;
            padding: 8px;
            border-radius: 6px;
            font-weight: 600;
            cursor: pointer;
            font-size: 12px;
            margin-top: 6px;
            transition: all 0.15s;
          }
          .btn-secondary:hover { background: #eff6ff; border-color: #93c5fd; }
          .btn-secondary:disabled { border-color: #e5e7eb; color: #9ca3af; cursor: not-allowed; }

          .btn-danger {
            width: 100%;
            background: #ffffff;
            color: #dc2626;
            border: 1px solid #fecaca;
            padding: 8px;
            border-radius: 6px;
            font-weight: 600;
            cursor: pointer;
            font-size: 12px;
            margin-top: 6px;
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 6px;
            transition: all 0.15s;
          }
          .btn-danger:hover { background: #fef2f2; border-color: #f87171; }
          .btn-danger:disabled { border-color: #e5e7eb; color: #9ca3af; cursor: not-allowed; }

          .status-box {
            font-size: 11px;
            margin-top: 10px;
            padding: 8px;
            border-radius: 6px;
            text-align: center;
            font-weight: 500;
            display: none;
          }
          .status-box.success {
            display: block;
            background: #ecfdf5;
            color: #047857;
            border: 1px solid #a7f3d0;
          }
          .status-box.error {
            display: block;
            background: #fef2f2;
            color: #b91c1c;
            border: 1px solid #fecaca;
          }
          .status-box.loading {
            display: block;
            background: #eff6ff;
            color: #1d4ed8;
            border: 1px solid #bfdbfe;
          }
        </style>
      </head>
      <body>

        <!-- Master 1-Click Smart Auto-Formatter Card -->
        <div style="background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%); border-radius: 8px; padding: 10px 12px; margin-bottom: 8px; color: #ffffff; text-align: center; box-shadow: 0 2px 4px rgba(37,99,235,0.2);">
          <div style="font-weight: 700; font-size: 13px; margin-bottom: 2px; display:flex; align-items:center; justify-content:center; gap:6px;">
            <span>🚀</span> Smart Auto-Format
          </div>
          <div style="font-size: 11px; opacity: 0.9; margin-bottom: 8px;">
            Auto-detects Title, Headings, Code &amp; Tables
          </div>
          <button id="btnSmartAutoFormat" onclick="runSmartAutoFormat()" style="width: 100%; border: none; background: #ffffff; color: #1e3a8a; font-weight: 700; font-size: 12px; padding: 7px 12px; border-radius: 6px; cursor: pointer; transition: all 0.2s ease;">
            ⚡ Auto-Format Entire Document
          </button>
          <div id="smartAutoStatus" class="status-box" style="display:none; margin-top: 8px; text-align: left; background:#ffffff; color:#1e293b;"></div>
        </div>

        <!-- Remove Bookmarks utility row -->
        <div style="display:flex; align-items:center; gap:6px; margin-bottom:12px;">
          <button id="btnRemoveBookmarks" onclick="runRemoveBookmarks()" style="flex:1; border:1px solid #cbd5e1; background:#f8fafc; color:#334155; font-size:11px; font-weight:600; padding:6px 10px; border-radius:6px; cursor:pointer; transition: all 0.2s ease;" title="Remove all bookmark flags from the document">
            🔖 Remove All Bookmarks
          </button>
          <div id="bookmarkStatus" style="font-size:10px; color:#64748b; flex:1; min-width:0; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;"></div>
        </div>

        <!-- Navigation Tabs -->
        <div class="tab-header" style="display:grid; grid-template-columns: repeat(4,1fr); gap:2px;">
          <button class="tab-btn active" id="tabBtnTypography" onclick="switchTab('typography')" style="font-size:10px; padding:5px 2px;">
            <span>✍️</span> Typo
          </button>
          <button class="tab-btn" id="tabBtnTables" onclick="switchTab('tables')" style="font-size:10px; padding:5px 2px;">
            <span>📊</span> Tables
          </button>
          <button class="tab-btn" id="tabBtnCode" onclick="switchTab('code')" style="font-size:10px; padding:5px 2px;">
            <span>⚡</span> Code
          </button>
          <button class="tab-btn" id="tabBtnCleanup" onclick="switchTab('cleanup')" style="font-size:10px; padding:5px 2px;">
            <span>✨</span> Cleanup
          </button>
        </div>

        <!-- ==========================================
             TAB 1: ✍️ DOCUMENT TYPOGRAPHY & HEADINGS
             ========================================== -->
        <div id="tabContentTypography" class="tab-content active">
          <div class="info-badge">
            ✓ <strong>Title &amp; Headings:</strong> Full color picking, fonts &amp; sizes<br>
            ✓ <strong>Background Styles:</strong> Full-width header banner or inline text highlight<br>
            ✓ <strong>Body Text:</strong> Font changing, size up to 20pt, color &amp; alignment
          </div>

          <div class="control-group">
            <label>Typography Preset</label>
            <select id="typoPresetSelect" onchange="onTypographyPresetChange()">
              <option value="executive-navy" selected>Executive Navy (Montserrat &amp; Roboto)</option>
              <option value="modern-tech">Modern Tech (Inter Clean)</option>
              <option value="emerald-forest">Emerald Forest (Montserrat &amp; Slate)</option>
              <option value="editorial-classic">Editorial Classic (Georgia Serif)</option>
              <option value="crimson-luxe">Crimson Luxe (Ruby &amp; Slate)</option>
              <option value="custom">Custom Typography Palette...</option>
            </select>
          </div>

          <!-- Document Title Card -->
          <div class="typo-card">
            <div class="typo-card-header">
              <span class="typo-card-title">📖 Document Title (TITLE)</span>
              <label class="checkbox-label"><input type="checkbox" id="typoTitleBgToggle" onchange="updateTypographyPreview()"> Add Background</label>
            </div>
            <div class="row-2col">
              <div>
                <label>Font</label>
                <select id="typoTitleFontSelect" onchange="updateTypographyPreview()">
                  <option value="Montserrat" selected>Montserrat</option>
                  <option value="Inter">Inter</option>
                  <option value="Roboto">Roboto</option>
                  <option value="Arial">Arial</option>
                  <option value="Lato">Lato</option>
                  <option value="Open Sans">Open Sans</option>
                  <option value="Georgia">Georgia</option>
                  <option value="Merriweather">Merriweather</option>
                  <option value="Times New Roman">Times New Roman</option>
                  <option value="Calibri">Calibri</option>
                  <option value="Trebuchet MS">Trebuchet MS</option>
                  <option value="JetBrains Mono">JetBrains Mono</option>
                  <option value="Consolas">Consolas</option>
                </select>
              </div>
              <div>
                <label>Size</label>
                <select id="typoTitleSizeSelect" onchange="updateTypographyPreview()">
                  <option value="18">18 pt</option>
                  <option value="20">20 pt</option>
                  <option value="22">22 pt</option>
                  <option value="24">24 pt</option>
                  <option value="26" selected>26 pt (Default)</option>
                  <option value="28">28 pt</option>
                  <option value="30">30 pt</option>
                  <option value="32">32 pt</option>
                  <option value="36">36 pt (Large)</option>
                </select>
              </div>
            </div>
            <div class="row-2col">
              <div>
                <label>Alignment</label>
                <select id="typoTitleAlignSelect" onchange="updateTypographyPreview()">
                  <option value="CENTER" selected>Center</option>
                  <option value="LEFT">Left</option>
                  <option value="RIGHT">Right</option>
                </select>
              </div>
              <div>
                <label>Background Style</label>
                <select id="typoTitleBgStyleSelect" onchange="updateTypographyPreview()">
                  <option value="banner" selected>Full-Width Banner</option>
                  <option value="inline">Inline Highlight</option>
                </select>
              </div>
            </div>
            <div class="color-grid" style="margin-bottom:0; padding:6px 8px;">
              <div class="color-row">
                <label>Title Text Color:</label>
                <input type="color" id="typoTitleColor" value="#1E3A8A" onchange="setTypoCustomMode()">
              </div>
              <div class="color-row">
                <label>Background Color:</label>
                <input type="color" id="typoTitleBgColor" value="#EFF6FF" onchange="setTypoCustomMode()">
              </div>
            </div>
          </div>

          <!-- Main Heading 1 Card -->
          <div class="typo-card">
            <div class="typo-card-header">
              <span class="typo-card-title">📌 Main Heading 1</span>
              <label class="checkbox-label"><input type="checkbox" id="typoH1BgToggle" checked onchange="updateTypographyPreview()"> Add Background</label>
            </div>
            <div class="row-2col">
              <div>
                <label>Font</label>
                <select id="typoH1FontSelect" onchange="updateTypographyPreview()">
                  <option value="Montserrat" selected>Montserrat</option>
                  <option value="Inter">Inter</option>
                  <option value="Roboto">Roboto</option>
                  <option value="Arial">Arial</option>
                  <option value="Lato">Lato</option>
                  <option value="Open Sans">Open Sans</option>
                  <option value="Georgia">Georgia</option>
                  <option value="Merriweather">Merriweather</option>
                  <option value="Times New Roman">Times New Roman</option>
                  <option value="Calibri">Calibri</option>
                  <option value="Trebuchet MS">Trebuchet MS</option>
                  <option value="JetBrains Mono">JetBrains Mono</option>
                  <option value="Consolas">Consolas</option>
                </select>
              </div>
              <div>
                <label>Size</label>
                <select id="typoH1SizeSelect" onchange="updateTypographyPreview()">
                  <option value="14">14 pt</option>
                  <option value="15">15 pt</option>
                  <option value="16">16 pt</option>
                  <option value="17">17 pt</option>
                  <option value="18" selected>18 pt (Recommended)</option>
                  <option value="19">19 pt</option>
                  <option value="20">20 pt</option>
                  <option value="22">22 pt</option>
                  <option value="24">24 pt</option>
                </select>
              </div>
            </div>
            <div class="row-2col">
              <div>
                <label>Alignment</label>
                <select id="typoH1AlignSelect" onchange="updateTypographyPreview()">
                  <option value="LEFT" selected>Left</option>
                  <option value="CENTER">Center</option>
                  <option value="RIGHT">Right</option>
                </select>
              </div>
              <div>
                <label>Background Style</label>
                <select id="typoH1BgStyleSelect" onchange="updateTypographyPreview()">
                  <option value="banner" selected>Full-Width Banner</option>
                  <option value="inline">Inline Highlight</option>
                </select>
              </div>
            </div>
            <div class="color-grid" style="margin-bottom:0; padding:6px 8px;">
              <div class="color-row">
                <label>Heading Color:</label>
                <input type="color" id="typoH1Color" value="#1E3A8A" onchange="setTypoCustomMode()">
              </div>
              <div class="color-row">
                <label>Background Color:</label>
                <input type="color" id="typoH1BgColor" value="#EFF6FF" onchange="setTypoCustomMode()">
              </div>
            </div>
          </div>

          <!-- Sub-Headings Card -->
          <div class="typo-card">
            <div class="typo-card-header">
              <span class="typo-card-title">📑 Sub-Headings (H2, H3, Subtitle)</span>
              <label class="checkbox-label"><input type="checkbox" id="typoSubBgToggle" onchange="updateTypographyPreview()"> Add Background</label>
            </div>
            <div class="row-2col">
              <div>
                <label>Font</label>
                <select id="typoSubFontSelect" onchange="updateTypographyPreview()">
                  <option value="Montserrat" selected>Montserrat</option>
                  <option value="Inter">Inter</option>
                  <option value="Roboto">Roboto</option>
                  <option value="Arial">Arial</option>
                  <option value="Lato">Lato</option>
                  <option value="Open Sans">Open Sans</option>
                  <option value="Georgia">Georgia</option>
                  <option value="Merriweather">Merriweather</option>
                  <option value="Times New Roman">Times New Roman</option>
                  <option value="Calibri">Calibri</option>
                  <option value="Trebuchet MS">Trebuchet MS</option>
                  <option value="JetBrains Mono">JetBrains Mono</option>
                  <option value="Consolas">Consolas</option>
                </select>
              </div>
              <div>
                <label>Size</label>
                <select id="typoSubSizeSelect" onchange="updateTypographyPreview()">
                  <option value="10">10 pt</option>
                  <option value="11">11 pt</option>
                  <option value="12">12 pt</option>
                  <option value="13">13 pt</option>
                  <option value="14" selected>14 pt (Recommended)</option>
                  <option value="15">15 pt</option>
                  <option value="16">16 pt</option>
                  <option value="17">17 pt</option>
                  <option value="18">18 pt</option>
                  <option value="20">20 pt (Large)</option>
                </select>
              </div>
            </div>
            <div class="row-2col">
              <div>
                <label>Alignment</label>
                <select id="typoSubAlignSelect" onchange="updateTypographyPreview()">
                  <option value="LEFT" selected>Left</option>
                  <option value="CENTER">Center</option>
                  <option value="RIGHT">Right</option>
                </select>
              </div>
              <div>
                <label>Background Style</label>
                <select id="typoSubBgStyleSelect" onchange="updateTypographyPreview()">
                  <option value="inline" selected>Inline Highlight</option>
                  <option value="banner">Full-Width Banner</option>
                </select>
              </div>
            </div>
            <div class="color-grid" style="margin-bottom:0; padding:6px 8px;">
              <div class="color-row">
                <label>Heading Color:</label>
                <input type="color" id="typoSubColor" value="#2563EB" onchange="setTypoCustomMode()">
              </div>
              <div class="color-row">
                <label>Background Color:</label>
                <input type="color" id="typoSubBgColor" value="#F1F5F9" onchange="setTypoCustomMode()">
              </div>
            </div>
          </div>

          <!-- General Body Text Card -->
          <div class="typo-card">
            <div class="typo-card-header">
              <span class="typo-card-title">📝 General Body Text (NORMAL)</span>
              <label class="checkbox-label"><input type="checkbox" id="typoBodyApplyToggle" checked onchange="updateTypographyPreview()"> Apply to Body</label>
            </div>
            <div class="row-2col">
              <div>
                <label>Font Family</label>
                <select id="typoBodyFontSelect" onchange="updateTypographyPreview()">
                  <option value="Roboto" selected>Roboto (Clean Modern)</option>
                  <option value="Inter">Inter (Executive)</option>
                  <option value="Arial">Arial (Standard)</option>
                  <option value="Lato">Lato (Balanced)</option>
                  <option value="Open Sans">Open Sans (Readable)</option>
                  <option value="Montserrat">Montserrat</option>
                  <option value="Georgia">Georgia (Editorial Serif)</option>
                  <option value="Merriweather">Merriweather (Classic Serif)</option>
                  <option value="Times New Roman">Times New Roman</option>
                  <option value="Calibri">Calibri</option>
                  <option value="Trebuchet MS">Trebuchet MS</option>
                  <option value="Consolas">Consolas (Code)</option>
                  <option value="JetBrains Mono">JetBrains Mono (Code)</option>
                </select>
              </div>
              <div>
                <label>Font Size</label>
                <select id="typoBodySizeSelect" onchange="updateTypographyPreview()">
                  <option value="8">8 pt</option>
                  <option value="8.5">8.5 pt</option>
                  <option value="9">9 pt</option>
                  <option value="9.5">9.5 pt</option>
                  <option value="10">10 pt</option>
                  <option value="10.5">10.5 pt</option>
                  <option value="11" selected>11 pt (Default)</option>
                  <option value="11.5">11.5 pt</option>
                  <option value="12">12 pt</option>
                  <option value="13">13 pt</option>
                  <option value="14">14 pt</option>
                  <option value="15">15 pt</option>
                  <option value="16">16 pt</option>
                  <option value="17">17 pt</option>
                  <option value="18">18 pt</option>
                  <option value="19">19 pt</option>
                  <option value="20">20 pt (Extra Large)</option>
                </select>
              </div>
            </div>
            <div class="row-2col">
              <div>
                <label>Alignment</label>
                <select id="typoBodyAlignSelect" onchange="updateTypographyPreview()">
                  <option value="LEFT" selected>Left</option>
                  <option value="JUSTIFY">Justify</option>
                  <option value="CENTER">Center</option>
                  <option value="RIGHT">Right</option>
                </select>
              </div>
              <div class="color-row" style="margin-top:14px;">
                <label>Text Color:</label>
                <input type="color" id="typoBodyColor" value="#1F2937" onchange="setTypoCustomMode()">
              </div>
            </div>
          </div>

          <!-- Live Typography Preview -->
          <div class="preview-container">
            <div class="preview-header">
              <span>LIVE TYPOGRAPHY PREVIEW</span>
              <span>Document Styles</span>
            </div>
            <div class="preview-body" style="background:#ffffff; padding:12px;">
              <div id="prevTitleBox" style="margin-bottom:8px; transition: all 0.15s;">
                <span id="prevTitleText">Document Title</span>
              </div>
              <div id="prevH1Box" style="margin-bottom:6px; transition: all 0.15s;">
                <span id="prevH1Text">Heading 1: Executive Overview</span>
              </div>
              <div id="prevSubBox" style="margin-bottom:6px; transition: all 0.15s;">
                <span id="prevSubText">Heading 2: Key Methodology &amp; Details</span>
              </div>
              <div id="prevBodyBox" style="margin-bottom:0; line-height: 1.4; transition: all 0.15s;">
                <span id="prevBodyText">Standard document body text styled with chosen font family, size and color settings.</span>
              </div>
            </div>
          </div>

          <div class="section-title">⚡ Quick Format for Selection</div>
          <button type="button" class="btn-primary" id="btnFormatSelectedTypo" style="margin-top:0; width:100%; margin-bottom:6px;" onclick="runFormatSelectedTypography()">✍️ Auto-Detect &amp; Format Selection</button>
          <div class="row-2col">
            <button type="button" class="btn-secondary" style="margin-top:0;" onclick="runFormatSelectedAs('title')">📖 Apply as Title</button>
            <button type="button" class="btn-secondary" style="margin-top:0;" onclick="runFormatSelectedAs('heading1')">📌 Apply as Heading 1</button>
          </div>
          <div class="row-2col" style="margin-top:4px;">
            <button type="button" class="btn-secondary" style="margin-top:0;" onclick="runFormatSelectedAs('sub')">📑 Apply as Sub-Heading</button>
            <button type="button" class="btn-secondary" style="margin-top:0;" onclick="runFormatSelectedAs('body')">📝 Apply as Body</button>
          </div>

          <button class="btn-primary" id="btnFormatDocTypo" onclick="runFormatDocumentTypography()" style="margin-top:10px;">
            <span>✍️</span> Format All Document Typography
          </button>
          <button class="btn-danger" id="btnUndoDocTypo" onclick="runUndoDocumentTypography()">
            <span>↩</span> Undo Document Text Formatting
          </button>
          <div id="typoStatus" class="status-box"></div>
        </div>

        <!-- ==========================================
             TAB 2: 📊 PROFESSIONAL TABLE FORMATTER
             ========================================== -->
        <div id="tabContentTables" class="tab-content">
          <div class="info-badge">
            ✓ <strong>Header:</strong> Middle-aligned &amp; Larger Bold Font<br>
            ✓ <strong>Table Text:</strong> Left-aligned &amp; Top-aligned (Always on Top)<br>
            ✓ <strong>Table:</strong> Centered on Page
          </div>

          <div class="control-group">
            <label>Color Theme Preset</label>
            <select id="tableThemeSelect" onchange="onTableThemeChange()">
              <option value="corporate-navy">Corporate Navy (Royal & Ice)</option>
              <option value="emerald-mint">Emerald Mint (Forest & Mint)</option>
              <option value="royal-indigo">Royal Indigo (Deep Purple & Lilac)</option>
              <option value="sunset-crimson">Sunset Crimson (Ruby & Rose)</option>
              <option value="ocean-teal">Ocean Teal (Deep Teal & Aqua)</option>
              <option value="executive-slate">Executive Slate (Charcoal & Cool Gray)</option>
              <option value="custom">Custom Color Palette...</option>
            </select>
          </div>

          <div class="control-group">
            <label>Font Family</label>
            <select id="tableFontFamilySelect" onchange="updateTablePreview()">
              <option value="Roboto" selected>Roboto (Modern Clean)</option>
              <option value="Arial">Arial (Standard)</option>
              <option value="Inter">Inter (Executive)</option>
              <option value="Open Sans">Open Sans (Readable)</option>
              <option value="Lato">Lato (Balanced)</option>
              <option value="Montserrat">Montserrat (Geometric)</option>
              <option value="Calibri">Calibri</option>
              <option value="Trebuchet MS">Trebuchet MS</option>
              <option value="Georgia">Georgia (Editorial Serif)</option>
              <option value="Merriweather">Merriweather (Classic Serif)</option>
              <option value="Times New Roman">Times New Roman</option>
              <option value="Consolas">Consolas (Data Monospace)</option>
              <option value="JetBrains Mono">JetBrains Mono (Code Monospace)</option>
            </select>
          </div>

          <div class="control-group">
            <label>Header Font Size (Bigger)</label>
            <select id="tableHeaderFontSelect" onchange="updateTablePreview()">
              <option value="8.5">8.5 pt</option>
              <option value="9">9 pt</option>
              <option value="9.5">9.5 pt</option>
              <option value="10">10 pt</option>
              <option value="10.5">10.5 pt</option>
              <option value="11" selected>11 pt (Recommended)</option>
              <option value="11.5">11.5 pt</option>
              <option value="12">12 pt (Prominent)</option>
              <option value="13">13 pt (Large)</option>
              <option value="14">14 pt</option>
              <option value="15">15 pt</option>
              <option value="16">16 pt</option>
              <option value="17">17 pt</option>
              <option value="18">18 pt</option>
              <option value="19">19 pt</option>
              <option value="20">20 pt (Extra Large)</option>
            </select>
          </div>

          <div class="control-group">
            <label>Data Rows Font Size</label>
            <select id="tableBodyFontSelect" onchange="updateTablePreview()">
              <option value="8">8 pt</option>
              <option value="8.5">8.5 pt (Compact)</option>
              <option value="9">9 pt</option>
              <option value="9.5" selected>9.5 pt (Default)</option>
              <option value="10">10 pt</option>
              <option value="10.5">10.5 pt</option>
              <option value="11">11 pt</option>
              <option value="11.5">11.5 pt</option>
              <option value="12">12 pt</option>
              <option value="13">13 pt</option>
              <option value="14">14 pt</option>
              <option value="15">15 pt</option>
              <option value="16">16 pt</option>
              <option value="17">17 pt</option>
              <option value="18">18 pt</option>
              <option value="19">19 pt</option>
              <option value="20">20 pt (Extra Large)</option>
            </select>
          </div>

          <div class="control-group">
            <label>Cell Padding</label>
            <select id="tablePaddingSelect" onchange="updateTablePreview()">
              <option value="compact">Compact (Tight data grid)</option>
              <option value="normal" selected>Normal (Balanced & Clean)</option>
              <option value="relaxed">Relaxed (Spacious executive)</option>
            </select>
          </div>

          <div class="control-group">
            <label>Border Width</label>
            <select id="tableBorderWidthSelect" onchange="updateTablePreview()">
              <option value="0">0 pt &mdash; Borderless (No lines)</option>
              <option value="0.5">0.5 pt &mdash; Hairline (Ultra-thin)</option>
              <option value="1" selected>1 pt &mdash; Standard (Default)</option>
              <option value="1.5">1.5 pt &mdash; Medium (Visible)</option>
              <option value="2">2 pt &mdash; Bold (Strong grid)</option>
              <option value="3">3 pt &mdash; Heavy (Prominent)</option>
            </select>
          </div>

          <!-- Color Adjustments -->
          <div class="section-title">Color Palette</div>
          <div class="color-grid">
            <div class="color-row">
              <label>Header Background:</label>
              <input type="color" id="tHeaderBg" value="${tablePrefs.headerBg}" onchange="setTableCustomMode()">
            </div>
            <div class="color-row">
              <label>Header Text:</label>
              <input type="color" id="tHeaderText" value="${tablePrefs.headerText}" onchange="setTableCustomMode()">
            </div>
            <div class="color-row">
              <label>Alternating Row (Zebra):</label>
              <input type="color" id="tAltRowBg" value="${tablePrefs.altRowBg}" onchange="setTableCustomMode()">
            </div>
            <div class="color-row">
              <label>Base Row Background:</label>
              <input type="color" id="tNormalRowBg" value="${tablePrefs.normalRowBg}" onchange="setTableCustomMode()">
            </div>
            <div class="color-row">
              <label>Border Line Color:</label>
              <input type="color" id="tBorderColor" value="${tablePrefs.borderColor}" onchange="setTableCustomMode()">
            </div>
          </div>

          <!-- Inline Code Highlighting in Table Cells -->
          <div class="section-title">🔤 Inline Code Highlighting</div>
          <div class="control-group" style="flex-direction:row;align-items:center;gap:10px;">
            <label style="margin:0;flex:0 0 auto;">Detect &amp; Highlight Code Tokens</label>
            <label class="toggle-switch" style="margin:0 0 0 auto;">
              <input type="checkbox" id="tInlineCodeToggle" onchange="updateTablePreview()" ${tablePrefs.inlineCodeHighlight !== false ? 'checked' : ''}>
              <span class="toggle-slider"></span>
            </label>
          </div>
          <div id="inlineCodeColorGroup" class="color-grid">
            <div class="color-row">
              <label>Code Token Background:</label>
              <input type="color" id="tInlineCodeBg" value="${tablePrefs.inlineCodeBg || '#EFF1F3'}" onchange="updateTablePreview()">
            </div>
            <div class="color-row">
              <label>Code Token Text Color:</label>
              <input type="color" id="tInlineCodeColor" value="${tablePrefs.inlineCodeColor || '#B45309'}" onchange="updateTablePreview()">
            </div>
          </div>

          <div class="preview-container">
            <div class="preview-header">
              <span>LIVE TABLE PREVIEW</span>
              <span>Header: Middle | Body: Left</span>
            </div>
            <div class="preview-body">
              <table id="previewTableEl" class="preview-table">
                <thead>
                  <tr id="prevHeaderRow">
                    <th>Function</th>
                    <th>Type</th>
                    <th>Description</th>
                  </tr>
                </thead>
                <tbody>
                  <tr id="prevRow1">
                    <td><code id="prevCodeToken" style="font-family:Consolas,monospace;padding:1px 5px;border-radius:3px;font-size:0.9em;">fetchData()</code></td>
                    <td>async</td>
                    <td>Fetches remote data</td>
                  </tr>
                  <tr id="prevRow2">
                    <td>parseJSON()</td>
                    <td>string → object</td>
                    <td>Parses JSON response</td>
                  </tr>
                </tbody>
              </table>
            </div>
          </div>

          <button class="btn-primary" id="btnFormatAllTables" onclick="runFormatAllTables()">
            <span>📊</span> Format & Center All Tables
          </button>
          <button class="btn-secondary" id="btnFormatSelectedTable" onclick="runFormatSelectedTable()">
            Format Selected Table Only
          </button>
          <button class="btn-danger" id="btnUndoAllTables" onclick="runUndoAllTables()">
            <span>↩</span> Undo All Tables Formatting
          </button>
          <button class="btn-secondary" id="btnUndoSelectedTable" onclick="runUndoSelectedTable()" style="color:#6b7280;border-color:#e5e7eb;">
            Undo Selected Table Only
          </button>
          <div id="tableStatus" class="status-box"></div>
        </div>

        <!-- ==========================================
             TAB 2: ⚡ CODE HIGHLIGHTER
             ========================================== -->
        <div id="tabContentCode" class="tab-content">
          <div class="control-group">
            <label>Theme Preset</label>
            <select id="codeThemeSelect" onchange="onCodeThemeChange()">
              <option value="github-light">GitHub Light</option>
              <option value="one-dark">One Dark Pro (VS Code)</option>
              <option value="dracula">Dracula Dark</option>
              <option value="monokai">Monokai Dark</option>
              <option value="solarized-light">Solarized Light</option>
              <option value="custom">Custom Colors...</option>
            </select>
          </div>

          <div class="control-group">
            <label>Font Size</label>
            <select id="codeFontSizeSelect" onchange="updateCodePreview()">
              <option value="8">8 pt</option>
              <option value="8.5">8.5 pt (Compact)</option>
              <option value="9">9 pt</option>
              <option value="9.5" selected>9.5 pt (Default)</option>
              <option value="10">10 pt</option>
              <option value="10.5">10.5 pt</option>
              <option value="11">11 pt</option>
              <option value="11.5">11.5 pt</option>
              <option value="12">12 pt (Large)</option>
              <option value="13">13 pt</option>
              <option value="14">14 pt</option>
              <option value="15">15 pt</option>
              <option value="16">16 pt</option>
              <option value="17">17 pt</option>
              <option value="18">18 pt</option>
              <option value="19">19 pt</option>
              <option value="20">20 pt (Extra Large)</option>
            </select>
          </div>

          <div class="control-group">
            <label>Font Family</label>
            <select id="codeFontFamilySelect" onchange="updateCodePreview()">
              <option value="Consolas" selected>Consolas (Windows Default)</option>
              <option value="JetBrains Mono">JetBrains Mono (Developer Favorite)</option>
              <option value="Roboto Mono">Roboto Mono (Modern)</option>
              <option value="Courier New">Courier New (Classic)</option>
              <option value="Inconsolata">Inconsolata (Clean)</option>
              <option value="Source Code Pro">Source Code Pro (Adobe)</option>
              <option value="Space Mono">Space Mono (Geometric)</option>
              <option value="PT Mono">PT Mono</option>
              <option value="Ubuntu Mono">Ubuntu Mono</option>
            </select>
          </div>

          <div class="control-group">
            <label>Code Indentation</label>
            <select id="codeIndentSelect" onchange="updateCodePreview()">
              <option value="auto-2">Smart Auto-Indent (2 Spaces)</option>
              <option value="auto-4">Smart Auto-Indent (4 Spaces)</option>
              <option value="tab-2">Convert Tabs → 2 Spaces</option>
              <option value="tab-4">Convert Tabs → 4 Spaces</option>
              <option value="keep">Preserve Original Indentation</option>
            </select>
          </div>

          <div class="section-title">Block Colors</div>
          <div class="color-grid">
            <div class="color-row">
              <label>Background:</label>
              <input type="color" id="cBgColor" value="${codePrefs.bgColor}" onchange="setCodeCustomMode()">
            </div>
            <div class="color-row">
              <label>Text Color:</label>
              <input type="color" id="cTextColor" value="${codePrefs.textColor}" onchange="setCodeCustomMode()">
            </div>
            <div class="color-row">
              <label>Border:</label>
              <input type="color" id="cBorderColor" value="${codePrefs.borderColor}" onchange="setCodeCustomMode()">
            </div>
          </div>

          <div class="section-title">🎨 Token Colors</div>
          <div class="color-grid">
            <div class="color-row">
              <label>Keywords <small style="opacity:.65">(if, return, class…)</small>:</label>
              <input type="color" id="cKwColor" value="${codePrefs.keywordColor || '#CF222E'}" onchange="setCodeCustomMode()">
            </div>
            <div class="color-row">
              <label>Functions & Methods <small style="opacity:.65">(calculateTotal, run…)</small>:</label>
              <input type="color" id="cFnColor" value="${codePrefs.functionColor || '#8250DF'}" onchange="setCodeCustomMode()">
            </div>
            <div class="color-row">
              <label>Types & Classes <small style="opacity:.65">(Invoice, String…)</small>:</label>
              <input type="color" id="cTypeColor" value="${codePrefs.typeColor || '#953800'}" onchange="setCodeCustomMode()">
            </div>
            <div class="color-row">
              <label>Strings <small style="opacity:.65">("text", 'value')</small>:</label>
              <input type="color" id="cStrColor" value="${codePrefs.stringColor || '#0A3069'}" onchange="setCodeCustomMode()">
            </div>
            <div class="color-row">
              <label>Comments <small style="opacity:.65">(// notes)</small>:</label>
              <input type="color" id="cComColor" value="${codePrefs.commentColor || '#6E7781'}" onchange="setCodeCustomMode()">
            </div>
            <div class="color-row">
              <label>Numbers <small style="opacity:.65">(42, 3.14)</small>:</label>
              <input type="color" id="cNumColor" value="${codePrefs.numberColor || '#0550AE'}" onchange="setCodeCustomMode()">
            </div>
          </div>

          <!-- Code Preview -->
          <div class="preview-container">
            <div class="preview-header">
              <span>LIVE CODE PREVIEW</span>
              <span>Monospace</span>
            </div>
            <div class="preview-body">
              <div id="codePreviewBox" class="preview-code-box">
                <span id="pKw" style="color: #0550AE; font-weight: normal;">function</span> <span id="pFn">renderChart</span>() {<br>
                <span id="pIndent">&nbsp;&nbsp;</span><span id="pCom" style="color: #6E7781;">// Align & format</span><br>
                <span id="pIndent2">&nbsp;&nbsp;</span><span id="pKw2" style="color: #0550AE; font-weight: normal;">return</span> <span id="pStr" style="color: #0A3069;">"Success!"</span>;<br>
                }
              </div>
            </div>
          </div>

          <button class="btn-primary" id="btnHighlightAllCode" onclick="runHighlightAllCode()">
            <span>⚡</span> Highlight All Code
          </button>
          <button class="btn-secondary" id="btnHighlightSelectedCode" onclick="runHighlightSelectedCode()">
            Format Selected Code
          </button>
          <button class="btn-secondary" id="btnIndentAllCode" onclick="runIndentAllCode()">
            <span>📐</span> Auto-Indent All Code Blocks
          </button>
          <button class="btn-danger" id="btnUndoAllCode" onclick="runUndoAllCode()">
            <span>↩</span> Undo All Code Blocks Formatting
          </button>
          <button class="btn-secondary" id="btnUndoSelectedCode" onclick="runUndoSelectedCode()" style="color:#6b7280;border-color:#e5e7eb;">
            Undo Selected Code Block
          </button>
          <div id="codeStatus" class="status-box"></div>
        </div>

        <!-- ==========================================
             TAB 4: ✨ TEXT CLEANUP & SPACING
             ========================================== -->
        <div id="tabContentCleanup" class="tab-content">

          <div class="info-badge">
            ✓ <strong>Text Cleanup:</strong> Whitespace, smart quotes, em-dashes<br>
            ✓ <strong>Spacing:</strong> Before/after headings, keep-with-next<br>
            ✓ <strong>Inline Markdown:</strong> **bold**, *italic*, ~~strike~~, ==highlight==<br>
            ✓ <strong>Heading Numbers:</strong> 1. / 1.1 / 1.1.1 auto-outline
          </div>

          <!-- Text Cleanup Card -->
          <div class="typo-card">
            <div class="typo-card-header"><span class="typo-card-title">🧹 Text Cleanup</span></div>
            <div style="padding:8px 10px;">
              <label style="display:flex;align-items:center;gap:6px;font-size:11px;margin-bottom:6px;">
                <input type="checkbox" id="cleanSmartTypo" checked>
                Smart typography (curly quotes, em-dash, ellipsis)
              </label>
              <label style="display:flex;align-items:center;gap:6px;font-size:11px;margin-bottom:6px;">
                <input type="checkbox" id="cleanDoubleSpaces" checked>
                Strip double spaces
              </label>
              <label style="display:flex;align-items:center;gap:6px;font-size:11px;margin-bottom:10px;">
                <input type="checkbox" id="cleanBlankLines" checked>
                Normalize excessive blank lines (max 2)
              </label>
              <button class="btn-primary" id="btnCleanupText" onclick="runCleanupText()" style="margin-top:0;">
                🧹 Run Text Cleanup
              </button>
            </div>
          </div>

          <!-- Paragraph Spacing Card -->
          <div class="typo-card">
            <div class="typo-card-header"><span class="typo-card-title">📐 Paragraph Spacing</span></div>
            <div style="padding:8px 10px;">
              <div class="row-2col">
                <div class="control-group" style="margin:0;">
                  <label>Space Before Heading (pt)</label>
                  <input type="number" id="spacingHBefore" value="14" min="0" max="72" style="width:100%;">
                </div>
                <div class="control-group" style="margin:0;">
                  <label>Space After Heading (pt)</label>
                  <input type="number" id="spacingHAfter" value="4" min="0" max="72" style="width:100%;">
                </div>
              </div>
              <div class="control-group" style="margin-top:6px;">
                <label>Space After Body Paragraph (pt)</label>
                <input type="number" id="spacingBodyAfter" value="6" min="0" max="72" style="width:100%;">
              </div>
              <label style="display:flex;align-items:center;gap:6px;font-size:11px;margin-bottom:10px;">
                <input type="checkbox" id="spacingKeepWithNext" checked>
                Pin headings to next paragraph (prevent orphans)
              </label>
              <button class="btn-primary" id="btnNormalizeSpacing" onclick="runNormalizeSpacing()" style="margin-top:0;">
                📐 Normalize Paragraph Spacing
              </button>
            </div>
          </div>

          <!-- Inline Markdown Card -->
          <div class="typo-card">
            <div class="typo-card-header"><span class="typo-card-title">✦ Inline Markdown Formatter</span></div>
            <div style="padding:8px 10px;">
              <div class="info-badge" style="margin-bottom:8px; font-size:10px;">
                <strong>**bold**</strong> &nbsp;·&nbsp; <em>*italic*</em> &nbsp;·&nbsp;
                <span style="text-decoration:line-through;">~~strike~~</span> &nbsp;·&nbsp;
                <span style="background:#ffff00;">==highlight==</span>
              </div>
              <button class="btn-primary" id="btnInlineMarkdown" onclick="runInlineMarkdown()" style="margin-top:0;">
                ✦ Apply Inline Markdown Formatting
              </button>
            </div>
          </div>

          <!-- Heading Numbers Card -->
          <div class="typo-card">
            <div class="typo-card-header"><span class="typo-card-title">🔢 Heading Numbering</span></div>
            <div style="padding:8px 10px;">
              <div class="info-badge" style="margin-bottom:8px; font-size:10px;">
                H1 → 1.&nbsp;&nbsp; H2 → 1.1&nbsp;&nbsp; H3 → 1.1.1
              </div>
              <button class="btn-primary" id="btnHeadingNumbers" onclick="runHeadingNumbers()" style="margin-top:0;margin-bottom:6px;">
                🔢 Auto-Number Headings
              </button>
            </div>
          </div>

          <!-- Heading Case Card -->
          <div class="typo-card">
            <div class="typo-card-header"><span class="typo-card-title">🔡 Heading Case Transform</span></div>
            <div style="padding:8px 10px;">
              <div class="row-2col">
                <button class="btn-secondary" style="margin-top:0;" onclick="runHeadingCase('upper')">UPPERCASE</button>
                <button class="btn-secondary" style="margin-top:0;" onclick="runHeadingCase('title')">Title Case</button>
              </div>
              <div class="row-2col" style="margin-top:4px;">
                <button class="btn-secondary" style="margin-top:0;" onclick="runHeadingCase('sentence')">Sentence case</button>
                <button class="btn-secondary" style="margin-top:0;" onclick="runHeadingCase('lower')">lowercase</button>
              </div>
            </div>
          </div>

          <!-- List Converter Card -->
          <div class="typo-card">
            <div class="typo-card-header"><span class="typo-card-title">📋 Markdown List Converter</span></div>
            <div style="padding:8px 10px;">
              <div class="info-badge" style="margin-bottom:8px;font-size:10px;">
                Converts <code>- item</code>, <code>• item</code>, <code>* item</code>, <code>1. item</code>
                lines into proper Google Docs list items.
              </div>
              <button class="btn-primary" id="btnConvertLists" onclick="runConvertLists()" style="margin-top:0;">
                📋 Convert Markdown Lists
              </button>
            </div>
          </div>

          <div id="cleanupStatus" class="status-box"></div>
        </div>

        <script>
          /* Tab Switching */
          function switchTab(tab) {
            document.getElementById('tabBtnTypography').classList.toggle('active', tab === 'typography');
            document.getElementById('tabBtnTables').classList.toggle('active', tab === 'tables');
            document.getElementById('tabBtnCode').classList.toggle('active', tab === 'code');
            document.getElementById('tabBtnCleanup').classList.toggle('active', tab === 'cleanup');
            document.getElementById('tabContentTypography').classList.toggle('active', tab === 'typography');
            document.getElementById('tabContentTables').classList.toggle('active', tab === 'tables');
            document.getElementById('tabContentCode').classList.toggle('active', tab === 'code');
            document.getElementById('tabContentCleanup').classList.toggle('active', tab === 'cleanup');
          }

          /* =========================================================
             CLEANUP TAB LOGIC
             ========================================================= */
          function setCleanupStatus(msg, type) {
            const el = document.getElementById('cleanupStatus');
            if (!el) return;
            el.style.display = msg ? 'block' : 'none';
            el.className = 'status-box' + (type ? ' ' + type : '');
            el.innerText = msg || '';
          }

          function runCleanupText() {
            setCleanupStatus('Cleaning up text...', 'loading');
            const btn = document.getElementById('btnCleanupText');
            if (btn) btn.disabled = true;
            const opts = {
              smartTypography: document.getElementById('cleanSmartTypo').checked,
              stripDoubleSpaces: document.getElementById('cleanDoubleSpaces').checked,
              normalizeBlankLines: document.getElementById('cleanBlankLines').checked
            };
            google.script.run
              .withSuccessHandler(res => {
                if (btn) btn.disabled = false;
                setCleanupStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                if (btn) btn.disabled = false;
                setCleanupStatus('Error: ' + err, 'error');
              })
              .cleanupDocumentText(opts);
          }

          function runNormalizeSpacing() {
            setCleanupStatus('Normalizing paragraph spacing...', 'loading');
            const btn = document.getElementById('btnNormalizeSpacing');
            if (btn) btn.disabled = true;
            const opts = {
              headingSpacingBefore: Number(document.getElementById('spacingHBefore').value) || 14,
              headingSpacingAfter:  Number(document.getElementById('spacingHAfter').value)  || 4,
              bodySpacingAfter:     Number(document.getElementById('spacingBodyAfter').value) || 6,
              keepHeadingsWithNext: document.getElementById('spacingKeepWithNext').checked
            };
            google.script.run
              .withSuccessHandler(res => {
                if (btn) btn.disabled = false;
                setCleanupStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                if (btn) btn.disabled = false;
                setCleanupStatus('Error: ' + err, 'error');
              })
              .normalizeDocumentSpacing(opts);
          }

          function runInlineMarkdown() {
            setCleanupStatus('Applying inline markdown formatting...', 'loading');
            const btn = document.getElementById('btnInlineMarkdown');
            if (btn) btn.disabled = true;
            google.script.run
              .withSuccessHandler(res => {
                if (btn) btn.disabled = false;
                setCleanupStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                if (btn) btn.disabled = false;
                setCleanupStatus('Error: ' + err, 'error');
              })
              .applyInlineMarkdown();
          }

          function runHeadingNumbers() {
            setCleanupStatus('Auto-numbering headings...', 'loading');
            const btn = document.getElementById('btnHeadingNumbers');
            if (btn) btn.disabled = true;
            google.script.run
              .withSuccessHandler(res => {
                if (btn) btn.disabled = false;
                setCleanupStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                if (btn) btn.disabled = false;
                setCleanupStatus('Error: ' + err, 'error');
              })
              .applyHeadingNumbers();
          }

          function runHeadingCase(caseType) {
            setCleanupStatus('Applying ' + caseType + ' case to headings...', 'loading');
            google.script.run
              .withSuccessHandler(res => setCleanupStatus(res.message, res.success ? 'success' : 'error'))
              .withFailureHandler(err => setCleanupStatus('Error: ' + err, 'error'))
              .transformHeadingCase(caseType);
          }

          function runConvertLists() {
            setCleanupStatus('Converting markdown lists...', 'loading');
            const btn = document.getElementById('btnConvertLists');
            if (btn) btn.disabled = true;
            google.script.run
              .withSuccessHandler(res => {
                if (btn) btn.disabled = false;
                setCleanupStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                if (btn) btn.disabled = false;
                setCleanupStatus('Error: ' + err, 'error');
              })
              .convertMarkdownLists();
          }

          /* =========================================================
             DOCUMENT TYPOGRAPHY LOGIC
             ========================================================= */
          const TYPOGRAPHY_PRESETS = {
            'executive-navy': {
              titleFont: 'Montserrat', titleSize: '26', titleColor: '#1E3A8A', titleAlign: 'CENTER', titleBgEnabled: false, titleBgColor: '#EFF6FF', titleBgStyle: 'banner',
              h1Font: 'Montserrat', h1Size: '18', h1Color: '#1E3A8A', h1Align: 'LEFT', h1BgEnabled: true, h1BgColor: '#EFF6FF', h1BgStyle: 'banner',
              subFont: 'Montserrat', subSize: '14', subColor: '#2563EB', subAlign: 'LEFT', subBgEnabled: false, subBgColor: '#F1F5F9', subBgStyle: 'inline',
              bodyFont: 'Roboto', bodySize: '11', bodyColor: '#1F2937', bodyAlign: 'LEFT', bodyApply: true
            },
            'modern-tech': {
              titleFont: 'Inter', titleSize: '26', titleColor: '#312E81', titleAlign: 'CENTER', titleBgEnabled: false, titleBgColor: '#EEF2FF', titleBgStyle: 'banner',
              h1Font: 'Inter', h1Size: '18', h1Color: '#4338CA', h1Align: 'LEFT', h1BgEnabled: true, h1BgColor: '#EEF2FF', h1BgStyle: 'banner',
              subFont: 'Inter', subSize: '14', subColor: '#6366F1', subAlign: 'LEFT', subBgEnabled: true, subBgColor: '#F5F3FF', subBgStyle: 'inline',
              bodyFont: 'Inter', bodySize: '10.5', bodyColor: '#111827', bodyAlign: 'LEFT', bodyApply: true
            },
            'emerald-forest': {
              titleFont: 'Montserrat', titleSize: '26', titleColor: '#064E3B', titleAlign: 'CENTER', titleBgEnabled: false, titleBgColor: '#ECFDF5', titleBgStyle: 'banner',
              h1Font: 'Montserrat', h1Size: '18', h1Color: '#065F46', h1Align: 'LEFT', h1BgEnabled: true, h1BgColor: '#ECFDF5', h1BgStyle: 'banner',
              subFont: 'Montserrat', subSize: '14', subColor: '#0D9488', subAlign: 'LEFT', subBgEnabled: false, subBgColor: '#F0FDFA', subBgStyle: 'inline',
              bodyFont: 'Roboto', bodySize: '11', bodyColor: '#1F2937', bodyAlign: 'LEFT', bodyApply: true
            },
            'editorial-classic': {
              titleFont: 'Georgia', titleSize: '28', titleColor: '#18181B', titleAlign: 'CENTER', titleBgEnabled: false, titleBgColor: '#F4F4F5', titleBgStyle: 'inline',
              h1Font: 'Georgia', h1Size: '18', h1Color: '#27272A', h1Align: 'LEFT', h1BgEnabled: true, h1BgColor: '#F4F4F5', h1BgStyle: 'inline',
              subFont: 'Georgia', subSize: '14', subColor: '#52525B', subAlign: 'LEFT', subBgEnabled: false, subBgColor: '#F4F4F5', subBgStyle: 'inline',
              bodyFont: 'Georgia', bodySize: '11', bodyColor: '#27272A', bodyAlign: 'LEFT', bodyApply: true
            },
            'crimson-luxe': {
              titleFont: 'Montserrat', titleSize: '26', titleColor: '#881337', titleAlign: 'CENTER', titleBgEnabled: false, titleBgColor: '#FFF1F2', titleBgStyle: 'banner',
              h1Font: 'Montserrat', h1Size: '18', h1Color: '#9F1239', h1Align: 'LEFT', h1BgEnabled: true, h1BgColor: '#FFF1F2', h1BgStyle: 'banner',
              subFont: 'Montserrat', subSize: '14', subColor: '#BE123C', subAlign: 'LEFT', subBgEnabled: false, subBgColor: '#FFE4E6', subBgStyle: 'inline',
              bodyFont: 'Roboto', bodySize: '11', bodyColor: '#1F2937', bodyAlign: 'LEFT', bodyApply: true
            }
          };

          const savedTypoPrefs = ${JSON.stringify(typoPrefs)};
          document.getElementById('typoPresetSelect').value = savedTypoPrefs.preset || 'executive-navy';
          if (savedTypoPrefs.title) {
            document.getElementById('typoTitleFontSelect').value = savedTypoPrefs.title.fontFamily || 'Montserrat';
            document.getElementById('typoTitleSizeSelect').value = savedTypoPrefs.title.fontSize || '26';
            document.getElementById('typoTitleAlignSelect').value = savedTypoPrefs.title.alignment || 'CENTER';
            document.getElementById('typoTitleBgToggle').checked = !!savedTypoPrefs.title.bgEnabled;
            document.getElementById('typoTitleBgStyleSelect').value = savedTypoPrefs.title.bgStyle || 'banner';
            document.getElementById('typoTitleColor').value = savedTypoPrefs.title.textColor || '#1E3A8A';
            document.getElementById('typoTitleBgColor').value = savedTypoPrefs.title.bgColor || '#EFF6FF';
          }
          if (savedTypoPrefs.heading1) {
            document.getElementById('typoH1FontSelect').value = savedTypoPrefs.heading1.fontFamily || 'Montserrat';
            document.getElementById('typoH1SizeSelect').value = savedTypoPrefs.heading1.fontSize || '18';
            document.getElementById('typoH1AlignSelect').value = savedTypoPrefs.heading1.alignment || 'LEFT';
            document.getElementById('typoH1BgToggle').checked = savedTypoPrefs.heading1.bgEnabled !== undefined ? savedTypoPrefs.heading1.bgEnabled : true;
            document.getElementById('typoH1BgStyleSelect').value = savedTypoPrefs.heading1.bgStyle || 'banner';
            document.getElementById('typoH1Color').value = savedTypoPrefs.heading1.textColor || '#1E3A8A';
            document.getElementById('typoH1BgColor').value = savedTypoPrefs.heading1.bgColor || '#EFF6FF';
          }
          if (savedTypoPrefs.subHeading) {
            document.getElementById('typoSubFontSelect').value = savedTypoPrefs.subHeading.fontFamily || 'Montserrat';
            document.getElementById('typoSubSizeSelect').value = savedTypoPrefs.subHeading.fontSize || '14';
            document.getElementById('typoSubAlignSelect').value = savedTypoPrefs.subHeading.alignment || 'LEFT';
            document.getElementById('typoSubBgToggle').checked = !!savedTypoPrefs.subHeading.bgEnabled;
            document.getElementById('typoSubBgStyleSelect').value = savedTypoPrefs.subHeading.bgStyle || 'inline';
            document.getElementById('typoSubColor').value = savedTypoPrefs.subHeading.textColor || '#2563EB';
            document.getElementById('typoSubBgColor').value = savedTypoPrefs.subHeading.bgColor || '#F1F5F9';
          }
          if (savedTypoPrefs.body) {
            document.getElementById('typoBodyFontSelect').value = savedTypoPrefs.body.fontFamily || 'Roboto';
            document.getElementById('typoBodySizeSelect').value = savedTypoPrefs.body.fontSize || '11';
            document.getElementById('typoBodyAlignSelect').value = savedTypoPrefs.body.alignment || 'LEFT';
            document.getElementById('typoBodyColor').value = savedTypoPrefs.body.textColor || '#1F2937';
            document.getElementById('typoBodyApplyToggle').checked = savedTypoPrefs.body.applyToBody !== undefined ? savedTypoPrefs.body.applyToBody : true;
          }

          function onTypographyPresetChange() {
            const val = document.getElementById('typoPresetSelect').value;
            if (val !== 'custom' && TYPOGRAPHY_PRESETS[val]) {
              const p = TYPOGRAPHY_PRESETS[val];
              document.getElementById('typoTitleFontSelect').value = p.titleFont;
              document.getElementById('typoTitleSizeSelect').value = p.titleSize;
              document.getElementById('typoTitleAlignSelect').value = p.titleAlign;
              document.getElementById('typoTitleBgToggle').checked = p.titleBgEnabled;
              document.getElementById('typoTitleBgStyleSelect').value = p.titleBgStyle;
              document.getElementById('typoTitleColor').value = p.titleColor;
              document.getElementById('typoTitleBgColor').value = p.titleBgColor;

              document.getElementById('typoH1FontSelect').value = p.h1Font;
              document.getElementById('typoH1SizeSelect').value = p.h1Size;
              document.getElementById('typoH1AlignSelect').value = p.h1Align;
              document.getElementById('typoH1BgToggle').checked = p.h1BgEnabled;
              document.getElementById('typoH1BgStyleSelect').value = p.h1BgStyle;
              document.getElementById('typoH1Color').value = p.h1Color;
              document.getElementById('typoH1BgColor').value = p.h1BgColor;

              document.getElementById('typoSubFontSelect').value = p.subFont;
              document.getElementById('typoSubSizeSelect').value = p.subSize;
              document.getElementById('typoSubAlignSelect').value = p.subAlign;
              document.getElementById('typoSubBgToggle').checked = p.subBgEnabled;
              document.getElementById('typoSubBgStyleSelect').value = p.subBgStyle;
              document.getElementById('typoSubColor').value = p.subColor;
              document.getElementById('typoSubBgColor').value = p.subBgColor;

              document.getElementById('typoBodyFontSelect').value = p.bodyFont;
              document.getElementById('typoBodySizeSelect').value = p.bodySize;
              document.getElementById('typoBodyAlignSelect').value = p.bodyAlign;
              document.getElementById('typoBodyColor').value = p.bodyColor;
              document.getElementById('typoBodyApplyToggle').checked = p.bodyApply;
            }
            updateTypographyPreview();
          }

          function setTypoCustomMode() {
            document.getElementById('typoPresetSelect').value = 'custom';
            updateTypographyPreview();
          }

          function updateTypographyPreview() {
            // Title Preview
            const tFont = document.getElementById('typoTitleFontSelect').value;
            const tSize = parseFloat(document.getElementById('typoTitleSizeSelect').value) || 26;
            const tAlign = document.getElementById('typoTitleAlignSelect').value;
            const tBgEnabled = document.getElementById('typoTitleBgToggle').checked;
            const tBgStyle = document.getElementById('typoTitleBgStyleSelect').value;
            const tColor = document.getElementById('typoTitleColor').value;
            const tBgColor = document.getElementById('typoTitleBgColor').value;

            const prevTitleBox = document.getElementById('prevTitleBox');
            const prevTitleText = document.getElementById('prevTitleText');
            prevTitleBox.style.textAlign = tAlign.toLowerCase();
            prevTitleText.style.fontFamily = tFont + ', sans-serif';
            prevTitleText.style.fontSize = Math.round(tSize * 0.72) + 'px';
            prevTitleText.style.fontWeight = 'bold';
            prevTitleText.style.color = tColor;

            if (tBgEnabled && tBgStyle === 'banner') {
              prevTitleBox.style.backgroundColor = tBgColor;
              prevTitleBox.style.padding = '6px 8px';
              prevTitleBox.style.borderRadius = '4px';
              prevTitleText.style.backgroundColor = 'transparent';
              prevTitleText.style.padding = '0';
            } else if (tBgEnabled && tBgStyle === 'inline') {
              prevTitleBox.style.backgroundColor = 'transparent';
              prevTitleBox.style.padding = '0';
              prevTitleText.style.backgroundColor = tBgColor;
              prevTitleText.style.padding = '1px 5px';
              prevTitleText.style.borderRadius = '3px';
            } else {
              prevTitleBox.style.backgroundColor = 'transparent';
              prevTitleBox.style.padding = '0';
              prevTitleText.style.backgroundColor = 'transparent';
              prevTitleText.style.padding = '0';
            }

            // H1 Preview
            const h1Font = document.getElementById('typoH1FontSelect').value;
            const h1Size = parseFloat(document.getElementById('typoH1SizeSelect').value) || 18;
            const h1Align = document.getElementById('typoH1AlignSelect').value;
            const h1BgEnabled = document.getElementById('typoH1BgToggle').checked;
            const h1BgStyle = document.getElementById('typoH1BgStyleSelect').value;
            const h1Color = document.getElementById('typoH1Color').value;
            const h1BgColor = document.getElementById('typoH1BgColor').value;

            const prevH1Box = document.getElementById('prevH1Box');
            const prevH1Text = document.getElementById('prevH1Text');
            prevH1Box.style.textAlign = h1Align.toLowerCase();
            prevH1Text.style.fontFamily = h1Font + ', sans-serif';
            prevH1Text.style.fontSize = Math.round(h1Size * 0.75) + 'px';
            prevH1Text.style.fontWeight = 'bold';
            prevH1Text.style.color = h1Color;

            if (h1BgEnabled && h1BgStyle === 'banner') {
              prevH1Box.style.backgroundColor = h1BgColor;
              prevH1Box.style.padding = '5px 8px';
              prevH1Box.style.borderRadius = '4px';
              prevH1Text.style.backgroundColor = 'transparent';
              prevH1Text.style.padding = '0';
            } else if (h1BgEnabled && h1BgStyle === 'inline') {
              prevH1Box.style.backgroundColor = 'transparent';
              prevH1Box.style.padding = '0';
              prevH1Text.style.backgroundColor = h1BgColor;
              prevH1Text.style.padding = '1px 5px';
              prevH1Text.style.borderRadius = '3px';
            } else {
              prevH1Box.style.backgroundColor = 'transparent';
              prevH1Box.style.padding = '0';
              prevH1Text.style.backgroundColor = 'transparent';
              prevH1Text.style.padding = '0';
            }

            // Sub-Heading Preview
            const subFont = document.getElementById('typoSubFontSelect').value;
            const subSize = parseFloat(document.getElementById('typoSubSizeSelect').value) || 14;
            const subAlign = document.getElementById('typoSubAlignSelect').value;
            const subBgEnabled = document.getElementById('typoSubBgToggle').checked;
            const subBgStyle = document.getElementById('typoSubBgStyleSelect').value;
            const subColor = document.getElementById('typoSubColor').value;
            const subBgColor = document.getElementById('typoSubBgColor').value;

            const prevSubBox = document.getElementById('prevSubBox');
            const prevSubText = document.getElementById('prevSubText');
            prevSubBox.style.textAlign = subAlign.toLowerCase();
            prevSubText.style.fontFamily = subFont + ', sans-serif';
            prevSubText.style.fontSize = Math.round(subSize * 0.8) + 'px';
            prevSubText.style.fontWeight = 'bold';
            prevSubText.style.color = subColor;

            if (subBgEnabled && subBgStyle === 'banner') {
              prevSubBox.style.backgroundColor = subBgColor;
              prevSubBox.style.padding = '4px 7px';
              prevSubBox.style.borderRadius = '3px';
              prevSubText.style.backgroundColor = 'transparent';
              prevSubText.style.padding = '0';
            } else if (subBgEnabled && subBgStyle === 'inline') {
              prevSubBox.style.backgroundColor = 'transparent';
              prevSubBox.style.padding = '0';
              prevSubText.style.backgroundColor = subBgColor;
              prevSubText.style.padding = '1px 4px';
              prevSubText.style.borderRadius = '2px';
            } else {
              prevSubBox.style.backgroundColor = 'transparent';
              prevSubBox.style.padding = '0';
              prevSubText.style.backgroundColor = 'transparent';
              prevSubText.style.padding = '0';
            }

            // Body Preview
            const bFont = document.getElementById('typoBodyFontSelect').value;
            const bSize = parseFloat(document.getElementById('typoBodySizeSelect').value) || 11;
            const bAlign = document.getElementById('typoBodyAlignSelect').value;
            const bColor = document.getElementById('typoBodyColor').value;

            const prevBodyBox = document.getElementById('prevBodyBox');
            const prevBodyText = document.getElementById('prevBodyText');
            prevBodyBox.style.textAlign = bAlign.toLowerCase();
            prevBodyText.style.fontFamily = bFont + ', sans-serif';
            prevBodyText.style.fontSize = Math.round(bSize * 0.9) + 'px';
            prevBodyText.style.color = bColor;
          }

          function getTypographyOptions() {
            return {
              preset: document.getElementById('typoPresetSelect').value,
              title: {
                fontFamily: document.getElementById('typoTitleFontSelect').value,
                fontSize: parseFloat(document.getElementById('typoTitleSizeSelect').value) || 26,
                textColor: document.getElementById('typoTitleColor').value,
                alignment: document.getElementById('typoTitleAlignSelect').value,
                bold: true,
                bgEnabled: document.getElementById('typoTitleBgToggle').checked,
                bgColor: document.getElementById('typoTitleBgColor').value,
                bgStyle: document.getElementById('typoTitleBgStyleSelect').value
              },
              heading1: {
                fontFamily: document.getElementById('typoH1FontSelect').value,
                fontSize: parseFloat(document.getElementById('typoH1SizeSelect').value) || 18,
                textColor: document.getElementById('typoH1Color').value,
                alignment: document.getElementById('typoH1AlignSelect').value,
                bold: true,
                bgEnabled: document.getElementById('typoH1BgToggle').checked,
                bgColor: document.getElementById('typoH1BgColor').value,
                bgStyle: document.getElementById('typoH1BgStyleSelect').value
              },
              subHeading: {
                fontFamily: document.getElementById('typoSubFontSelect').value,
                fontSize: parseFloat(document.getElementById('typoSubSizeSelect').value) || 14,
                textColor: document.getElementById('typoSubColor').value,
                alignment: document.getElementById('typoSubAlignSelect').value,
                bold: true,
                bgEnabled: document.getElementById('typoSubBgToggle').checked,
                bgColor: document.getElementById('typoSubBgColor').value,
                bgStyle: document.getElementById('typoSubBgStyleSelect').value
              },
              body: {
                applyToBody: document.getElementById('typoBodyApplyToggle').checked,
                fontFamily: document.getElementById('typoBodyFontSelect').value,
                fontSize: parseFloat(document.getElementById('typoBodySizeSelect').value) || 11,
                textColor: document.getElementById('typoBodyColor').value,
                alignment: document.getElementById('typoBodyAlignSelect').value
              }
            };
          }

          function setTypographyStatus(msg, type) {
            const el = document.getElementById('typoStatus');
            el.className = 'status-box ' + type;
            el.innerText = msg;
          }

          function runSmartAutoFormat() {
            const statusEl = document.getElementById('smartAutoStatus');
            if (statusEl) {
              statusEl.style.display = 'block';
              statusEl.className = 'status-box loading';
              statusEl.innerText = 'Scanning & auto-formatting entire document...';
            }
            const btn = document.getElementById('btnSmartAutoFormat');
            if (btn) btn.disabled = true;

            google.script.run
              .withSuccessHandler(res => {
                if (btn) btn.disabled = false;
                if (statusEl) {
                  statusEl.className = 'status-box ' + (res.success ? 'success' : 'error');
                  statusEl.innerText = res.message;
                }
              })
              .withFailureHandler(err => {
                if (btn) btn.disabled = false;
                if (statusEl) {
                  statusEl.className = 'status-box error';
                  statusEl.innerText = 'Error: ' + err;
                }
              })
              .smartAutoFormatEntireDocument(getCodeOptions(), getTableOptions(), getTypographyOptions());
          }

          function runRemoveBookmarks() {
            const statusEl = document.getElementById('bookmarkStatus');
            const btn = document.getElementById('btnRemoveBookmarks');
            if (btn) btn.disabled = true;
            if (statusEl) statusEl.innerText = 'Removing...';
            google.script.run
              .withSuccessHandler(res => {
                if (btn) btn.disabled = false;
                if (statusEl) {
                  statusEl.innerText = res.count > 0
                    ? '✓ Removed ' + res.count + ' bookmark(s)'
                    : '✓ No bookmarks found';
                }
              })
              .withFailureHandler(err => {
                if (btn) btn.disabled = false;
                if (statusEl) statusEl.innerText = '✗ Error: ' + err;
              })
              .removeAllBookmarks();
          }

          function runFormatDocumentTypography() {
            setTypographyStatus('Formatting document typography...', 'loading');
            document.getElementById('btnFormatDocTypo').disabled = true;
            google.script.run
              .withSuccessHandler(res => {
                document.getElementById('btnFormatDocTypo').disabled = false;
                setTypographyStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                document.getElementById('btnFormatDocTypo').disabled = false;
                setTypographyStatus('Error: ' + err, 'error');
              })
              .formatDocumentTypography(getTypographyOptions());
          }

          function runFormatSelectedTypography() {
            setTypographyStatus('Auto-formatting selected text...', 'loading');
            const btn = document.getElementById('btnFormatSelectedTypo');
            if (btn) btn.disabled = true;
            google.script.run
              .withSuccessHandler(res => {
                if (btn) btn.disabled = false;
                setTypographyStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                if (btn) btn.disabled = false;
                setTypographyStatus('Error: ' + err, 'error');
              })
              .formatSelectedTypography(getTypographyOptions());
          }

          function runUndoDocumentTypography() {
            if (!confirm('Are you sure you want to reset all document headings and body text to standard defaults?')) {
              return;
            }
            setTypographyStatus('Resetting typography to defaults...', 'loading');
            document.getElementById('btnUndoDocTypo').disabled = true;
            google.script.run
              .withSuccessHandler(res => {
                document.getElementById('btnUndoDocTypo').disabled = false;
                setTypographyStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                document.getElementById('btnUndoDocTypo').disabled = false;
                setTypographyStatus('Error: ' + err, 'error');
              })
              .undoDocumentTypography();
          }

          function runFormatSelectedAs(targetType) {
            setTypographyStatus('Applying ' + targetType + ' formatting to selected text...', 'loading');
            google.script.run
              .withSuccessHandler(res => {
                setTypographyStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                setTypographyStatus('Error: ' + err, 'error');
              })
              .formatSelectedAs(targetType, getTypographyOptions());
          }

          /* =========================================================
             TABLE FORMATTER LOGIC
             ========================================================= */
          const TABLE_THEMES = {
            'corporate-navy': {
              headerBg: '#1E3A8A', headerText: '#FFFFFF',
              altRowBg: '#F0F7FF', normalRowBg: '#FFFFFF',
              borderColor: '#93C5FD', textColor: '#1E293B'
            },
            'emerald-mint': {
              headerBg: '#065F46', headerText: '#FFFFFF',
              altRowBg: '#ECFDF5', normalRowBg: '#FFFFFF',
              borderColor: '#A7F3D0', textColor: '#064E3B'
            },
            'royal-indigo': {
              headerBg: '#4C1D95', headerText: '#FFFFFF',
              altRowBg: '#F5F3FF', normalRowBg: '#FFFFFF',
              borderColor: '#DDD6FE', textColor: '#2E1065'
            },
            'sunset-crimson': {
              headerBg: '#881337', headerText: '#FFFFFF',
              altRowBg: '#FFF1F2', normalRowBg: '#FFFFFF',
              borderColor: '#FECDD3', textColor: '#4C0519'
            },
            'ocean-teal': {
              headerBg: '#0F766E', headerText: '#FFFFFF',
              altRowBg: '#F0FDFA', normalRowBg: '#FFFFFF',
              borderColor: '#99F6E4', textColor: '#134E4A'
            },
            'executive-slate': {
              headerBg: '#1E293B', headerText: '#FFFFFF',
              altRowBg: '#F8FAFC', normalRowBg: '#FFFFFF',
              borderColor: '#CBD5E1', textColor: '#0F172A'
            }
          };

          const initialTableTheme = '${tablePrefs.theme || 'corporate-navy'}';
          document.getElementById('tableThemeSelect').value = initialTableTheme;
          document.getElementById('tableFontFamilySelect').value = '${tablePrefs.fontFamily || 'Roboto'}';
          document.getElementById('tableHeaderFontSelect').value = '${tablePrefs.headerFontSize || 11}';
          document.getElementById('tableBodyFontSelect').value = '${tablePrefs.bodyFontSize || 9.5}';
          document.getElementById('tablePaddingSelect').value = '${tablePrefs.padding || 'normal'}';
          document.getElementById('tableBorderWidthSelect').value = '${tablePrefs.borderWidth !== undefined ? tablePrefs.borderWidth : 1}';

          function onTableThemeChange() {
            const val = document.getElementById('tableThemeSelect').value;
            if (val !== 'custom' && TABLE_THEMES[val]) {
              const t = TABLE_THEMES[val];
              document.getElementById('tHeaderBg').value = t.headerBg;
              document.getElementById('tHeaderText').value = t.headerText;
              document.getElementById('tAltRowBg').value = t.altRowBg;
              document.getElementById('tNormalRowBg').value = t.normalRowBg;
              document.getElementById('tBorderColor').value = t.borderColor;
            }
            updateTablePreview();
          }

          function setTableCustomMode() {
            document.getElementById('tableThemeSelect').value = 'custom';
            updateTablePreview();
          }

          function updateTablePreview() {
            const hBg = document.getElementById('tHeaderBg').value;
            const hText = document.getElementById('tHeaderText').value;
            const altBg = document.getElementById('tAltRowBg').value;
            const normBg = document.getElementById('tNormalRowBg').value;
            const borderCol = document.getElementById('tBorderColor').value;
            const tFont = document.getElementById('tableFontFamilySelect').value;
            const hFontSize = document.getElementById('tableHeaderFontSelect').value;
            const bFontSize = document.getElementById('tableBodyFontSelect').value;
            const padMode = document.getElementById('tablePaddingSelect').value;

            const previewTable = document.getElementById('previewTableEl');
            const headerRow = document.getElementById('prevHeaderRow');
            const row1 = document.getElementById('prevRow1');
            const row2 = document.getElementById('prevRow2');

            // Apply selected font family to table preview
            previewTable.style.fontFamily = tFont + ', -apple-system, sans-serif';

            // Header Row: Middle-aligned, ALWAYS BOLD & Larger Font
            headerRow.style.backgroundColor = hBg;
            headerRow.style.color = hText;
            headerRow.style.textAlign = 'center';
            headerRow.style.fontWeight = 'bold';
            headerRow.style.fontSize = (parseFloat(hFontSize) + 1) + 'px';

            // Data Rows: strictly LEFT-aligned & standard font size
            row1.style.backgroundColor = normBg;
            row1.style.color = '#1E293B';
            row1.style.textAlign = 'left';
            row1.style.fontSize = parseFloat(bFontSize) + 'px';

            row2.style.backgroundColor = altBg;
            row2.style.color = '#1E293B';
            row2.style.textAlign = 'left';
            row2.style.fontSize = parseFloat(bFontSize) + 'px';

            // Padding
            let padPx = '7px 9px';
            if (padMode === 'compact') padPx = '4px 6px';
            if (padMode === 'relaxed') padPx = '10px 12px';

            const borderWidthPx = document.getElementById('tableBorderWidthSelect').value;
            const allCells = document.querySelectorAll('#previewTableEl th, #previewTableEl td');
            allCells.forEach(cell => {
              cell.style.borderColor = borderCol;
              cell.style.padding = padPx;
              cell.style.borderWidth = (parseFloat(borderWidthPx) || 0) + 'px';
              cell.style.borderStyle = parseFloat(borderWidthPx) > 0 ? 'solid' : 'none';
              cell.style.verticalAlign = 'top';
            });

            // Live preview: inline code token highlight
            const inlineEnabled = document.getElementById('tInlineCodeToggle').checked;
            const inlineBg = document.getElementById('tInlineCodeBg').value;
            const inlineColor = document.getElementById('tInlineCodeColor').value;
            const codeToken = document.getElementById('prevCodeToken');
            if (codeToken) {
              codeToken.style.backgroundColor = inlineEnabled ? inlineBg : 'transparent';
              codeToken.style.color = inlineEnabled ? inlineColor : 'inherit';
              codeToken.style.fontFamily = inlineEnabled ? 'Consolas, monospace' : (tFont + ', sans-serif');
              codeToken.style.border = inlineEnabled ? ('1px solid ' + inlineBg) : 'none';
              codeToken.style.borderRadius = inlineEnabled ? '3px' : '0';
              codeToken.style.padding = inlineEnabled ? '1px 5px' : '0';
            }
          }

          function getTableOptions() {
            return {
              theme: document.getElementById('tableThemeSelect').value,
              headerBg: document.getElementById('tHeaderBg').value,
              headerText: document.getElementById('tHeaderText').value,
              altRowBg: document.getElementById('tAltRowBg').value,
              normalRowBg: document.getElementById('tNormalRowBg').value,
              borderColor: document.getElementById('tBorderColor').value,
              fontFamily: document.getElementById('tableFontFamilySelect').value,
              headerFontSize: document.getElementById('tableHeaderFontSelect').value,
              bodyFontSize: document.getElementById('tableBodyFontSelect').value,
              padding: document.getElementById('tablePaddingSelect').value,
              textColor: '#1E293B',
              borderWidth: parseFloat(document.getElementById('tableBorderWidthSelect').value) || 0,
              inlineCodeHighlight: document.getElementById('tInlineCodeToggle').checked,
              inlineCodeBg: document.getElementById('tInlineCodeBg').value,
              inlineCodeColor: document.getElementById('tInlineCodeColor').value
            };
          }

          function setTableStatus(msg, type) {
            const el = document.getElementById('tableStatus');
            el.className = 'status-box ' + type;
            el.innerText = msg;
          }

          function runFormatAllTables() {
            setTableStatus('Formatting and centering all tables...', 'loading');
            document.getElementById('btnFormatAllTables').disabled = true;
            google.script.run
              .withSuccessHandler(res => {
                document.getElementById('btnFormatAllTables').disabled = false;
                setTableStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                document.getElementById('btnFormatAllTables').disabled = false;
                setTableStatus('Error: ' + err, 'error');
              })
              .formatAllDocumentTables(getTableOptions());
          }

          function runFormatSelectedTable() {
            setTableStatus('Formatting selected table...', 'loading');
            document.getElementById('btnFormatSelectedTable').disabled = true;
            google.script.run
              .withSuccessHandler(res => {
                document.getElementById('btnFormatSelectedTable').disabled = false;
                setTableStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                document.getElementById('btnFormatSelectedTable').disabled = false;
                setTableStatus('Error: ' + err, 'error');
              })
              .formatSelectedTable(getTableOptions());
          }

          function runUndoAllTables() {
            if (!confirm('Are you sure you want to reset all tables to document defaults?')) {
              return;
            }
            setTableStatus('Resetting all tables to default...', 'loading');
            document.getElementById('btnUndoAllTables').disabled = true;
            google.script.run
              .withSuccessHandler(res => {
                document.getElementById('btnUndoAllTables').disabled = false;
                setTableStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                document.getElementById('btnUndoAllTables').disabled = false;
                setTableStatus('Error: ' + err, 'error');
              })
              .undoAllTableFormatting();
          }

          function runUndoSelectedTable() {
            setTableStatus('Resetting selected table...', 'loading');
            document.getElementById('btnUndoSelectedTable').disabled = true;
            google.script.run
              .withSuccessHandler(res => {
                document.getElementById('btnUndoSelectedTable').disabled = false;
                setTableStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                document.getElementById('btnUndoSelectedTable').disabled = false;
                setTableStatus('Error: ' + err, 'error');
              })
              .undoSelectedTableFormatting();
          }

          /* =========================================================
             CODE HIGHLIGHTER LOGIC
             ========================================================= */
          const CODE_THEMES = {
            'github-light': {
              bg: '#F6F8FA', text: '#24292F', border: '#D0D7DE',
              kw: '#CF222E', str: '#0A3069', com: '#6E7781', num: '#0550AE',
              fn: '#8250DF', type: '#953800'
            },
            'one-dark': {
              bg: '#21252B', text: '#ABB2BF', border: '#3B4048',
              kw: '#C678DD', str: '#98C379', com: '#5C6370', num: '#D19A66',
              fn: '#61AFEF', type: '#E5C07B'
            },
            'dracula': {
              bg: '#282A36', text: '#F8F8F2', border: '#44475A',
              kw: '#FF79C6', str: '#F1FA8C', com: '#6272A4', num: '#BD93F9',
              fn: '#50FA7B', type: '#8BE9FD'
            },
            'monokai': {
              bg: '#272822', text: '#F8F8F2', border: '#3E3D32',
              kw: '#F92672', str: '#E6DB74', com: '#75715E', num: '#AE81FF',
              fn: '#A6E22E', type: '#66D9EF'
            },
            'solarized-light': {
              bg: '#FDF6E3', text: '#657B83', border: '#EEE8D5',
              kw: '#859900', str: '#2AA198', com: '#93A1A1', num: '#D33682',
              fn: '#268BD2', type: '#B58900'
            }
          };

          const initialCodeTheme = '${codePrefs.theme || 'github-light'}';
          document.getElementById('codeThemeSelect').value = initialCodeTheme;
          document.getElementById('codeFontSizeSelect').value = '${codePrefs.fontSize || '9.5'}';
          document.getElementById('codeFontFamilySelect').value = '${codePrefs.fontFamily || 'Consolas'}';
          document.getElementById('codeIndentSelect').value = '${codePrefs.indentStyle || 'auto-2'}';

          function onCodeThemeChange() {
            const val = document.getElementById('codeThemeSelect').value;
            if (val !== 'custom' && CODE_THEMES[val]) {
              const t = CODE_THEMES[val];
              document.getElementById('cBgColor').value = t.bg;
              document.getElementById('cTextColor').value = t.text;
              document.getElementById('cBorderColor').value = t.border;
              document.getElementById('cKwColor').value = t.kw;
              document.getElementById('cStrColor').value = t.str;
              document.getElementById('cComColor').value = t.com;
              document.getElementById('cNumColor').value = t.num;
              if (document.getElementById('cFnColor')) document.getElementById('cFnColor').value = t.fn;
              if (document.getElementById('cTypeColor')) document.getElementById('cTypeColor').value = t.type;
            }
            updateCodePreview();
          }

          function setCodeCustomMode() {
            document.getElementById('codeThemeSelect').value = 'custom';
            updateCodePreview();
          }

          function updateCodePreview() {
            const bg = document.getElementById('cBgColor').value;
            const text = document.getElementById('cTextColor').value;
            const border = document.getElementById('cBorderColor').value;
            const font = document.getElementById('codeFontFamilySelect').value;

            const box = document.getElementById('codePreviewBox');
            box.style.backgroundColor = bg;
            box.style.color = text;
            box.style.borderColor = border;
            box.style.fontFamily = font + ', monospace';
            box.style.fontWeight = 'normal';
            const cFontSize = document.getElementById('codeFontSizeSelect').value;
            box.style.fontSize = parseFloat(cFontSize) + 'px';

            // Indentation preview
            const indentVal = document.getElementById('codeIndentSelect').value;
            const indentSpaces = (indentVal === 'auto-4' || indentVal === 'tab-4') ? '&nbsp;&nbsp;&nbsp;&nbsp;' : '&nbsp;&nbsp;';
            const pInd1 = document.getElementById('pIndent');
            const pInd2 = document.getElementById('pIndent2');
            if (pInd1) pInd1.innerHTML = indentSpaces;
            if (pInd2) pInd2.innerHTML = indentSpaces;

            // Use token color pickers directly (updated by theme OR manual pick)
            document.getElementById('pKw').style.color  = document.getElementById('cKwColor').value;
            document.getElementById('pKw2').style.color = document.getElementById('cKwColor').value;
            document.getElementById('pStr').style.color = document.getElementById('cStrColor').value;
            document.getElementById('pCom').style.color = document.getElementById('cComColor').value;
            if (document.getElementById('pFn')) {
              document.getElementById('pFn').style.color = document.getElementById('cFnColor') ? document.getElementById('cFnColor').value : '#8250DF';
            }
          }

          function getCodeOptions() {
            return {
              theme: document.getElementById('codeThemeSelect').value,
              fontSize: document.getElementById('codeFontSizeSelect').value,
              fontFamily: document.getElementById('codeFontFamilySelect').value,
              indentStyle: document.getElementById('codeIndentSelect').value,
              bgColor: document.getElementById('cBgColor').value,
              textColor: document.getElementById('cTextColor').value,
              borderColor: document.getElementById('cBorderColor').value,
              keywordColor: document.getElementById('cKwColor').value,
              stringColor: document.getElementById('cStrColor').value,
              commentColor: document.getElementById('cComColor').value,
              numberColor: document.getElementById('cNumColor').value,
              functionColor: document.getElementById('cFnColor') ? document.getElementById('cFnColor').value : undefined,
              typeColor: document.getElementById('cTypeColor') ? document.getElementById('cTypeColor').value : undefined
            };
          }

          function setCodeStatus(msg, type) {
            const el = document.getElementById('codeStatus');
            el.className = 'status-box ' + type;
            el.innerText = msg;
          }

          function runHighlightAllCode() {
            setCodeStatus('Processing code blocks...', 'loading');
            document.getElementById('btnHighlightAllCode').disabled = true;
            google.script.run
              .withSuccessHandler(res => {
                document.getElementById('btnHighlightAllCode').disabled = false;
                setCodeStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                document.getElementById('btnHighlightAllCode').disabled = false;
                setCodeStatus('Error: ' + err, 'error');
              })
              .highlightAllCodeBlocks(getCodeOptions());
          }

          function runHighlightSelectedCode() {
            setCodeStatus('Formatting selected text...', 'loading');
            document.getElementById('btnHighlightSelectedCode').disabled = true;
            google.script.run
              .withSuccessHandler(res => {
                document.getElementById('btnHighlightSelectedCode').disabled = false;
                setCodeStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                document.getElementById('btnHighlightSelectedCode').disabled = false;
                setCodeStatus('Error: ' + err, 'error');
              })
              .formatSelectedCodeBlock(getCodeOptions());
          }

          function runIndentAllCode() {
            setCodeStatus('Indenting all code blocks...', 'loading');
            document.getElementById('btnIndentAllCode').disabled = true;
            google.script.run
              .withSuccessHandler(res => {
                document.getElementById('btnIndentAllCode').disabled = false;
                setCodeStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                document.getElementById('btnIndentAllCode').disabled = false;
                setCodeStatus('Error: ' + err, 'error');
              })
              .indentAllCodeBlocks(getCodeOptions());
          }

          function runUndoAllCode() {
            if (!confirm('Are you sure you want to revert all code blocks to plain document text?')) {
              return;
            }
            setCodeStatus('Reverting all code blocks...', 'loading');
            document.getElementById('btnUndoAllCode').disabled = true;
            google.script.run
              .withSuccessHandler(res => {
                document.getElementById('btnUndoAllCode').disabled = false;
                setCodeStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                document.getElementById('btnUndoAllCode').disabled = false;
                setCodeStatus('Error: ' + err, 'error');
              })
              .undoAllCodeBlocksFormatting();
          }

          function runUndoSelectedCode() {
            setCodeStatus('Reverting selected code block...', 'loading');
            document.getElementById('btnUndoSelectedCode').disabled = true;
            google.script.run
              .withSuccessHandler(res => {
                document.getElementById('btnUndoSelectedCode').disabled = false;
                setCodeStatus(res.message, res.success ? 'success' : 'error');
              })
              .withFailureHandler(err => {
                document.getElementById('btnUndoSelectedCode').disabled = false;
                setCodeStatus('Error: ' + err, 'error');
              })
              .undoSelectedCodeBlockFormatting();
          }

          // Initialize previews on load
          updateTypographyPreview();
          onTableThemeChange();
          onCodeThemeChange();
        </script>
      </body>
    </html>
  `;
}
