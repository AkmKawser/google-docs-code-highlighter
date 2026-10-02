/**
 * Smart Unified Code Highlighter with Custom Sidebar & Themes for Google Docs
 *
 * @license MIT
 * @repository https://github.com/
 */

function onOpen() {
  DocumentApp.getUi()
    .createMenu('⚡ Code Highlighter')
    .addItem('Open Sidebar (Styles & Colors)', 'showSidebar')
    .addItem('Highlight All Code (Quick Run)', 'quickHighlightAll')
    .addItem('Highlight Selected Code', 'formatSelectedCodeBlock')
    .addToUi();
}

/**
 * Opens the styling sidebar in Google Docs
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutput(getSidebarHtml())
    .setTitle('⚡ Code Highlighter')
    .setWidth(320);
  DocumentApp.getUi().showSidebar(html);
}

function quickHighlightAll() {
  const prefs = getUserPreferences();
  highlightAllCodeBlocks(prefs);
}

/**
 * Core scanning & formatting function
 * Uses curly brace tracking ({ and }) to prevent code splitting.
 */
function highlightAllCodeBlocks(options) {
  options = options || getUserPreferences();
  saveUserPreferences(options);

  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const paragraphs = body.getParagraphs();
  
  let codeGroups = [];
  let inCode = false;
  let braceDepth = 0;
  let currentGroup = [];
  let blankBuffer = [];

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    
    // Skip items already inside a code block
    if (p.getParent().getType() === DocumentApp.ElementType.TABLE_CELL) {
      if (inCode) finishGroup();
      continue;
    }

    const text = p.getText();
    const trimmed = text.trim();

    if (!inCode) {
      if (isCodeStart(text)) {
        inCode = true;
        currentGroup.push(p);
        blankBuffer = [];
        braceDepth = Math.max(0, getNetBraceCount(text));
      }
    } else {
      // 1. LOCKED BY CURLY BRACES
      if (braceDepth > 0) {
        if (blankBuffer.length > 0) {
          currentGroup = currentGroup.concat(blankBuffer);
          blankBuffer = [];
        }
        currentGroup.push(p);
        braceDepth = Math.max(0, braceDepth + getNetBraceCount(text));
        continue;
      }

      // 2. OUTSIDE BRACES (braceDepth === 0)
      if (trimmed.length === 0) {
        blankBuffer.push(p);
        if (blankBuffer.length > 2) finishGroup();
      } else if (isStrongProse(text)) {
        finishGroup();
      } else if (isCodeLine(text)) {
        if (blankBuffer.length > 0) {
          currentGroup = currentGroup.concat(blankBuffer);
          blankBuffer = [];
        }
        currentGroup.push(p);
        braceDepth = Math.max(0, braceDepth + getNetBraceCount(text));
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
    braceDepth = 0;
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

/**
 * Calculates net curly braces ignoring comments/strings
 */
function getNetBraceCount(text) {
  if (!text) return 0;
  let clean = text.replace(/\/\*[\s\S]*?\*\/|\/\/.*$|#.*$/gm, '');
  clean = clean.replace(/(["'`])(?:(?=(\\?))\2[\s\S])*?\1/g, '');
  const openCount = (clean.match(/\{/g) || []).length;
  const closeCount = (clean.match(/\}/g) || []).length;
  return openCount - closeCount;
}

function isCodeStart(text) {
  const trimmed = text.trim();
  if (!trimmed || isStrongProse(trimmed)) return false;

  const startPatterns = [
    /^(def|class|function|const|let|var|import|export|public|private|protected|static|package|namespace)\b/,
    /^(if|for|while|switch|catch|with|elif)\s*(\(|:)/,
    /^(try|except|finally|else):?$/,
    /^(SELECT|INSERT|UPDATE|DELETE|CREATE|ALTER|DROP|FROM|WHERE)\b/i,
    /^\s*(\/\/|#|\/\*|\*\/|<!--)/,
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
  if (isStrongProse(trimmed)) return false;
  if (/^\s*[}\])];?$/.test(trimmed)) return true;
  if (/^\s*\.[A-Za-z0-9_$]+/.test(trimmed)) return true;
  if (/^\s{2,}|\t/.test(text)) return true;
  if (/^(return|break|continue|pass|throw|export|default)\b/.test(trimmed)) return true;
  if (/[{};]$/.test(trimmed)) return true;
  return isCodeStart(text);
}

function isStrongProse(text) {
  const trimmed = text.trim();
  if (!trimmed) return false;
  if (/^(def|class|function|const|let|var|import|export|return|public|private|static|if|for|while)\b/.test(trimmed)) return false;
  if (/[{};]$/.test(trimmed) || /(=>|===|!==)/.test(trimmed)) return false;
  if (/^\s{2,}|\t/.test(text)) return false;

  const words = trimmed.toLowerCase().split(/\s+/);
  const stopWords = ['the', 'this', 'that', 'with', 'and', 'are', 'can', 'you', 'we', 'our', 'will', 'should', 'about', 'there', 'from', 'which', 'because', 'also', 'here', 'when', 'then', 'into', 'have', 'been', 'would', 'could'];
  const proseCount = words.filter(w => stopWords.includes(w.replace(/[^a-z]/g, ''))).length;

  if (proseCount >= 2 && /[.!?:]$/.test(trimmed)) return true;
  if (proseCount >= 4) return true;
  return false;
}

/**
 * Creates and styles the code table container
 */
function convertParagraphsToCodeBlock(body, paragraphGroup, options) {
  const firstParagraph = paragraphGroup[0];
  const insertIndex = body.getChildIndex(firstParagraph);

  const table = body.insertTable(insertIndex);
  table.setBorderWidth(1);
  table.setBorderColor(options.borderColor || '#D0D7DE');
  
  const cell = table.appendTableRow().appendTableCell();
  cell.setBackgroundColor(options.bgColor || '#F6F8FA');
  cell.setPaddingTop(8);
  cell.setPaddingBottom(8);
  cell.setPaddingLeft(12);
  cell.setPaddingRight(12);
  cell.setText('');

  const fontSize = Number(options.fontSize) || 9.5;
  const fontFamily = options.fontFamily || 'Consolas';
  const textColor = options.textColor || '#24292F';

  paragraphGroup.forEach(p => {
    const line = cell.appendParagraph(p.getText());
    line.setFontFamily(fontFamily);
    line.setFontSize(fontSize);
    line.setLineSpacing(1.15);
    line.setForegroundColor(textColor);
    
    applySyntaxHighlight(line.editAsText(), options);
    p.removeFromParent();
  });
}

/**
 * Manual selection formatter
 */
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
 * Apply syntax colors based on the chosen theme
 */
function applySyntaxHighlight(textObj, options) {
  const text = textObj.getText();
  const kwColor = options.keywordColor || '#CF222E';
  const strColor = options.stringColor || '#0A3069';
  const comColor = options.commentColor || '#6E7781';
  const numColor = options.numberColor || '#953800';

  const rules = [
    { regex: /(\/\/.*$|#.*$)/gm, color: comColor },
    { regex: /(["'`])(?:(?=(\\?))\2[\s\S])*?\1/g, color: strColor },
    { regex: /\b(def|class|function|const|let|var|return|if|else|elif|for|while|import|from|export|public|private|protected|async|await|try|catch|finally|new|this|print|self|lambda)\b/g, color: kwColor },
    { regex: /\b(true|false|null|undefined|None|True|False)\b/g, color: '#D97706' },
    { regex: /\b\d+(\.\d+)?\b/g, color: numColor }
  ];

  rules.forEach(rule => {
    let match;
    while ((match = rule.regex.exec(text)) !== null) {
      textObj.setForegroundColor(match.index, match.index + match[0].length - 1, rule.color);
    }
  });
}

/**
 * Preferences Storage
 */
function saveUserPreferences(prefs) {
  PropertiesService.getUserProperties().setProperty('code_highlighter_prefs', JSON.stringify(prefs));
}

function getUserPreferences() {
  const saved = PropertiesService.getUserProperties().getProperty('code_highlighter_prefs');
  if (saved) {
    try { return JSON.parse(saved); } catch(e) {}
  }
  return {
    theme: 'github-light',
    fontSize: '9.5',
    fontFamily: 'Consolas',
    bgColor: '#F6F8FA',
    textColor: '#24292F',
    borderColor: '#D0D7DE',
    keywordColor: '#0550AE',
    stringColor: '#0A3069',
    commentColor: '#6E7781',
    numberColor: '#953800'
  };
}

/**
 * Returns the HTML for the sidebar
 */
function getSidebarHtml() {
  const current = getUserPreferences();
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; font-size: 13px; color: #333; padding: 14px; margin: 0; }
          h3 { margin-top: 0; color: #1a73e8; font-size: 16px; display: flex; align-items: center; gap: 6px; }
          .control-group { margin-bottom: 12px; }
          label { display: block; font-weight: 600; margin-bottom: 5px; font-size: 12px; color: #444; }
          select, input[type="text"] { width: 100%; padding: 7px; border: 1px solid #ccc; border-radius: 4px; box-sizing: border-box; font-size: 12px; }
          .color-row { display: flex; align-items: center; justify-content: space-between; margin-bottom: 8px; }
          .color-row label { margin-bottom: 0; font-weight: normal; }
          .color-row input[type="color"] { width: 34px; height: 26px; border: 1px solid #ccc; border-radius: 4px; cursor: pointer; padding: 0; background: none; }
          
          /* Live Preview Box */
          .preview-box { border-radius: 4px; padding: 10px; margin: 14px 0; font-family: 'Consolas', monospace; font-size: 10px; line-height: 1.3; overflow: hidden; border: 1px solid #ddd; }
          
          /* Buttons */
          .btn-primary { width: 100%; background: #1a73e8; color: white; border: none; padding: 10px; border-radius: 4px; font-weight: 600; cursor: pointer; font-size: 13px; margin-top: 6px; }
          .btn-primary:hover { background: #1557b0; }
          .btn-secondary { width: 100%; background: white; color: #1a73e8; border: 1px solid #1a73e8; padding: 8px; border-radius: 4px; font-weight: 600; cursor: pointer; font-size: 12px; margin-top: 8px; }
          .btn-secondary:hover { background: #f1f8ff; }
          .status { font-size: 12px; margin-top: 10px; text-align: center; color: #059669; font-weight: 500; min-height: 16px; }
        </style>
      </head>
      <body>
        <h3>⚡ Code Highlighter</h3>
        
        <div class="control-group">
          <label>Theme Preset</label>
          <select id="themeSelect" onchange="onThemeChange()">
            <option value="github-light">GitHub Light</option>
            <option value="dracula">Dracula Dark</option>
            <option value="monokai">Monokai Dark</option>
            <option value="solarized-light">Solarized Light</option>
            <option value="custom">Custom Colors...</option>
          </select>
        </div>

        <div class="control-group">
          <label>Font Size</label>
          <select id="fontSizeSelect" onchange="updatePreview()">
            <option value="8.5">8.5 pt (Compact)</option>
            <option value="9.5" selected>9.5 pt (Default)</option>
            <option value="10">10 pt</option>
            <option value="11">11 pt (Large)</option>
            <option value="12">12 pt</option>
          </select>
        </div>

        <div class="control-group">
          <label>Font Family</label>
          <select id="fontFamilySelect" onchange="updatePreview()">
            <option value="Consolas" selected>Consolas</option>
            <option value="Courier New">Courier New</option>
            <option value="Roboto Mono">Roboto Mono</option>
          </select>
        </div>

        <div id="customColorsSection">
          <label style="margin-bottom: 8px;">Colors</label>
          <div class="color-row">
            <label>Background:</label>
            <input type="color" id="bgColor" value="${current.bgColor}" onchange="setCustomMode()">
          </div>
          <div class="color-row">
            <label>Text Color:</label>
            <input type="color" id="textColor" value="${current.textColor}" onchange="setCustomMode()">
          </div>
          <div class="color-row">
            <label>Border:</label>
            <input type="color" id="borderColor" value="${current.borderColor}" onchange="setCustomMode()">
          </div>
        </div>

        <label style="margin-top: 10px;">Live Preview:</label>
        <div id="previewBox" class="preview-box">
          <span id="pKw" style="color: #0550AE; font-weight: bold;">function</span> <span id="pFn">helloWorld</span>() {<br>
          &nbsp;&nbsp;<span id="pCom" style="color: #6E7781;">// greeting</span><br>
          &nbsp;&nbsp;console.log(<span id="pStr" style="color: #0A3069;">"Welcome!"</span>);<br>
          }
        </div>

        <button class="btn-primary" onclick="runHighlightAll()">⚡ Highlight All Code</button>
        <button class="btn-secondary" onclick="runHighlightSelected()">Format Selected Text</button>
        <div id="status" class="status"></div>

        <script>
          const THEMES = {
            'github-light': { bg: '#F6F8FA', text: '#24292F', border: '#D0D7DE', kw: '#CF222E', str: '#0A3069', com: '#6E7781', num: '#953800' },
            'dracula':      { bg: '#282A36', text: '#F8F8F2', border: '#44475A', kw: '#FF79C6', str: '#F1FA8C', com: '#6272A4', num: '#BD93F9' },
            'monokai':      { bg: '#272822', text: '#F8F8F2', border: '#3E3D32', kw: '#F92672', str: '#E6DB74', com: '#75715E', num: '#AE81FF' },
            'solarized-light': { bg: '#FDF6E3', text: '#657B83', border: '#EEE8D5', kw: '#268BD2', str: '#2AA198', com: '#93A1A1', num: '#D33682' }
          };

          let currentTheme = '${current.theme || 'github-light'}';
          document.getElementById('themeSelect').value = currentTheme;
          document.getElementById('fontSizeSelect').value = '${current.fontSize || '9.5'}';
          document.getElementById('fontFamilySelect').value = '${current.fontFamily || 'Consolas'}';

          function onThemeChange() {
            const val = document.getElementById('themeSelect').value;
            if (val !== 'custom') {
              const t = THEMES[val];
              document.getElementById('bgColor').value = t.bg;
              document.getElementById('textColor').value = t.text;
              document.getElementById('borderColor').value = t.border;
            }
            updatePreview();
          }

          function setCustomMode() {
            document.getElementById('themeSelect').value = 'custom';
            updatePreview();
          }

          function updatePreview() {
            const bg = document.getElementById('bgColor').value;
            const text = document.getElementById('textColor').value;
            const border = document.getElementById('borderColor').value;
            const font = document.getElementById('fontFamilySelect').value;
            const theme = document.getElementById('themeSelect').value;

            const box = document.getElementById('previewBox');
            box.style.backgroundColor = bg;
            box.style.color = text;
            box.style.borderColor = border;
            box.style.fontFamily = font;

            const t = THEMES[theme] || THEMES['github-light'];
            document.getElementById('pKw').style.color = t.kw;
            document.getElementById('pStr').style.color = t.str;
            document.getElementById('pCom').style.color = t.com;
          }

          function getOptions() {
            const theme = document.getElementById('themeSelect').value;
            const t = THEMES[theme] || {};
            return {
              theme: theme,
              fontSize: document.getElementById('fontSizeSelect').value,
              fontFamily: document.getElementById('fontFamilySelect').value,
              bgColor: document.getElementById('bgColor').value,
              textColor: document.getElementById('textColor').value,
              borderColor: document.getElementById('borderColor').value,
              keywordColor: t.kw || '#2563EB',
              stringColor: t.str || '#059669',
              commentColor: t.com || '#64748B',
              numberColor: t.num || '#9333EA'
            };
          }

          function runHighlightAll() {
            showStatus('Processing document...');
            google.script.run
              .withSuccessHandler(res => showStatus(res.message))
              .withFailureHandler(err => showStatus('Error: ' + err))
              .highlightAllCodeBlocks(getOptions());
          }

          function runHighlightSelected() {
            showStatus('Formatting selection...');
            google.script.run
              .withSuccessHandler(res => showStatus(res.message))
              .withFailureHandler(err => showStatus('Error: ' + err))
              .formatSelectedCodeBlock(getOptions());
          }

          function showStatus(msg) {
            document.getElementById('status').innerText = msg;
          }

          onThemeChange();
        </script>
      </body>
    </html>
  `;
}
