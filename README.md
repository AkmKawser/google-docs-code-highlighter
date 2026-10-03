# ⚡ Google Docs Code, Table & Typography Suite

> A smart Google Apps Script suite that formats document typography (Title, Heading 1, Sub-headings, Body), highlights code blocks with rich token syntax and professional spacing normalization, and styles data tables with professional themes & zebra striping — featuring an interactive 3-tab sidebar and zero markdown backticks required.

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![Google Docs](https://img.shields.io/badge/Google%20Docs-0F9D58?style=for-the-badge&logo=googledocs&logoColor=white)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)

---

## ⚡ Google Quotas, Rate Limits & Troubleshooting ("Service Documents failed")

When working with heavy automation across large Google Docs, you may encounter Google's internal service limits or the generic error:
> `Service Documents failed while accessing document with id <YOUR_DOCUMENT_ID>`

Understanding how Google Docs enforces quotas will help you diagnose and bypass these limits.

### 1. 📊 Official Google Apps Script Platform Quotas

Google enforces user-level and account-level quotas reset every 24 hours:

| Quota / Constraint | Personal Accounts (`@gmail.com`) | Google Workspace Accounts |
|---|---|---|
| **Max Execution Time** | **6 minutes / execution** | **6 minutes / execution** |
| **API Request Rate** | ~100 requests / 100 seconds / user | ~100 requests / 100 seconds / user |
| **Daily Trigger Runtime** | 90 minutes / day | 6 hours / day |
| **Document Size Limit** | 1.02 million characters | 1.02 million characters |
| **Max Simultaneous Executions** | 30 concurrent scripts | 30 concurrent scripts |

---

### 2. 🛡️ Hidden Per-Document Engine Limits

Even if your daily account quota is healthy, Google Docs enforces internal real-time engine constraints on **individual documents**:

1. **Real-Time Collaboration & Burst Rate Limits**:
   - Google Docs runs on a real-time collaborative Operational Transformation (OT) engine designed for human keystrokes.
   - When an Apps Script suite performs hundreds of structural DOM mutations in a few seconds (e.g. converting 10+ code groups into tables, styling multi-row tables, and re-styling dozens of paragraphs), it floods the synchronization buffer for that specific document.
   - When Google's server-side document engine is temporarily throttled or unable to reconcile changes fast enough, it rejects the batch transaction with `Service Documents failed while accessing document with id...`.

2. **Dirty Revision Journal & History Bloat**:
   - Every mutation and script run is recorded in the document's internal revision tree.
   - If a script crashes or encounters an issue mid-mutation, Google's server retains uncommitted or fragmented revision journal states for that specific file ID. Subsequent heavy batch writes must reconcile against all prior changes, frequently timing out or failing on that specific document.

3. **Inline Element Slicing Conflicts**:
   - Paragraphs containing embedded drawings, equations, horizontal rules, or bookmarks cannot accept arbitrary character range styling (e.g., `(0, textLen - 1)`). Attempting to style across these boundaries causes Google Docs' server validator to abort the transaction.

---

### 3. 🔍 Why Does It Work on One Document but Fail on Another?

* **Clean Documents**: A newly created document or fresh copy has an empty revision journal and full burst quota headroom. The suite executes seamlessly from start to finish.
* **Problem Documents**: Documents that have undergone dozens of rapid script executions, extensive edits, or crashed runs have accumulated a heavy revision backlog on Google's backend, making them vulnerable to transient service errors.

---

### 4. 🚀 Proven Solutions & Workarounds

If you ever encounter `Service Documents failed`:

#### Option A: Clean Copy (30-Second Permanent Reset)
1. In Google Docs, click **File** > **Make a copy**.
2. Open the newly created copy.
3. Run **`🚀 Smart Auto-Format Entire Document`**.
> **Why it works**: Making a copy generates a brand-new file ID with a 100% clean revision journal and full burst quota, while preserving all of your text and tables.

#### Option B: Run Phases Individually (Prevents Burst Throttling)
Instead of running the combined 1-click formatter, execute each module with a 3-second gap:
1. Click **`⚡ Highlight All Code (Quick Run)`** — wait 3 seconds.
2. Click **`📊 Format & Center All Tables`** — wait 3 seconds.
3. Click **`✍️ Format All Document Typography`**.
> **Why it works**: Breaking the process into 3 discrete transactions keeps your write volume well within Google's real-time rate limit.

#### Option C: Built-in Diagnostic Tool
Use **`⚡ Code, Table & Typography Suite`** > **`🔍 Diagnose Smart Format (Debug)`**:
- Runs isolated health checks for Document Access, Paragraphs, Tables, Bookmarks, and runs live tests of each formatting module (Code, Tables, Typography).
- Reports exact pass/fail status for every step so you can identify if a specific table or paragraph is causing an issue.

---

## ✨ Features

### ⚡ Smart Code Highlighter & Professional Spacing Normalizer
- **🎨 Rich Multi-Token Syntax Highlighting**:
  - **Functions & Methods**: Colors function calls (e.g. `calculateTotal()`, `insert()`, `console.log()`) and pseudocode method declarations.
  - **Types & Classes**: Highlights primitive types (`int`, `float`, `string`, `bool`, `void`) and PascalCase class/interface names (`Invoice`, `Customer`, `DocumentApp`).
  - **Keywords & Declarations**: Control flow (`if`, `else`, `return`, `try`, `catch`), storage declarations (`class`, `function`, `const`, `let`, `var`), and SQL statements (`SELECT`, `INSERT`, `UPDATE`, `DELETE`).
  - **Special Identifiers**: Dedicated styling for `this`, `self`, and `super`.
  - **Booleans & Constants**: Highlights `true`, `false`, `null`, `undefined`, `None`, and `ALL_CAPS` constants.
  - **Strings & Comments**: Multi-line strings, template literals, single-line comments (`//`, `#`), and block comments (`/* ... */`).
  - **Operators & Punctuation**: Colorizes `=>`, `===`, `!==`, `&&`, `||`, `+`, `-`, `*`, `/`, etc.
- **🛡️ Non-Overlapping Token Masking**:
  - Comments and strings claim character spans first, guaranteeing keywords and operators never incorrectly colorize text inside comments (e.g. `// class note`) or inside string literals.
- **📐 Professional Code Spacing Normalization**:
  - **Collapses Consecutive Empty Lines**: Normalizes runaway blank lines inside code blocks down to at most one clean empty line.
  - **No Awkward Gaps**: Automatically strips blank lines immediately following class/function declarations (e.g., directly under `class Invoice:`) and before closing brackets.
  - **Clean Pseudocode Continuations**: Tightens spacing around continuation markers (`...` or `…`) so they don't produce multi-line voids.
  - **Trims Outer Padding**: Automatically removes empty lines from the top and bottom of every code block.
- **📏 Zero Paragraph Margins inside Code Containers**:
  - Explicitly sets `spacingBefore = 0` and `spacingAfter = 0` with `1.15` line spacing on all paragraphs within the code block table, eliminating Google Docs' default paragraph margins for a crisp, authentic IDE editor feel.
- **🧠 Advanced Code & Pseudocode Detection**:
  - **Class & Pseudocode Context**: Tracks class blocks to keep bare field names (`items`, `customer`), method signatures, and ellipsis markers together as a single unified code block without splitting.
  - **Prose Rejection**: Ensures conversational sentences and narrative body paragraphs (e.g. "Issue: Invoice has four reasons to change...") always break out into normal text.
  - **Curly Brace Nesting Engine**: Uses net `{` and `}` balance counting to guarantee functions, loops, and nested classes never get split into fragmented chunks.
  - **🚫 No Markdown Required**: Triple backticks (```) are optional. Simply paste raw code or pseudocode anywhere in the document.
- **📐 Smart Indentation Engine**:
  - **Smart Auto-Indent (2 or 4 spaces)**: Syntax-aware indentation that calculates bracket nesting and colon-based pseudocode/Python headers (`calculateTotal():`), indenting method bodies while keeping sibling methods aligned.
  - **Tab Normalization**: Converts hard tabs to 2-space or 4-space indentations.
- **🎨 5 Built-in Code Themes + Custom**:
  - **GitHub Light**: Clean light theme with royal purple functions, warm amber types, crimson keywords, and navy strings.
  - **One Dark Pro (VS Code)**: Developer favorite with purple keywords, sky blue functions, gold types, and sage green strings.
  - **Dracula Dark**: High-contrast dark theme with pink keywords, neon green functions, cyan types, and purple numbers.
  - **Monokai Dark**: Iconic dark theme with hot pink keywords, lime green functions, and electric cyan types.
  - **Solarized Light**: Classic warm theme with blue functions, gold types, and olive keywords.
  - **Custom Palette**: Full RGB pickers for Background, Text, Border, Keywords, Functions, Types, Strings, Comments, and Numbers.
- **↩ 1-Click Code Blocks Formatting Undo**: Safely unwraps 1x1 code block tables back into standard document paragraphs, preserving your original code text.

---

### 📊 Professional Table Formatter & Middle-of-Page Alignment
- **📐 Automatic Middle Alignment**: Calculates page width and margins (`pageWidth - marginLeft - marginRight`) to center-align all tables symmetrically across the printable page with proportional column sizing.
- **🎨 Colorful Professional Themes**:
  - **Corporate Navy**: Deep royal navy headers with crisp ice-blue zebra striping.
  - **Emerald Mint**: Modern forest green headers with fresh mint alternating rows.
  - **Royal Indigo**: Deep violet headers with subtle lilac accents.
  - **Sunset Crimson**: Rich ruby red headers with warm rose striping.
  - **Ocean Teal**: Deep cyan/teal headers with aqua accents.
  - **Executive Slate**: Sleek charcoal headers with cool gray striping.
  - **Custom Palette**: Full RGB control over Header Background, Header Text, Alternating Row, Base Row, and Border colors.
- **🦓 Smart Alternating Rows (Zebra Striping)**: Enhanced visual rhythm for scanning rows and dense data grids.
- **🎯 Intelligent Cell Typography & Alignment**:
  - **Font Families**: Roboto, Arial, Inter, Open Sans, Lato, Montserrat, Calibri, Trebuchet MS, Georgia, Merriweather, Times New Roman, Consolas, or JetBrains Mono.
  - **Header Row**: Middle-aligned (centered) with larger bold typography (8.5–20pt) &mdash; table header row is always bold.
  - **Table Text (Data Rows)**: Strictly left-aligned with top vertical alignment (always on top, never middle-aligned) and clean body typography (8–20pt).
  - Refined cell padding (Compact, Normal, Relaxed) and soft borders.
- **🛡️ Code Block Safety**: Automatically differentiates between regular data tables and code containers, keeping your syntax highlighting intact while centering both!
- **↩ 1-Click Table Formatting Undo**: Reset all tables (or the selected table) back to clean standard document defaults with one click.

---

### ✍️ Document Typography & Heading Formatter
- **📖 Document Title Formatting (`TITLE`)**:
  - Independent typography controls: Font family, size (18–36pt), bold, text alignment (Center, Left, Right).
  - **Full Color Picking**: Individual pickers for text color and background color.
  - **Background Style Switch**: Switch between **Full-Width Header Banner** (seamless edge-to-edge block) or **Inline Text Highlight**.
- **📌 Main Headings Formatting (`HEADING_1`)**:
  - Font family, font size (14–24pt), bold, alignment.
  - Heading text color picker and background color picker.
  - Full-width colored banner table or inline highlight with toggle.
- **📑 Sub-Headings Formatting (`HEADING_2`, `HEADING_3`, `SUBTITLE`)**:
  - Font family, font size (10–20pt), bold, alignment.
  - Sub-heading text color picker and background color picker.
  - Full-width banner or inline highlight with toggle.
- **📝 General Body Text Formatting (`NORMAL`)**:
  - Pure font-changing controls without clutter: Font family, font size (8–20pt), text color, and alignment (Left, Justify, Center, Right).
  - **"Apply to General Body Text" Toggle**: Allows formatting headings only, or both headings and body text together.
- **🎨 Interactive Typography Presets**:
  - **Executive Navy**: Montserrat headers & Roboto body with royal navy banners.
  - **Modern Tech**: Inter clean typography with indigo banners.
  - **Emerald Forest**: Montserrat headers with fresh forest green banners.
  - **Editorial Classic**: Georgia serif styling with refined slate highlights.
  - **Crimson Luxe**: Montserrat headers with rich ruby banners.
- **👁️ Live Interactive Typography Preview**: Shows real-time updates for Title, Heading 1, Sub-headings, and body paragraphs as you adjust fonts and colors.
- **↩ 1-Click Document Typography Undo**: Reverts all headings, titles, and body paragraphs back to Google Docs defaults and unrolls header banners cleanly back into normal paragraphs.

---

### 🧹 Document Utilities & Cleaners
- **🧹 Text Cleanup**: Normalizes multiple spaces into single spaces, trims trailing line whitespace, and converts straight quotes to curly quotes, double hyphens to em dashes (`—`), and `...` to `…`.
- **📐 Paragraph Spacing Normalizer**: Standardizes heading spacing before/after, body spacing after, and sets `keepWithNext` on headings to prevent orphan headings at page breaks.
- **🔖 Bookmark Cleaner**: 1-click removal of all blue named anchor bookmark flags from the document.
- **✦ Inline Markdown Formatter**: Converts markdown markers (`**bold**`, `*italic*`, `~~strikethrough~~`, `==highlight==`) in body text into native styled Google Docs text.

---

## 🚀 Quick Setup Guide

### Method 1: Manual Copy & Paste (Recommended)

1. Open your document in **Google Docs**.
2. In the top menu, go to **Extensions** > **Apps Script**.
3. Clear out everything in `Code.gs`.
4. Copy the entire content of [`Code.js`](./Code.js) and paste it into `Code.gs`.
5. Press **`Ctrl + S`** (or click the disk icon) to save.
6. Return to your Google Docs tab and **refresh the page** (`Ctrl + R` or `F5`).
7. You will now see a new menu: **⚡ Code, Table & Typography Suite** in the toolbar!

### Method 2: Google Clasp CLI

If you use [`@google/clasp`](https://github.com/google/clasp):
```bash
npm install -g @google/clasp
clasp login
clasp clone <YOUR_SCRIPT_ID>
clasp push
```

---

## 📖 How to Use

### 1. 🚀 1-Click Smart Auto-Format Entire Document
- Click **⚡ Code, Table & Typography Suite** > **🚀 Smart Auto-Format Entire Document** (or click the top banner button in the sidebar).
- Intelligently scans the entire document with safe, isolated error boundaries:
  - **Bookmarks**: Cleans up all document bookmark anchors.
  - **Code Blocks**: Formats and indents code blocks with professional spacing and rich syntax colors.
  - **Data Tables**: Centers data tables across margins with colorful headers and zebra striping.
  - **Typography**: Auto-detects and styles Document Title, Headings, Sub-headings, Body text, and inline code highlights.

### 2. Open the Interactive 3-Tab Sidebar
- Click **⚡ Code, Table & Typography Suite** > **Open Sidebar (Styles & Colors)**.
- Switch between **✍️ Typography**, **📊 Tables**, and **⚡ Code** tabs:
  - **Typography Tab**: Select a preset or customize title/heading/subheading colors, fonts, sizes, and banner styles with a live document preview.
  - **Tables Tab**: Choose your table color theme, cell alignment, and padding with a real-time live preview.
  - **Code Tab**: Pick syntax themes (GitHub Light, One Dark Pro, Dracula, Monokai, Solarized Light), configure token colors for Keywords, Functions, Types, Strings, Comments, Numbers, and adjust auto-indentation.

### 3. Format & Indent Code Blocks
- Click **⚡ Code, Table & Typography Suite** > **⚡ Highlight All Code (Quick Run)**, or select code and click **Highlight Selected Code**.
- Code blocks are automatically formatted with zero paragraph gaps, standardized spacing, and syntax highlighted.

### 4. Format & Center All Tables
- Click **⚡ Code, Table & Typography Suite** > **📊 Format & Center All Tables**.
- All data tables are aligned in the middle of the page and styled with colorful headers, zebra striping, and clean borders.

### 5. Undo Formatting Anytime
- Click **↩ Undo Document Text Formatting** to revert titles, headings, and body paragraphs back to clean Google Docs defaults.
- Click **↩ Undo All Tables Formatting** to reset tables to standard document defaults.
- Click **↩ Undo All Code Blocks Formatting** to convert 1x1 code block containers back to normal paragraphs.

---

## 🎨 Supported Code Themes

| Theme | Background | Keywords | Functions | Types / Classes | Strings | Comments |
|---|---|---|---|---|---|---|
| **GitHub Light** | `#F6F8FA` | `#CF222E` (Red) | `#8250DF` (Purple) | `#953800` (Amber) | `#0A3069` (Navy) | `#6E7781` (Slate) |
| **One Dark Pro** | `#21252B` | `#C678DD` (Purple) | `#61AFEF` (Sky Blue) | `#E5C07B` (Gold) | `#98C379` (Green) | `#5C6370` (Gray) |
| **Dracula Dark** | `#282A36` | `#FF79C6` (Pink) | `#50FA7B` (Neon Green) | `#8BE9FD` (Cyan) | `#F1FA8C` (Yellow) | `#6272A4` (Lavender) |
| **Monokai Dark** | `#272822` | `#F92672` (Hot Pink) | `#A6E22E` (Lime Green) | `#66D9EF` (Cyan) | `#E6DB74` (Yellow) | `#75715E` (Warm Gray) |
| **Solarized Light** | `#FDF6E3` | `#859900` (Olive) | `#268BD2` (Blue) | `#B58900` (Gold) | `#2AA198` (Teal) | `#93A1A1` (Gray) |
| **Custom Palette** | Custom RGB | Custom RGB | Custom RGB | Custom RGB | Custom RGB | Custom RGB |

---

## 🎨 Supported Table Themes

| Theme | Header Color | Alternating Row | Best For |
|---|---|---|---|
| **Corporate Navy** | Royal Navy (`#1E3A8A`) | Ice Blue (`#F0F7FF`) | Executive summaries, business plans, reports |
| **Emerald Mint** | Forest Green (`#065F46`) | Mint Tint (`#ECFDF5`) | Financial reports, spreadsheets, environment notes |
| **Royal Indigo** | Deep Violet (`#4C1D95`) | Lavender (`#F5F3FF`) | Tech roadmaps, research papers, modern docs |
| **Sunset Crimson** | Ruby Red (`#881337`) | Soft Rose (`#FFF1F2`) | Marketing plans, audits, status reports |
| **Ocean Teal** | Cyan Teal (`#0F766E`) | Aqua Tint (`#F0FDFA`) | Analytics briefs, scientific docs, clean notes |
| **Executive Slate** | Charcoal (`#1E293B`) | Off-White Slate (`#F8FAFC`) | Minimalist docs, formal agreements, whitepapers |
| **Custom Palette** | Any RGB Color | Any RGB Color | Brand guides & custom color schemes |

---



## 🛠️ Project Structure

```text
.
├── Code.js            # Unified Apps Script engine for Code, Tables, Typography & Sidebar
├── appsscript.json    # Apps Script manifest file
├── README.md          # Project documentation & user guide
├── LICENSE            # MIT License
└── .gitignore         # Ignored files
```

---

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! Feel free to open an issue or pull request.

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
