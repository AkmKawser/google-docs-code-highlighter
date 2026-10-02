# ⚡ Google Docs Code, Table & Typography Suite

> A smart Google Apps Script suite that formats document typography (Title, Heading 1, Sub-headings, Body), highlights code blocks with auto-indentation, and styles data tables with professional themes & zebra striping — with an interactive 3-tab sidebar and zero markdown backticks required.

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![Google Docs](https://img.shields.io/badge/Google%20Docs-0F9D58?style=for-the-badge&logo=googledocs&logoColor=white)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)

---

## ✨ Features

### ✍️ Document Typography & Heading Formatter
- **📖 Document Title Formatting (`TITLE`)**:
  - Independent typography controls: Font family, size (18–36pt), bold, text alignment (Center, Left, Right).
  - **Full Color Picking**: Individual pickers for text color and background color.
  - **Background Style Switch**: Switch between **Full-Width Header Banner** (seamless edge-to-edge block) or **Inline Text Highlight**.
  - Background enable/disable toggle.
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
  - **Custom Configuration**: Full granular control over every font, color, and size.
- **👁️ Live Interactive Typography Preview**: Shows real-time updates for Title, Heading 1, Sub-headings, and body paragraphs as you adjust fonts and colors.
- **↩ 1-Click Document Typography Undo**: Reverts all headings, titles, and body paragraphs back to Google Docs defaults and unrolls header banners cleanly back into normal paragraphs.

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
  - **Header Row**: Middle-aligned (centered) with larger bold typography (8.5–20pt) &mdash; table title row is always bold.
  - **Table Text (Data Rows)**: Strictly left-aligned with top vertical alignment (always on top, never middle-aligned) and clean body typography (8–20pt).
  - **Table Element**: Centered on page across document margins.
  - Refined cell padding (Compact, Normal, Relaxed) and soft borders.
- **🛡️ Code Block Safety**: Automatically differentiates between regular data tables and code containers, keeping your syntax highlighting intact while centering both!
- **↩ 1-Click Table Formatting Undo**: Reset all tables (or the selected table) back to clean standard document defaults (white backgrounds, 1pt black borders, left alignment, default font) with one click.

---

### ⚡ Smart Code Highlighter & Indentation
- **⚡ 1-Click Formatting**: Formats every code block in your entire document in one click.
- **📐 Smart Code Indentation Engine**:
  - **Smart Auto-Indent (2 or 4 spaces)**: Syntax-aware indentation that automatically calculates brace nesting, function blocks, loops, and tags.
  - **Tab Normalization**: Converts hard tabs into clean 2-space or 4-space indentations and normalizes irregular spacing.
  - **Dedicated Indent Action**: Re-indent all existing code blocks or selected blocks directly from the sidebar or menu.
- **↩ 1-Click Code Blocks Formatting Undo**: Safely unwraps 1x1 code block tables back into standard document paragraphs, removing background shading and syntax colors while preserving your code text.
- **🔒 Curly Brace Tracking Engine**: Uses net `{` and `}` balance counting to guarantee functions, loops, and nested classes never get split into fragmented chunks, even across multiple blank lines.
- **🚫 No Markdown Required**: You don't need to wrap code in triple backticks (\`\`\`). Simply paste your code anywhere in the document.
- **🎨 Interactive Sidebar Themes**: GitHub Light, Dracula Dark, Monokai Dark, Solarized Light, or Custom.
- **🔤 Expanded Monospace Fonts**: Consolas, JetBrains Mono, Roboto Mono, Courier New, Inconsolata, Source Code Pro, Space Mono, PT Mono, and Ubuntu Mono with font sizes from 8pt to 20pt (strictly clean regular weight, never bold).
- **🧠 Prose Rejection Heuristics**: Distinguishes between actual code lines and natural conversational English sentences so regular text is never converted.
- **💾 Auto-Saved Preferences**: Automatically remembers your chosen fonts, colors, padding, indentation, and font sizes for future sessions.

---

## 🚀 Quick Setup Guide

### Method 1: Manual Copy & Paste (Recommended)

1. Open your document in **Google Docs**.
2. In the top menu, go to **Extensions** > **Apps Script**.
3. Clear out everything in `Code.gs`.
4. Copy the entire content of [`Code.js`](./Code.js) and paste it into `Code.gs`.
5. Press **`Ctrl + S`** (or click the disk icon) to save.
6. Return to your Google Docs tab and **refresh the page** (`Ctrl + R` or `F5`).
7. You will now see a new menu: **⚡ Code & Table Tools** in the toolbar!

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

### 1. 🚀 1-Click Smart Auto-Format Entire Document (All-in-One)
- Click **⚡ Code, Table & Typography Suite** > **🚀 Smart Auto-Format Entire Document** (or click the top banner button in the sidebar).
- Intelligently scans the entire document to auto-detect what formatting to apply:
  - **Code Blocks**: Detects markdown fenced blocks (\`\`\`) and programming syntax &rarr; converts to styled, auto-indented code block tables.
  - **Data Tables**: Detects data tables &rarr; centers across margins and applies colorful header and zebra striping.
  - **Document Titles**: Detects titles by position, \`#\` prefix, or \`Title:\` tags &rarr; applies selected Title typography & banners.
  - **Headings & Sub-Headings**: Detects \`#\`, \`##\`, \`###\`, numbered sections (\`1. Introduction\`, \`1.1 Overview\`, \`1.1.1 Details\`, \`Section 1:\`, \`Step 1:\`), outline letters (\`A.\`, \`a)\`, \`(1)\`), ALL-CAPS headers, and bold standalone lines &rarr; styles as Heading 1 or Sub-headings.
  - **Body Text**: Applies clean body typography and automatically detects inline code tokens (\`code\`) with monospace highlight.

### 2. Format Document Typography
- Click **⚡ Code, Table & Typography Suite** > **✍️ Format Document Typography**.
- Your document Title, Heading 1s, Sub-headings (H2, H3, Subtitle), and body paragraphs will be auto-detected and styled with your chosen theme, fonts, colors, and background banners.
- Or highlight any passage and click **✍️ Format Selected Text Auto** to auto-detect and format only the selected lines.

### 3. Format & Center All Tables
- Click **⚡ Code, Table & Typography Suite** > **📊 Format & Center All Tables**.
- All data tables will instantly be aligned in the middle of the page and styled with colorful headers, zebra striping, and clean borders.

### 4. Open the Interactive 3-Tab Sidebar
- Click **⚡ Code, Table & Typography Suite** > **Open Sidebar (Styles & Colors)**.
- Switch between **✍️ Typography**, **📊 Tables**, and **⚡ Code** tabs:
  - **Top Banner**: 1-Click **⚡ Auto-Format Entire Document**.
  - **Typography Tab**: Select a preset (Executive Navy, Modern Tech, etc.) or customize title/heading/subheading colors, fonts, sizes, and banner styles with a live document preview. Includes **✍️ Auto-Detect & Format Selection**.
  - **Tables Tab**: Choose your color theme (Corporate Navy, Emerald Mint, Royal Indigo, etc.), cell alignment, and padding with a real-time live table preview. Includes **↩ Undo All Tables Formatting** and **Undo Selected Table Only**.
  - **Code Tab**: Pick syntax themes (GitHub Light, Dracula, Monokai), monospace font, and **Code Indentation** (Smart 2-Space, 4-Space, or Tab conversion). Includes **📐 Auto-Indent All Code Blocks**, **↩ Undo All Code Blocks Formatting**, and **Undo Selected Code Block**.

### 5. Highlight & Indent Code Blocks
- Click **⚡ Code, Table & Typography Suite** > **⚡ Highlight All Code (Quick Run)**, or select code and click **Highlight Selected Code**.
- Code blocks are automatically indented according to your chosen indentation preference. Fenced code blocks (\`\`\`) are cleanly parsed with fence markers stripped.

### 6. Undo Formatting Anytime
- Click **↩ Undo Document Text Formatting** to revert titles, headings, and body paragraphs back to clean Google Docs defaults and unroll banner tables.
- Click **↩ Undo All Tables Formatting** to revert tables back to standard document defaults.
- Click **↩ Undo All Code Blocks Formatting** to convert 1x1 code block containers back to normal paragraphs.

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
├── Code.js            # Unified Apps Script logic for Tables & Code Blocks + Dual-tab UI Sidebar
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
