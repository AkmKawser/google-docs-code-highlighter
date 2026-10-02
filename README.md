# ⚡ Google Docs Code Highlighter & Table Formatter

> A smart Google Apps Script tool that automatically highlights code blocks, formats tables with professional colorful themes & zebra striping, and aligns everything in the middle of your Google Docs — with a dual-tab sidebar and zero markdown backticks required.

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![Google Docs](https://img.shields.io/badge/Google%20Docs-0F9D58?style=for-the-badge&logo=googledocs&logoColor=white)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)

---

## ✨ Features

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
  - **Font Families**: Roboto, Arial, Inter, Open Sans, Lato, Montserrat, Calibri, Trebuchet MS, Georgia, Merriweather, Times New Roman, or Consolas.
  - **Header Row**: Middle-aligned (centered) with larger bold typography (10.5–13pt).
  - **Table Text (Data Rows)**: Strictly left-aligned (not centered or right-aligned) with clean body typography (8.5–10pt).
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
- **🔤 Expanded Monospace Fonts**: Consolas, Roboto Mono, Courier New, Inconsolata, Source Code Pro, Space Mono, PT Mono, and Ubuntu Mono with font sizes from 8.5pt to 12pt.
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

### 1. Format & Center All Tables
- Click **⚡ Code & Table Tools** > **📊 Format & Center All Tables**.
- All data tables will instantly be aligned in the middle of the page and styled with colorful headers, zebra striping, and clean borders.

### 2. Format a Selected Table Only
- Place your cursor inside any table in your document.
- Click **⚡ Code & Table Tools** > **📊 Format Selected Table**.

### 3. Open the Interactive Sidebar
- Click **⚡ Code & Table Tools** > **Open Sidebar (Styles & Colors)**.
- Switch between **📊 Tables** and **⚡ Code Blocks** tabs:
  - **Tables Tab**: Choose your color theme (Corporate Navy, Emerald Mint, Royal Indigo, etc.), cell alignment, and padding with a real-time live table preview. Includes **↩ Undo All Tables Formatting** and **Undo Selected Table Only**.
  - **Code Tab**: Pick syntax themes (GitHub Light, Dracula, Monokai), monospace font, and **Code Indentation** (Smart 2-Space, 4-Space, or Tab conversion). Includes **📐 Auto-Indent All Code Blocks**, **↩ Undo All Code Blocks Formatting**, and **Undo Selected Code Block**.

### 4. Highlight & Indent Code Blocks
- Click **⚡ Code & Table Tools** > **⚡ Highlight All Code (Quick Run)**, or select code and click **Highlight Selected Code**.
- Code blocks are automatically indented according to your chosen indentation preference.
- Use **📐 Auto-Indent All Code Blocks** to re-indent existing blocks anytime.

### 5. Undo Formatting Anytime
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
