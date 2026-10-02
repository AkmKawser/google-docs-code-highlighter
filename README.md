# ⚡ Google Docs Code Highlighter

> A smart Google Apps Script tool that automatically detects, highlights, and styles code blocks inside Google Docs — with a custom styling sidebar, preset themes, and zero markdown backticks required.

![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![Google Docs](https://img.shields.io/badge/Google%20Docs-0F9D58?style=for-the-badge&logo=googledocs&logoColor=white)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)

---

## ✨ Features

- **⚡ 1-Click Formatting**: Formats every code block in your entire document in one click.
- **🔒 Curly Brace Tracking Engine**: Uses net `{` and `}` balance counting to guarantee functions, loops, and nested classes never get split into fragmented chunks, even across multiple blank lines.
- **🚫 No Markdown Required**: You don't need to wrap code in triple backticks (\`\`\`). Simply paste your code anywhere in the document.
- **🎨 Interactive Sidebar**:
  - **Themes**: GitHub Light, Dracula Dark, Monokai Dark, Solarized Light, or Custom.
  - **Font Size**: 8.5pt, 9.5pt, 10pt, 11pt, 12pt.
  - **Font Families**: Consolas, Courier New, Roboto Mono.
  - **Color Pickers**: Full custom control over Background, Text, and Border colors.
  - **Live Preview Box**: See your code theme live before applying it.
- **🧠 Prose Rejection Heuristics**: Distinguishes between actual code lines and natural conversational English sentences so regular text is never converted.
- **💾 Auto-Saved Preferences**: Automatically remembers your chosen colors and font sizes for future sessions.

---

## 🚀 Quick Setup Guide

### Method 1: Manual Copy & Paste (Recommended)

1. Open your document in **Google Docs**.
2. In the top menu, go to **Extensions** > **Apps Script**.
3. Clear out everything in `Code.gs`.
4. Copy the entire content of [`Code.js`](./Code.js) and paste it into `Code.gs`.
5. Press **`Ctrl + S`** (or click the disk icon) to save.
6. Return to your Google Docs tab and **refresh the page** (`Ctrl + R` or `F5`).
7. You will now see a new menu: **⚡ Code Highlighter** in the toolbar!

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

### 1. Open the Styles & Colors Sidebar
- Click **⚡ Code Highlighter** > **Open Sidebar (Styles & Colors)**.
- Choose your favorite theme (e.g. **Dracula Dark** or **GitHub Light**), select your font size, and preview it live.
- Click **⚡ Highlight All Code**.

### 2. Quick Highlight
- If you already set your preferences, simply click **⚡ Code Highlighter** > **Highlight All Code (Quick Run)**.

### 3. Highlight Selected Text Only
- If you have a single snippet or terminal command you want to format immediately, highlight it with your cursor and click **Format Selected Text**.

---

## 🎨 Supported Themes

| Theme | Preview Style | Best For |
|---|---|---|
| **GitHub Light** | Light gray background, red/blue keywords, dark text | Clean notes & technical reports |
| **Dracula Dark** | Dark purple background, pink/yellow syntax | Modern dark aesthetic |
| **Monokai Dark** | Deep charcoal background, vibrant pink & green | Sublime/VS Code lovers |
| **Solarized Light** | Warm cream background, cyan/blue syntax | High readability on white pages |
| **Custom** | Fully customizable via RGB color pickers | Tailored brand palettes |

---

## 🛠️ Project Structure

```text
.
├── Code.js            # Main Google Apps Script logic & UI sidebar
├── appsscript.json    # Apps Script manifest file
├── README.md          # Project documentation
├── LICENSE            # MIT License
└── .gitignore         # Ignored files
```

---

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! Feel free to check the [issues page](../../issues).

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
