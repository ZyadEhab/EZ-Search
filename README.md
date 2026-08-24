# 🔍 EZ-Search V6

**EZ-Search** is a fast, automated Python desktop application built to streamline workflow by searching corporate OneDrive directories for specific engineering reference files (.pdf, .dwg, .dxf) and generating structured Excel reports instantly.

---

## ✨ Features

- **⚡ High-Speed Scanning:** Efficiently traverses deep OneDrive folder trees to map required files.
- **🧠 Smart Fallback Search Logic:** Uses a hierarchical search algorithm:
  1. **Exact Match:** Searches for the exact reference ID.
  2. **Fallback 1:** Strips suffixes (e.g., `-IH`) to find the base reference.
  3. **Fallback 2:** Intelligently replaces trailing characters to locate zero-padded versions.
- **🎯 Precise Regex Matching:** Ensures accurate matching by treating special characters (like `-` and `.`) as boundaries, preventing false positives.
- **📊 Automated Excel Reporting:** Integrates with `xlwings` to generate a formatted Excel workbook containing search results, statuses (Done, Missing, Alt), and direct clickable file hyperlinks.
- **🎨 Modern UI/UX:** Built with `Flet` (Flutter for Python), featuring Dark/Light modes, bilingual support (English/Arabic), and responsive progress indicators.

---

## 🛠️ Tech Stack & Tools

- **Language:** Python 3.x
- **GUI Framework:** [Flet](https://flet.dev/) (Based on Flutter)
- **Data Integration:** [xlwings](https://www.xlwings.org/) (MS Excel Automation)
- **Core Modules:** `os`, `re` (Regular Expressions), `threading`, `ctypes`

---

## 🚀 How It Works (The Logic)

Instead of manually searching for hundreds of references, the user pastes a list into the app. The system operates entirely locally, utilizing the machine's `ONEDRIVE` environment variable. It processes the text, compares it against the local directory map, and automatically structures a color-coded Excel sheet so engineers can open the exact CAD/PDF files with a single click.

---

## 👨‍💻 Developer
**Ziad Ehab** 
- Mathematics & Computer Science
- Passionate about workflow automation...
**Ziad Ehab** 
- Mathematics & Computer Science
- Passionate about workflow automation and building practical desktop solutions.
