# Instant Transcript & Subtitle Cleaner 🧹

[![Python 3.13+](https://shields.io)](https://python.org)
[![License: MIT](https://shields.io)](https://opensource.org)
[![Platform: Windows](https://shields.io)](https://microsoft.com)

A lightweight, fully automated GUI application built with Python to instantly clean up raw video transcripts, subtitles, and YouTube exports. Using smart pattern matching, it strips out distracting clutter like timestamps and duration lines, providing clean, readable text ready for books, articles, formatting, or note-taking.

---

## 🌟 Key Features

* **🚀 Instant Zero-Click Automation**: No "Start" button needed. Processing begins the exact millisecond you drop a file.
* **📥 Drag-and-Drop Interface**: Toss your file directly into the drop zone window without digging through file explorer dialogs.
* **💾 Smart Auto-Save**: Cleaned files automatically save to the **exact same directory** as the source file. It adds a `CLEANED_` prefix to guarantee your original data remains safe.
* **🟢 Real-Time Status Monitor**: Keeps you updated visually with explicit status triggers (`IDLE` ➡️ `STARTED` ➡️ `FINISHED`).
* **📄 Multi-Format Support**: Seamlessly parses and cleans native `.txt`, modern `.docx`, and legacy Microsoft Word `.doc` files.
* **🔒 100% Secure & Private**: Processes all documents locally on your machine. Your proprietary text and data are never sent to external servers or APIs.

---

## 🧹 What It Automatically Removes

The software filters out typical transcript debris, including:
* **Timestamps**: Single lines matching patterns like `12:34` or `1:23:45`.
* **Duration Metrics**: System-generated tracking statements such as `5 seconds`, `10 minutes`, or `2 hours` (case-insensitive).
* **Whitespace Bloat**: Removes empty paragraphs, padding, and unnecessary blank line breaks.

---

## 🛠️ Installation & Setup

### 1. Clone the Repository
```bash
git clone https://github.com
cd YOUR_REPO_NAME
```

### 2. Install Required Dependencies
Open your Command Prompt (CMD) or PowerShell and execute the following pip command:
```bash
pip install python-docx pywin32 tkinterdnd2
```

---

## 🚀 How to Run

Execute the main script through Python:
```bash
python Youtube_Transcript_Cleaner.py
```

### 🎮 Workflow Steps:
1. Launch the application.
2. Drag your dirty transcript file (`.txt`, `.docx`, or `.doc`) from your desktop/folder.
3. Drop it directly into the yellow **【 Drop File Here to Start 】** box.
4. Watch the status switch to `STARTED`, then click **OK** on the completion popup when it triggers `FINISHED`.
5. Check your original file folder for the newly compiled `CLEANED_` document.

---

## 📦 Compiling to a Standalone Executable (.exe)

If you want to package this tool into a standalone Windows app so that it can run on systems without Python installed, use `pyinstaller`.

1. Install PyInstaller:
   ```bash
   pip install pyinstaller
   ```
2. Compile into a single, windowed executable:
   ```bash
   pyinstaller --noconsole --onefile --hidden-import=tkinterdnd2 Youtube_Transcript_Cleaner.py
   ```
3. Locate your ready-to-run `.exe` file inside the newly generated `dist` folder.

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
