# Batch PPT to PDF Converter 🚀 (v1.1.0)

A lightweight and fast Python automation tool that converts all PowerPoint presentations (.ppt, .pptx) in a folder to PDF format simultaneously.

## 🌟 Features
* **Batch Processing:** Converts all files in the directory with one click.
* **Smart Skipping:** Automatically skips files that are already converted to save time.
* **Temporary File Protection:** Ignores PowerPoint's temporary owner files (`~$`) to prevent crashes.
* **Robust Error Handling:** Provides detailed logs for failed conversions using tracebacks.
* **Original Quality:** Uses the native Microsoft PowerPoint engine for perfect layout and font preservation.

## 🛠️ How to Use (For Regular Users)
If you don't have Python or VS Code installed, follow these steps:
1. Go to the **Releases** section of this repository.
2. Download the latest **`convert.exe`** file (v1.1.0).
3. Place the `.exe` file into the folder containing your PowerPoint files.
4. Double-click to run. Your PDFs will be ready in seconds!

## 💻 For Developers
If you want to run or modify the script:
1. Clone this repository.
2. Install the required library:
   ```bash
   pip install comtypes

## 🛡️ Safety & Data Integrity
Non-Destructive: This tool operates in Read-Only mode for your source files.
No Deletion: Your original PowerPoint presentations remain untouched and will not be deleted or modified
