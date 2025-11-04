[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python 3.9+](https://img.shields.io/badge/python-3.9+-blue.svg)](https://www.python.org/downloads/)

# docx2md-comments

Convert Word documents (.docx) to Markdown while preserving inline comments from Word's review feature. Perfect for processing editorial feedback, code reviews, collaborative document edits, or preparing documents for AI/LLM analysis.

## When to Use This Tool

**Use docx2md-comments when:**
- You need Word review comments extracted and preserved in Markdown
- You want simple, zero-config operation with clear `[FEEDBACK: ...]` format
- You're on Windows and want right-click context menu integration

**Use [MarkItDown](https://github.com/microsoft/markitdown) when:**
- You need to convert multiple file formats (PDF, PowerPoint, Excel, images, etc.)
- You don't specifically need Word comments preserved
- You need plugin support or Azure Document Intelligence integration

## Why docx2md-comments?

Unlike general-purpose converters like [Microsoft's MarkItDown](https://github.com/microsoft/markitdown) or [commentary](https://github.com/hdb/commentary), this tool is purpose-built for extracting Word review comments with:

✅ Simple, zero-config operation  
✅ Clear `[FEEDBACK: ...]` format (not HTML comments)  
✅ Windows context menu integration  
✅ Optimized for editorial/review workflows  

## Features

📝 Extracts comments from `word/comments.xml`  
📍 Inserts `[FEEDBACK: ...]` inline at correct locations  
💻 Cross-platform CLI (Windows, macOS, Linux)  
🖱️ Optional Windows Explorer context menu integration  

## Quick Start
```bash
# Clone the repository
git clone https://github.com/naji44/docx2md-comments.git
cd docx2md-comments

# Install dependencies
pip install -r requirements.txt

# Convert a document
python src/word_to_md_with_comments.py "path/to/document.docx"
```

## Example Output

**Input (Word document with comment):**
```
This is some text with a comment attached.
```

**Output (document_feedback.md):**
```markdown
This is some text [FEEDBACK: Consider rephrasing this for clarity] with a comment attached.
```

## Installation

**Requirements:** Python 3.9 or higher
```bash
# Clone the repository
git clone https://github.com/naji44/docx2md-comments.git
cd docx2md-comments

# Install dependencies
pip install -r requirements.txt
```

## Usage

### Command Line

**Windows:**
```bash
python src\word_to_md_with_comments.py "C:\path\to\file.docx"
```

**macOS/Linux:**
```bash
python src/word_to_md_with_comments.py "/path/to/file.docx"
```

This creates `file_feedback.md` in the same directory as the original.

### Windows Context Menu Integration

Right-click any `.docx` file and select "Convert to Markdown with Comments" for instant conversion.

#### Setup Instructions

1. **Install dependencies:**
```bash
   pip install -r requirements.txt
```

2. **Locate your Python installation:**
```bash
   where python
```
   Example output: `C:\Python311\python.exe`

3. **Place the script:**
   - Create folder: `C:\Scripts\`
   - Copy `src\word_to_md_with_comments.py` to `C:\Scripts\word_to_md_with_comments.py`

4. **Edit the registry file:**
   - Open `registry\add_context_menu.reg` in a text editor
   - Update the Python path if yours differs from `C:\Python311\python.exe`
   - Save the file

5. **Install the context menu:**
   - Double-click `registry\add_context_menu.reg`
   - Click **Yes** when Windows prompts for permission

#### Using the Context Menu

1. Right-click any `.docx` file in Windows Explorer
2. Select "Convert to Markdown with Comments"
3. **Windows 11 users:** Click "Show more options" first to see the full menu

#### Uninstalling

To remove the context menu integration, double-click `registry\remove_context_menu.reg`.

## Troubleshooting

**"Python is not recognized" error:**
- Ensure Python is in your PATH, or use the full Python path in the registry file

**No comments appear in output:**
- Verify your Word document contains comments (Review → New Comment in Word)
- Comments must be saved in the document

**Permission denied errors (Windows context menu):**
- Right-click the `.reg` file and select "Run as administrator"

**Context menu doesn't appear:**
- Verify the Python path in `add_context_menu.reg` is correct
- Try logging out and back in to Windows

## How It Works

1. Unzips the `.docx` file (Word files are ZIP archives)
2. Parses `word/comments.xml` to extract all comments
3. Maps comment IDs to their locations in `word/document.xml`
4. Converts document content to Markdown
5. Inserts `[FEEDBACK: ...]` tags at comment anchor points

## Notes

- Your original `.docx` file is never modified
- Output files use the `_feedback.md` suffix
- The `.reg` files can be deleted after installation (registry entries persist)
- Cross-platform: Works on Windows, macOS, and Linux (context menu is Windows-only)

## Contributing

Contributions are welcome! Feel free to submit issues or pull requests.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
