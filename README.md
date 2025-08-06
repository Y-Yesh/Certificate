# Word Document Editor

A Python script to edit Word documents using the `python-docx` library.

## Features

- Open existing Word documents
- Create new Word documents
- Add headings, paragraphs, and tables
- Replace text in documents
- Save documents
- Get document information

## Installation

1. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage

Run the script to automatically process the existing `t.docx` file:

```bash
python word_editor.py
```

### Using the WordDocumentEditor Class

```python
from word_editor import WordDocumentEditor

# Create an editor instance
editor = WordDocumentEditor()

# Open an existing document
editor.open_document("t.docx")

# Get document information
editor.get_document_info()

# Add content
editor.add_heading("New Section", 1)
editor.add_paragraph("This is a new paragraph.")

# Save the document
editor.save_document("modified_document.docx")
```

### Available Methods

- `open_document(file_path)` - Open an existing Word document
- `create_new_document()` - Create a new Word document
- `add_heading(text, level=1)` - Add a heading (level 0-9)
- `add_paragraph(text, style=None)` - Add a paragraph
- `add_table(rows, cols, data=None)` - Add a table
- `add_page_break()` - Add a page break
- `replace_text(old_text, new_text)` - Replace text in the document
- `get_document_info()` - Get information about the current document
- `save_document(file_path=None)` - Save the document

## Example

The script includes an example that:
1. Opens the existing `t.docx` file (if it exists)
2. Adds a title heading
3. Adds sample paragraphs
4. Creates a sample table
5. Saves the result as `edited_document.docx`

## Requirements

- Python 3.6+
- python-docx library

## Notes

- The script will automatically detect and open `t.docx` if it exists in the current directory
- If no existing document is found, it will create a new one
- All operations are non-destructive to the original file 