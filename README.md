# PDF Editor for Windows

Desktop PDF editor built with **Python**, **PySide6**, and **PyMuPDF**.  
It is designed for local Windows use and provides practical editing tools for viewing PDFs, adding annotations, replacing paragraph text, editing page content regions, drawing figures, and managing bookmarks.

## Features

- Open and view PDF files
- Navigate pages and zoom in or out
- Add text comments/annotations
- Add symbols onto the page
- Edit text by clicking a paragraph and replacing it in place
- Erase content regions from a page
- Replace or add images in a selected region
- Draw figures such as rectangles, ellipses, and arrows
- View, add, rename, delete, and jump to bookmarks
- Save the edited document as a new PDF

## Project structure

| File | Purpose |
| --- | --- |
| `main.py` | Main desktop application and PDF editing logic |
| `requirements.txt` | Python dependencies |
| `README.md` | Project overview and usage guide |

## Requirements

- Windows
- Python 3.10+ recommended

## Installation

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

## Run the editor

```powershell
.\.venv\Scripts\python main.py
```

## How to use

### Open and view a PDF

1. Launch the app.
2. Click **Open**.
3. Choose a `.pdf` file.
4. Use **Previous**, **Next**, the page number box, **Zoom In**, **Zoom Out**, and **Fit Width** to navigate.

### Add a comment

1. Click **Add Comment**.
2. Click the target point on the page.
3. Enter the comment text.

### Add a symbol

1. Click **Add Symbol**.
2. Choose the symbol, color, and font size.
3. Click the position where the symbol should be placed.

### Edit paragraph text

1. Click **Edit Text**.
2. Move the mouse over page text to preview the paragraph selection.
3. Click the paragraph to edit.
4. Replace the text in the dialog.
5. Adjust font size, alignment, or color if needed.

The editor replaces the selected paragraph inside its original bounds. If needed, the text size is reduced slightly so the new content still fits.

### Erase a content region

1. Click **Erase Region**.
2. Drag a box around the page content you want to remove.

### Replace or add an image

1. Click **Replace/Add Image**.
2. Drag a box over the target region.
3. Choose an image file.

The existing content in that region is removed and the selected image is inserted there.

### Add a figure/object

1. Click **Add Figure**.
2. Choose the figure type and colors.
3. Drag a region on the page to place it.

### Manage bookmarks

Use the bookmark panel on the left side of the window:

- **Add** creates a new bookmark
- **Add Child** creates a nested bookmark
- **Rename** updates the selected bookmark
- **Delete** removes the selected bookmark
- Double-click a bookmark to jump to its page

### Save the result

Click **Save As** to write the edited PDF to a new file.

## Notes and limitations

- PDF editing is not the same as editing a Word document. Some changes are implemented by replacing visible content in a region while preserving a practical editing workflow.
- Text editing works best on PDFs with detectable text content. Image-only scanned PDFs may require OCR before paragraph editing is possible.
- The app currently saves changes through **Save As** rather than overwriting the original file automatically.

## Main dependencies

- [PyMuPDF](https://pymupdf.readthedocs.io/)
- [PySide6](https://doc.qt.io/qtforpython/)
