# Simple Acrobat

A simple Adobe Acrobat style PDF app written in **one Python file** (`acrobat.py`, about 4,050 lines, no comments). Run it and a Tkinter window opens. It works on Windows, Mac and Linux.

## Features

**Read**
- Continuous scrolling, zoom (25% to 800%), fit width, fit page, full screen
- Page thumbnails, bookmarks panel, clickable links
- Find text across the whole document, select and copy text

**Comment**
- Highlight, underline, strikethrough, squiggly underline
- Sticky notes, text boxes, freehand drawing, rectangles, ovals, lines, arrows
- Stamps (Approved, Draft, Confidential and more, or your own text)
- Comments panel to jump to, edit or delete any comment

**Edit**
- Edit a line of text, add new text, add images, add links
- Organize pages: move, rotate, duplicate, insert blank, delete, extract
- Insert pages from another PDF, crop, header and footer, page numbers, watermark, background color
- Edit bookmarks and document properties
- Undo and redo

**Forms and signing**
- Fill form fields, create text fields and check boxes
- Sign by drawing, typing or using a picture of your signature

**Tools**
- Merge, split, compress, password protect (AES-256), remove password
- Redact (mark, then apply to remove content for good)
- OCR to make scanned pages searchable
- Compare two PDFs
- Flatten comments and forms, remove hidden information

**Create and export**
- Create a PDF from images or a text file, or start a blank one
- Export pages as PNG, export text, HTML or a Word document, extract pictures
- Print

## Install

You need Python 3.8 or newer (Tkinter comes with most Python installs).

```
pip install pymupdf pypdf pillow reportlab python-docx pytesseract
```

OCR also needs the Tesseract program installed on your computer: https://github.com/tesseract-ocr/tesseract

If `reportlab`, `python-docx` or `pytesseract` are missing, the app still starts and only those features show a message.

## Run

```
python acrobat.py
```

To open a file straight away:

```
python acrobat.py myfile.pdf
```

## Quick guide

1. Open a PDF with **File > Open** (Ctrl+O).
2. Pick a tool in the toolbar (Hand, Select, Edit Text, Highlight, Note, Text Box, Draw, Redact) or use the **Comment** and **Tools** menus.
3. Drag on the page to use the tool. Press **Esc** to go back to the Hand tool.
4. Save with Ctrl+S. Undo with Ctrl+Z.

Press **Help > Keyboard Shortcuts** inside the app for the full list. On a Mac, use Command instead of Ctrl.

## Libraries

| Library | Used for |
| --- | --- |
| tkinter | The window and all dialogs |
| PyMuPDF | Showing pages, comments, text editing, redaction, forms, saving |
| pypdf | Merging files |
| Pillow | Turning pages into pictures for the screen, signatures |
| reportlab | Creating a PDF from a text file |
| python-docx | Exporting to Word |
| pytesseract | OCR |

## Limits

This is a learning project, not the real Adobe Acrobat. It does not have:
- Adobe Sign, cloud storage, or AI features
- Certificate based digital signatures
- Accessibility checking or JavaScript in forms
- Several documents open in tabs (one document at a time)

Editing text uses standard fonts, so the new text may look slightly different from the original. Word export keeps text only, not layout or pictures.
