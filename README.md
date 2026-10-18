# Simple Acrobat

<img width="1200" height="750" alt="screenshot" src="https://github.com/user-attachments/assets/5ae5d8b3-7c2d-4750-b6ab-79a003e70845" />

A PDF app that works a bit like Adobe Acrobat. It's one Python file (`acrobat.py`, about 4,000 lines) and opens a Tkinter window. Works on Windows, Mac and Linux.

## What it can do

- **Read:** scroll through pages, zoom, thumbnails, bookmarks, clickable links, search and copy text
- **Comment:** highlight, underline, strikethrough, sticky notes, text boxes, drawing, shapes, arrows, stamps
- **Edit:** change a line of text, add text, images and links, and organize pages (move, rotate, delete, extract, insert, crop)
- **Page extras:** header and footer, page numbers, watermark, background color, bookmarks, document properties
- **Forms and signing:** fill forms, make text fields and check boxes, sign by drawing, typing or with an image
- **Tools:** merge, split, compress, password protect and remove password, redact, OCR, compare two PDFs, flatten comments and forms
- **Create and export:** make a PDF from images or a text file, export pages as PNG, export text, HTML or Word, print
- Undo and redo work for most things

## Install

You need Python 3.8+ (Tkinter comes with most installs).

```
pip install pymupdf pypdf pillow reportlab python-docx pytesseract
```

OCR also needs Tesseract installed on your computer: https://github.com/tesseract-ocr/tesseract

If `reportlab`, `python-docx` or `pytesseract` are missing, the app still starts. Only those features will show a message.

## Run

```
python acrobat.py
```

To open a file right away:

```
python acrobat.py myfile.pdf
```

## How to use it

1. Open a PDF with Ctrl+O.
2. Pick a tool from the toolbar or the Comment and Tools menus.
3. Drag on the page to use it. Esc goes back to the Hand tool.
4. Save with Ctrl+S, undo with Ctrl+Z.

Help > Keyboard Shortcuts has the full list. On Mac use Command instead of Ctrl.

## Libraries

| Library | What I used it for |
| --- | --- |
| tkinter | The window and dialogs |
| PyMuPDF | Showing pages, comments, text editing, redaction, forms, saving |
| pypdf | Merging files |
| Pillow | Turning pages into images, signatures |
| reportlab | Making a PDF from a text file |
| python-docx | Word export |
| pytesseract | OCR |

## AI usage

I wrote the app myself: the Tkinter window, the layout, the toolbar and menus, and how everything fits together. AI helped me with the PDF side only. That means which functions to call in PyMuPDF and pypdf, which modules to use, and how to use them.

## Limits

This is a learning project, not the real Adobe Acrobat. It doesn't have:

- Adobe Sign, cloud storage or AI features
- Certificate based digital signatures
- Accessibility checking or JavaScript in forms
- Multiple documents in tabs (one at a time)

Text editing uses standard fonts, so new text can look a bit different from the original. Word export keeps text only, not layout or pictures.
