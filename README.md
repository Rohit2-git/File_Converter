# FileShift — Local File Converter

**100% free, runs locally, no cloud needed.**

## Supported Conversions

| From → | PDF | MD | DOCX | TXT | HTML |
|--------|-----|----|------|-----|------|
| **PDF** | — | ✅ | ✅ | ✅ | |
| **MD** | ✅ | — | ✅ | ✅ | ✅ |
| **DOCX** | ✅ | ✅ | — | ✅ | |
| **TXT** | ✅ | ✅ | ✅ | — | |
| **HTML** | ✅ | ✅ | ✅ | | — |

## Setup (one time)

```bash
pip install fastapi uvicorn python-multipart pymupdf markdown weasyprint python-docx pdf2docx pypandoc
```

Also install **Pandoc** (used for DOCX conversions):
- **Ubuntu/Debian:** `sudo apt install pandoc`
- **Windows:** https://pandoc.org/installing.html
- **macOS:** `brew install pandoc`

## Run

```bash
cd file_converter
python server.py
```

Then open **http://localhost:8000** in your browser.

## How it works

- Drop a file → pick a target format → click Convert → auto-downloads
- History panel shows all converted files in the session
- Everything runs on your machine, nothing is uploaded anywhere

## Libraries used (all free/open source)

| Library | Purpose |
|---------|---------|
| `pymupdf` | PDF text extraction |
| `weasyprint` | HTML/MD → PDF rendering |
| `python-docx` | Read/write DOCX |
| `pdf2docx` | PDF → DOCX layout conversion |
| `pypandoc` | Universal format bridge (wraps Pandoc) |
| `markdown` | MD → HTML parsing |
| `fastapi` + `uvicorn` | Local web server |
