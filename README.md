# eml-msg-signature-parser-ocr

Extracts sender name/email and phone numbers from email signatures in `.eml` and `.msg`.

## Features
- Latest-reply trimming (thread-aware)
- Signature-window extraction
- Phone extraction with:
  - Fax filtering
  - Date filtering
  - SiliconExpert phone blacklist
  - Country-code recovery
  - Handles lines containing both phone and fax
- Optional OCR (checkbox) to extract phones from signature images

## Run
```bash
py extract_phone.py
