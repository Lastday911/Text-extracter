# PDF Formatting Extractor

This project hosts a small Express server that accepts PDF uploads and builds an HTML-friendly representation of the text, including font styles and line breaks. The UI lives in `public/` and talks to `/api/extract`.

## Running with Mistral OCR

The extractor now forwards PDFs to Mistral’s OCR by default using this key baked into the server:

```
POoaOK8EX3VDLmKC2So1PTGPjBpRP5Na
```

and the default endpoint:

```
https://api.mistral.ai/v1/ocr
```

You can override via env vars if needed:

- `MISTRAL_API_KEY` – bearer token (optional override)
- `MISTRAL_OCR_API_URL` – endpoint URL (optional override)

Wenn der OCR-Call fehlschlägt, gibt die API jetzt einen Fehler zurück – es gibt keinen PDF.js-Fallback mehr, damit Probleme mit der Mistral-Integration direkt sichtbar werden.

## Local development

```bash
npm install
MISTRAL_API_KEY=<your key> MISTRAL_OCR_API_URL=<your endpoint> npm run dev
```

Navigate to `http://localhost:3000` and upload a PDF to see the formatted preview and copy buffers.
