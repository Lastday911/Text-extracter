const fs = require('fs');
const path = require('path');
const express = require('express');
const multer = require('multer');
const { readFile } = fs.promises;
const { marked } = require('marked');

const app = express();
const uploadDir = path.join(__dirname, 'uploads');

fs.mkdirSync(uploadDir, { recursive: true });

app.use(express.json({ limit: '25mb' }));

const storage = multer.diskStorage({
  destination(req, file, cb) {
    cb(null, uploadDir);
  },
  filename(req, file, cb) {
    cb(null, `${Date.now()}-${file.originalname}`);
  },
});

const upload = multer({
  storage,
  limits: {
    fileSize: 50 * 1024 * 1024,
  },
  fileFilter(req, file, cb) {
    if (file.mimetype !== 'application/pdf') {
      cb(new Error('Nur PDF-Dateien erlaubt.'));
      return;
    }
    cb(null, true);
  },
});

const MISTRAL_OCR_API_URL = process.env.MISTRAL_OCR_API_URL || 'https://api.mistral.ai/v1/ocr';
const MISTRAL_CHAT_API_URL =
  process.env.MISTRAL_CHAT_API_URL || 'https://api.mistral.ai/v1/chat/completions';
const MISTRAL_VISION_MODEL = process.env.MISTRAL_VISION_MODEL || 'pixtral-large-latest';
const MISTRAL_USAGE_API_URL = process.env.MISTRAL_USAGE_API_URL || '';
const defaultSegmentStyle = {
  fontSize: '16px',
  fontWeight: 400,
  fontStyle: 'normal',
  fontFamily: 'Inter, sans-serif',
  letterSpacing: '0.15px',
};
const fallbackFontSize = 16;
const mistralEnabled = Boolean(MISTRAL_OCR_API_URL);
const markdownRenderer = new marked.Renderer();

marked.setOptions({
  breaks: true,
  gfm: true,
  renderer: markdownRenderer,
});

const toPixelString = (value, fallback = fallbackFontSize) => {
  if (typeof value === 'number') {
    return `${value}px`;
  }
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (trimmed.endsWith('px')) {
      return trimmed;
    }
    if (/^\d+(\.\d+)?$/.test(trimmed)) {
      return `${trimmed}px`;
    }
    return trimmed;
  }
  return `${fallback}px`;
};

const buildSimpleSegment = (text) => ({
  text,
  style: { ...defaultSegmentStyle },
  meta: {},
});

const normalizeLine = (lineData, lineIndex) => {
  const rawSegments =
    lineData.segments ??
    lineData.text_runs ??
    lineData.words ??
    lineData.chunks ??
    (lineData.text ? [{ text: lineData.text }] : []);

  const segments = rawSegments
    .map((segment) => {
      if (!segment) {
        return null;
      }

      const rawText = (segment.text ?? segment.content ?? segment.value ?? '').replace(/\u00A0/g, ' ');
      if (!rawText.trim()) {
        return null;
      }

      const isBold = /bold/i.test(segment.fontWeight ?? segment.style?.fontWeight ?? '') || segment.bold;
      const isItalic =
        /italic/i.test(segment.fontStyle ?? segment.style?.fontStyle ?? '') || segment.italic || segment.oblique;

      const fontFamilyValue = (segment.fontFamily ?? segment.fontName ?? defaultSegmentStyle.fontFamily)
        .replace(/[^a-zA-Z0-9 ,'-]/g, '')
        .trim();

      return {
        text: rawText,
        style: {
          fontSize: toPixelString(segment.fontSize ?? segment.style?.fontSize ?? fallbackFontSize),
          fontWeight: isBold ? 600 : defaultSegmentStyle.fontWeight,
          fontStyle: isItalic ? 'italic' : defaultSegmentStyle.fontStyle,
          fontFamily: `'${fontFamilyValue || defaultSegmentStyle.fontFamily}', ${defaultSegmentStyle.fontFamily}`,
          letterSpacing: defaultSegmentStyle.letterSpacing,
        },
        meta: {
          original: segment,
          position: {
            x: lineData.position?.x ?? lineData.x ?? 0,
            y: lineData.position?.y ?? lineData.y ?? lineIndex * 18,
          },
        },
      };
    })
    .filter(Boolean);

  if (!segments.length) {
    return null;
  }

  return {
    y: lineData.position?.y ?? lineData.y ?? lineIndex * 18,
    segments,
  };
};

const convertSimpleTextToPages = (text) => {
  if (!text) {
    return [];
  }

  const cleaned = text.replace(/\r/g, '');
  const lines = cleaned
    .split('\n')
    .map((line, index) => {
      const normalizedLine = line.replace(/\u00A0/g, ' ');
      if (!normalizedLine.trim()) {
        return null;
      }
      return {
        y: index * 20,
        segments: [buildSimpleSegment(normalizedLine)],
      };
    })
    .filter(Boolean);

  if (!lines.length) {
    return [];
  }

  return [
    {
      number: 1,
      lines,
    },
  ];
};

const pagesToPlainText = (pages = []) =>
  pages
    .map((page) =>
      (page.lines || [])
        .map((line) => (line?.segments || []).map((segment) => segment?.text || '').join(''))
        .map((lineText) => lineText.trimEnd())
        .filter(Boolean)
        .join('\n')
    )
    .filter(Boolean)
    .join('\n\n');

const markdownToHtml = (markdownText) => {
  if (!markdownText || typeof markdownText !== 'string') {
    return '';
  }
  return marked.parse(markdownText);
};

const markdownToPages = (markdownText) => {
  if (!markdownText || typeof markdownText !== 'string') {
    return [];
  }
  const asPlainText = markdownText.replace(/\r/g, '').split('\n').map((line) => line).join('\n');
  return convertSimpleTextToPages(asPlainText);
};

const getApiKey = (req) => {
  const headerKey = req.headers['x-mistral-api-key'] || req.headers['x-mistral-api-key'.toLowerCase()];
  if (typeof headerKey === 'string' && headerKey.trim()) {
    return headerKey.trim();
  }
  return null;
};

const normalizeMistralResponse = (payload) => {
  if (!payload) {
    return { pages: [], html: null, raw: {} };
  }

  const pages = [];
  let html = null;
  const htmlParts = [];
  const raw = payload;

  const normalizePage = (pageData, pageIndex) => {
    const pageNumber = pageData.number ?? pageData.page ?? pageIndex + 1;
    let lines = (pageData.lines ?? pageData.text_lines ?? pageData.blocks ?? [])
      .map((line, idx) => normalizeLine(line, idx))
      .filter(Boolean);
    const images =
      pageData.images?.map((image) => ({
        id: image.id ?? image.image_id ?? null,
        base64:
          image.image_base64 ??
          image.image ??
          (image.image_url?.startsWith('data:') ? image.image_url.split(',')[1] : null),
        position: {
          topLeft: {
            x: image.top_left_x ?? image.x ?? 0,
            y: image.top_left_y ?? image.y ?? 0,
          },
          bottomRight: {
            x: image.bottom_right_x ?? 0,
            y: image.bottom_right_y ?? 0,
          },
        },
      })) || [];

    if (!lines.length && typeof pageData.markdown === 'string' && pageData.markdown.trim()) {
      const mdPages = markdownToPages(pageData.markdown);
      if (mdPages?.length) {
        lines = mdPages[0].lines || [];
      }
      htmlParts.push(markdownToHtml(pageData.markdown));
    }

    if (!lines.length && pageData.text) {
      const fauxPage = convertSimpleTextToPages(pageData.text)[0];
      if (fauxPage?.lines?.length) {
        lines = fauxPage.lines;
      }
    }

    if (!lines.length) {
      return null;
    }

    return {
      number: pageNumber,
      lines,
      images,
    };
  };

  if (Array.isArray(payload.pages) && payload.pages.length) {
    payload.pages.forEach((page, index) => {
      const normalized = normalizePage(page, index);
      if (normalized) {
        pages.push(normalized);
      }
    });
  } else if (Array.isArray(payload.results) && payload.results.length) {
    payload.results.forEach((result, index) => {
      const normalized = normalizePage(result, index);
      if (normalized) {
        pages.push(normalized);
      }
    });
  }

  const markdownText =
    payload.markdown ??
    payload.text_markdown ??
    payload.document_markdown ??
    payload.output_markdown ??
    payload.result_markdown;
  if (markdownText) {
    html = markdownToHtml(markdownText);
  }

  if (htmlParts.length) {
    html = htmlParts.join('\n');
  }

  const pagePlainText = pagesToPlainText(pages);
  if (!html && pagePlainText) {
    html = markdownToHtml(pagePlainText);
  }

  if (!pages.length) {
    const rawText =
      payload.text ??
      payload.document_text ??
      payload.output ??
      payload.result ??
      payload.ocr_text ??
      payload.data ??
      payload.markdown ??
      '';

    if (typeof rawText === 'string' && rawText.trim()) {
      const simplePages = convertSimpleTextToPages(rawText);
      const fallbackHtml = html || markdownToHtml(rawText);
      return { pages: simplePages, html: fallbackHtml, raw };
    }
  }

  if (!html && pagePlainText) {
    html = markdownToHtml(pagePlainText);
  }

  return { pages, html, raw };
};

const callMistralOcr = async (buffer, apiKey) => {
  const base64Doc = buffer.toString('base64');
  const payload = {
    model: 'mistral-ocr-latest',
    document: {
      type: 'document_url',
      document_url: `data:application/pdf;base64,${base64Doc}`,
    },
    include_image_base64: true,
  };

  const response = await fetch(MISTRAL_OCR_API_URL, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${apiKey}`,
      Accept: 'application/json',
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const payloadText = await response.text().catch(() => null);
    throw new Error(
      `Mistral OCR ${response.status}: ${payloadText || 'Keine Details'}`
    );
  }

  return response.json();
};

app.use(express.static(path.join(__dirname, 'public')));

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok' });
});

const escapeHtml = (value) =>
  String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');

const describeImageWithVision = async (base64, apiKey) => {
  if (!base64 || !apiKey || !MISTRAL_VISION_MODEL) {
    return null;
  }

  const payload = {
    model: MISTRAL_VISION_MODEL,
    messages: [
      {
        role: 'user',
        content: [
          {
            type: 'text',
            text: 'Erzeuge eine kurze, sachliche Bildbeschreibung auf Deutsch für ein Word-Dokument. Maximal 2 Sätze.',
          },
          { type: 'image_url', image_url: `data:image/jpeg;base64,${base64}` },
        ],
      },
    ],
    max_tokens: 120,
  };

  const response = await fetch(MISTRAL_CHAT_API_URL, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${apiKey}`,
      Accept: 'application/json',
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const errText = await response.text().catch(() => '');
    throw new Error(`Vision-Call fehlgeschlagen: ${response.status} ${errText}`);
  }

  const data = await response.json();
  const choice = data?.choices?.[0]?.message?.content;
  if (Array.isArray(choice)) {
    const textPart = choice.find((c) => c.type === 'text') || choice[0];
    return textPart?.text?.trim() || null;
  }
  if (typeof choice === 'string') {
    return choice.trim();
  }
  return null;
};

app.post('/api/extract', upload.single('pdf'), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'Bitte eine PDF-Datei hochladen.' });
  }

  try {
    const apiKey = getApiKey(req);
    if (!apiKey) {
      return res.status(400).json({ error: 'Bitte eigenen Mistral API-Key angeben.' });
    }
    if (!mistralEnabled) {
      return res.status(500).json({ error: 'Mistral OCR ist nicht konfiguriert.' });
    }

    const fileBuffer = await readFile(req.file.path);
    const mistralPayload = await callMistralOcr(fileBuffer, apiKey);
    const normalized = normalizeMistralResponse(mistralPayload);
    const pages = normalized?.pages || [];
    const html = normalized?.html || null;

    if (!pages.length && !html) {
      console.error('Mistral OCR keine Daten. Keys:', Object.keys(normalized.raw || {}));
      console.error('Mistral OCR Payload (gekürzt):', JSON.stringify(mistralPayload || {}).slice(0, 1200));
      return res.status(502).json({ error: 'Mistral OCR lieferte keine auswertbaren Daten.' });
    }

    res.json({ pages, pageCount: pages.length, source: 'mistral', html });
  } catch (error) {
    console.error('PDF-Parsing-Fehler:', error);
    res.status(502).json({ error: error?.message || 'Die PDF-Datei konnte nicht verarbeitet werden.' });
  } finally {
    if (req.file?.path) {
      fs.promises.unlink(req.file.path).catch(() => {});
    }
  }
});

app.post('/api/export-docx', async (req, res) => {
  try {
    const { disableDescriptions = false } = req.body || {};
    const apiKey = getApiKey(req);
    if (!disableDescriptions && !apiKey) {
      return res.status(400).json({ error: 'Bitte eigenen Mistral API-Key angeben.' });
    }

    const htmlInput = typeof req.body?.html === 'string' ? req.body.html.trim() : '';
    const pages = Array.isArray(req.body?.pages) ? req.body.pages : [];
    if (!pages.length && !htmlInput) {
      return res.status(400).json({ error: 'Keine Inhalte zum Export übergeben.' });
    }

    const escapeLine = (line) =>
      (line?.segments || [])
        .map((segment) => escapeHtml(segment.text || ''))
        .join('')
        .trim();
    const includePageText = !htmlInput;

    const descriptions = {};
    if (!disableDescriptions) {
      for (const page of pages) {
        if (!Array.isArray(page.images)) continue;
        // Sequential to stay under rate limits
        /* eslint-disable no-await-in-loop */
        for (const image of page.images) {
          if (!image?.base64 || image._removed) continue;
          if (image._replaceWithDescription && image._description) {
            const key = image.id || image.base64.slice(0, 16);
            descriptions[key] = image._description;
            continue;
          }
          const key = image.id || image.base64.slice(0, 16);
          if (descriptions[key]) {
            continue;
          }
          try {
            const desc =
              (await describeImageWithVision(image.base64, apiKey)) ||
              'Bildbeschreibung konnte nicht erzeugt werden.';
            descriptions[key] = desc;
          } catch (err) {
            console.error('Vision-Beschreibung fehlgeschlagen:', err?.message || err);
            descriptions[key] = 'Bildbeschreibung konnte nicht erzeugt werden.';
          }
        }
        /* eslint-enable no-await-in-loop */
      }
    }

    const bodyParts = [];
    bodyParts.push('<h1>PDF-Export</h1>');
    if (htmlInput) {
      bodyParts.push(htmlInput);
    }
    pages.forEach((page) => {
      const hasImages = Array.isArray(page.images) && page.images.length;
      const hasText = includePageText && Array.isArray(page.lines) && page.lines.length;
      if (!hasText && !hasImages) {
        return;
      }

      bodyParts.push(`<h2>Seite ${escapeHtml(page.number || '')}</h2>`);
      if (hasText) {
        (page.lines || []).forEach((line) => {
          const text = escapeLine(line);
          if (text) {
            bodyParts.push(`<p>${text}</p>`);
          }
        });
      }
      if (hasImages) {
        page.images.forEach((img, idx) => {
          if (img._removed) {
            return;
          }
          const key = img.id || img.base64?.slice(0, 16) || `${page.number}-${idx}`;
          if (disableDescriptions) {
            if (img.base64 && !img._replaceWithDescription) {
              bodyParts.push(
                `<p><img alt="Bild ${idx + 1}" style="max-width:100%;height:auto;" src="data:image/jpeg;base64,${img.base64}"/></p>`
              );
            } else if (img._description) {
              bodyParts.push(`<p><strong>Bild:</strong> ${escapeHtml(img._description)}</p>`);
            }
            return;
          }
          const desc =
            (img._replaceWithDescription && img._description) ||
            descriptions[key] ||
            'Bildbeschreibung nicht verfügbar.';
          bodyParts.push(`<p><strong>Bild:</strong> ${escapeHtml(desc)}</p>`);
        });
      }
      bodyParts.push('<hr />');
    });

    const htmlDoc = `
      <!DOCTYPE html>
      <html lang="de">
        <head>
          <meta charset="UTF-8" />
          <style>
            body { font-family: "Segoe UI", Arial, sans-serif; line-height: 1.5; color: #111; }
            h1, h2 { color: #0f172a; }
            hr { border: 0; border-top: 1px solid #e2e8f0; margin: 1.5rem 0; }
            p { margin: 0.35rem 0; }
          </style>
        </head>
        <body>
          ${bodyParts.join('\n')}
        </body>
      </html>
    `;

    res.setHeader(
      'Content-Disposition',
      `attachment; filename="export-${Date.now().toString().slice(-6)}.doc"`
    );
    res.setHeader('Content-Type', 'application/msword');
    res.send(htmlDoc);
  } catch (error) {
    console.error('Word-Export-Fehler:', error);
    res.status(500).json({ error: error?.message || 'Export fehlgeschlagen.' });
  }
});

app.post('/api/describe-image', async (req, res) => {
  try {
    const apiKey = getApiKey(req);
    if (!apiKey) {
      return res.status(400).json({ error: 'Bitte eigenen Mistral API-Key angeben.' });
    }
    const base64 = req.body?.base64;
    if (!base64) {
      return res.status(400).json({ error: 'Keine Bilddaten übergeben.' });
    }
    const description = await describeImageWithVision(base64, apiKey);
    if (!description) {
      return res.status(502).json({ error: 'Keine Bildbeschreibung erhalten.' });
    }
    res.json({ description });
  } catch (error) {
    console.error('Vision-Endpoint-Fehler:', error);
    res.status(500).json({ error: error?.message || 'Beschreibung fehlgeschlagen.' });
  }
});

app.get('/api/usage', async (req, res) => {
  try {
    const apiKey = getApiKey(req);
    if (!apiKey) {
      return res.status(400).json({ error: 'Bitte eigenen Mistral API-Key angeben.' });
    }
    if (!MISTRAL_USAGE_API_URL) {
      return res.status(501).json({
        error:
          'Usage-Endpoint nicht konfiguriert. Bitte MISTRAL_USAGE_API_URL setzen oder im Mistral-Dashboard prüfen.',
        unsupported: true,
      });
    }

    const response = await fetch(MISTRAL_USAGE_API_URL, {
      method: 'GET',
      headers: {
        Authorization: `Bearer ${apiKey}`,
        Accept: 'application/json',
      },
    });

    const text = await response.text().catch(() => '');
    if (!response.ok) {
      return res.status(response.status).json({
        error: `Usage-Request fehlgeschlagen (${response.status})`,
        details: text?.slice(0, 500) || null,
      });
    }

    let data = null;
    try {
      data = text ? JSON.parse(text) : null;
    } catch (err) {
      return res.status(502).json({ error: 'Usage-Antwort konnte nicht gelesen werden.' });
    }

    res.json({ data });
  } catch (error) {
    console.error('Usage-Endpoint-Fehler:', error);
    res.status(502).json({ error: error?.message || 'Usage-Request fehlgeschlagen.' });
  }
});

app.use((error, req, res, next) => {
  if (res.headersSent) {
    return next(error);
  }
  console.error('Upload-Fehler:', error?.message || error);
  res.status(400).json({ error: error?.message || 'Beim Upload ist ein Fehler aufgetreten.' });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  // eslint-disable-next-line no-console
  console.log(`PDF-Extractor läuft auf http://localhost:${PORT}`);
});
