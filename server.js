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
  textDecoration: 'none',
};
const fallbackFontSize = 16;
const mistralEnabled = Boolean(MISTRAL_OCR_API_URL);
const markdownRenderer = new marked.Renderer();

const normalizeAlignment = (value) => {
  if (!value) return null;
  const raw = String(value).trim().toLowerCase();
  if (['left', 'start', 'l', 'align_left'].includes(raw)) return 'left';
  if (['right', 'end', 'r', 'align_right'].includes(raw)) return 'right';
  if (['center', 'centre', 'middle', 'c', 'align_center'].includes(raw)) return 'center';
  if (['justify', 'justified', 'full', 'distributed', 'block'].includes(raw)) return 'justify';
  return null;
};

const normalizeTextDecoration = (segment = {}) => {
  const rawDecoration =
    segment.textDecoration ||
    segment.text_decoration ||
    segment.textDecorationLine ||
    segment.text_decoration_line ||
    segment.decoration ||
    segment.style?.textDecoration ||
    segment.style?.text_decoration;

  const textDecorStr = typeof rawDecoration === 'string' ? rawDecoration.toLowerCase() : '';
  const fromString = [];
  if (textDecorStr.includes('underline')) fromString.push('underline');
  if (textDecorStr.includes('line-through') || textDecorStr.includes('strikethrough')) {
    fromString.push('line-through');
  }

  const hasUnderline = Boolean(
    segment.underline ||
      segment.isUnderline ||
      segment.underlined ||
      fromString.includes('underline')
  );
  const hasStrike = Boolean(
    segment.strikethrough ||
      segment.strike ||
      segment.isStrike ||
      segment.isStrikethrough ||
      fromString.includes('line-through')
  );

  if (hasUnderline && hasStrike) return 'underline line-through';
  if (hasUnderline) return 'underline';
  if (hasStrike) return 'line-through';
  return 'none';
};

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

const buildSimpleSegment = (text, styleOverrides = {}) => ({
  text,
  style: { ...defaultSegmentStyle, ...styleOverrides },
  meta: {},
});

const parseMarkdownToSegments = (text) => {
  if (!text) return [];
  
  // Regex for bold (** or __), italic (* or _), and underline (<u>...</u>)
  // This is a simplified parser and might not handle nested tags perfectly in all edge cases,
  // but covers the 99% use case for OCR output.
  // We split by tags and keep delimiters to identify them.
  
  const segments = [];
  let currentStyle = {
    fontWeight: defaultSegmentStyle.fontWeight,
    fontStyle: defaultSegmentStyle.fontStyle,
    textDecoration: defaultSegmentStyle.textDecoration,
  };

  // Strategy: Scan string and process tokens. 
  // Because regex split is tricky with overlapping, we'll use a tokenizing loop.
  // Supported: **bold**, __bold__, *italic*, _italic_, <u>underline</u>
  
  let remaining = text;
  
  while (remaining.length > 0) {
    // Find earliest special token
    const bold1 = remaining.indexOf('**');
    const bold2 = remaining.indexOf('__');
    const italic1 = remaining.indexOf('*');
    const italic2 = remaining.indexOf('_');
    const underlineStart = remaining.indexOf('<u>');
    const underlineEnd = remaining.indexOf('</u>');

    // Filter out -1 and find min
    const indices = [bold1, bold2, italic1, italic2, underlineStart, underlineEnd]
      .filter(i => i !== -1)
      .sort((a, b) => a - b);

    if (indices.length === 0) {
      // No more tokens
      segments.push(buildSimpleSegment(remaining, currentStyle));
      break;
    }

    const nextIndex = indices[0];
    
    // Push text before token
    if (nextIndex > 0) {
      segments.push(buildSimpleSegment(remaining.substring(0, nextIndex), currentStyle));
    }

    // Process token
    if (nextIndex === bold1) {
      // Toggle bold
      const isBold = currentStyle.fontWeight === 600;
      currentStyle = { ...currentStyle, fontWeight: isBold ? 400 : 600 };
      remaining = remaining.substring(nextIndex + 2);
    } else if (nextIndex === bold2) {
      const isBold = currentStyle.fontWeight === 600;
      currentStyle = { ...currentStyle, fontWeight: isBold ? 400 : 600 };
      remaining = remaining.substring(nextIndex + 2);
    } else if (nextIndex === italic1) {
      // Toggle italic (check if it's not part of **)
      // If we hit * and it's actually part of **, bold1 would have been min index? 
      // Wait, if string is "**text**", bold1 is 0, italic1 is 0. 
      // We need to prioritize longer tokens.
      
      // Refined check:
      if (remaining.startsWith('**')) {
        const isBold = currentStyle.fontWeight === 600;
        currentStyle = { ...currentStyle, fontWeight: isBold ? 400 : 600 };
        remaining = remaining.substring(2);
      } else {
        const isItalic = currentStyle.fontStyle === 'italic';
        currentStyle = { ...currentStyle, fontStyle: isItalic ? 'normal' : 'italic' };
        remaining = remaining.substring(1);
      }
    } else if (nextIndex === italic2) {
      if (remaining.startsWith('__')) {
        const isBold = currentStyle.fontWeight === 600;
        currentStyle = { ...currentStyle, fontWeight: isBold ? 400 : 600 };
        remaining = remaining.substring(2);
      } else {
        const isItalic = currentStyle.fontStyle === 'italic';
        currentStyle = { ...currentStyle, fontStyle: isItalic ? 'normal' : 'italic' };
        remaining = remaining.substring(1);
      }
    } else if (nextIndex === underlineStart) {
       currentStyle = { ...currentStyle, textDecoration: 'underline' };
       remaining = remaining.substring(3);
    } else if (nextIndex === underlineEnd) {
       currentStyle = { ...currentStyle, textDecoration: 'none' }; // Or restore previous? Simplified to none for now or 'default'
       // Better: if we have complex nesting, we might need a stack. 
       // But for simple OCR output, toggling off is usually safe.
       // Let's assume plain text default is none.
       remaining = remaining.substring(4);
    } else {
      // Should not happen
      remaining = remaining.substring(1);
    }
  }
  
  return segments.filter(s => s.text);
};

const escapeHtml = (value) =>
  String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');

const normalizeTable = (tableData, index = 0) => {
  if (!tableData) return null;
  const rows =
    tableData.rows ||
    tableData.data ||
    tableData.cells ||
    tableData.content ||
    tableData.table ||
    [];

  if (!Array.isArray(rows) || !rows.length) {
    return null;
  }

  // Attempt to find geometry for positioning
  // Mistral often returns 'geometry' with 'bounding_box' or 'top_left'
  let yPos = 0;
  let bbox = null;

  if (tableData.geometry?.bounding_box) {
    // [x_min, y_min, x_max, y_max] usually
    bbox = tableData.geometry.bounding_box;
    yPos = bbox[1] || 0;
  } else if (tableData.geometry?.top_left) {
    yPos = tableData.geometry.top_left.y || 0;
  } else if (typeof tableData.y === 'number') {
    yPos = tableData.y;
  } else if (typeof tableData.top_left_y === 'number') {
    yPos = tableData.top_left_y;
  }

  const normalizedRows = rows
    .map((row) => {
      if (Array.isArray(row)) return row;
      if (Array.isArray(row?.cells)) return row.cells;
      return null;
    })
    .filter(Boolean);

  if (!normalizedRows.length) {
    return null;
  }

  const buildCellText = (cell) => {
    if (cell == null) return '';
    if (typeof cell === 'string') return String(cell);
    const raw = cell.text || cell.content || cell.value || cell.plain_text || '';
    return String(raw);
  };

  const buildCellHtml = (cell) => escapeHtml(buildCellText(cell));

  const htmlRows = normalizedRows
    .map((row) => {
      const cells = row
        .map((cell) => {
          if (cell == null) return '<td></td>';
          const colspan = cell.colspan || cell.col_span || cell.span_cols;
          const rowspan = cell.rowspan || cell.row_span || cell.span_rows;
          const tag = cell.header || cell.is_header || cell.th ? 'th' : 'td';
          const attrs = [];
          if (colspan && Number(colspan) > 1) attrs.push(`colspan="${Number(colspan)}"`);
          if (rowspan && Number(rowspan) > 1) attrs.push(`rowspan="${Number(rowspan)}"`);
          const attrStr = attrs.length ? ` ${attrs.join(' ')}` : '';
          return `<${tag}${attrStr}>${buildCellHtml(cell)}</${tag}>`;
        })
        .join('');
      return `<tr>${cells}</tr>`;
    })
    .join('');

  return {
    id: tableData.id || tableData.table_id || `table-${index}`,
    html: `<table>${htmlRows}</table>`,
    text: normalizedRows
      .map((row) => row.map((cell) => buildCellText(cell)).join('\t'))
      .join('\n'),
    y: yPos,
    boundingBox: bbox,
  };
};

const normalizeLine = (lineData, lineIndex) => {
// ... existing normalizeLine code ...
// I will keep the previous implementation of normalizeLine exactly as is, but just re-declare it to be safe if context was lost.
// Actually, I'll rely on the existing `normalizeLine` if I can, but the tool requires replacing `old_string`.
// Since I am replacing `normalizeTable`, I need to be careful.
// The previous `normalizeTable` was modified in the last turn.
// I will replace the block from `const normalizeTable ...` down to `const normalizeLine ...` to be safe, 
// but wait, `normalizeLine` is large. 
// I will just replace `normalizeTable` specifically.
// But I also need to update `api/export-docx` which is further down.
// I will split this into two replacements for safety.
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

      const textDecoration = normalizeTextDecoration(segment.style || segment);

      return {
        text: rawText,
        style: {
          fontSize: toPixelString(segment.fontSize ?? segment.style?.fontSize ?? fallbackFontSize),
          fontWeight: isBold ? 600 : defaultSegmentStyle.fontWeight,
          fontStyle: isItalic ? 'italic' : defaultSegmentStyle.fontStyle,
          textDecoration: textDecoration || defaultSegmentStyle.textDecoration,
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

  const alignment =
    normalizeAlignment(lineData.text_alignment) ||
    normalizeAlignment(lineData.alignment) ||
    normalizeAlignment(lineData.textAlign) ||
    normalizeAlignment(lineData.text_align) ||
    normalizeAlignment(lineData.align) ||
    normalizeAlignment(lineData.justification) ||
    normalizeAlignment(lineData.justify) ||
    normalizeAlignment(lineData.style?.textAlign) ||
    normalizeAlignment(lineData.style?.text_alignment) ||
    null;

  return {
    y: lineData.position?.y ?? lineData.y ?? lineIndex * 18,
    segments,
    align: alignment || 'left',
    x: lineData.position?.x ?? lineData.x ?? 0, // Ensure X is accessible at line level
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
      
      // Use the new parser
      const segments = parseMarkdownToSegments(normalizedLine);
      
      return {
        y: index * 20,
        segments: segments.length ? segments : [buildSimpleSegment(normalizedLine)],
        align: 'left',
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
    .map((page) => {
      const lineText = (page.lines || [])
        .map((line) => (line?.segments || []).map((segment) => segment?.text || '').join(''))
        .map((lineText) => lineText.trimEnd())
        .filter(Boolean)
        .join('\n');
      const tableText = (page.tables || [])
        .map((tbl) => tbl.text)
        .filter(Boolean)
        .join('\n\n');
      return [lineText, tableText].filter(Boolean).join('\n\n');
    })
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
    const tablesRaw = pageData.tables || pageData.table || [];
    const tables = Array.isArray(tablesRaw)
      ? tablesRaw.map((tbl, idx) => normalizeTable(tbl, idx)).filter(Boolean)
      : [];
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
      tables,
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

  const topLevelTables = Array.isArray(payload.tables)
    ? payload.tables.map((tbl, idx) => normalizeTable(tbl, idx)).filter(Boolean)
    : [];
  if (topLevelTables.length) {
    if (pages.length) {
      pages[0].tables = [...(pages[0].tables || []), ...topLevelTables];
    } else {
      pages.push({ number: 1, lines: [], images: [], tables: topLevelTables });
    }
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

    const pages = Array.isArray(req.body?.pages) ? req.body.pages : [];
    const htmlInput = typeof req.body?.html === 'string' ? req.body.html.trim() : '';

    // Strategy: "1 to 1" Fidelity
    // We prioritize the detailed 'pages' object model (Mistral OCR JSON) because it contains
    // exact font sizes, styles, and positioning that Markdown lacks.
    // However, to ensure tables render correctly (as tables, not just text), we must:
    // 1. Use the 'tables' array from the JSON.
    // 2. Filter out the text lines that correspond to the table content (using bounding boxes) to avoid duplication.
    const usePagesLayout = pages.length > 0;

    if (!usePagesLayout && !htmlInput) {
      return res.status(400).json({ error: 'Keine Inhalte zum Export übergeben.' });
    }

    const escapeLine = (line) =>
      (line?.segments || [])
        .map((segment) => escapeHtml(segment.text || ''))
        .join('')
        .trim();
        
    const getLineAlign = (line) =>
      normalizeAlignment(
        line?.align ||
          line?.textAlign ||
          line?.text_align ||
          line?.text_alignment ||
          line?.justification ||
          line?.justify
      );
      
    const segmentToHtml = (segment) => {
      const text = escapeHtml(segment?.text || '');
      if (!text) return '';
      const allowedKeys = {
        fontSize: true,
        fontWeight: true,
        fontStyle: true,
        fontFamily: true,
        letterSpacing: true,
        textDecoration: true,
      };
      const styleString = Object.entries(segment?.style || {})
        .filter(([key]) => allowedKeys[key])
        .map(([key, value]) => {
          const cssKey = key.replace(/[A-Z]/g, (m) => `-${m.toLowerCase()}`);
          return `${cssKey}:${escapeHtml(value)}`;
        })
        .join(';');
      return styleString ? `<span style="${styleString}">${text}</span>` : text;
    };

    const renderLineHtml = (line) =>
      (line?.segments || [])
        .map((segment) => segmentToHtml(segment))
        .join('')
        .trim();

    const descriptions = {};
    if (!disableDescriptions) {
      for (const page of pages) {
        if (!Array.isArray(page.images)) continue;
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

    if (usePagesLayout) {
      pages.forEach((page) => {
        bodyParts.push(`<h2>Seite ${escapeHtml(page.number || '')}</h2>`);
        
        const tables = (page.tables || []).map(t => ({ ...t, type: 'table' }));
        
        // Calculate table exclusion zones
        // Mistral bbox is [x_min, y_min, x_max, y_max]
        const tableZones = tables
          .map(t => t.boundingBox)
          .filter(bbox => Array.isArray(bbox) && bbox.length === 4)
          .map(bbox => ({ xMin: bbox[0], yMin: bbox[1], xMax: bbox[2], yMax: bbox[3] }));

        // Filter lines: Keep only lines that are NOT inside a table zone
        // A line is "inside" if its center point falls within the box
        const lines = (page.lines || []).filter(line => {
           const lx = line.x ?? line.meta?.position?.x ?? 0;
           const ly = line.y ?? line.meta?.position?.y ?? 0;
           // Simple point check. If inside any table zone, drop it.
           // We can add a small margin of error if needed, but exact check is usually fine.
           const isInsideTable = tableZones.some(z => 
             lx >= z.xMin && lx <= z.xMax && ly >= z.yMin && ly <= z.yMax
           );
           return !isInsideTable;
        }).map(l => ({ ...l, type: 'line' }));

        // Interleaved Rendering: Combine and Sort by Y
        const elements = [...lines, ...tables].sort((a, b) => {
           const ay = a.y ?? a.boundingBox?.[1] ?? 0;
           const by = b.y ?? b.boundingBox?.[1] ?? 0;
           return ay - by;
        });

        elements.forEach(el => {
          if (el.type === 'table') {
             if (el.html) {
               bodyParts.push(`<div class="table-block">${el.html}</div>`);
             }
          } else {
            const lineHtml = renderLineHtml(el);
            if (lineHtml) {
              const align = getLineAlign(el);
              const xPos = el.x || el.meta?.position?.x || 0;
              let styleAttr = '';
              const styles = [];
              
              if (align && align !== 'left') {
                styles.push(`text-align:${align}`);
              }
              if (align === 'left' && xPos > 10) {
                 const indent = Math.min(xPos, 400); 
                 styles.push(`margin-left:${indent}px`);
              }

              if (styles.length) {
                styleAttr = ` style="${styles.join(';')}"`;
              }

              bodyParts.push(`<p${styleAttr}>${lineHtml}</p>`);
            }
          }
        });
        
        if (Array.isArray(page.images) && page.images.length) {
           const validImages = page.images.filter(img => !img._removed);
           if (validImages.length) {
             bodyParts.push('<br/>');
             validImages.forEach((img, idx) => {
                const key = img.id || img.base64?.slice(0, 16) || `${page.number}-${idx}`;
                if (disableDescriptions) {
                  if (img.base64 && !img._replaceWithDescription) {
                    bodyParts.push(
                      `<p><img alt="Bild" style="max-width:100%;height:auto;" src="data:image/jpeg;base64,${img.base64}"/></p>`
                    );
                  } else if (img._description) {
                     bodyParts.push(`<p><strong>Bild:</strong> ${escapeHtml(img._description)}</p>`);
                  }
                } else {
                   const desc = (img._replaceWithDescription && img._description) || descriptions[key] || 'Bildbeschreibung nicht verfügbar.';
                   bodyParts.push(`<p><strong>Bild:</strong> ${escapeHtml(desc)}</p>`);
                }
             });
           }
        }
        bodyParts.push('<hr />');
      });
    } else {
      // Fallback to Markdown HTML
      if (htmlInput) {
        bodyParts.push(htmlInput);
      }
      const allImages = [];
      pages.forEach(p => {
        if (Array.isArray(p.images)) allImages.push(...p.images);
      });
      const validImages = allImages.filter(img => !img._removed);
      if (validImages.length) {
         bodyParts.push('<hr/><h2>Bilder</h2>');
         validImages.forEach((img, idx) => {
             // ... image rendering fallback ...
            const key = img.id || img.base64?.slice(0, 16);
            if (disableDescriptions) {
              if (img.base64 && !img._replaceWithDescription) {
                bodyParts.push(
                  `<p><img alt="Bild" style="max-width:100%;height:auto;" src="data:image/jpeg;base64,${img.base64}"/></p>`
                );
              } else if (img._description) {
                 bodyParts.push(`<p><strong>Bild:</strong> ${escapeHtml(img._description)}</p>`);
              }
            } else {
               const desc = (img._replaceWithDescription && img._description) || descriptions[key] || 'Bildbeschreibung nicht verfügbar.';
               bodyParts.push(`<p><strong>Bild:</strong> ${escapeHtml(desc)}</p>`);
            }
         });
      }
    }

    const htmlDoc = `
      <!DOCTYPE html>
      <html lang="de">
        <head>
          <meta charset="UTF-8" />
          <style>
            body { font-family: "Segoe UI", Arial, sans-serif; line-height: 1.5; color: #111; }
            h1, h2, h3, h4 { color: #0f172a; margin-top: 1.2em; margin-bottom: 0.6em; }
            hr { border: 0; border-top: 1px solid #e2e8f0; margin: 1.5rem 0; }
            p { margin: 0.35rem 0; }
            /* Stronger Table Styling for Word */
            table { 
              border-collapse: collapse; 
              width: 100%; 
              margin: 1rem 0; 
              border: 1px solid #000;
            }
            th, td { 
              border: 1px solid #000; 
              padding: 0.5rem; 
              text-align: left; 
              vertical-align: top;
            }
            th { 
              background-color: #f1f5f9; 
              font-weight: bold; 
            }
            ul, ol { margin-left: 1.5rem; padding-left: 0; }
            li { margin-bottom: 0.25rem; }
            img { max-width: 100%; height: auto; }
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
