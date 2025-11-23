const uploadForm = document.getElementById('upload-form');
const fileInput = document.getElementById('pdf-input');
const statusEl = document.getElementById('status');
const dropArea = document.getElementById('drop-area');
const previewEl = document.getElementById('preview');
const textOutput = document.getElementById('text-output');
const copyHtmlBtn = document.getElementById('copy-html');
const copyTextBtn = document.getElementById('copy-text');
const exportWordBtn = document.getElementById('export-word');
const toggleImageDesc = document.getElementById('toggle-image-desc');
const loadingIndicator = document.getElementById('loading-indicator');
const audioEl = document.getElementById('convert-audio');
const apiKeyInput = document.getElementById('api-key');
const saveApiKeyBtn = document.getElementById('save-api-key');
const convertBtn = document.getElementById('convert-btn');

const API_KEY_STORAGE_KEY = 'mistral_api_key';

let audioInterval = null;
let lastResult = null;

const setStatus = (message, tone = 'neutral') => {
  statusEl.textContent = message;
  statusEl.dataset.tone = tone;
};

const startLoadingFeedback = () => {
  if (loadingIndicator) {
    loadingIndicator.hidden = false;
  }
  const playPulse = async () => {
    if (!audioEl) return;
    try {
      audioEl.currentTime = 0;
      await audioEl.play();
    } catch (err) {
      // autoplay kann blockiert sein; still akzeptieren
    }
  };
  clearInterval(audioInterval);
  playPulse();
  audioInterval = setInterval(playPulse, 2000);
};

const stopLoadingFeedback = () => {
  if (loadingIndicator) {
    loadingIndicator.hidden = true;
  }
  if (audioInterval) {
    clearInterval(audioInterval);
    audioInterval = null;
  }
  if (audioEl) {
    audioEl.pause();
    audioEl.currentTime = 0;
  }
};

const getStoredApiKey = () => {
  try {
    return localStorage.getItem(API_KEY_STORAGE_KEY) || '';
  } catch (err) {
    return '';
  }
};

const hasSavedApiKey = () => Boolean(getStoredApiKey());

const persistApiKey = (value) => {
  try {
    localStorage.setItem(API_KEY_STORAGE_KEY, value);
  } catch (err) {
    // ignore storage errors
  }
  updateConvertAccess();
};

const ensureApiKey = () => {
  const key = getStoredApiKey();
  if (!key) {
    setStatus('Bitte API-Key speichern, um zu konvertieren.', 'error');
    apiKeyInput?.focus();
    return null;
  }
  return key;
};

const updateConvertAccess = () => {
  const savedKey = getStoredApiKey();
  if (apiKeyInput && savedKey && !apiKeyInput.value) {
    apiKeyInput.value = savedKey;
  }
  if (convertBtn) {
    convertBtn.disabled = !savedKey;
  }
  if (dropArea) {
    dropArea.classList.toggle('is-locked', !savedKey);
  }
};

const requestImageDescription = async (img) => {
  const apiKey = ensureApiKey();
  if (!apiKey) {
    img._description = 'API-Key fehlt. Bitte zuerst Schlüssel speichern.';
    img._replaceWithDescription = true;
    renderPreview(lastResult.pages);
    return;
  }
  if (!img?.base64) {
    img._description = 'Keine Bilddaten verfügbar.';
    renderPreview(lastResult.pages);
    return;
  }
  try {
    const resp = await fetch('/api/describe-image', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-mistral-api-key': apiKey,
      },
      body: JSON.stringify({ base64: img.base64 }),
    });
    if (!resp.ok) {
      throw new Error('Beschreibung fehlgeschlagen.');
    }
    const data = await resp.json();
    img._description = data?.description || 'Keine Beschreibung erhalten.';
    img._replaceWithDescription = true;
  } catch (err) {
    img._description = err?.message || 'Beschreibung fehlgeschlagen.';
    img._replaceWithDescription = true;
  } finally {
    renderPreview(lastResult.pages);
  }
};

const buildPlainText = (pages) =>
  pages
    .map((page) =>
      page.lines.map((line) => line.segments.map((segment) => segment.text).join('')).join('\n')
    )
    .join('\n\n');

const renderPreview = (pages = []) => {
  previewEl.innerHTML = '';
  if (!pages.length) {
    previewEl.innerHTML = '<p class="empty">Noch keine Extraktion.</p>';
    return;
  }

  pages.forEach((page) => {
    const article = document.createElement('article');
    const title = document.createElement('h3');
    title.textContent = `Seite ${page.number}`;
    article.appendChild(title);

    (page.lines || []).forEach((line) => {
      const lineEl = document.createElement('div');
      lineEl.className = 'line';
      line.segments?.forEach((segment) => {
        const span = document.createElement('span');
        Object.entries(segment.style).forEach(([key, value]) => {
          span.style[key] = value;
        });
        span.textContent = segment.text;
        lineEl.appendChild(span);
      });
      article.appendChild(lineEl);
    });

    if (Array.isArray(page.images) && page.images.length) {
      const imageGrid = document.createElement('div');
      imageGrid.className = 'image-grid';
      page.images.forEach((img, idx) => {
        if (img._removed) return;
        const card = document.createElement('div');
        card.className = 'image-card';
        if (!img._replaceWithDescription) {
          const image = document.createElement('img');
          image.src = `data:image/jpeg;base64,${img.base64}`;
          image.alt = `Bild ${idx + 1}`;
          card.appendChild(image);
        }
        if (img._description) {
          const desc = document.createElement('div');
          desc.className = 'desc';
          desc.textContent = img._description;
          card.appendChild(desc);
        }

        const actions = document.createElement('div');
        actions.className = 'image-actions';
        const removeBtn = document.createElement('button');
        removeBtn.textContent = 'Nicht exportieren';
        removeBtn.addEventListener('click', () => {
          img._removed = true;
          renderPreview(lastResult.pages);
        });
        const describeBtn = document.createElement('button');
        describeBtn.textContent = 'Bildbeschreibung einfügen';
        describeBtn.addEventListener('click', () => {
          img._description = 'Beschreibung wird geladen...';
          img._replaceWithDescription = true;
          renderPreview(lastResult.pages);
          requestImageDescription(img);
        });
        actions.appendChild(removeBtn);
        actions.appendChild(describeBtn);
        card.appendChild(actions);

        imageGrid.appendChild(card);
      });
      article.appendChild(imageGrid);
    }

    previewEl.appendChild(article);
  });
};

const htmlToPlainText = (html) => {
  if (!html) return '';
  const temp = document.createElement('div');
  temp.innerHTML = html;
  return temp.textContent || temp.innerText || '';
};

const handleResult = async (response) => {
  if (!response.ok) {
    const payload = await response.json().catch(() => ({}));
    throw new Error(payload.error || 'PDF konnte nicht entziffert werden.');
  }
  const data = await response.json();
  const hasHtml = typeof data.html === 'string' && data.html.trim();
  const hasPages = Array.isArray(data.pages) && data.pages.length;
  lastResult = {
    ...data,
    html: hasHtml ? data.html : null,
    pages: hasPages ? data.pages : [],
  };

  if (hasHtml) {
    previewEl.innerHTML = data.html;
  } else if (hasPages) {
    renderPreview(data.pages);
  } else {
    renderPreview([]);
  }

  if (hasHtml) {
    textOutput.value = htmlToPlainText(data.html);
  } else if (hasPages) {
    textOutput.value = buildPlainText(data.pages);
  } else {
    textOutput.value = '';
  }
  const sourceLabel = data.source === 'mistral' ? 'Mistral OCR' : 'PDF.js';
  setStatus(`Extrahiert ${data.pageCount} Seiten (${sourceLabel})`, 'success');
};

const uploadPdf = async () => {
  if (!fileInput.files.length) {
    setStatus('Bitte wähle eine PDF-Datei aus.', 'error');
    return;
  }
  const apiKey = ensureApiKey();
  if (!apiKey) return;

  const formData = new FormData();
  formData.append('pdf', fileInput.files[0]);
  setStatus('Extrahiere Text...');
  startLoadingFeedback();

  try {
    const response = await fetch('/api/extract', {
      method: 'POST',
      headers: {
        'x-mistral-api-key': apiKey,
      },
      body: formData,
    });
    await handleResult(response);
    stopLoadingFeedback();
  } catch (error) {
    setStatus(error.message || 'Beim Extrahieren ist ein Fehler aufgetreten.', 'error');
    renderPreview([]);
    textOutput.value = '';
    stopLoadingFeedback();
  }
};

const exportWord = async () => {
  if (!lastResult?.pages?.length && !lastResult?.html) {
    setStatus('Bitte zuerst eine PDF extrahieren.', 'error');
    return;
  }
  const apiKey = ensureApiKey();
  if (!apiKey) return;

  const payloadPages = lastResult.pages.map((page) => ({
    ...page,
    images: (page.images || []).map((img) => ({
      ...img,
      _removed: img._removed,
      _description: img._description,
    })),
  }));

  try {
    setStatus('Exportiere Word...', 'neutral');
    const response = await fetch('/api/export-docx', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-mistral-api-key': apiKey,
      },
      body: JSON.stringify({
        pages: payloadPages,
        html: lastResult.html,
        disableDescriptions: toggleImageDesc?.checked || false,
      }),
    });
    if (!response.ok) {
      const payload = await response.json().catch(() => ({}));
      throw new Error(payload.error || 'Export fehlgeschlagen.');
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'export.doc';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    setStatus('Word-Export erstellt.', 'success');
  } catch (error) {
    setStatus(error.message || 'Beim Export ist ein Fehler aufgetreten.', 'error');
  }
};

uploadForm.addEventListener('submit', (event) => {
  event.preventDefault();
  uploadPdf();
});

['dragenter', 'dragover'].forEach((eventType) => {
  dropArea.addEventListener(eventType, (event) => {
    event.preventDefault();
    if (!hasSavedApiKey()) return;
    dropArea.classList.add('is-active');
  });
});

['dragleave', 'drop'].forEach((eventType) => {
  dropArea.addEventListener(eventType, (event) => {
    event.preventDefault();
    dropArea.classList.remove('is-active');

    if (eventType === 'drop' && event.dataTransfer?.files?.length) {
      if (!hasSavedApiKey()) {
        setStatus('Bitte zuerst deinen API-Key speichern.', 'error');
        apiKeyInput?.focus();
        return;
      }
      const [file] = event.dataTransfer.files;
      if (file.type !== 'application/pdf') {
        setStatus('Bitte nur PDF-Dateien ablegen.', 'error');
        return;
      }

      const dataTransfer = new DataTransfer();
      dataTransfer.items.add(file);
      fileInput.files = dataTransfer.files;
      uploadPdf();
    }
  });
});

copyHtmlBtn.addEventListener('click', async () => {
  if (!previewEl.innerHTML.trim()) {
    setStatus('Keine Vorschau zum Kopieren vorhanden.', 'error');
    return;
  }
  try {
    await navigator.clipboard.writeText(previewEl.innerHTML);
    setStatus('Formatiertes HTML kopiert.', 'success');
  } catch (error) {
    console.error('Clipboard-Fehler', error);
    setStatus('Clipboard nicht verfügbar.', 'error');
  }
});

copyTextBtn.addEventListener('click', async () => {
  if (!textOutput.value.trim()) {
    setStatus('Der reine Textbereich ist leer.', 'error');
    return;
  }
  try {
    await navigator.clipboard.writeText(textOutput.value);
    setStatus('Reiner Text kopiert.', 'success');
  } catch (error) {
    console.error('Clipboard-Fehler', error);
    setStatus('Clipboard nicht verfügbar.', 'error');
  }
});

exportWordBtn.addEventListener('click', exportWord);

if (saveApiKeyBtn) {
  saveApiKeyBtn.addEventListener('click', () => {
    const key = apiKeyInput?.value?.trim();
    if (!key) {
      setStatus('Bitte einen API-Key eingeben.', 'error');
      return;
    }
    persistApiKey(key);
    setStatus('API-Key gespeichert. Jetzt kannst du extrahieren.', 'success');
  });
}

(() => {
  updateConvertAccess();
  const stored = getStoredApiKey();
  if (stored) {
    setStatus('API-Key aus dem Browser geladen. Bereit für deine PDF.', 'success');
    return;
  }
  setStatus('Bitte zuerst deinen API-Key speichern, um zu konvertieren.', 'neutral');
})();
