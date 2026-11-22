(() => {
  'use strict';

  const $ = (id) => document.getElementById(id);
  const fileInput = $('file');
  const drop = $('drop');
  const preview = $('preview');
  const hint = $('hint');
  const out = $('out');
  const statusEl = $('status');
  const resultEl = $('result');
  const actions = $('actions');
  const copyBtn = $('copyBtn');
  const openBtn = $('openBtn');
  const shareBtn = $('shareBtn');
  const pasteBtn = $('pasteBtn');
  const historyEl = $('history');
  const emptyHist = $('emptyHist');
  const clearHist = $('clearHist');
  const toastEl = $('toast');

  const HIST_KEY = 'leitor-qr-historico';
  const MAX_HIST = 30;
  let currentText = '';
  let previewUrl = null;

  // ---------- Decodificação ----------

  function loadImage(blob) {
    // createImageBitmap respeita a orientação EXIF em Safari moderno; fallback para <img>.
    if ('createImageBitmap' in window) {
      return createImageBitmap(blob, { imageOrientation: 'from-image' })
        .catch(() => loadViaImg(blob));
    }
    return loadViaImg(blob);
  }

  function loadViaImg(blob) {
    return new Promise((resolve, reject) => {
      const url = URL.createObjectURL(blob);
      const img = new Image();
      img.onload = () => { URL.revokeObjectURL(url); resolve(img); };
      img.onerror = () => { URL.revokeObjectURL(url); reject(new Error('Formato de imagem não suportado.')); };
      img.src = url;
    });
  }

  function scanAt(source, w, h, maxSide) {
    const scale = Math.min(1, maxSide / Math.max(w, h));
    const cw = Math.max(1, Math.round(w * scale));
    const ch = Math.max(1, Math.round(h * scale));
    const canvas = document.createElement('canvas');
    canvas.width = cw;
    canvas.height = ch;
    const ctx = canvas.getContext('2d', { willReadFrequently: true });
    ctx.fillStyle = '#fff';
    ctx.fillRect(0, 0, cw, ch); // fundo branco para PNGs transparentes
    ctx.drawImage(source, 0, 0, cw, ch);
    const data = ctx.getImageData(0, 0, cw, ch);
    return jsQR(data.data, cw, ch, { inversionAttempts: 'attemptBoth' });
  }

  async function decode(blob) {
    const img = await loadImage(blob);
    const w = img.width || img.naturalWidth;
    const h = img.height || img.naturalHeight;
    // Tenta em várias resoluções: fotos grandes do iPhone (12–48 MP) excedem o limite de
    // canvas do Safari e QR pequenos dentro de fotos precisam de mais resolução.
    const sizes = [1600, 2400, 1000, 600, 3200];
    for (const s of sizes) {
      if (s > Math.max(w, h) * 1.01 && s !== sizes[0]) continue;
      const code = scanAt(img, w, h, s);
      if (code && code.data !== undefined) return code;
      await new Promise((r) => setTimeout(r, 0)); // não travar a interface
    }
    return null;
  }

  async function handleBlob(blob) {
    if (!blob || !blob.type.startsWith('image/')) {
      showError('Isso não parece ser uma imagem.');
      return;
    }
    if (previewUrl) URL.revokeObjectURL(previewUrl);
    previewUrl = URL.createObjectURL(blob);
    preview.src = previewUrl;
    preview.classList.remove('hidden');
    hint.classList.add('hidden');

    out.classList.remove('hidden');
    statusEl.className = 'status';
    statusEl.textContent = 'Lendo…';
    resultEl.classList.add('hidden');
    actions.classList.add('hidden');

    try {
      const code = await decode(blob);
      if (!code) {
        showError('Nenhum QR code encontrado. Tente uma imagem mais nítida ou recortada no código.');
        return;
      }
      const text = code.data || '(QR code sem texto legível)';
      showResult(text);
      addHistory(text);
    } catch (e) {
      showError(e && e.message ? e.message : 'Não foi possível ler a imagem.');
    }
  }

  // ---------- Resultado ----------

  function asUrl(text) {
    try {
      const u = new URL(text.trim());
      return ['http:', 'https:', 'mailto:', 'tel:', 'sms:'].includes(u.protocol) ? u.href : null;
    } catch { return null; }
  }

  function showResult(text) {
    currentText = text;
    statusEl.className = 'status ok';
    statusEl.textContent = 'QR code lido:';
    resultEl.textContent = text;
    resultEl.classList.remove('hidden');
    actions.classList.remove('hidden');
    const url = asUrl(text);
    if (url) { openBtn.href = url; openBtn.classList.remove('hidden'); }
    else { openBtn.classList.add('hidden'); openBtn.removeAttribute('href'); }
    shareBtn.classList.toggle('hidden', !navigator.share);
  }

  function showError(msg) {
    out.classList.remove('hidden');
    statusEl.className = 'status err';
    statusEl.textContent = msg;
    resultEl.classList.add('hidden');
    actions.classList.add('hidden');
  }

  async function copy(text) {
    try {
      await navigator.clipboard.writeText(text);
    } catch {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      document.execCommand('copy');
      ta.remove();
    }
    toast('Copiado');
  }

  let toastTimer;
  function toast(msg) {
    toastEl.textContent = msg;
    toastEl.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => toastEl.classList.remove('show'), 1400);
  }

  // ---------- Histórico ----------

  function readHist() {
    try { return JSON.parse(localStorage.getItem(HIST_KEY)) || []; } catch { return []; }
  }
  function writeHist(list) {
    try { localStorage.setItem(HIST_KEY, JSON.stringify(list)); } catch { /* armazenamento indisponível */ }
  }
  function addHistory(text) {
    const list = readHist().filter((i) => i.text !== text);
    list.unshift({ text, at: Date.now() });
    writeHist(list.slice(0, MAX_HIST));
    renderHist();
  }
  function renderHist() {
    const list = readHist();
    historyEl.textContent = '';
    for (const item of list) {
      const li = document.createElement('li');
      const span = document.createElement('span');
      span.className = 'txt';
      span.textContent = item.text;
      span.title = new Date(item.at).toLocaleString('pt-BR');
      span.addEventListener('click', () => { out.classList.remove('hidden'); showResult(item.text); window.scrollTo({ top: 0, behavior: 'smooth' }); });
      const b = document.createElement('button');
      b.type = 'button';
      b.className = 'secondary';
      b.textContent = 'Copiar';
      b.addEventListener('click', () => copy(item.text));
      li.append(span, b);
      historyEl.appendChild(li);
    }
    emptyHist.classList.toggle('hidden', list.length > 0);
    clearHist.classList.toggle('hidden', list.length === 0);
  }

  // ---------- Eventos ----------

  fileInput.addEventListener('change', () => {
    const f = fileInput.files && fileInput.files[0];
    if (f) handleBlob(f);
    fileInput.value = ''; // permite escolher a mesma imagem de novo
  });

  copyBtn.addEventListener('click', () => copy(currentText));
  shareBtn.addEventListener('click', () => {
    navigator.share({ text: currentText }).catch(() => {});
  });
  clearHist.addEventListener('click', () => { writeHist([]); renderHist(); });

  pasteBtn.addEventListener('click', async () => {
    if (!navigator.clipboard || !navigator.clipboard.read) {
      toast('Use ⌘V ou “Escolher imagem”');
      return;
    }
    try {
      const items = await navigator.clipboard.read();
      for (const item of items) {
        const type = item.types.find((t) => t.startsWith('image/'));
        if (type) { handleBlob(await item.getType(type)); return; }
      }
      toast('Não há imagem na área de transferência');
    } catch {
      toast('Permissão para colar negada');
    }
  });

  document.addEventListener('paste', (e) => {
    const items = e.clipboardData ? e.clipboardData.items : [];
    for (const it of items) {
      if (it.type.startsWith('image/')) { handleBlob(it.getAsFile()); e.preventDefault(); return; }
    }
  });

  ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', (e) => {
    const f = e.dataTransfer && e.dataTransfer.files[0];
    if (f) handleBlob(f);
  });
  drop.addEventListener('click', () => fileInput.click());

  renderHist();

  // Funciona offline depois da primeira visita.
  if ('serviceWorker' in navigator && location.protocol === 'https:') {
    navigator.serviceWorker.register('sw.js').catch(() => {});
  }
})();
