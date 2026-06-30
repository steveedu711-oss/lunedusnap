'use strict';

// ── State ─────────────────────────────────────────────────────
const state = { photos: [] };

// ── DOM ───────────────────────────────────────────────────────
const photoInput  = document.getElementById('photoInput');
const photoList   = document.getElementById('photoList');
const emptyState  = document.getElementById('emptyState');
const photoCount  = document.getElementById('photoCount');
const clearBtn    = document.getElementById('clearBtn');
const genWordBtn  = document.getElementById('genWordBtn');
const overlay     = document.getElementById('overlay');
const overlayMsg  = document.getElementById('overlayMsg');
const toast       = document.getElementById('toast');
const helpModal   = document.getElementById('helpModal');

// ── Events ────────────────────────────────────────────────────
photoInput.addEventListener('change', handleFiles);
document.getElementById('cameraInput').addEventListener('change', handleFiles);
clearBtn.addEventListener('click', clearAll);
genWordBtn.addEventListener('click', () => driveUploadFlow('docx'));
document.getElementById('genPdfBtn').addEventListener('click', () => driveUploadFlow('pdf'));
document.getElementById('genImgBtn').addEventListener('click', () => driveUploadFlow('img'));
document.getElementById('helpBtn').addEventListener('click',  () => { helpModal.hidden = false; });
document.getElementById('closeHelp').addEventListener('click', () => { helpModal.hidden = true; });
helpModal.addEventListener('click', e => { if (e.target === helpModal) helpModal.hidden = true; });

// ── File handling ─────────────────────────────────────────────
async function handleFiles(e) {
  const files = Array.from(e.target.files);
  e.target.value = '';
  if (!files.length) return;
  for (const file of files) {
    const date     = await getExifDate(file);
    const thumbUrl = URL.createObjectURL(file);
    const id       = `p_${Date.now()}_${Math.random().toString(36).slice(2)}`;
    state.photos.push({ id, file, date, thumbUrl });
    renderCard({ id, file, date, thumbUrl });
  }
  updateUI();
}

// ── EXIF date ─────────────────────────────────────────────────
async function getExifDate(file) {
  try {
    const data = await exifr.parse(file, ['DateTimeOriginal']);
    if (data?.DateTimeOriginal) {
      const d = data.DateTimeOriginal instanceof Date
        ? data.DateTimeOriginal : new Date(data.DateTimeOriginal);
      return fmtDate(d);
    }
  } catch (_) {}
  return fmtDate(new Date(file.lastModified));
}

function fmtDate(d) {
  const p = n => String(n).padStart(2, '0');
  return `${d.getFullYear()}/${p(d.getMonth() + 1)}/${p(d.getDate())}`;
}

// ── Card rendering ────────────────────────────────────────────
function renderCard(photo) {
  const div = document.createElement('div');
  div.className = 'photo-card';
  div.dataset.id = photo.id;
  div.innerHTML = `
    <img class="thumb" src="${photo.thumbUrl}" alt="">
    <div class="card-body">
      <div class="fname">${esc(photo.file.name)}</div>
      <div class="inputs">
        <div class="inp-row"><span class="lbl">日期</span><input type="text" class="f-date" value="${esc(photo.date)}"></div>
        <div class="inp-row"><span class="lbl">地點</span><input type="text" class="f-loc" placeholder="填寫地點"></div>
        <div class="inp-row inp-full"><span class="lbl">說明</span><input type="text" class="f-desc" placeholder="填寫說明"></div>
      </div>
    </div>
    <button class="del-btn" aria-label="刪除">✕</button>
  `;
  div.querySelector('.del-btn').addEventListener('click', () => removePhoto(photo.id));
  photoList.appendChild(div);
}

function getCardValues(id) {
  const c = photoList.querySelector(`[data-id="${id}"]`);
  if (!c) return {};
  return {
    date:        c.querySelector('.f-date').value,
    location:    c.querySelector('.f-loc').value,
    description: c.querySelector('.f-desc').value,
  };
}

// ── Photo management ──────────────────────────────────────────
function removePhoto(id) {
  const p = state.photos.find(x => x.id === id);
  if (p) URL.revokeObjectURL(p.thumbUrl);
  state.photos = state.photos.filter(x => x.id !== id);
  photoList.querySelector(`[data-id="${id}"]`)?.remove();
  updateUI();
}

function clearAll() {
  if (!state.photos.length) return;
  if (!confirm('確定要清除所有照片？')) return;
  state.photos.forEach(p => URL.revokeObjectURL(p.thumbUrl));
  state.photos = [];
  photoList.innerHTML = '';
  updateUI();
}

function updateUI() {
  const n = state.photos.length;
  const pages = Math.ceil(n / 3) || 0;
  photoCount.textContent = n ? `${n} 張（${pages} 頁）` : '0 張';
  emptyState.hidden  = n > 0;
  clearBtn.style.display = n ? 'inline-flex' : 'none';
}

function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ── Image for Word ────────────────────────────────────────────
async function getImageForWord(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      const MAX = 1400;
      let w = img.naturalWidth, h = img.naturalHeight;
      if (w > MAX || h > MAX) {
        const s = MAX / Math.max(w, h);
        w = Math.round(w * s); h = Math.round(h * s);
      }
      const cv = document.createElement('canvas');
      cv.width = w; cv.height = h;
      const ctx = cv.getContext('2d');
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, w, h);
      ctx.drawImage(img, 0, 0, w, h);
      URL.revokeObjectURL(url);
      cv.toBlob(blob => {
        const fr = new FileReader();
        fr.onload = ev => resolve({ data: new Uint8Array(ev.target.result), w, h });
        fr.onerror = reject;
        fr.readAsArrayBuffer(blob);
      }, 'image/jpeg', 0.88);
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error(`圖片載入失敗：${file.name}`)); };
    img.src = url;
  });
}

// ── Word generation ───────────────────────────────────────────
// docx v7: transformation 單位為 px at 96dpi
const DISP_W     = Math.round(112 * 96 / 25.4); // 423px = 112mm
const DISP_H_MAX = Math.round(68  * 96 / 25.4); // 257px = 68mm
const mmToDxa    = mm => Math.round(mm * 1440 / 25.4);

async function buildDoc(school, title) {
  const {
    Document, Table, TableRow, TableCell,
    Paragraph, TextRun, ImageRun,
    AlignmentType, VerticalAlign, WidthType, HeightRule,
    convertMillimetersToTwip,
  } = docx;

  const pages = [];
  for (let i = 0; i < state.photos.length; i += 3) pages.push(state.photos.slice(i, i + 3));

  const children = [];

  for (let pi = 0; pi < pages.length; pi++) {
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      pageBreakBefore: pi > 0,
      spacing: { before: 0, after: 80 },
      children: [new TextRun({ text: school, size: 32, bold: true })],
    }));
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
      children: [new TextRun({ text: title, size: 26 })],
    }));

    const rows = [];
    for (let ri = 0; ri < pages[pi].length; ri++) {
      const ph = pages[pi][ri];
      const fields = getCardValues(ph.id);
      overlayMsg.textContent = `處理照片 ${pi * 3 + ri + 1} / ${state.photos.length}...`;

      let photoChild;
      try {
        const img = await getImageForWord(ph.file);
        let dw = DISP_W, dh = Math.round(DISP_W * img.h / img.w);
        if (dh > DISP_H_MAX) { dh = DISP_H_MAX; dw = Math.round(dh * img.w / img.h); }
        photoChild = new ImageRun({ data: img.data, transformation: { width: dw, height: dh } });
      } catch (_) {
        photoChild = new TextRun({ text: `[${ph.file.name}]`, size: 18 });
      }

      rows.push(new TableRow({
        height: { value: mmToDxa(74), rule: HeightRule.EXACT },
        children: [
          new TableCell({
            width: { size: mmToDxa(120), type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 0 },
              children: [photoChild],
            })],
          }),
          new TableCell({
            width: { size: mmToDxa(52), type: WidthType.DXA },
            verticalAlign: VerticalAlign.CENTER,
            children: ['日期', '地點', '說明'].map((label, i) => new Paragraph({
              spacing: { before: 80, after: 80 },
              children: [
                new TextRun({ text: `${label}：`, bold: true, size: 20 }),
                new TextRun({ text: [fields.date, fields.location, fields.description][i] || '', size: 20 }),
              ],
            })),
          }),
        ],
      }));
    }

    children.push(new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows,
    }));
  }

  return new Document({
    sections: [{
      properties: {
        page: {
          size: { width: convertMillimetersToTwip(210), height: convertMillimetersToTwip(297) },
          margin: {
            top:    convertMillimetersToTwip(14),
            bottom: convertMillimetersToTwip(14),
            left:   convertMillimetersToTwip(18),
            right:  convertMillimetersToTwip(18),
          },
        },
      },
      children,
    }],
  });
}

async function generateWord() {
  if (!state.photos.length) { showToast('請先新增照片'); return; }
  const school = document.getElementById('schoolName').value.trim() || '學校名稱';
  const title  = document.getElementById('formTitle').value.trim()  || '照片紀錄';
  showOverlay('準備中...');
  try {
    const doc  = await buildDoc(school, title);
    const blob = await docx.Packer.toBlob(doc);
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = `${title}.docx`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    setTimeout(() => URL.revokeObjectURL(link.href), 8000);
    showToast('Word 文件已下載！', 'ok');
  } catch (err) {
    console.error(err);
    showToast('生成失敗：' + err.message, 'err');
  } finally {
    hideOverlay();
  }
}

// ── Canvas page rendering (for PDF & Image export) ────────────
async function getResizedImage(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      const MAX = 1400;
      let w = img.naturalWidth, h = img.naturalHeight;
      if (w > MAX || h > MAX) { const s = MAX / Math.max(w, h); w = Math.round(w * s); h = Math.round(h * s); }
      const cv = document.createElement('canvas');
      cv.width = w; cv.height = h;
      const ctx = cv.getContext('2d');
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, w, h);
      ctx.drawImage(img, 0, 0, w, h);
      URL.revokeObjectURL(url);
      resolve({ canvas: cv, w, h });
    };
    img.onerror = () => { URL.revokeObjectURL(url); reject(new Error(`載入失敗：${file.name}`)); };
    img.src = url;
  });
}

function wrapText(ctx, text, x, y, maxW, lineH) {
  if (!text) return;
  let line = '';
  for (const ch of text) {
    const test = line + ch;
    if (ctx.measureText(test).width > maxW && line) { ctx.fillText(line, x, y); y += lineH; line = ch; }
    else line = test;
  }
  if (line) ctx.fillText(line, x, y);
}

async function renderPageToCanvas(school, title, pagePhotos) {
  const S = 2;
  const px = mm => Math.round(mm * 96 / 25.4) * S;
  const cv = document.createElement('canvas');
  cv.width = Math.round(210 * 96 / 25.4) * S;
  cv.height = Math.round(297 * 96 / 25.4) * S;
  const ctx = cv.getContext('2d');
  const F = '"Noto Sans TC","Microsoft JhengHei",sans-serif';

  ctx.fillStyle = '#ffffff';
  ctx.fillRect(0, 0, cv.width, cv.height);

  let y = px(14);
  ctx.textAlign = 'center';
  ctx.fillStyle = '#000000';
  ctx.font = `bold ${px(5.6)}px ${F}`;
  ctx.fillText(school, cv.width / 2, y + px(5.5));
  y += px(9);
  ctx.font = `${px(4.2)}px ${F}`;
  ctx.fillText(title, cv.width / 2, y + px(4));
  y += px(8);

  const mL = px(18), photoW = px(120), infoW = px(52), gap = px(2), rowH = px(74);
  const txtX = mL + photoW + gap + px(3);
  const fs = px(3.5);

  for (const ph of pagePhotos) {
    const fields = getCardValues(ph.id);
    ctx.strokeStyle = '#bbbbbb'; ctx.lineWidth = S;
    ctx.strokeRect(mL, y, photoW, rowH);
    ctx.strokeRect(mL + photoW + gap, y, infoW, rowH);

    try {
      const { canvas: ic, w: iw, h: ih } = await getResizedImage(ph.file);
      const sc = Math.min((photoW - px(4)) / iw, (rowH - px(4)) / ih);
      const dw = iw * sc, dh = ih * sc;
      ctx.drawImage(ic, mL + (photoW - dw) / 2, y + (rowH - dh) / 2, dw, dh);
    } catch(_) {}

    ctx.textAlign = 'left';
    ctx.fillStyle = '#111111';
    let ty = y + px(12);
    for (const [label, val] of [['日期', fields.date], ['地點', fields.location], ['說明', fields.description]]) {
      ctx.font = `bold ${fs}px ${F}`;
      const lw = ctx.measureText(`${label}：`).width;
      ctx.fillText(`${label}：`, txtX, ty);
      ctx.font = `${fs}px ${F}`;
      wrapText(ctx, val || '', txtX + lw + px(0.5), ty, infoW - px(6) - lw, px(5.5));
      ty += px(10);
    }
    y += rowH + px(2);
  }
  return cv;
}

// ── PDF generation ─────────────────────────────────────────────
async function generatePDF() {
  if (!state.photos.length) { showToast('請先新增照片'); return; }
  if (!window.jspdf) { showToast('PDF 模組載入中，請稍後重試'); return; }
  const school = document.getElementById('schoolName').value.trim() || '學校名稱';
  const title  = document.getElementById('formTitle').value.trim() || '照片紀錄';
  showOverlay('準備 PDF...');
  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait' });
    const pages = [];
    for (let i = 0; i < state.photos.length; i += 3) pages.push(state.photos.slice(i, i + 3));
    for (let pi = 0; pi < pages.length; pi++) {
      if (pi > 0) doc.addPage();
      overlayMsg.textContent = `PDF 第 ${pi + 1} / ${pages.length} 頁...`;
      const pageCanvas = await renderPageToCanvas(school, title, pages[pi]);
      doc.addImage(pageCanvas.toDataURL('image/jpeg', 0.9), 'JPEG', 0, 0, 210, 297);
    }
    doc.save(`${title}.pdf`);
    showToast('PDF 已下載！', 'ok');
  } catch(err) {
    console.error(err);
    showToast('PDF 失敗：' + err.message, 'err');
  } finally {
    hideOverlay();
  }
}

// ── Image download ──────────────────────────────────────────────
async function downloadImages() {
  if (!state.photos.length) { showToast('請先新增照片'); return; }
  const school = document.getElementById('schoolName').value.trim() || '學校名稱';
  const title  = document.getElementById('formTitle').value.trim() || '照片紀錄';
  showOverlay('準備圖片...');
  try {
    const pages = [];
    for (let i = 0; i < state.photos.length; i += 3) pages.push(state.photos.slice(i, i + 3));

    const download = (blob, filename) => {
      const a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(a.href), 8000);
    };

    if (pages.length === 1) {
      overlayMsg.textContent = '產生圖片...';
      const pageCanvas = await renderPageToCanvas(school, title, pages[0]);
      const blob = await new Promise(res => pageCanvas.toBlob(res, 'image/png'));
      download(blob, `${title}.png`);
      showToast('圖片已下載！', 'ok');
    } else {
      if (!window.JSZip) { showToast('ZIP 模組未載入'); return; }
      const zip = new JSZip();
      for (let pi = 0; pi < pages.length; pi++) {
        overlayMsg.textContent = `圖片 ${pi + 1} / ${pages.length} 頁...`;
        const pageCanvas = await renderPageToCanvas(school, title, pages[pi]);
        const blob = await new Promise(res => pageCanvas.toBlob(res, 'image/png'));
        zip.file(`${title}_p${String(pi + 1).padStart(2, '0')}.png`, blob);
      }
      const zipBlob = await zip.generateAsync({ type: 'blob' });
      download(zipBlob, `${title}_圖片.zip`);
      showToast(`${pages.length} 頁圖片已下載！`, 'ok');
    }
  } catch(err) {
    console.error(err);
    showToast('圖片下載失敗：' + err.message, 'err');
  } finally {
    hideOverlay();
  }
}

// ── Google Drive upload ────────────────────────────────────────
const DRIVE_CLIENT_ID = '1001894711281-oq75uqan9c8dev2miugon3s1cgt8i97g.apps.googleusercontent.com';
const DRIVE_SCOPE = 'https://www.googleapis.com/auth/drive.file';

let driveTokenClient = null;
let driveAccessToken = null;
let pendingFormat = null;

function initDriveClient() {
  driveTokenClient = google.accounts.oauth2.initTokenClient({
    client_id: DRIVE_CLIENT_ID,
    scope: DRIVE_SCOPE,
    callback: async (resp) => {
      if (resp.error) { showToast('授權失敗：' + resp.error, 'err'); hideOverlay(); return; }
      driveAccessToken = resp.access_token;
      await performDriveUpload(pendingFormat);
    },
  });
}

async function driveUploadFlow(format) {
  if (!state.photos.length) { showToast('請先新增照片'); return; }
  if (!window.google?.accounts?.oauth2) { showToast('Google 模組未載入，請重新整理頁面'); return; }
  pendingFormat = format;
  showOverlay('請完成 Google 授權...');
  try {
    if (!driveTokenClient) initDriveClient();
    driveTokenClient.requestAccessToken();
  } catch (err) {
    showToast('啟動授權失敗：' + err.message, 'err');
    hideOverlay();
  }
}

async function performDriveUpload(format) {
  const school     = document.getElementById('schoolName').value.trim() || '學校名稱';
  const title      = document.getElementById('formTitle').value.trim() || '照片紀錄';
  const folderName = document.getElementById('driveFolder').value.trim() || 'LunEduSnap';
  try {
    overlayMsg.textContent = `取得資料夾「${folderName}」...`;
    const folderId = await getOrCreateDriveFolder(folderName);

    if (format === 'pdf') {
      if (!window.jspdf) { showToast('PDF 模組未載入'); hideOverlay(); return; }
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({ unit: 'mm', format: 'a4', orientation: 'portrait' });
      const pages = [];
      for (let i = 0; i < state.photos.length; i += 3) pages.push(state.photos.slice(i, i + 3));
      for (let pi = 0; pi < pages.length; pi++) {
        if (pi > 0) doc.addPage();
        overlayMsg.textContent = `PDF 第 ${pi + 1} / ${pages.length} 頁...`;
        const pageCanvas = await renderPageToCanvas(school, title, pages[pi]);
        doc.addImage(pageCanvas.toDataURL('image/jpeg', 0.9), 'JPEG', 0, 0, 210, 297);
      }
      overlayMsg.textContent = '上傳 PDF...';
      await uploadFileToDrive(doc.output('blob'), `${title}.pdf`, 'application/pdf', folderId);
      showToast(`PDF 已上傳到「${folderName}」`, 'ok');

    } else if (format === 'docx') {
      overlayMsg.textContent = '產生文件...';
      const wordDoc = await buildDoc(school, title);
      const docxBlob = await docx.Packer.toBlob(wordDoc);
      overlayMsg.textContent = '上傳文件...';
      await uploadFileToDrive(docxBlob, `${title}.docx`, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', folderId);
      showToast(`文件已上傳到「${folderName}」`, 'ok');

    } else if (format === 'img') {
      const pages = [];
      for (let i = 0; i < state.photos.length; i += 3) pages.push(state.photos.slice(i, i + 3));
      for (let pi = 0; pi < pages.length; pi++) {
        overlayMsg.textContent = `上傳圖片 ${pi + 1} / ${pages.length}...`;
        const pageCanvas = await renderPageToCanvas(school, title, pages[pi]);
        const blob = await new Promise(res => pageCanvas.toBlob(res, 'image/png'));
        const fileName = pages.length === 1 ? `${title}.png` : `${title}_p${String(pi + 1).padStart(2, '0')}.png`;
        await uploadFileToDrive(blob, fileName, 'image/png', folderId);
      }
      showToast(`圖片已上傳到「${folderName}」`, 'ok');
    }
  } catch (err) {
    console.error(err);
    showToast('上傳失敗：' + err.message, 'err');
  } finally {
    hideOverlay();
  }
}

async function getOrCreateDriveFolder(name) {
  const q = encodeURIComponent(`name='${name}' and mimeType='application/vnd.google-apps.folder' and trashed=false`);
  const res = await fetch(`https://www.googleapis.com/drive/v3/files?q=${q}&fields=files(id,name)`, {
    headers: { Authorization: `Bearer ${driveAccessToken}` },
  });
  if (!res.ok) { const t = await res.text(); throw new Error(`查詢資料夾失敗 ${res.status}: ${t}`); }
  const data = await res.json();
  if (data.files?.length) return data.files[0].id;

  const createRes = await fetch('https://www.googleapis.com/drive/v3/files', {
    method: 'POST',
    headers: { Authorization: `Bearer ${driveAccessToken}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ name, mimeType: 'application/vnd.google-apps.folder' }),
  });
  if (!createRes.ok) { const t = await createRes.text(); throw new Error(`建立資料夾失敗 ${createRes.status}: ${t}`); }
  const folder = await createRes.json();
  return folder.id;
}

async function uploadFileToDrive(blob, fileName, mimeType, folderId) {
  const metadata = { name: fileName, mimeType, ...(folderId ? { parents: [folderId] } : {}) };
  const form = new FormData();
  form.append('metadata', new Blob([JSON.stringify(metadata)], { type: 'application/json' }));
  form.append('file', blob);
  const res = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart', {
    method: 'POST',
    headers: { Authorization: `Bearer ${driveAccessToken}` },
    body: form,
  });
  if (!res.ok) { const t = await res.text(); throw new Error(`HTTP ${res.status}: ${t}`); }
  return res.json();
}

// ── UI helpers ────────────────────────────────────────────────
function showOverlay(msg) { overlayMsg.textContent = msg; overlay.hidden = false; }
function hideOverlay()    { overlay.hidden = true; }

let toastTimer;
function showToast(msg, type = '') {
  clearTimeout(toastTimer);
  toast.textContent = msg;
  toast.className = 'toast' + (type ? ' ' + type : '');
  toast.hidden = false;
  toastTimer = setTimeout(() => { toast.hidden = true; }, 3500);
}
