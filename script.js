/**
 * 파일 형식 변환 (브라우저 전용)
 * - PDF 입력: 변환 없이 원본 그대로 저장
 * - 이미지 → PNG / JPG (Canvas)
 * - 엑셀 → PDF(캔버스 렌더) / DOCX (docx)
 * - HWPX → 텍스트 추출 후 PDF / DOCX (.hwp 바이너리는 미지원 안내)
 * - 저장 폴더: File System Access API + IndexedDB(선택 유지)
 */

const docxMod = await import("https://esm.sh/docx@8.5.0");

const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType } =
  docxMod;

const IMAGE_EXT = new Set(["png", "jpg", "jpeg", "gif", "webp", "bmp", "svg", "ico", "avif"]);
const HANGUL_EXT = new Set(["hwpx"]);
const HWP_BINARY_EXT = new Set(["hwp"]);
const EXCEL_EXT = new Set(["xlsx", "xls", "xlsm", "xlsb", "csv", "ods"]);

const DB_NAME = "file-converter-v1";
const DB_STORE = "settings";

/** @type {FileSystemDirectoryHandle | null} */
let outputDirHandle = null;

const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const logEl = document.getElementById("log");
const folderStatusEl = document.getElementById("folderStatus");
const btnPickFolder = document.getElementById("btnPickFolder");
const btnClearFolder = document.getElementById("btnClearFolder");

function safeFileName(name) {
  const n = String(name).replace(/[/\\]/g, "_").replace(/^\.+/, "");
  return n || "download";
}

function openSettingsDB() {
  return new Promise((resolve, reject) => {
    const r = indexedDB.open(DB_NAME, 1);
    r.onupgradeneeded = (e) => {
      const db = e.target.result;
      if (!db.objectStoreNames.contains(DB_STORE)) db.createObjectStore(DB_STORE);
    };
    r.onsuccess = () => resolve(r.result);
    r.onerror = () => reject(r.error);
  });
}

async function persistDirHandle(handle) {
  const db = await openSettingsDB();
  await new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, "readwrite");
    tx.oncomplete = () => {
      db.close();
      resolve();
    };
    tx.onerror = () => reject(tx.error);
    const store = tx.objectStore(DB_STORE);
    if (handle) store.put(handle, "outputDir");
    else store.delete("outputDir");
  });
}

async function loadPersistedDirHandle() {
  try {
    const db = await openSettingsDB();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, "readonly");
      tx.onerror = () => {
        db.close();
        reject(tx.error);
      };
      const store = tx.objectStore(DB_STORE);
      const g = store.get("outputDir");
      g.onsuccess = () => {
        db.close();
        resolve(g.result ?? null);
      };
      g.onerror = () => {
        db.close();
        resolve(null);
      };
    });
  } catch {
    return null;
  }
}

function updateFolderLabel() {
  if (!folderStatusEl) return;
  if (outputDirHandle) {
    const folderName = outputDirHandle.split(/[\/\\]/).pop();
    folderStatusEl.textContent = `저장 위치: ${folderName} — 변환·PDF 복사 결과가 여기에 자동 저장됩니다.`;
  } else {
    folderStatusEl.textContent =
      "저장 위치: 미설정 — 완료 시 브라우저 기본 다운로드 폴더로 받습니다.";
  }
}

async function pickOutputFolder() {
  try {
    const folderPath = await window.electronAPI.selectFolder();
    if (!folderPath) return;
    outputDirHandle = folderPath;
    await persistDirHandle(folderPath);
    updateFolderLabel();
    const folderName = folderPath.split(/[\\/]/).pop();
    log(`저장 폴더를 설정했습니다: ${folderName}`, "ok");
  } catch (err) {
    log(err?.message || String(err), "err");
  }
}

async function clearOutputFolder() {
  outputDirHandle = null;
  try {
    await persistDirHandle(null);
  } catch (_) {}
  updateFolderLabel();
  log("저장 폴더를 해제했습니다. 이후에는 다운로드로 저장됩니다.", "ok");
}

async function writeBlobToDirectory(folderPath, filename, blob) {
  const name = safeFileName(filename);
  const buffer = await blob.arrayBuffer();
  await window.electronAPI.saveFile(folderPath, name, buffer);
}

/**
 * 폴더가 설정되어 있으면 그곳에 저장하고, 아니면(또는 실패 시) 다운로드합니다.
 */
async function saveOutput(blob, filename) {
  const name = uniqueFileName(filename);
  if (outputDirHandle) {
    try {
      await writeBlobToDirectory(outputDirHandle, name, blob);
      return;
    } catch (err) {
      log(`폴더 저장 실패: ${err?.message}. 다운로드로 저장합니다.`, "warn");
    }
  }
  downloadBlob(blob, name);
}

async function passThroughPdf(file) {
  const buf = await file.arrayBuffer();
  const blob = new Blob([buf], { type: "application/pdf" });
  const outName = extOf(file.name) === "pdf" ? file.name : `${baseName(file.name)}.pdf`;
  await saveOutput(blob, outName);
}

function extOf(name) {
  const i = name.lastIndexOf(".");
  return i >= 0 ? name.slice(i + 1).toLowerCase() : "";
}

function baseName(name) {
  const i = name.lastIndexOf(".");
  return i >= 0 ? name.slice(0, i) : name;
}

function getSelectedFormat() {
  return document.getElementById("select-format")?.value || "pdf";
}

function log(msg, type = "info") {
  const div = document.createElement("div");
  div.className = `log-entry ${type === "ok" ? "ok" : type === "warn" ? "warn" : type === "err" ? "err" : ""}`;
  div.textContent = msg;
  logEl.appendChild(div);
  logEl.scrollTop = logEl.scrollHeight;
}

// 탭 제거됨 - 파일 타입 자동 감지

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    fileInput.click();
  }
});

["dragenter", "dragover"].forEach((ev) => {
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    dropzone.classList.add("dragover");
  });
});
["dragleave", "drop"].forEach((ev) => {
  dropzone.addEventListener(ev, (e) => {
    e.preventDefault();
    dropzone.classList.remove("dragover");
  });
});

dropzone.addEventListener("drop", (e) => {
  const files = e.dataTransfer?.files;
  if (files?.length) handleFiles(Array.from(files));
});

fileInput.addEventListener("change", () => {
  if (fileInput.files?.length) handleFiles(Array.from(fileInput.files));
  fileInput.value = "";
});

function classify(file) {
  const ext = extOf(file.name);
  if (ext === "pdf" || file.type === "application/pdf") return "pdf";
  if (file.type.startsWith("image/") || IMAGE_EXT.has(ext)) return "image";
  if (HWP_BINARY_EXT.has(ext)) return "hwp";
  if (HANGUL_EXT.has(ext)) return "hangul";
  if (EXCEL_EXT.has(ext) || file.type.includes("sheet") || file.type.includes("excel")) return "excel";
  return "unknown";
}



// 대기 중인 파일 목록
let pendingFiles = [];

// 파일명 중복 카운터 (세션 내)
const fileNameCounter = {};

function uniqueFileName(filename) {
  const base = baseName(filename);
  const ext = extOf(filename);
  const key = filename.toLowerCase();
  if (!fileNameCounter[key]) {
    fileNameCounter[key] = 1;
    return filename;
  }
  const n = fileNameCounter[key]++;
  return ext ? `${base}(${n}).${ext}` : `${base}(${n})`;
}

function handleFiles(files) {
  pendingFiles = Array.from(files);
  if (!pendingFiles.length) return;

  // 파일 타입 감지
  const kinds = new Set(pendingFiles.map(classify));
  const hasImage = kinds.has("image");
  const hasDoc = kinds.has("hangul") || kinds.has("excel");
  const hasPdf = kinds.has("pdf") && !hasImage && !hasDoc;
  const hasHwpBinary = kinds.has("hwp");

  // 변환 섹션 표시
  const section = document.getElementById("convertSection");
  const select = document.getElementById("select-format");
  const previewList = document.getElementById("filePreviewList");
  section.style.display = "";

  // 파일 목록 미리보기
  renderFilePreview();

  // 선택 가능한 형식 구성
  select.innerHTML = "";
  if (hasPdf) {
    const o = document.createElement("option");
    o.value = "passthrough"; o.textContent = "PDF 그대로 저장";
    select.appendChild(o);
  } else {
    if (hasImage) {
      [["png","PNG (이미지)"],["jpg","JPG (이미지)"]].forEach(([v,t]) => {
        const o = document.createElement("option"); o.value=v; o.textContent=t; select.appendChild(o);
      });
    }
    if (hasDoc) {
      [["pdf","PDF (문서)"],["docx","DOCX (문서)"]].forEach(([v,t]) => {
        const o = document.createElement("option"); o.value=v; o.textContent=t; select.appendChild(o);
      });
    }
  }

  // HWP 바이너리 경고
  if (hasHwpBinary) {
    const hwpFiles = pendingFiles.filter(f => classify(f) === "hwp").map(f => f.name).join(", ");
    log(`⚠️ HWP 파일은 변환할 수 없습니다: ${hwpFiles}`, "warn");
    log("한컴오피스에서 해당 파일을 열고 → 다른 이름으로 저장 → HWPX 형식으로 저장 후 다시 올려주세요.", "warn");
    // hwp만 있으면 변환 버튼 숨김
    if (!hasImage && !hasDoc && !hasPdf) {
      document.getElementById("btnConvert").style.display = "none";
      select.style.display = "none";
      return;
    }
  }
  document.getElementById("btnConvert").style.display = "";
  select.style.display = "";
}

function renderFilePreview() {
  const previewList = document.getElementById("filePreviewList");
  previewList.innerHTML = "";
  pendingFiles.forEach((f, i) => {
    const item = document.createElement("div");
    item.className = "file-preview-item";
    item.innerHTML = `
      <span class="file-preview-icon">${iconForKind(classify(f))}</span>
      <span class="file-preview-name">${f.name}</span>
      <button class="btn-remove-file" title="삭제" data-index="${i}">✕</button>
    `;
    previewList.appendChild(item);
  });

  // 개별 삭제 버튼 이벤트
  previewList.querySelectorAll(".btn-remove-file").forEach(btn => {
    btn.addEventListener("click", () => {
      const idx = parseInt(btn.dataset.index);
      pendingFiles.splice(idx, 1);
      if (pendingFiles.length === 0) {
        document.getElementById("convertSection").style.display = "none";
      } else {
        renderFilePreview();
      }
    });
  });
}

function iconForKind(kind) {
  if (kind === "image") return "🖼";
  if (kind === "excel") return "📊";
  if (kind === "hangul") return "📄";
  if (kind === "hwp") return "⚠️";
  if (kind === "pdf") return "📕";
  return "📎";
}

async function convertPendingFiles() {
  const fmt = getSelectedFormat();
  const btn = document.getElementById("btnConvert");
  btn.disabled = true;
  btn.textContent = "변환 중...";

  for (const file of pendingFiles) {
    const kind = classify(file);
    try {
      if (kind === "pdf" || fmt === "passthrough") {
        await passThroughPdf(file);
        log(`완료 (변환 없음): ${file.name} → PDF`, "ok");
      } else if (kind === "image") {
        await convertImage(file, fmt);
        log(`완료: ${file.name} → ${fmt.toUpperCase()}`, "ok");
      } else if (kind === "excel") {
        await convertExcel(file, fmt);
        log(`완료: ${file.name} → ${fmt.toUpperCase()}`, "ok");
      } else if (kind === "hangul") {
        await convertHangul(file, fmt);
        log(`완료: ${file.name} → ${fmt.toUpperCase()}`, "ok");
      } else if (kind === "hwp") {
        log(`건너뜀: ${file.name} — HWP는 변환 불가. HWPX로 저장 후 다시 올려주세요.`, "warn");
      } else {
        log(`건너뜀 (지원 형식 아님): ${file.name}`, "warn");
      }
    } catch (err) {
      console.error(err);
      log(`실패: ${file.name} — ${err.message || String(err)}`, "err");
    }
  }

  btn.disabled = false;
  btn.textContent = "변환하기";
}

function downloadBlob(blob, filename) {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = filename;
  a.click();
  URL.revokeObjectURL(a.href);
}

async function convertImage(file, fmt) {
  const bitmap = await loadImageToBitmap(file);
  const canvas = document.createElement("canvas");
  canvas.width = bitmap.width;
  canvas.height = bitmap.height;
  const ctx = canvas.getContext("2d");
  if (fmt === "jpg" || fmt === "jpeg") {
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
  }
  ctx.drawImage(bitmap, 0, 0);
  const mime = fmt === "png" ? "image/png" : "image/jpeg";
  const quality = fmt === "png" ? undefined : 0.92;
  const blob = await new Promise((resolve, reject) => {
    canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("이미지 인코딩 실패"))), mime, quality);
  });
  const outExt = fmt === "png" ? "png" : "jpg";
  await saveOutput(blob, `${baseName(file.name)}.${outExt}`);
  bitmap.close?.();
}

function loadImageToBitmap(file) {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = async () => {
      try {
        const bmp = await createImageBitmap(img);
        URL.revokeObjectURL(url);
        resolve(bmp);
      } catch (e) {
        URL.revokeObjectURL(url);
        reject(e);
      }
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("이미지를 불러올 수 없습니다."));
    };
    img.crossOrigin = "anonymous";
    img.src = url;
  });
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, { type: "array" });
        resolve(wb);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = () => reject(new Error("파일 읽기 실패"));
    reader.readAsArrayBuffer(file);
  });
}

function workbookToTables(wb) {
  const tables = [];
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", raw: false });
    const strRows = rows.map((r) => r.map((c) => (c === null || c === undefined ? "" : String(c))));
    tables.push({ name: sheetName, rows: strRows });
  }
  return tables;
}

async function convertExcel(file, fmt) {
  const wb = await readWorkbook(file);
  const tables = workbookToTables(wb);
  if (fmt === "docx") {
    const doc = buildDocxFromTables(tables, `엑셀: ${file.name}`);
    const blob = await Packer.toBlob(doc);
    await saveOutput(blob, `${baseName(file.name)}.docx`);
    return;
  }
  const pngPages = await renderTablesToPngPages(tables, { title: file.name });
  const pdfBytes = await buildPdfFromPngPages(pngPages);
  await saveOutput(new Blob([pdfBytes], { type: "application/pdf" }), `${baseName(file.name)}.pdf`);
}

async function convertHangul(file, fmt) {
  const ext = extOf(file.name);
  const buf = await file.arrayBuffer();

  // HWP 바이너리는 제외
  if (ext === "hwp") {
    throw new Error("HWP 바이너리는 지원하지 않습니다. HWPX로 저장 후 변환해주세요.");
  }

  // HWPX 블록 추출
  const blocks = await extractHwpxBlocks(buf);

  if (!blocks.length) {
    throw new Error("HWPX 내용을 읽을 수 없습니다.");
  }

  // DOCX 변환
  if (fmt === "docx") {
    const tables = [{
      name: "본문",
      rows: blocks
        .filter(b => b.type === "text")
        .map(b => [b.value])
    }];

    const doc = buildDocxFromTables(tables, `HWPX: ${file.name}`);
    const blob = await Packer.toBlob(doc);

    await saveOutput(blob, `${baseName(file.name)}.docx`);
    return;
  }

  // PDF 변환 (이미지 포함)
  const pngPages = await renderBlocksToPngPages(blocks, file.name);

  const pdfBytes = await buildPdfFromPngPages(pngPages);

  await saveOutput(
    new Blob([pdfBytes], { type: "application/pdf" }),
    `${baseName(file.name)}.pdf`
  );
}

async function extractHwpxText(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  let names = Object.keys(zip.files).filter((n) => /^Contents\/section\d+\.xml$/i.test(n));
  if (!names.length) {
    names = Object.keys(zip.files).filter(
      (n) => /^Contents\/.+\.xml$/i.test(n) && !/(header|footer|chart|revision)/i.test(n)
    );
  }
  names.sort();
  const chunks = [];
  for (const n of names) {
    const xml = await zip.file(n).async("string");
    const stripped = xml
      .replace(/<script[\s\S]*?<\/script>/gi, "")
      .replace(/<style[\s\S]*?<\/style>/gi, "")
      .replace(/<[^>]+>/g, " ")
      .replace(/&nbsp;/g, " ")
      .replace(/&lt;/g, "<")
      .replace(/&gt;/g, ">")
      .replace(/&amp;/g, "&")
      .replace(/\s+/g, " ")
      .trim();
    if (stripped) chunks.push(stripped);
  }
  return chunks.join("\n\n");
}

// HWPX에서 콘텐츠 블록(텍스트/이미지) 순서대로 추출
async function extractHwpxBlocks(arrayBuffer) {
  const zip = await JSZip.loadAsync(arrayBuffer);

  // BinData 이미지 미리 로드 - 파일명(확장자포함)과 베이스명 둘 다 키로 등록
  const binImages = {};
  for (const [path, file] of Object.entries(zip.files)) {
    if (/^BinData\//i.test(path) && !file.dir) {
      const ext = extOf(path).toLowerCase();
      if (["png","jpg","jpeg","gif","bmp","tif","tiff"].includes(ext)) {
        const bytes = await file.async("uint8array");
        const mime = (ext === "jpg" || ext === "jpeg") ? "image/jpeg"
          : ext === "png" ? "image/png"
          : ext === "gif" ? "image/gif" : "image/bmp";
        const blob = new Blob([bytes], { type: mime });
        const url = URL.createObjectURL(blob);
        const fullName = path.split("/").pop().toLowerCase(); // "image1.bmp"
        const base = fullName.replace(/\.[^.]+$/, "");        // "image1"
        binImages[fullName] = url;
        binImages[base] = url; // binaryItemIDRef="image1" 대응
      }
    }
  }

  // 섹션 XML 파싱
  let sectionNames = Object.keys(zip.files).filter(n => /^Contents\/section\d+\.xml$/i.test(n));
  if (!sectionNames.length) {
    sectionNames = Object.keys(zip.files).filter(
      n => /^Contents\/.+\.xml$/i.test(n) && !/(header|footer|chart|revision)/i.test(n)
    );
  }
  sectionNames.sort();

  const blocks = [];

  for (const name of sectionNames) {
    const xmlStr = await zip.file(name).async("string");
    const parser = new DOMParser();
    const doc = parser.parseFromString(xmlStr, "text/xml");

    // HWPX는 hp:p 네임스페이스 사용
    const paras = Array.from(doc.getElementsByTagName("hp:p"));
    if (paras.length === 0) {
      const text = xmlStr.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
      if (text) blocks.push({ type: "text", value: text });
      continue;
    }

    for (const para of paras) {
      // 이미지: hc:img 태그의 binaryItemIDRef 속성
      const imgEls = Array.from(para.getElementsByTagName("hc:img"));
      for (const imgEl of imgEls) {
        const ref = imgEl.getAttribute("binaryItemIDRef");
        if (ref) {
          const url = binImages[ref.toLowerCase()];
          if (url) blocks.push({ type: "image", value: url });
        }
      }

      // 텍스트 추출: hp:t 태그만 수집 (메타데이터/alt텍스트 제외)
      const tEls = Array.from(para.getElementsByTagName("hp:t"));
      const text = tEls
        .map(el => el.textContent)
        .join("")
        .replace(/\s+/g, " ")
        .trim();

      if (text) blocks.push({ type: "text", value: text });
    }
  }

  return blocks;
}

async function renderHwpxBlocksToPdf(blocks, title) {
  const { PDFDocument, rgb, StandardFonts } = PDFLib;
  const pdfDoc = await PDFDocument.create();

  const PAGE_W = 595, PAGE_H = 842, MARGIN = 50;
  const contentW = PAGE_W - MARGIN * 2;
  const LINE_HEIGHT = 18, FONT_SIZE = 11;

  let page = pdfDoc.addPage([PAGE_W, PAGE_H]);
  let y = PAGE_H - MARGIN;
  const font = await pdfDoc.embedFont(StandardFonts.Helvetica);

  function newPage() {
    page = pdfDoc.addPage([PAGE_W, PAGE_H]);
    y = PAGE_H - MARGIN;
  }

  function ensureSpace(h) {
    if (y - h < MARGIN) newPage();
  }

  // 제목
  ensureSpace(LINE_HEIGHT * 2);
  page.drawText(title.slice(0, 60), {
    x: MARGIN, y,
    size: 13, font, color: rgb(0, 0, 0),
    maxWidth: contentW,
  });
  y -= LINE_HEIGHT * 2;

  for (const block of blocks) {
    if (block.type === "text") {
      // 텍스트 줄바꿈 처리
      const words = block.value;
      const charsPerLine = Math.floor(contentW / (FONT_SIZE * 0.55));
      const lines = [];
      for (let i = 0; i < words.length; i += charsPerLine) {
        lines.push(words.slice(i, i + charsPerLine));
      }
      for (const line of lines) {
        ensureSpace(LINE_HEIGHT);
        try {
          page.drawText(line, {
            x: MARGIN, y,
            size: FONT_SIZE, font, color: rgb(0.1, 0.1, 0.1),
            maxWidth: contentW,
          });
        } catch(_) {}
        y -= LINE_HEIGHT;
      }
      y -= 4;

    } else if (block.type === "image") {
      try {
        const resp = await fetch(block.value);
        const imgBytes = new Uint8Array(await resp.arrayBuffer());
        let embedded;
        try { embedded = await pdfDoc.embedPng(imgBytes); }
        catch(_) { embedded = await pdfDoc.embedJpg(imgBytes); }

        const maxW = contentW;
        const maxH = Math.min(300, PAGE_H - MARGIN * 2);
        const scale = Math.min(maxW / embedded.width, maxH / embedded.height, 1);
        const dw = embedded.width * scale;
        const dh = embedded.height * scale;

        ensureSpace(dh + 12);
        page.drawImage(embedded, {
          x: MARGIN + (contentW - dw) / 2,
          y: y - dh,
          width: dw, height: dh,
        });
        y -= dh + 12;
      } catch(e) {
        // 이미지 로드 실패 시 무시
      }
    }
  }

  return pdfDoc.save();
}

function buildDocxFromTables(tables, heading) {
  const children = [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: heading, bold: true, size: 28 })],
    }),
    new Paragraph({ children: [new TextRun({ text: "" })] }),
  ];
  for (const t of tables) {
    children.push(
      new Paragraph({
        children: [new TextRun({ text: `시트: ${t.name}`, bold: true, size: 24 })],
      })
    );
    if (!t.rows.length) {
      children.push(new Paragraph({ children: [new TextRun({ text: "(빈 시트)" })] }));
      continue;
    }
    const colCount = Math.max(1, ...t.rows.map((r) => r.length));
    const tableRows = t.rows.map(
      (row) =>
        new TableRow({
          children: Array.from({ length: colCount }, (_, j) => {
            const cellText = row[j] ?? "";
            return new TableCell({
              width: { size: 100 / colCount, type: WidthType.PERCENTAGE },
              children: [new Paragraph({ children: [new TextRun(String(cellText))] })],
            });
          }),
        })
    );
    children.push(
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: tableRows,
      })
    );
    children.push(new Paragraph({ children: [new TextRun({ text: "" })] }));
  }

  return new Document({
    sections: [{ properties: {}, children }],
  });
}

const PDF_W_PT = 595;
const PDF_H_PT = 842;
const MARGIN = 40;
const FONT = '14px "Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans KR", sans-serif';
const LINE_H = 22;
const HEADER_H = 36;

/** maxWidth(px) 안에 들어가도록 글자 단위로 줄 나눔 (한글·영문 공통) */
function wrapLines(ctx, text, maxWidth) {
  const s = String(text).replace(/\r?\n/g, " ");
  if (!s) return [""];
  const lines = [];
  let line = "";
  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    const test = line + ch;
    if (ctx.measureText(test).width <= maxWidth) {
      line = test;
    } else {
      if (line) lines.push(line);
      line = ch;
      if (ctx.measureText(line).width > maxWidth && maxWidth > 0) {
        lines.push(line);
        line = "";
      }
    }
  }
  if (line) lines.push(line);
  return lines.length ? lines : [""];
}

async function renderTablesToPngPages(tables, options = {}) {
  const { title = "", monoColumn = false } = options;
  const scale = 1.25;
  const pageW = Math.round((PDF_W_PT - MARGIN * 2) * scale);
  const maxPageH = Math.round((PDF_H_PT - MARGIN * 2) * scale);

  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
      ctx.font = FONT;

  const pages = [];

  async function flushPage() {
    pages.push(await canvasToPngBlob(canvas));
  }

  function newPage() {
    canvas.width = pageW;
    canvas.height = maxPageH;
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = "#111827";
    ctx.font = FONT;
    return 12;
  }

  let y = newPage();

  function drawHeaderLine(text) {
    ctx.save();
    ctx.font = `bold 15px "Malgun Gothic", "Apple SD Gothic Neo", sans-serif`;
    ctx.fillText(text.slice(0, 80), 8, y);
    ctx.restore();
    y += LINE_H;
  }

  if (title) drawHeaderLine(title);

  async function ensureSpace(h) {
    if (y + h > maxPageH - 8) {
  await flushPage();
      y = newPage();
}
  }

  for (const t of tables) {
    await ensureSpace(HEADER_H);
    ctx.save();
    ctx.font = `bold 13px "Malgun Gothic", sans-serif`;
    ctx.fillText(`[${t.name}]`, 8, y);
    ctx.restore();
    y += LINE_H + 4;

    if (!t.rows.length) {
      await ensureSpace(LINE_H);
      ctx.fillText("(비어 있음)", 8, y);
      y += LINE_H + 8;
      continue;
}

    const colCount = monoColumn ? 1 : Math.min(12, Math.max(1, ...t.rows.map((r) => r.length)));
    const colW = (pageW - 16) / colCount;
    const padX = 6;
    const innerW = colW - padX * 2;

    for (const row of t.rows) {
      const cells = monoColumn ? [row.join(" ")] : row.slice(0, colCount);
      while (cells.length < colCount) cells.push("");

      ctx.font = FONT;
      const lineBlocks = cells.map((txt) => wrapLines(ctx, txt ?? "", Math.max(20, innerW)));
      const maxLines = Math.max(1, ...lineBlocks.map((ls) => ls.length));
      const rowH = maxLines * LINE_H + 12;

      await ensureSpace(rowH + 4);

      ctx.strokeStyle = "#e5e7eb";
      for (let c = 0; c <= colCount; c++) {
        const lx = 8 + c * colW;
        ctx.beginPath();
        ctx.moveTo(lx, y);
        ctx.lineTo(lx, y + rowH);
        ctx.stroke();
      }
      ctx.beginPath();
      ctx.moveTo(8, y);
      ctx.lineTo(pageW - 8, y);
      ctx.moveTo(8, y + rowH);
      ctx.lineTo(pageW - 8, y + rowH);
      ctx.stroke();

      for (let c = 0; c < colCount; c++) {
        const x = 8 + c * colW + padX;
        const lines = lineBlocks[c];
        let ly = y + LINE_H - 2;
        for (const ln of lines) {
          ctx.fillText(ln, x, ly);
          ly += LINE_H;
        }
      }
      y += rowH + 2;
    }
    y += 12;
  }

  await flushPage();
  return pages;
}

// 표 없이 일반 문단+이미지로 캔버스 페이지 렌더링
async function renderBlocksToPngPages(blocks, title) {
  const PAGE_W = 744;  // A4 비율 (595pt * 1.25)
  const PAGE_H = 1052;
  const MARGIN = 60;
  const CONTENT_W = PAGE_W - MARGIN * 2;
  const FONT = '14px "Malgun Gothic", "Apple SD Gothic Neo", "Noto Sans KR", sans-serif';
  const TITLE_FONT = 'bold 16px "Malgun Gothic", "Apple SD Gothic Neo", sans-serif';
  const LINE_H = 22;
  const PARA_GAP = 10;

  const pages = [];
  const canvas = document.createElement("canvas");
  canvas.width = PAGE_W;
  canvas.height = PAGE_H;
  const ctx = canvas.getContext("2d");

  function newPage() {
    canvas.width = PAGE_W;
    canvas.height = PAGE_H;
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, PAGE_W, PAGE_H);
    ctx.fillStyle = "#1a1a1a";
    ctx.font = FONT;
    return MARGIN;
  }

  async function flushPage() {
    pages.push(await canvasToPngBlob(canvas));
  }

  async function ensureSpace(h) {
    if (y + h > PAGE_H - MARGIN) {
      await flushPage();
      y = newPage();
    }
  }

  let y = newPage();

  ctx.font = FONT;
  ctx.fillStyle = "#1a1a1a";

  for (const block of blocks) {
    if (block.type === "text") {
      const text = block.value.trim();
      if (!text) { y += LINE_H / 2; continue; }

      // 글자 단위 줄바꿈
      ctx.font = FONT;
      const lines = [];
      let line = "";
      for (const ch of text) {
        const test = line + ch;
        if (ctx.measureText(test).width <= CONTENT_W) {
          line = test;
        } else {
          if (line) lines.push(line);
          line = ch;
        }
      }
      if (line) lines.push(line);

      for (const ln of lines) {
        await ensureSpace(LINE_H);
        ctx.font = FONT;
        ctx.fillStyle = "#1a1a1a";
        ctx.fillText(ln, MARGIN, y);
        y += LINE_H;
      }
      y += PARA_GAP;

    } else if (block.type === "image") {
      try {
        const img = await new Promise((resolve, reject) => {
          const i = new Image();
          i.onload = () => resolve(i);
          i.onerror = reject;
          i.src = block.value;
        });

        const maxW = CONTENT_W;
        const maxH = Math.min(400, PAGE_H - MARGIN * 2 - 20);
        const scale = Math.min(maxW / img.naturalWidth, maxH / img.naturalHeight, 1);
        const dw = img.naturalWidth * scale;
        const dh = img.naturalHeight * scale;

        await ensureSpace(dh + PARA_GAP * 2);
        const x = MARGIN + (CONTENT_W - dw) / 2;
        ctx.drawImage(img, x, y, dw, dh);
        y += dh + PARA_GAP * 2;
      } catch (e) {
        // 이미지 로드 실패 무시
      }
    }
  }

  await flushPage();
  return pages;
}

function canvasToPngBlob(canvas) {
  return new Promise((resolve, reject) => {
    canvas.toBlob((b) => (b ? resolve(b) : reject(new Error("PNG 생성 실패"))), "image/png");
  });
}

async function buildPdfFromPngPages(pngBlobs) {
  const { PDFDocument } = PDFLib;
  const pdfDoc = await PDFDocument.create();
  for (const blob of pngBlobs) {
    const bytes = new Uint8Array(await blob.arrayBuffer());
    const img = await pdfDoc.embedPng(bytes);
    const page = pdfDoc.addPage([PDF_W_PT, PDF_H_PT]);
    const iw = img.width;
    const ih = img.height;
    const maxW = PDF_W_PT - MARGIN * 2;
    const maxH = PDF_H_PT - MARGIN * 2;
    const r = Math.min(maxW / iw, maxH / ih);
    const dw = iw * r;
    const dh = ih * r;
    const x = (PDF_W_PT - dw) / 2;
    const y = (PDF_H_PT - dh) / 2;
    page.drawImage(img, { x, y, width: dw, height: dh });
  }
  return pdfDoc.save();
}

btnPickFolder?.addEventListener("click", () => {
  void pickOutputFolder();
});

document.getElementById("btnConvert")?.addEventListener("click", () => {
  void convertPendingFiles();
});

document.getElementById("btnCancel")?.addEventListener("click", () => {
  pendingFiles = [];
  document.getElementById("convertSection").style.display = "none";
  document.getElementById("filePreviewList").innerHTML = "";
  document.getElementById("select-format").innerHTML = "";
});
btnClearFolder?.addEventListener("click", () => {
  void clearOutputFolder();
});

(async () => {
  const saved = await loadPersistedDirHandle();
  if (saved && typeof saved === "string") {
    outputDirHandle = saved;
  }
  updateFolderLabel();
})();