import { validateBatch } from "./batchValidate.js";
import { buildGradesWorkbook } from "./buildGradesWorkbook.js";
import {
  extractEssayQuestionCount,
  parseExamText,
} from "./parseExam.js";
import { extractExamPdfText } from "./pdfText.js";

const dropzone = document.getElementById("dropzone");
const fileInput = document.getElementById("fileInput");
const pickBtn = document.getElementById("pickBtn");
const filePanel = document.getElementById("filePanel");
const fileList = document.getElementById("fileList");
const errorPanel = document.getElementById("errorPanel");
const errorList = document.getElementById("errorList");
const warnPanel = document.getElementById("warnPanel");
const warnList = document.getElementById("warnList");
const downloadBtn = document.getElementById("downloadBtn");
const statusEl = document.getElementById("status");

/** @type {File[]} */
let selectedFiles = [];

/** @type {{ ok: boolean, errors: string[], warnings: string[], reconciledCourse: string, nSa: number | null, exams: Map<string, Record<string, string>> } | null} */
let lastValidation = null;

const pdfjsModule = await import(
  new URL("../vendor/pdfjs/pdf.mjs", import.meta.url).href
);
const workerSrc = new URL(
  "../vendor/pdfjs/pdf.worker.mjs",
  import.meta.url
).href;

function setStatus(text, ready = false) {
  statusEl.textContent = text;
  statusEl.classList.toggle("ready", ready);
}

function renderFileList() {
  fileList.innerHTML = "";
  if (selectedFiles.length === 0) {
    filePanel.hidden = true;
    return;
  }
  filePanel.hidden = false;
  for (const f of selectedFiles) {
    const li = document.createElement("li");
    li.textContent = `${f.name} (${Math.round(f.size / 1024)} KB)`;
    fileList.appendChild(li);
  }
}

function renderMessages(errors, warnings) {
  if (errors.length) {
    errorPanel.hidden = false;
    errorList.innerHTML = "";
    for (const e of errors) {
      const li = document.createElement("li");
      li.textContent = e;
      errorList.appendChild(li);
    }
  } else {
    errorPanel.hidden = true;
  }

  if (warnings.length) {
    warnPanel.hidden = false;
    warnList.innerHTML = "";
    for (const w of warnings) {
      const li = document.createElement("li");
      li.textContent = w;
      warnList.appendChild(li);
    }
  } else {
    warnPanel.hidden = true;
  }
}

async function parseAllPdfs(files) {
  /** @type {{ fileName: string, course: string, exam_no: string, mc: Record<string, string>, essayCount: number | null }[]} */
  const rows = [];
  for (const file of files) {
    const buf = await file.arrayBuffer();
    const { page1Text, fullText } = await extractExamPdfText(
      buf,
      pdfjsModule,
      workerSrc
    );
    const parsed = parseExamText(fullText);
    const essayCount = extractEssayQuestionCount(page1Text);
    rows.push({
      fileName: file.name,
      course: parsed.course,
      exam_no: parsed.exam_no,
      mc: parsed.mc,
      essayCount,
    });
  }
  return rows;
}

async function runPipeline() {
  downloadBtn.disabled = true;
  lastValidation = null;
  renderMessages([], []);

  if (selectedFiles.length === 0) {
    setStatus("Add one or more PDF files.");
    return;
  }

  setStatus("Parsing PDFs…");
  try {
    const rows = await parseAllPdfs(selectedFiles);
    const v = validateBatch(rows);
    lastValidation = v;
    renderMessages(v.errors, v.warnings);

    if (v.ok && v.nSa !== null) {
      setStatus(
        `Ready: ${v.exams.size} student(s), ${v.nSa} SA column(s), course "${v.reconciledCourse}".`,
        true
      );
      downloadBtn.disabled = false;
    } else {
      setStatus("Fix errors above, then re-select files.");
    }
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    renderMessages([msg], []);
    setStatus("Parsing failed.");
  }
}

function setFiles(fileArray) {
  selectedFiles = [...fileArray].filter((f) =>
    f.name.toLowerCase().endsWith(".pdf")
  );
  renderFileList();
  void runPipeline();
}

pickBtn.addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", () => {
  if (fileInput.files?.length) {
    setFiles(fileInput.files);
  }
});

dropzone.addEventListener("dragover", (e) => {
  e.preventDefault();
  dropzone.classList.add("dragover");
});
dropzone.addEventListener("dragleave", () => {
  dropzone.classList.remove("dragover");
});
dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  dropzone.classList.remove("dragover");
  if (e.dataTransfer?.files?.length) {
    setFiles(e.dataTransfer.files);
  }
});

dropzone.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    fileInput.click();
  }
});

downloadBtn.addEventListener("click", async () => {
  const v = lastValidation;
  if (!v?.ok || v.nSa === null) return;
  setStatus("Building workbook…");
  downloadBtn.disabled = true;
  try {
    const blob = await buildGradesWorkbook({
      reconciledCourse: v.reconciledCourse,
      nSa: v.nSa,
      exams: v.exams,
    });
    const name = `${sanitizeFilename(v.reconciledCourse)}.xlsx`;
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    a.click();
    URL.revokeObjectURL(a.href);
    setStatus(`Downloaded ${name}.`, true);
    downloadBtn.disabled = false;
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    renderMessages([...(lastValidation?.errors ?? []), msg], lastValidation?.warnings ?? []);
    setStatus("Export failed.");
    downloadBtn.disabled = false;
  }
});

function sanitizeFilename(course) {
  return course.replace(/[/\\?%*:|"<>]/g, "-").trim() || "grades";
}

setStatus("Add one or more PDF files.");
