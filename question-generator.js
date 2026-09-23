(() => {
  "use strict";

  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
  const TEMPLATE_URL = "question-template.docx";
  const DEFAULT_SET_COUNT = 4;

  const $ = (id) => document.getElementById(id);
  const fileInput = $("docxFile");
  const dropZone = $("dropZone");
  const generateBtn = $("generateBtn");
  const resetBtn = $("resetBtn");
  const statusEl = $("status");
  const selectedFileEl = $("selectedFile");
  const detectedInfoEl = $("detectedInfo");
  const resultsEl = $("results");
  const resultsNote = $("resultsNote");
  const setGrid = $("setGrid");
  const downloadAllBtn = $("downloadAllBtn");
  const debugEl = $("debug");

  let selectedFile = null;
  let sourceDoc = null;
  let sourceRecords = [];
  let templateBytes = null;
  let generatedFiles = [];
  let lastGenerationLabel = "Question-Sets";
  let busy = false;

  fileInput.addEventListener("change", () => {
    const file = fileInput.files && fileInput.files[0] ? fileInput.files[0] : null;
    selectFile(file);
  });

  ["dragenter", "dragover"].forEach((type) => {
    dropZone.addEventListener(type, (event) => {
      event.preventDefault();
      dropZone.classList.add("dragover");
    });
  });

  ["dragleave", "drop"].forEach((type) => {
    dropZone.addEventListener(type, (event) => {
      event.preventDefault();
      dropZone.classList.remove("dragover");
    });
  });

  dropZone.addEventListener("drop", (event) => {
    const file = event.dataTransfer && event.dataTransfer.files && event.dataTransfer.files[0]
      ? event.dataTransfer.files[0]
      : null;
    selectFile(file, true);
  });

  dropZone.addEventListener("click", (event) => {
    if (event.target.closest("label") || event.target === fileInput) return;
    fileInput.click();
  });

  dropZone.addEventListener("keydown", (event) => {
    if (event.key === "Enter" || event.key === " ") {
      event.preventDefault();
      fileInput.click();
    }
  });

  generateBtn.addEventListener("click", () => generateSets(false));
  resetBtn.addEventListener("click", resetAll);
  downloadAllBtn.addEventListener("click", downloadAllAsZip);

  async function selectFile(file, fromDrop = false) {
    if (busy) return;

    if (!file) {
      clearSource();
      updateFileState();
      return;
    }

    if (!/\.docx$/i.test(file.name)) {
      clearSource();
      fileInput.value = "";
      updateFileState();
      setStatus("শুধু .docx ফাইল ব্যবহার করুন।", "error");
      return;
    }

    selectedFile = file;

    if (fromDrop) {
      try {
        const transfer = new DataTransfer();
        transfer.items.add(file);
        fileInput.files = transfer.files;
      } catch (_) {
        // selectedFile variable is enough in browsers that block FileList assignment.
      }
    }

    updateFileState();
    await prepareSourceAndAutoGenerate();
  }

  function clearSource() {
    selectedFile = null;
    sourceDoc = null;
    sourceRecords = [];
    generatedFiles = [];
  }

  function updateFileState() {
    if (selectedFile) {
      selectedFileEl.textContent = `✓ ${selectedFile.name} (${formatBytes(selectedFile.size)})`;
      selectedFileEl.classList.add("ready");
    } else {
      selectedFileEl.textContent = "কোনো ফাইল নির্বাচন করা হয়নি";
      selectedFileEl.classList.remove("ready");
      detectedInfoEl.hidden = true;
      detectedInfoEl.textContent = "";
      generateBtn.disabled = true;
    }
  }

  async function prepareSourceAndAutoGenerate() {
    setBusy(true);
    resetResults();
    debugEl.style.display = "none";
    debugEl.textContent = "";

    try {
      ensureJSZip();
      setStatus("Admin DOCX table পড়া হচ্ছে...");

      const bytes = await selectedFile.arrayBuffer();
      const parsed = await parseQuestionBank(bytes);
      sourceDoc = parsed.doc;
      sourceRecords = parsed.records;

      if (!sourceRecords.length) {
        throw new Error("Question | A | B | C | D | Answer | Explanation structure-এর কোনো valid table row পাওয়া যায়নি।");
      }

      const count = sourceRecords.length;
      $("questionsPerSet").max = String(count);
      $("questionsPerSet").value = String(count);
      $("setCount").value = String(DEFAULT_SET_COUNT);

      const derivedExam = deriveExamName(selectedFile.name);
      if (derivedExam) $("examName").value = derivedExam;

      const mathCount = sourceRecords.filter((record) => record.hasMath).length;
      detectedInfoEl.hidden = false;
      detectedInfoEl.textContent = `✓ ${count}টি প্রশ্ন পাওয়া গেছে${mathCount ? ` · ${mathCount}টি প্রশ্নে Word equation আছে` : ""}`;
      generateBtn.disabled = false;

      setStatus(`✓ ${count}টি প্রশ্ন পাওয়া গেছে। Default ${DEFAULT_SET_COUNT}টি Set স্বয়ংক্রিয়ভাবে তৈরি হচ্ছে...`);
      await generateSets(true);
    } catch (error) {
      console.error(error);
      sourceDoc = null;
      sourceRecords = [];
      generateBtn.disabled = true;
      setStatus(error && error.message ? error.message : "DOCX পড়ার সময় সমস্যা হয়েছে।", "error");
    } finally {
      setBusy(false);
    }
  }

  async function parseQuestionBank(arrayBuffer) {
    const zip = await window.JSZip.loadAsync(arrayBuffer);
    const documentFile = zip.file("word/document.xml");
    if (!documentFile) throw new Error("এটি valid DOCX নয়: word/document.xml পাওয়া যায়নি।");

    const xml = await documentFile.async("text");
    const doc = parseXml(xml, "Admin DOCX");
    const tables = Array.from(doc.getElementsByTagNameNS(W_NS, "tbl"));

    let bestRecords = [];
    for (const table of tables) {
      const rows = directChildren(table, "tr");
      const records = [];
      for (const row of rows) {
        const record = recordFromSourceRow(row);
        if (record) records.push(record);
      }
      if (records.length > bestRecords.length) bestRecords = records;
    }

    return { doc, records: bestRecords };
  }

  function recordFromSourceRow(row) {
    const cells = directChildren(row, "tc");
    if (cells.length < 7) return null;

    const mappings = [];
    if (cells.length >= 8) mappings.push({ q: 1, a: 2, b: 3, c: 4, d: 5, ans: 6, exp: 7 });
    mappings.push({ q: 0, a: 1, b: 2, c: 3, d: 4, ans: 5, exp: 6 });

    for (const map of mappings) {
      if (Math.max(...Object.values(map)) >= cells.length) continue;
      const question = cellText(cells[map.q]);
      const options = [map.a, map.b, map.c, map.d].map((index) => cellText(cells[index]));
      const answer = normalizeAnswer(cellText(cells[map.ans]));
      const explanation = cellText(cells[map.exp]);

      if (!question || options.filter(Boolean).length < 2 || !answer) continue;

      const sourceNodes = [cells[map.q], cells[map.a], cells[map.b], cells[map.c], cells[map.d], cells[map.exp]];
      const hasMath = sourceNodes.some((cell) => cell.getElementsByTagNameNS(M_NS, "oMath").length > 0);
      const hasUnsupportedDrawing = sourceNodes.some((cell) =>
        cell.getElementsByTagNameNS(W_NS, "drawing").length > 0 ||
        cell.getElementsByTagNameNS(W_NS, "pict").length > 0
      );

      return {
        questionCell: cells[map.q],
        optionCells: [cells[map.a], cells[map.b], cells[map.c], cells[map.d]],
        answer,
        explanationCell: cells[map.exp],
        questionText: question,
        explanationText: explanation,
        hasMath,
        hasUnsupportedDrawing
      };
    }

    return null;
  }

  async function generateSets(isAutomatic) {
    if (busy && !isAutomatic) return;
    if (!sourceDoc || !sourceRecords.length) {
      setStatus("আগে একটি valid Admin DOCX upload করুন।", "error");
      return;
    }

    const manageBusy = !isAutomatic;
    if (manageBusy) setBusy(true);
    resetResults();

    try {
      ensureJSZip();
      await loadTemplate();

      const setCount = clamp(positiveInt($("setCount").value, DEFAULT_SET_COUNT), 1, 26);
      const questionsPerSet = clamp(positiveInt($("questionsPerSet").value, sourceRecords.length), 1, sourceRecords.length);
      $("setCount").value = String(setCount);
      $("questionsPerSet").value = String(questionsPerSet);

      const exam = cleanText($("examName").value) || deriveExamName(selectedFile ? selectedFile.name : "") || "Exam";
      const subject = cleanText($("subjectName").value);
      const warningCount = sourceRecords.filter((record) => record.hasUnsupportedDrawing).length;

      lastGenerationLabel = sanitizeFilename([exam, subject, `${setCount}-Sets`].filter(Boolean).join("-"));
      setStatus(`${setCount}টি Set তৈরি হচ্ছে · প্রতি Set-এ ${questionsPerSet}টি প্রশ্ন...`);

      generatedFiles = [];
      setGrid.innerHTML = "";

      for (let index = 0; index < setCount; index += 1) {
        const setLabel = indexToSetLabel(index);
        const shuffled = secureShuffle(sourceRecords.slice()).slice(0, questionsPerSet);
        const blob = await buildSetDocx(setLabel, shuffled);
        const name = makeSetFilename(exam, subject, setLabel);
        const file = { blob, name, setLabel, count: shuffled.length };
        generatedFiles.push(file);
        renderSetCard(file);

        setStatus(`Set ${setLabel} তৈরি হয়েছে (${index + 1}/${setCount})...`);
        await nextFrame();
      }

      resultsNote.textContent = `${sourceRecords.length}টি source প্রশ্ন থেকে ${setCount}টি shuffled Set তৈরি হয়েছে · প্রতি Set-এ ${questionsPerSet}টি প্রশ্ন · Answer ও Explanation একই প্রশ্নের সঙ্গে রাখা হয়েছে।`;
      resultsEl.classList.add("visible");
      downloadAllBtn.disabled = false;

      if (warningCount) {
        setStatus(`✓ সম্পন্ন। ${warningCount}টি source row-এ embedded drawing/image পাওয়া গেছে; text ও Word equation থাকবে, কিন্তু embedded image copy নাও হতে পারে।`, "warning");
      } else {
        setStatus(`✓ সম্পন্ন। ${setCount}টি DOCX Set download-এর জন্য প্রস্তুত।`);
      }

      if (!isAutomatic) resultsEl.scrollIntoView({ behavior: "smooth", block: "start" });
    } catch (error) {
      console.error(error);
      setStatus(error && error.message ? error.message : "Set তৈরির সময় সমস্যা হয়েছে।", "error");
    } finally {
      if (manageBusy) setBusy(false);
    }
  }

  async function loadTemplate() {
    if (templateBytes) return;
    const response = await fetch(TEMPLATE_URL, { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`question-template.docx পাওয়া যায়নি (HTTP ${response.status})। Template file-টি repository root-এ upload করুন।`);
    }
    templateBytes = await response.arrayBuffer();
  }

  async function buildSetDocx(setLabel, records) {
    const zip = await window.JSZip.loadAsync(templateBytes.slice(0));
    const documentFile = zip.file("word/document.xml");
    if (!documentFile) throw new Error("Template DOCX invalid: document.xml নেই।");

    const xml = await documentFile.async("text");
    const doc = parseXml(xml, "Template DOCX");
    const body = doc.getElementsByTagNameNS(W_NS, "body")[0];
    const bodyTables = directChildren(body, "tbl");
    if (bodyTables.length < 2) throw new Error("Template structure invalid: question/answer table পাওয়া যায়নি।");

    const questionTable = bodyTables[0];
    const answerTable = bodyTables[1];
    const questionRows = directChildren(questionTable, "tr");
    const answerRows = directChildren(answerTable, "tr");
    if (!questionRows.length || answerRows.length < 2) throw new Error("Template prototype row পাওয়া যায়নি।");

    const questionPrototype = questionRows[0].cloneNode(true);
    const answerPrototype = answerRows[1].cloneNode(true);

    for (const row of questionRows) questionTable.removeChild(row);
    for (const row of answerRows.slice(1)) answerTable.removeChild(row);

    records.forEach((record, index) => {
      const questionRow = questionPrototype.cloneNode(true);
      fillQuestionRow(doc, questionRow, record, index + 1);
      questionTable.appendChild(questionRow);

      const answerRow = answerPrototype.cloneNode(true);
      fillAnswerRow(doc, answerRow, record, index + 1);
      answerTable.appendChild(answerRow);
    });

    replaceAllText(doc, "{{SET}}", setLabel);

    const serialized = new XMLSerializer().serializeToString(doc);
    zip.file("word/document.xml", ensureXmlDeclaration(serialized));

    return zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      compression: "DEFLATE",
      compressionOptions: { level: 6 }
    });
  }

  function fillQuestionRow(doc, row, record, number) {
    const cells = directChildren(row, "tc");
    if (cells.length < 2) throw new Error("Template question row invalid।");

    setCellText(cells[0], `${number}.`);

    const nestedTable = directChildren(cells[1], "tbl")[0];
    if (!nestedTable) throw new Error("Template nested question table পাওয়া যায়নি।");
    const nestedRows = directChildren(nestedTable, "tr");
    if (nestedRows.length < 3) throw new Error("Template option rows invalid।");

    const questionCell = directChildren(nestedRows[0], "tc")[0];
    const optionRowAB = directChildren(nestedRows[1], "tc");
    const optionRowCD = directChildren(nestedRows[2], "tc");
    if (!questionCell || optionRowAB.length < 4 || optionRowCD.length < 4) {
      throw new Error("Template option cell structure invalid।");
    }

    copyCellBody(doc, record.questionCell, questionCell);
    copyCellBody(doc, record.optionCells[0], optionRowAB[1]);
    copyCellBody(doc, record.optionCells[1], optionRowAB[3]);
    copyCellBody(doc, record.optionCells[2], optionRowCD[1]);
    copyCellBody(doc, record.optionCells[3], optionRowCD[3]);
  }

  function fillAnswerRow(doc, row, record, number) {
    const cells = directChildren(row, "tc");
    if (cells.length < 3) throw new Error("Template answer row invalid।");
    setCellText(cells[0], `${String(number).padStart(2, "0")}.`);
    setCellText(cells[1], record.answer);
    copyCellBody(doc, record.explanationCell, cells[2]);
  }

  function copyCellBody(targetDoc, sourceCell, targetCell) {
    const preservedTcPr = directChildren(targetCell, "tcPr")[0] || null;
    while (targetCell.firstChild) targetCell.removeChild(targetCell.firstChild);
    if (preservedTcPr) targetCell.appendChild(preservedTcPr);

    const sourceChildren = Array.from(sourceCell.childNodes).filter((node) => {
      return !(node.nodeType === 1 && node.namespaceURI === W_NS && node.localName === "tcPr");
    });

    for (const child of sourceChildren) {
      const imported = targetDoc.importNode(child, true);
      stripUnsafeRelationshipContent(imported);
      targetCell.appendChild(imported);
    }

    if (!directChildren(targetCell, "p").length && !directChildren(targetCell, "tbl").length) {
      targetCell.appendChild(targetDoc.createElementNS(W_NS, "w:p"));
    }
  }

  function stripUnsafeRelationshipContent(root) {
    if (!root || root.nodeType !== 1) return;
    const drawings = [
      ...Array.from(root.getElementsByTagNameNS(W_NS, "drawing")),
      ...Array.from(root.getElementsByTagNameNS(W_NS, "pict"))
    ];
    for (const node of drawings) node.remove();

    const attrsToDrop = ["paraId", "textId"];
    const elements = [root, ...Array.from(root.getElementsByTagName("*"))];
    for (const element of elements) {
      for (const name of attrsToDrop) {
        const matches = Array.from(element.attributes || []).filter((attr) => attr.localName === name);
        matches.forEach((attr) => element.removeAttributeNode(attr));
      }
    }
  }

  function setCellText(cell, value) {
    let paragraph = directChildren(cell, "p")[0];
    if (!paragraph) {
      paragraph = cell.ownerDocument.createElementNS(W_NS, "w:p");
      cell.appendChild(paragraph);
    }

    const pPr = directChildren(paragraph, "pPr")[0] || null;
    let prototypeRun = directChildren(paragraph, "r")[0];
    let rPr = prototypeRun ? directChildren(prototypeRun, "rPr")[0] : null;
    if (rPr) rPr = rPr.cloneNode(true);

    while (paragraph.firstChild) paragraph.removeChild(paragraph.firstChild);
    if (pPr) paragraph.appendChild(pPr);

    const run = paragraph.ownerDocument.createElementNS(W_NS, "w:r");
    if (rPr) run.appendChild(rPr);
    const text = paragraph.ownerDocument.createElementNS(W_NS, "w:t");
    text.setAttribute("xml:space", "preserve");
    text.textContent = value;
    run.appendChild(text);
    paragraph.appendChild(run);
  }

  function renderSetCard(file) {
    const article = document.createElement("article");
    article.className = "set-card";
    article.innerHTML = `
      <div class="set-badge">SET ${escapeHtml(file.setLabel)}</div>
      <h3>Set ${escapeHtml(file.setLabel)}</h3>
      <p>${file.count}টি shuffled প্রশ্ন + Answer/Explanation</p>
      <button class="set-download" type="button">⬇ ${escapeHtml(file.name)}</button>
    `;
    article.querySelector("button").addEventListener("click", () => triggerDownload(file.blob, file.name));
    setGrid.appendChild(article);
  }

  async function downloadAllAsZip() {
    if (!generatedFiles.length) return;
    try {
      ensureJSZip();
      downloadAllBtn.disabled = true;
      const oldText = downloadAllBtn.textContent;
      downloadAllBtn.textContent = "ZIP তৈরি হচ্ছে...";
      const zip = new window.JSZip();
      generatedFiles.forEach((file) => zip.file(file.name, file.blob));
      const blob = await zip.generateAsync({ type: "blob", compression: "DEFLATE", compressionOptions: { level: 6 } });
      triggerDownload(blob, `${lastGenerationLabel || "Question-Sets"}.zip`);
      downloadAllBtn.textContent = oldText;
    } catch (error) {
      console.error(error);
      setStatus("ZIP তৈরি করা যায়নি। Set card থেকে DOCX আলাদাভাবে download করুন।", "error");
    } finally {
      downloadAllBtn.disabled = generatedFiles.length === 0;
    }
  }

  function resetResults() {
    generatedFiles = [];
    setGrid.innerHTML = "";
    resultsNote.textContent = "";
    resultsEl.classList.remove("visible");
    downloadAllBtn.disabled = true;
  }

  function resetAll() {
    if (busy) return;
    clearSource();
    fileInput.value = "";
    $("examName").value = "T-08";
    $("subjectName").value = "Meghna";
    $("setCount").value = String(DEFAULT_SET_COUNT);
    $("questionsPerSet").value = "100";
    resetResults();
    debugEl.style.display = "none";
    debugEl.textContent = "";
    setStatus("");
    updateFileState();
  }

  function setBusy(value) {
    busy = value;
    fileInput.disabled = value;
    resetBtn.disabled = value;
    generateBtn.disabled = value || !sourceRecords.length;
    if (value) dropZone.classList.add("busy");
    else dropZone.classList.remove("busy");
  }

  function setStatus(message, type = "") {
    statusEl.textContent = message;
    statusEl.className = `status${type ? ` ${type}` : ""}`;
  }

  function ensureJSZip() {
    if (!window.JSZip) throw new Error("JSZip load হয়নি। Internet connection পরীক্ষা করে page reload করুন।");
  }

  function parseXml(xml, label) {
    const doc = new DOMParser().parseFromString(xml, "application/xml");
    const parserError = doc.getElementsByTagName("parsererror")[0];
    if (parserError) throw new Error(`${label} XML parse করা যায়নি।`);
    return doc;
  }

  function directChildren(parent, localName) {
    return Array.from(parent.childNodes).filter((node) =>
      node.nodeType === 1 && node.namespaceURI === W_NS && node.localName === localName
    );
  }

  function cellText(cell) {
    return cleanText(cell.textContent || "");
  }

  function normalizeAnswer(value) {
    const text = cleanText(value).replace(/[.()\-:।]/g, "").toUpperCase();
    const bengali = { "ক": "A", "খ": "B", "গ": "C", "ঘ": "D" };
    if (bengali[text]) return bengali[text];
    const match = text.match(/\b([ABCD])\b/);
    return match ? match[1] : "";
  }

  function replaceAllText(doc, needle, replacement) {
    const nodes = Array.from(doc.getElementsByTagNameNS(W_NS, "t"));
    for (const node of nodes) {
      if (node.textContent && node.textContent.includes(needle)) {
        node.textContent = node.textContent.split(needle).join(replacement);
      }
    }
  }

  function secureShuffle(array) {
    for (let i = array.length - 1; i > 0; i -= 1) {
      const j = secureRandomInt(i + 1);
      [array[i], array[j]] = [array[j], array[i]];
    }
    return array;
  }

  function secureRandomInt(maxExclusive) {
    if (window.crypto && window.crypto.getRandomValues) {
      const maxUint = 0x100000000;
      const limit = maxUint - (maxUint % maxExclusive);
      const buffer = new Uint32Array(1);
      do {
        window.crypto.getRandomValues(buffer);
      } while (buffer[0] >= limit);
      return buffer[0] % maxExclusive;
    }
    return Math.floor(Math.random() * maxExclusive);
  }

  function indexToSetLabel(index) {
    return String.fromCharCode(65 + index);
  }

  function makeSetFilename(exam, subject, setLabel) {
    const base = subject ? `${exam} (${subject}) Set-${setLabel}` : `${exam} Set-${setLabel}`;
    return `${sanitizeFilename(base)}.docx`;
  }

  function deriveExamName(filename) {
    return cleanText(String(filename || "").replace(/\.docx$/i, ""));
  }

  function sanitizeFilename(value) {
    return String(value || "")
      .replace(/[\\/:*?"<>|]/g, "-")
      .replace(/\s+/g, " ")
      .trim();
  }

  function cleanText(value) {
    return String(value || "").replace(/\u00a0/g, " ").replace(/\s+/g, " ").trim();
  }

  function positiveInt(value, fallback) {
    const parsed = Number.parseInt(value, 10);
    return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
  }

  function clamp(value, min, max) {
    return Math.min(Math.max(value, min), max);
  }

  function formatBytes(bytes) {
    if (!bytes) return "0 B";
    const units = ["B", "KB", "MB", "GB"];
    const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
    const number = bytes / (1024 ** index);
    return `${number.toFixed(index === 0 ? 0 : 1)} ${units[index]}`;
  }

  function triggerDownload(blob, name) {
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = name;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
  }

  function ensureXmlDeclaration(xml) {
    return xml.startsWith("<?xml") ? xml : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${xml}`;
  }

  function nextFrame() {
    return new Promise((resolve) => requestAnimationFrame(() => resolve()));
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }
})();
