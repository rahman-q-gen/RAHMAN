(() => {
  "use strict";

  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
  const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
  const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
  const TEMPLATE_URL = "question-template.docx";
  const DEFAULT_SET_COUNT = 4;
  const GENERATION_MODE = Object.freeze({
    QUESTIONS: "questions",
    QUESTIONS_OPTIONS: "questions-options"
  });
  const OPTION_LETTERS = ["A", "B", "C", "D"];

  const $ = (id) => document.getElementById(id);
  const fileInput = $("docxFile");
  const dropZone = $("dropZone");
  const generateQuestionsBtn = $("generateQuestionsBtn");
  const generateQuestionsOptionsBtn = $("generateQuestionsOptionsBtn");
  const resetBtn = $("resetBtn");
  const statusEl = $("status");
  const selectedFileEl = $("selectedFile");
  const detectedInfoEl = $("detectedInfo");
  const resultsEl = $("results");
  const resultsNote = $("resultsNote");
  const setGrid = $("setGrid");
  const downloadAllBtn = $("downloadAllBtn");
  const debugEl = $("debug");
  const messageModal = $("messageModal");
  const modalCard = messageModal.querySelector(".qg-modal-card");
  const modalIcon = $("modalIcon");
  const modalTitle = $("modalTitle");
  const modalMessage = $("modalMessage");
  const modalCancelBtn = $("modalCancelBtn");
  const modalConfirmBtn = $("modalConfirmBtn");

  let selectedFile = null;
  let sourceDoc = null;
  let sourcePackage = null;
  let sourceRecords = [];
  let sourceValidation = { blockingErrors: [], missingExplanations: [] };
  let templateBytes = null;
  let generatedFiles = [];
  let lastGenerationLabel = "Question-Sets";
  let busy = false;
  let activeModalResolve = null;
  let modalIsConfirmation = false;

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

  generateQuestionsBtn.addEventListener("click", () => generateSets(GENERATION_MODE.QUESTIONS));
  generateQuestionsOptionsBtn.addEventListener("click", () => generateSets(GENERATION_MODE.QUESTIONS_OPTIONS));
  resetBtn.addEventListener("click", resetAll);
  downloadAllBtn.addEventListener("click", downloadAllAsZip);
  modalConfirmBtn.addEventListener("click", () => finishModal(true));
  modalCancelBtn.addEventListener("click", () => finishModal(false));
  messageModal.addEventListener("click", (event) => {
    if (event.target.matches("[data-modal-close]")) finishModal(modalIsConfirmation ? false : true);
  });
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && !messageModal.hidden) finishModal(modalIsConfirmation ? false : true);
  });

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
      await showMessageModal({
        title: "ভুল ফাইল টাইপ",
        message: "শুধু .docx Question Bank ফাইল আপলোড করুন।",
        tone: "error"
      });
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
    await prepareSource();
  }

  function clearSource() {
    selectedFile = null;
    sourceDoc = null;
    sourcePackage = null;
    sourceRecords = [];
    sourceValidation = { blockingErrors: [], missingExplanations: [] };
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
      setGenerateButtonsDisabled(true);
    }
  }

  async function prepareSource() {
    setBusy(true);
    resetResults();
    debugEl.style.display = "none";
    debugEl.textContent = "";

    try {
      ensureJSZip();
      setStatus("Admin DOCX table পড়া ও যাচাই করা হচ্ছে...");

      const bytes = await selectedFile.arrayBuffer();
      const parsed = await parseQuestionBank(bytes);
      sourceDoc = parsed.doc;
      sourcePackage = parsed.package;
      sourceRecords = parsed.records;

      if (!sourceRecords.length) {
        throw new Error("Question | A | B | C | D | Answer | Explanation structure-এর কোনো data row পাওয়া যায়নি।");
      }

      const count = sourceRecords.length;
      $("questionsPerSet").max = String(count);
      $("questionsPerSet").value = String(count);
      $("setCount").value = String(DEFAULT_SET_COUNT);

      const derivedExam = deriveExamName(selectedFile.name);
      if (derivedExam) $("examName").value = derivedExam;

      sourceValidation = validateSourceRecords(sourceRecords);
      const mathCount = sourceRecords.filter((record) => record.hasMath).length;
      const imageCount = sourceRecords.filter((record) => record.imageRelIds && record.imageRelIds.length).length;
      const explanationCount = sourceValidation.missingExplanations.length;

      detectedInfoEl.hidden = false;
      detectedInfoEl.textContent = `✓ ${count}টি প্রশ্নের row পাওয়া গেছে${mathCount ? ` · ${mathCount}টিতে Word equation` : ""}${imageCount ? ` · ${imageCount}টিতে embedded image` : ""}${explanationCount ? ` · ${explanationCount}টিতে ব্যাখ্যা নেই` : ""}`;

      if (sourceValidation.blockingErrors.length) {
        setGenerateButtonsDisabled(true);
        setStatus(`Generate বন্ধ: ${sourceValidation.blockingErrors.length}টি সমস্যা পাওয়া গেছে।`, "error");
        await showMessageModal({
          title: "Question Bank ঠিক করতে হবে",
          message: buildBlockingValidationMessage(sourceValidation.blockingErrors),
          tone: "error",
          confirmText: "ঠিক আছে"
        });
        return;
      }

      setGenerateButtonsDisabled(false);
      if (explanationCount) {
        setStatus(`✓ প্রশ্ন ও Option ঠিক আছে। ${explanationCount}টি প্রশ্নে ব্যাখ্যা নেই; Generate করার সময় confirmation দেখাবে।`, "warning");
      } else {
        setStatus(`✓ ${count}টি প্রশ্ন যাচাই সম্পন্ন। এখন Shuffle mode নির্বাচন করে Generate করুন।`);
      }
    } catch (error) {
      console.error(error);
      sourceDoc = null;
      sourcePackage = null;
      sourceRecords = [];
      sourceValidation = { blockingErrors: [], missingExplanations: [] };
      setGenerateButtonsDisabled(true);
      const message = error && error.message ? error.message : "DOCX পড়ার সময় সমস্যা হয়েছে।";
      setStatus(message, "error");
      await showMessageModal({ title: "DOCX পড়া যায়নি", message, tone: "error", confirmText: "ঠিক আছে" });
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

    const relsFile = zip.file("word/_rels/document.xml.rels");
    const relsDoc = relsFile
      ? parseXml(await relsFile.async("text"), "Admin DOCX relationships")
      : createRelationshipsDocument();
    const relationships = relationshipMapFromDoc(relsDoc);

    const contentTypesFile = zip.file("[Content_Types].xml");
    const contentTypesDoc = contentTypesFile
      ? parseXml(await contentTypesFile.async("text"), "Admin DOCX content types")
      : null;

    const tables = Array.from(doc.getElementsByTagNameNS(W_NS, "tbl"));
    let best = { records: [], score: -Infinity };

    for (const table of tables) {
      const rows = directChildren(table, "tr");
      if (!rows.length) continue;
      const detected = detectBestMapping(rows, relationships);
      if (!detected) continue;

      const records = [];
      let ordinal = 0;
      for (const row of rows) {
        const record = recordFromSourceRow(row, relationships, detected.map, ordinal + 1);
        if (!record) continue;
        ordinal += 1;
        record.sourceOrder = ordinal;
        if (!record.sourceNumber) record.sourceNumber = String(ordinal);
        records.push(record);
      }

      const dataQuality = records.reduce((sum, record) => {
        const validOptions = 4 - record.missingOptions.length;
        return sum + (record.questionMissing ? 0 : 4) + validOptions + (record.answerMissing ? 0 : 3);
      }, 0);
      const tableScore = detected.score + records.length * 20 + dataQuality;
      if (records.length && tableScore > best.score) best = { records, score: tableScore };
    }

    return {
      doc,
      records: best.records,
      package: { zip, relationships, contentTypesDoc }
    };
  }

  function detectBestMapping(rows, relationships) {
    const maxCells = rows.reduce((max, row) => Math.max(max, directChildren(row, "tc").length), 0);
    const candidates = [];
    if (maxCells >= 8) candidates.push({ serial: 0, q: 1, a: 2, b: 3, c: 4, d: 5, ans: 6, exp: 7 });
    if (maxCells >= 7) candidates.push({ serial: null, q: 0, a: 1, b: 2, c: 3, d: 4, ans: 5, exp: 6 });
    if (!candidates.length) return null;

    let best = null;
    for (const map of candidates) {
      let score = 0;
      let considered = 0;
      for (const row of rows) {
        const cells = directChildren(row, "tc");
        if (!mappingFitsCells(map, cells)) continue;
        if (looksLikeHeaderRow(cells, map)) {
          score += 40;
          continue;
        }

        const serialCell = map.serial === null ? null : cells[map.serial];
        const serialText = serialCell ? cellText(serialCell) : "";
        const mappedCells = [cells[map.q], cells[map.a], cells[map.b], cells[map.c], cells[map.d], cells[map.ans], cells[map.exp]];
        const hasAnything = mappedCells.some((cell) => cellHasMeaningfulContent(cell, relationships))
          || cellHasMeaningfulContent(serialCell, relationships);
        if (!hasAnything) continue;

        considered += 1;
        if (cellHasMeaningfulContent(cells[map.q], relationships)) score += 5;
        [map.a, map.b, map.c, map.d].forEach((index) => {
          if (cellHasMeaningfulContent(cells[index], relationships)) score += 2;
        });
        if (cellHasMeaningfulContent(cells[map.ans], relationships)) score += 5;
        if (normalizeAnswer(cellText(cells[map.ans]))) score += 1;
        if (cellHasMeaningfulContent(cells[map.exp], relationships)) score += 1;
        if (map.serial !== null && cellHasMeaningfulContent(serialCell, relationships)) {
          score += /^[0-9০-৯]+[.)।-]*$/.test(serialText) ? 2 : 1;
        }
      }
      score += considered;
      if (!best || score > best.score) best = { map, score };
    }
    return best;
  }

  function mappingFitsCells(map, cells) {
    return Math.max(map.q, map.a, map.b, map.c, map.d, map.ans, map.exp, map.serial === null ? 0 : map.serial) < cells.length;
  }

  function looksLikeHeaderRow(cells, map) {
    if (!mappingFitsCells(map, cells)) return false;
    const q = normalizeHeaderToken(cellText(cells[map.q]));
    const ans = normalizeHeaderToken(cellText(cells[map.ans]));
    const options = [map.a, map.b, map.c, map.d].map((index) => normalizeHeaderToken(cellText(cells[index])));

    const qHeader = q === "q" || q.includes("question") || q.includes("প্রশ্ন");
    const ansHeader = ans.includes("answer") || ans.includes("ans") || ans.includes("উত্তর");
    const optionMatches = options.reduce((count, token, index) => {
      const letter = OPTION_LETTERS[index].toLowerCase();
      const bn = ["ক", "খ", "গ", "ঘ"][index];
      return count + (token === letter || token === `option${letter}` || token === bn || token === `বিকল্প${bn}` ? 1 : 0);
    }, 0);

    return (qHeader && ansHeader) || (optionMatches >= 3 && (qHeader || ansHeader));
  }

  function normalizeHeaderToken(value) {
    return cleanText(value).toLowerCase().replace(/[\s.:()\-_/।]+/g, "");
  }

  function recordFromSourceRow(row, relationships, map, ordinal) {
    const cells = directChildren(row, "tc");
    if (!mappingFitsCells(map, cells) || looksLikeHeaderRow(cells, map)) return null;

    const serialCell = map.serial === null ? null : cells[map.serial];
    const serialText = serialCell ? cellText(serialCell) : "";
    const questionCell = cells[map.q];
    const optionCells = [cells[map.a], cells[map.b], cells[map.c], cells[map.d]];
    const answerCell = cells[map.ans];
    const explanationCell = cells[map.exp];
    const mappedCells = [questionCell, ...optionCells, answerCell, explanationCell];
    const rowHasContent = mappedCells.some((cell) => cellHasMeaningfulContent(cell, relationships))
      || cellHasMeaningfulContent(serialCell, relationships);
    if (!rowHasContent) return null;

    const questionText = cellText(questionCell);
    const optionTexts = optionCells.map((cell) => cellText(cell));
    const answerText = cellText(answerCell);
    const answer = normalizeAnswer(answerText);
    const explanationText = cellText(explanationCell);
    const questionMissing = !cellHasMeaningfulContent(questionCell, relationships);
    const missingOptions = optionCells
      .map((cell, index) => cellHasMeaningfulContent(cell, relationships) ? null : OPTION_LETTERS[index])
      .filter(Boolean);
    const answerMissing = !cellHasMeaningfulContent(answerCell, relationships);
    const explanationMissing = !cellHasMeaningfulContent(explanationCell, relationships);

    const sourceNodes = [questionCell, ...optionCells, answerCell, explanationCell];
    const imageRelIds = collectImageRelationshipIds(sourceNodes, relationships);
    const hasMath = sourceNodes.some((cell) => cell.getElementsByTagNameNS(M_NS, "oMath").length > 0);
    const drawingNodes = sourceNodes.flatMap((cell) => [
      ...Array.from(cell.getElementsByTagNameNS(W_NS, "drawing")),
      ...Array.from(cell.getElementsByTagNameNS(W_NS, "pict"))
    ]);
    const hasUnsupportedDrawing = drawingNodes.some((drawing) =>
      collectImageRelationshipIds([drawing], relationships).length === 0
    );

    return {
      sourceNumber: cleanQuestionNumber(serialText) || String(ordinal),
      sourceOrder: ordinal,
      questionCell,
      optionCells,
      answerCell,
      answer,
      answerText,
      answerWasRemapped: false,
      optionShuffleSkipped: false,
      explanationCell,
      questionText,
      optionTexts,
      explanationText,
      questionMissing,
      missingOptions,
      answerMissing,
      explanationMissing,
      hasMath,
      imageRelIds,
      hasUnsupportedDrawing
    };
  }

  function cellHasMeaningfulContent(cell, relationships) {
    if (!cell) return false;

    // Validation is intentionally cell-based, not character-based. If Word stores
    // something visibly inside the cell (text, Symbol-font glyph, equation, image,
    // object, line break/tab, automatic numbering, etc.), the field is considered present.
    const textNodes = [
      ...Array.from(cell.getElementsByTagNameNS(W_NS, "t")),
      ...Array.from(cell.getElementsByTagNameNS(W_NS, "delText")),
      ...Array.from(cell.getElementsByTagNameNS(W_NS, "instrText")),
      ...Array.from(cell.getElementsByTagNameNS(M_NS, "t"))
    ];
    if (textNodes.some((node) => cleanText(node.textContent || ""))) return true;

    const wordContentElements = [
      "sym", "drawing", "pict", "object", "tab", "br", "cr",
      "noBreakHyphen", "softHyphen", "ptab", "fldSimple", "numPr",
      "footnoteReference", "endnoteReference"
    ];
    if (wordContentElements.some((name) => cell.getElementsByTagNameNS(W_NS, name).length > 0)) return true;

    if (cell.getElementsByTagNameNS(M_NS, "oMath").length || cell.getElementsByTagNameNS(M_NS, "oMathPara").length) return true;
    if (relationships && collectImageRelationshipIds([cell], relationships).length) return true;
    return false;
  }

  function cleanQuestionNumber(value) {
    return cleanText(value).replace(/[.)।:-]+$/g, "").trim();
  }

  function validateSourceRecords(records) {
    const blockingErrors = [];
    const missingExplanations = [];

    for (const record of records) {
      const number = record.sourceNumber || String(record.sourceOrder || "?");
      if (record.questionMissing) {
        blockingErrors.push(`${number} নম্বর প্রশ্নে প্রশ্ন নেই। প্রশ্ন যুক্ত করে Generate করুন।`);
      }
      if (record.missingOptions.length) {
        blockingErrors.push(`${number} নম্বর প্রশ্নে Option ${record.missingOptions.join(", ")} নেই। উক্ত Option যুক্ত করে Generate করুন।`);
      }
      if (record.answerMissing) {
        blockingErrors.push(`${number} নম্বর প্রশ্নে Answer cell ফাঁকা। Answer cell-এ কিছু যুক্ত করে Generate করুন।`);
      }
      if (record.explanationMissing) missingExplanations.push(number);
    }

    return { blockingErrors, missingExplanations };
  }

  function buildBlockingValidationMessage(errors) {
    const limit = 14;
    const shown = errors.slice(0, limit).map((item) => `• ${item}`).join("\n");
    const more = errors.length > limit ? `\n\nআরও ${errors.length - limit}টি সমস্যা আছে। DOCX ঠিক করে আবার upload করুন।` : "";
    return `${shown}${more}`;
  }

  function buildExplanationMessage(numbers) {
    const limit = 20;
    const shown = numbers.slice(0, limit).join(", ");
    const more = numbers.length > limit ? ` এবং আরও ${numbers.length - limit}টি` : "";
    return `${shown}${more} নম্বর প্রশ্নে ব্যাখ্যা নেই।\n\nব্যাখ্যা ছাড়াই প্রশ্নগুলো Generate করতে চান?`;
  }

  async function generateSets(mode) {
    if (busy) return;
    if (!sourceDoc || !sourceRecords.length) {
      setStatus("আগে একটি valid Admin DOCX upload করুন।", "error");
      await showMessageModal({ title: "Question Bank প্রয়োজন", message: "আগে একটি valid Admin DOCX upload করুন।", tone: "error" });
      return;
    }

    if (![GENERATION_MODE.QUESTIONS, GENERATION_MODE.QUESTIONS_OPTIONS].includes(mode)) return;

    sourceValidation = validateSourceRecords(sourceRecords);
    if (sourceValidation.blockingErrors.length) {
      setGenerateButtonsDisabled(true);
      setStatus("Question/Option/Answer-এর সমস্যা ঠিক না করা পর্যন্ত Generate করা যাবে না।", "error");
      await showMessageModal({
        title: "Generate করা যাবে না",
        message: buildBlockingValidationMessage(sourceValidation.blockingErrors),
        tone: "error",
        confirmText: "ঠিক আছে"
      });
      return;
    }

    if (sourceValidation.missingExplanations.length) {
      const confirmed = await showMessageModal({
        title: "ব্যাখ্যা পাওয়া যায়নি",
        message: buildExplanationMessage(sourceValidation.missingExplanations),
        tone: "confirm",
        confirmText: "হ্যাঁ, Generate করুন",
        cancelText: "না",
        showCancel: true
      });
      if (!confirmed) {
        setStatus("Generate বাতিল করা হয়েছে। প্রয়োজন হলে ব্যাখ্যা যোগ করে আবার চেষ্টা করুন।", "warning");
        return;
      }
    }

    setBusy(true);
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
      const imageCount = sourceRecords.filter((record) => record.imageRelIds && record.imageRelIds.length).length;
      const modeLabel = mode === GENERATION_MODE.QUESTIONS_OPTIONS ? "প্রশ্ন + Option shuffle" : "শুধু প্রশ্ন shuffle";
      const unparsedAnswerCount = mode === GENERATION_MODE.QUESTIONS_OPTIONS
        ? sourceRecords.filter((record) => !record.answerMissing && !OPTION_LETTERS.includes(record.answer)).length
        : 0;

      lastGenerationLabel = sanitizeFilename([exam, subject, `${setCount}-Sets`].filter(Boolean).join("-"));
      setStatus(`${setCount}টি Set তৈরি হচ্ছে · ${modeLabel} · প্রতি Set-এ ${questionsPerSet}টি প্রশ্ন...`);

      generatedFiles = [];
      setGrid.innerHTML = "";

      // Keep every generated Set different from the Admin file order and from
      // every other generated Set. A source question must not keep the same
      // numeric serial position in the generated paper (e.g. source 01 -> output 01).
      const usedQuestionOrders = new Set();

      for (let index = 0; index < setCount; index += 1) {
        const setLabel = indexToSetLabel(index);
        let shuffled = createDistinctQuestionOrder(sourceRecords, questionsPerSet, usedQuestionOrders);
        if (mode === GENERATION_MODE.QUESTIONS_OPTIONS) {
          shuffled = shuffled.map((record) => shuffleRecordOptions(record));
        }
        const blob = await buildSetDocx(setLabel, shuffled);
        const name = makeSetFilename(exam, subject, setLabel);
        const file = { blob, name, setLabel, count: shuffled.length, mode };
        generatedFiles.push(file);
        renderSetCard(file);

        setStatus(`Set ${setLabel} তৈরি হয়েছে (${index + 1}/${setCount})...`);
        await nextFrame();
      }

      const answerNote = unparsedAnswerCount
        ? ` · ${unparsedAnswerCount}টি Answer cell-এ A/B/C/D শনাক্ত না হওয়ায় ওই প্রশ্নগুলোর option order অপরিবর্তিত রাখা হয়েছে`
        : "";
      resultsNote.textContent = `${sourceRecords.length}টি source প্রশ্ন থেকে ${setCount}টি আলাদা Set তৈরি হয়েছে · প্রতি Set-এ ${questionsPerSet}টি প্রশ্ন · source serial ও generated serial একই রাখা হয়নি · Mode: ${modeLabel} · Answer ও Explanation সঠিক প্রশ্নের সঙ্গে রাখা হয়েছে${answerNote}।`;
      resultsEl.classList.add("visible");
      downloadAllBtn.disabled = false;

      if (unparsedAnswerCount) {
        setStatus(`✓ সম্পন্ন। ${unparsedAnswerCount}টি Answer cell-এ content আছে, কিন্তু A/B/C/D শনাক্ত করা যায়নি; তাই সেসব প্রশ্নে option shuffle না করে source option order রাখা হয়েছে।`, "warning");
      } else if (warningCount) {
        setStatus(`✓ সম্পন্ন। ${modeLabel} ব্যবহার হয়েছে। Embedded image copy করা হয়েছে। তবে ${warningCount}টি source row-এ non-image drawing/shape/chart আছে; সেগুলো পুরোপুরি copy নাও হতে পারে।`, "warning");
      } else if (imageCount) {
        setStatus(`✓ সম্পন্ন। ${modeLabel} ব্যবহার হয়েছে এবং ${imageCount}টি source row-এর embedded image generated DOCX-এ copy করা হয়েছে।`);
      } else {
        setStatus(`✓ সম্পন্ন। ${modeLabel} ব্যবহার করে ${setCount}টি DOCX Set প্রস্তুত।`);
      }

      resultsEl.scrollIntoView({ behavior: "smooth", block: "start" });
    } catch (error) {
      console.error(error);
      const message = error && error.message ? error.message : "Set তৈরির সময় সমস্যা হয়েছে।";
      setStatus(message, "error");
      await showMessageModal({ title: "Generate ব্যর্থ হয়েছে", message, tone: "error", confirmText: "ঠিক আছে" });
    } finally {
      setBusy(false);
    }
  }

  function createDistinctQuestionOrder(records, takeCount, usedOrders) {
    const source = Array.isArray(records) ? records.slice() : [];
    const count = clamp(positiveInt(takeCount, source.length), 1, source.length);
    if (!source.length) return [];

    const maxAttempts = Math.max(2500, source.length * 60);

    // Random attempts preserve the existing shuffle behaviour while enforcing:
    // 1) no source serial is kept at the same output serial, and
    // 2) no two Sets use the same question sequence.
    for (let attempt = 0; attempt < maxAttempts; attempt += 1) {
      const candidate = secureShuffle(source.slice()).slice(0, count);
      if (questionOrderHasSerialCollision(candidate)) continue;

      const key = questionOrderKey(candidate);
      if (usedOrders.has(key)) continue;

      usedOrders.add(key);
      return candidate;
    }

    // Deterministic rotation fallback. This is especially useful for smaller
    // banks where random retries may repeatedly hit an already-used order.
    for (let shift = 1; shift < source.length; shift += 1) {
      const rotated = source.slice(shift).concat(source.slice(0, shift)).slice(0, count);
      if (questionOrderHasSerialCollision(rotated)) continue;

      const key = questionOrderKey(rotated);
      if (usedOrders.has(key)) continue;

      usedOrders.add(key);
      return rotated;
    }

    throw new Error(
      "Admin serial-এর সঙ্গে generated serial না মেলানো এবং প্রতিটি Set-এর question order আলাদা রাখার শর্তে আর নতুন Set তৈরি করা সম্ভব হয়নি। Set সংখ্যা কমান অথবা Question Bank-এ আরও প্রশ্ন যোগ করুন।"
    );
  }

  function questionOrderHasSerialCollision(records) {
    return records.some((record, index) => {
      const outputNumber = index + 1;
      const sourceNumber = sourceSerialNumber(record);
      return sourceNumber === outputNumber;
    });
  }

  function questionOrderKey(records) {
    return records.map((record) => String(record.sourceOrder || record.sourceNumber || "?")).join("|");
  }

  function sourceSerialNumber(record) {
    if (!record) return null;

    const fromCell = parseQuestionSerial(record.sourceNumber);
    if (fromCell !== null) return fromCell;

    const fallback = Number(record.sourceOrder);
    return Number.isInteger(fallback) && fallback > 0 ? fallback : null;
  }

  function parseQuestionSerial(value) {
    const englishDigits = String(value == null ? "" : value)
      .replace(/[০-৯]/g, (digit) => String("০১২৩৪৫৬৭৮৯".indexOf(digit)))
      .trim();

    const match = englishDigits.match(/^0*(\d+)$/);
    if (!match) return null;

    const number = Number.parseInt(match[1], 10);
    return Number.isInteger(number) && number > 0 ? number : null;
  }

  function shuffleRecordOptions(record) {
    const correctOriginalIndex = OPTION_LETTERS.indexOf(record.answer);

    // The Answer cell is allowed to contain any Word content. If that content cannot
    // be interpreted as A/B/C/D, keep this question's option order unchanged so the
    // source answer is never corrupted merely to satisfy option-shuffle mode.
    if (correctOriginalIndex < 0) {
      return {
        ...record,
        answerWasRemapped: false,
        optionShuffleSkipped: true
      };
    }

    const order = secureShuffle([0, 1, 2, 3]);
    const correctNewIndex = order.indexOf(correctOriginalIndex);
    return {
      ...record,
      optionCells: order.map((index) => record.optionCells[index]),
      optionTexts: order.map((index) => record.optionTexts[index]),
      answer: OPTION_LETTERS[correctNewIndex],
      answerWasRemapped: true,
      optionShuffleSkipped: false
    };
  }

  function showMessageModal({
    title,
    message,
    tone = "warning",
    confirmText = "ঠিক আছে",
    cancelText = "না",
    showCancel = false
  }) {
    if (activeModalResolve) finishModal(false);
    modalIsConfirmation = showCancel;
    modalCard.className = `qg-modal-card ${tone}`;
    modalIcon.textContent = tone === "error" ? "!" : tone === "confirm" ? "?" : "i";
    modalTitle.textContent = title || "বার্তা";
    modalMessage.textContent = message || "";
    modalConfirmBtn.textContent = confirmText;
    modalCancelBtn.textContent = cancelText;
    modalCancelBtn.hidden = !showCancel;
    messageModal.hidden = false;
    document.body.classList.add("modal-open");
    setTimeout(() => modalConfirmBtn.focus(), 0);

    return new Promise((resolve) => {
      activeModalResolve = resolve;
    });
  }

  function finishModal(result) {
    if (messageModal.hidden) return;
    messageModal.hidden = true;
    document.body.classList.remove("modal-open");
    const resolve = activeModalResolve;
    activeModalResolve = null;
    modalIsConfirmation = false;
    if (resolve) resolve(Boolean(result));
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

    const relsFile = zip.file("word/_rels/document.xml.rels");
    const relsDoc = relsFile
      ? parseXml(await relsFile.async("text"), "Template DOCX relationships")
      : createRelationshipsDocument();

    const contentTypesFile = zip.file("[Content_Types].xml");
    if (!contentTypesFile) throw new Error("Template DOCX invalid: [Content_Types].xml নেই।");
    const contentTypesDoc = parseXml(await contentTypesFile.async("text"), "Template DOCX content types");

    const relationshipIdMap = await copyImagesForRecords(zip, records, relsDoc, contentTypesDoc);
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
      fillQuestionRow(doc, questionRow, record, index + 1, relationshipIdMap);
      questionTable.appendChild(questionRow);

      const answerRow = answerPrototype.cloneNode(true);
      fillAnswerRow(doc, answerRow, record, index + 1, relationshipIdMap);
      answerTable.appendChild(answerRow);
    });

    replaceAllText(doc, "{{SET}}", setLabel);

    const serialized = new XMLSerializer().serializeToString(doc);
    zip.file("word/document.xml", ensureXmlDeclaration(serialized));
    zip.file("word/_rels/document.xml.rels", ensureXmlDeclaration(new XMLSerializer().serializeToString(relsDoc)));
    zip.file("[Content_Types].xml", ensureXmlDeclaration(new XMLSerializer().serializeToString(contentTypesDoc)));

    return zip.generateAsync({
      type: "blob",
      mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      compression: "DEFLATE",
      compressionOptions: { level: 6 }
    });
  }

  function fillQuestionRow(doc, row, record, number, relationshipIdMap) {
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

    copyCellBody(doc, record.questionCell, questionCell, relationshipIdMap);
    copyCellBody(doc, record.optionCells[0], optionRowAB[1], relationshipIdMap);
    copyCellBody(doc, record.optionCells[1], optionRowAB[3], relationshipIdMap);
    copyCellBody(doc, record.optionCells[2], optionRowCD[1], relationshipIdMap);
    copyCellBody(doc, record.optionCells[3], optionRowCD[3], relationshipIdMap);
  }

  function fillAnswerRow(doc, row, record, number, relationshipIdMap) {
    const cells = directChildren(row, "tc");
    if (cells.length < 3) throw new Error("Template answer row invalid।");
    setCellText(cells[0], `${String(number).padStart(2, "0")}.`);

    if (record.answerWasRemapped) {
      setCellText(cells[1], record.answer);
    } else if (record.answerCell) {
      copyCellBody(doc, record.answerCell, cells[1], relationshipIdMap);
    } else {
      setCellText(cells[1], record.answer || record.answerText || "");
    }

    copyCellBody(doc, record.explanationCell, cells[2], relationshipIdMap);
  }

  function copyCellBody(targetDoc, sourceCell, targetCell, relationshipIdMap) {
    const preservedTcPr = directChildren(targetCell, "tcPr")[0] || null;
    while (targetCell.firstChild) targetCell.removeChild(targetCell.firstChild);
    if (preservedTcPr) targetCell.appendChild(preservedTcPr);

    const sourceChildren = Array.from(sourceCell.childNodes).filter((node) => {
      return !(node.nodeType === 1 && node.namespaceURI === W_NS && node.localName === "tcPr");
    });

    for (const child of sourceChildren) {
      const imported = targetDoc.importNode(child, true);
      sanitizeImportedContent(imported, relationshipIdMap);
      targetCell.appendChild(imported);
    }

    if (!directChildren(targetCell, "p").length && !directChildren(targetCell, "tbl").length) {
      targetCell.appendChild(targetDoc.createElementNS(W_NS, "w:p"));
    }
  }

  function sanitizeImportedContent(root, relationshipIdMap) {
    if (!root || root.nodeType !== 1) return;

    const drawings = [
      ...(root.namespaceURI === W_NS && (root.localName === "drawing" || root.localName === "pict") ? [root] : []),
      ...Array.from(root.getElementsByTagNameNS(W_NS, "drawing")),
      ...Array.from(root.getElementsByTagNameNS(W_NS, "pict"))
    ];
    for (const drawing of drawings) {
      const elementsInDrawing = [drawing, ...Array.from(drawing.getElementsByTagName("*"))];
      const relationshipAttrs = elementsInDrawing.flatMap((element) =>
        Array.from(element.attributes || []).filter((attr) => attr.namespaceURI === R_NS)
      );
      const hasMappedRelationship = relationshipAttrs.some((attr) => relationshipIdMap && relationshipIdMap.has(attr.value));
      const hasUnmappedRelationship = relationshipAttrs.some((attr) => !relationshipIdMap || !relationshipIdMap.has(attr.value));
      if (!hasMappedRelationship && hasUnmappedRelationship) drawing.remove();
    }

    const attrsToDrop = ["paraId", "textId"];
    const elements = [root, ...Array.from(root.getElementsByTagName("*"))];
    for (const element of elements) {
      const attrs = Array.from(element.attributes || []);
      for (const attr of attrs) {
        if (attr.namespaceURI === R_NS) {
          const mappedId = relationshipIdMap && relationshipIdMap.get(attr.value);
          if (mappedId) attr.value = mappedId;
          else element.removeAttributeNode(attr);
        }
      }

      for (const name of attrsToDrop) {
        const matches = Array.from(element.attributes || []).filter((attr) => attr.localName === name);
        matches.forEach((attr) => element.removeAttributeNode(attr));
      }
    }
  }

  function createRelationshipsDocument() {
    return parseXml(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${PKG_REL_NS}"></Relationships>`,
      "DOCX relationships"
    );
  }

  function relationshipMapFromDoc(doc) {
    const map = new Map();
    if (!doc) return map;

    const relationships = Array.from(doc.getElementsByTagNameNS(PKG_REL_NS, "Relationship"));
    for (const node of relationships) {
      const id = node.getAttribute("Id");
      if (!id) continue;
      map.set(id, {
        id,
        type: node.getAttribute("Type") || "",
        target: node.getAttribute("Target") || "",
        targetMode: node.getAttribute("TargetMode") || ""
      });
    }
    return map;
  }

  function isImageRelationship(relationship) {
    return Boolean(relationship && /\/image$/i.test(relationship.type || ""));
  }

  function collectImageRelationshipIds(nodes, relationships) {
    const ids = new Set();
    for (const node of nodes || []) {
      if (!node || node.nodeType !== 1) continue;
      const elements = [node, ...Array.from(node.getElementsByTagName("*"))];
      for (const element of elements) {
        for (const attr of Array.from(element.attributes || [])) {
          if (attr.namespaceURI !== R_NS) continue;
          const relationship = relationships.get(attr.value);
          if (isImageRelationship(relationship)) ids.add(attr.value);
        }
      }
    }
    return Array.from(ids);
  }

  async function copyImagesForRecords(targetZip, records, targetRelsDoc, targetContentTypesDoc) {
    const idMap = new Map();
    if (!sourcePackage || !sourcePackage.zip || !sourcePackage.relationships) return idMap;

    const sourceIds = new Set();
    for (const record of records) {
      for (const id of record.imageRelIds || []) sourceIds.add(id);
    }

    for (const sourceId of sourceIds) {
      const sourceRelationship = sourcePackage.relationships.get(sourceId);
      if (!isImageRelationship(sourceRelationship)) continue;

      const newId = nextRelationshipId(targetRelsDoc);

      if (/^external$/i.test(sourceRelationship.targetMode || "")) {
        appendRelationship(targetRelsDoc, {
          id: newId,
          type: sourceRelationship.type,
          target: sourceRelationship.target,
          targetMode: "External"
        });
        idMap.set(sourceId, newId);
        continue;
      }

      const sourcePart = resolvePackageTarget("word", sourceRelationship.target);
      const sourceEntry = sourcePackage.zip.file(sourcePart);
      if (!sourceEntry) {
        throw new Error(`Embedded image পাওয়া গেছে, কিন্তু DOCX-এর ${sourcePart} file-টি পাওয়া যায়নি।`);
      }

      const targetPart = uniqueTargetMediaPart(targetZip, sourcePart);
      const imageBytes = await sourceEntry.async("uint8array");
      targetZip.file(targetPart, imageBytes);

      const contentType = contentTypeForPart(sourcePackage.contentTypesDoc, sourcePart) || fallbackImageContentType(sourcePart);
      ensureContentTypeOverride(targetContentTypesDoc, targetPart, contentType);

      appendRelationship(targetRelsDoc, {
        id: newId,
        type: sourceRelationship.type,
        target: targetPart.replace(/^word\//, "")
      });
      idMap.set(sourceId, newId);
    }

    return idMap;
  }

  function nextRelationshipId(relsDoc) {
    const used = new Set(
      Array.from(relsDoc.getElementsByTagNameNS(PKG_REL_NS, "Relationship"))
        .map((node) => node.getAttribute("Id"))
        .filter(Boolean)
    );
    let index = 1;
    while (used.has(`rIdQG${index}`)) index += 1;
    return `rIdQG${index}`;
  }

  function appendRelationship(relsDoc, relationship) {
    const root = relsDoc.documentElement;
    const node = relsDoc.createElementNS(PKG_REL_NS, "Relationship");
    node.setAttribute("Id", relationship.id);
    node.setAttribute("Type", relationship.type);
    node.setAttribute("Target", relationship.target);
    if (relationship.targetMode) node.setAttribute("TargetMode", relationship.targetMode);
    root.appendChild(node);
  }

  function resolvePackageTarget(baseDir, target) {
    let value = String(target || "").replace(/\\/g, "/");
    try { value = decodeURI(value); } catch (_) { /* keep original path */ }

    const absolute = value.startsWith("/");
    const pieces = (absolute ? value.slice(1) : `${baseDir}/${value}`).split("/");
    const normalized = [];
    for (const piece of pieces) {
      if (!piece || piece === ".") continue;
      if (piece === "..") {
        normalized.pop();
        continue;
      }
      normalized.push(piece);
    }
    return normalized.join("/");
  }

  function uniqueTargetMediaPart(zip, sourcePart) {
    const sourceName = String(sourcePart || "image.bin").split("/").pop() || "image.bin";
    const safeName = sourceName.replace(/[^A-Za-z0-9._-]/g, "-");
    let index = 1;
    let candidate = `word/media/qg-${index}-${safeName}`;
    while (zip.file(candidate)) {
      index += 1;
      candidate = `word/media/qg-${index}-${safeName}`;
    }
    return candidate;
  }

  function contentTypeForPart(contentTypesDoc, partPath) {
    if (!contentTypesDoc) return "";
    const normalizedPart = `/${String(partPath || "").replace(/^\/+/, "")}`;

    const overrides = Array.from(contentTypesDoc.getElementsByTagNameNS(CT_NS, "Override"));
    for (const node of overrides) {
      if (normalizePartName(node.getAttribute("PartName")) === normalizePartName(normalizedPart)) {
        return node.getAttribute("ContentType") || "";
      }
    }

    const extension = fileExtension(partPath);
    if (!extension) return "";
    const defaults = Array.from(contentTypesDoc.getElementsByTagNameNS(CT_NS, "Default"));
    for (const node of defaults) {
      if ((node.getAttribute("Extension") || "").toLowerCase() === extension) {
        return node.getAttribute("ContentType") || "";
      }
    }
    return "";
  }

  function ensureContentTypeOverride(contentTypesDoc, targetPart, contentType) {
    const partName = `/${String(targetPart || "").replace(/^\/+/, "")}`;
    const normalized = normalizePartName(partName);
    const existing = Array.from(contentTypesDoc.getElementsByTagNameNS(CT_NS, "Override"))
      .some((node) => normalizePartName(node.getAttribute("PartName")) === normalized);
    if (existing) return;

    const node = contentTypesDoc.createElementNS(CT_NS, "Override");
    node.setAttribute("PartName", partName);
    node.setAttribute("ContentType", contentType);
    contentTypesDoc.documentElement.appendChild(node);
  }

  function normalizePartName(value) {
    let text = String(value || "");
    try { text = decodeURI(text); } catch (_) { /* keep original */ }
    return text.replace(/\\/g, "/").toLowerCase();
  }

  function fileExtension(path) {
    const name = String(path || "").split("/").pop() || "";
    const dot = name.lastIndexOf(".");
    return dot >= 0 ? name.slice(dot + 1).toLowerCase() : "";
  }

  function fallbackImageContentType(path) {
    const extension = fileExtension(path);
    const types = {
      png: "image/png",
      jpg: "image/jpeg",
      jpeg: "image/jpeg",
      gif: "image/gif",
      bmp: "image/bmp",
      tif: "image/tiff",
      tiff: "image/tiff",
      svg: "image/svg+xml",
      emf: "image/x-emf",
      wmf: "image/x-wmf"
    };
    return types[extension] || "application/octet-stream";
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
    const modeText = file.mode === GENERATION_MODE.QUESTIONS_OPTIONS
      ? "প্রশ্ন + Option shuffle"
      : "শুধু প্রশ্ন shuffle";
    article.innerHTML = `
      <div class="set-badge">SET ${escapeHtml(file.setLabel)}</div>
      <h3>Set ${escapeHtml(file.setLabel)}</h3>
      <p>${file.count}টি প্রশ্ন · ${escapeHtml(modeText)} · Answer/Explanation সংযুক্ত</p>
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

  function setGenerateButtonsDisabled(disabled) {
    generateQuestionsBtn.disabled = disabled;
    generateQuestionsOptionsBtn.disabled = disabled;
  }

  function setBusy(value) {
    busy = value;
    fileInput.disabled = value;
    resetBtn.disabled = value;
    const cannotGenerate = value || !sourceRecords.length || sourceValidation.blockingErrors.length > 0;
    setGenerateButtonsDisabled(cannotGenerate);
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
