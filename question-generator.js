(() => {
  "use strict";

  const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math";
  const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
  const CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
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
  const historyList = $("historyList");
  const previewDialog = $("previewDialog");
  const previewContent = $("previewContent");
  const issueDialog = $("issueDialog");
  const validationIssues = $("validationIssues");
  const HISTORY_DB = "rahman-question-generator";

  let selectedFile = null;
  let sourceDoc = null;
  let sourcePackage = null;
  let sourceRecords = [];
  let sourceSlots = [];
  let sourceIssues = [];
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
  $("closePreview").addEventListener("click", () => previewDialog.close());
  $("closeIssues").addEventListener("click", () => issueDialog.close());
  issueDialog.addEventListener("click", (event) => { if (event.target === issueDialog) issueDialog.close(); });
  previewDialog.addEventListener("click", (event) => { if (event.target === previewDialog) previewDialog.close(); });
  renderHistory();

  async function selectFile(file, fromDrop = false, savedSettings = null, autoGenerate = false) {
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
    await prepareSourceAndAutoGenerate(savedSettings, autoGenerate);
  }

  function clearSource() {
    selectedFile = null;
    sourceDoc = null;
    sourcePackage = null;
    sourceRecords = [];
    sourceSlots = [];
    sourceIssues = [];
    validationIssues.hidden = true;
    validationIssues.innerHTML = "";
    if (issueDialog.open) issueDialog.close();
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

  async function prepareSourceAndAutoGenerate(savedSettings = null, autoGenerate = false) {
    setBusy(true);
    resetResults();
    sourceIssues = [];
    validationIssues.hidden = true;
    validationIssues.innerHTML = "";
    if (issueDialog.open) issueDialog.close();
    debugEl.style.display = "none";
    debugEl.textContent = "";

    try {
      ensureJSZip();
      setStatus("Admin DOCX table পড়া হচ্ছে...");

      const bytes = await selectedFile.arrayBuffer();
      const parsed = await parseQuestionBank(bytes);
      sourceDoc = parsed.doc;
      sourcePackage = parsed.package;
      sourceRecords = parsed.records;
      sourceSlots = parsed.slots;
      sourceIssues = parsed.issues;

      if (!sourceSlots.length) {
        throw new Error("Question | A | B | C | D | Answer | Explanation structure-এর কোনো question row পাওয়া যায়নি।");
      }

      const count = sourceRecords.length;
      const rowCount = sourceSlots.length;
      $("questionsPerSet").removeAttribute("max");
      $("setCount").value = String(DEFAULT_SET_COUNT);

      const derivedExam = deriveExamName(selectedFile.name);
      if (derivedExam) $("examName").value = derivedExam;
      if (savedSettings) {
        $("examName").value = savedSettings.exam;
        $("subjectName").value = savedSettings.subject;
        $("setCount").value = String(savedSettings.sets.length);
        $("questionsPerSet").value = String(savedSettings.questionsPerSet);
        $("shuffleOptions").value = savedSettings.shuffleOptions === true || savedSettings.shuffleOptions === "yes" ? "yes" : "no";
      }

      const mathCount = sourceRecords.filter((record) => record.hasMath).length;
      const imageCount = sourceRecords.filter((record) => record.imageRelIds && record.imageRelIds.length).length;
      detectedInfoEl.hidden = false;
      detectedInfoEl.textContent = `✓ ${rowCount}টি question row শনাক্ত হয়েছে · ${count}টি সম্পূর্ণ valid${sourceIssues.length ? ` · ${sourceIssues.length}টিতে সংশোধন প্রয়োজন` : ""}${mathCount ? ` · ${mathCount}টি প্রশ্নে Word equation আছে` : ""}${imageCount ? ` · ${imageCount}টি প্রশ্নে embedded image আছে` : ""}`;
      generateBtn.disabled = false;

      if (autoGenerate) {
        setStatus(`✓ ${rowCount}টি question row পাওয়া গেছে। Generate validation চলছে...`);
        await generateSets(true);
      } else {
        setStatus(`✓ ${rowCount}টি question row পাওয়া গেছে · ${count}টি সম্পূর্ণ valid। প্রয়োজনমতো setting বদলে Generate করুন।`, sourceIssues.length ? "warning" : "");
        $("generatorTitle").scrollIntoView({ behavior: "smooth" });
      }
    } catch (error) {
      console.error(error);
      sourceDoc = null;
      sourcePackage = null;
      sourceRecords = [];
      sourceSlots = [];
      sourceIssues = [];
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

    let bestRecords = [];
    let bestIssues = [];
    let bestSlots = [];
    for (const table of tables) {
      const rows = directChildren(table, "tr");
      const records = [];
      const issues = [];
      const slots = [];
      for (const [index, row] of rows.entries()) {
        const result = recordFromSourceRow(row, relationships, index + 1);
        if (!result) continue;

        const ordinal = slots.length + 1;
        const number = result.serial || String(ordinal);
        const numericSerial = /^\d+$/.test(String(result.serial || "")) ? Number.parseInt(result.serial, 10) : null;

        if (result.record) {
          result.record.sourceNumber = number;
          result.record.sourceOrdinal = ordinal;
          result.record.sourceNumericSerial = numericSerial;
          records.push(result.record);
        }
        if (result.issue) {
          result.issue.number = number;
          result.issue.ordinal = ordinal;
          result.issue.numericSerial = numericSerial;
          issues.push(result.issue);
        }

        slots.push({
          ordinal,
          number,
          numericSerial,
          rowNumber: index + 1,
          record: result.record || null,
          issue: result.issue || null
        });
      }

      if (slots.length > bestSlots.length) {
        bestRecords = records;
        bestIssues = issues;
        bestSlots = slots;
      }
    }

    return {
      doc,
      records: bestRecords,
      issues: bestIssues,
      slots: bestSlots,
      package: { zip, relationships, contentTypesDoc }
    };
  }

  function showValidationIssues(issues, requestedCount = null) {
    const list = issues.map((issue) => {
      const number = escapeHtml(issue.number || issue.ordinal || "?");
      const rowInfo = issue.rowNumber ? ` <small>(DOCX সারি ${escapeHtml(issue.rowNumber)})</small>` : "";
      const details = [];

      if (issue.missingWholeQuestion) details.push(`প্রশ্ন ${number} পাওয়া যায়নি।`);
      if (issue.missingQuestion) details.push(`প্রশ্ন ${number}-এর Question সেল খালি।`);
      if (issue.missingOptions && issue.missingOptions.length) {
        details.push(`প্রশ্ন ${number}-এর Option ${issue.missingOptions.map(escapeHtml).join(", ")} পাওয়া যায়নি।`);
      }
      if (issue.invalidAnswer) {
        details.push(`প্রশ্ন ${number}-এর Answer সেলে শুধু A, B, C বা D থাকতে হবে${issue.answer ? ` (বর্তমানে “${escapeHtml(issue.answer)}”)` : " (সেল খালি)"}।`);
      }
      if (issue.validCountShort) {
        details.push(`চাওয়া ${escapeHtml(issue.requested)}টি প্রশ্নের জন্য মাত্র ${escapeHtml(issue.available)}টি সম্পূর্ণ valid প্রশ্ন পাওয়া গেছে।`);
      }

      return `<li><strong>${details.join(" ")}</strong>${rowInfo}</li>`;
    }).join("");

    const intro = requestedCount
      ? `<p>প্রতি Set-এ ${escapeHtml(requestedCount)}টি প্রশ্ন জেনারেট করতে হলে প্রয়োজনীয় প্রতিটি প্রশ্নের Question সেল এবং A, B, C, D — চারটি Option সেলেই content থাকতে হবে। কোনোটি খালি থাকলে Generate বন্ধ থাকবে।</p>`
      : `<p>Admin DOCX-এ কিছু প্রশ্ন/অপশন/উত্তর সংশোধন প্রয়োজন।</p>`;
    const message = `${intro}<p><strong>মূল Admin file-এ প্রয়োজনীয় প্রশ্ন/অপশন যুক্ত বা সংশোধন করে আবার Generate করুন।</strong></p><ol>${list}</ol>`;

    validationIssues.innerHTML = `<h3>সংশোধন প্রয়োজন</h3>${message}`;
    validationIssues.hidden = false;
    $("issueTitle").textContent = "প্রশ্ন/অপশনে সমস্যা পাওয়া গেছে";
    $("issueContent").innerHTML = message;
    if (!issueDialog.open) issueDialog.showModal();
  }

  function recordFromSourceRow(row, relationships, rowNumber) {
    const cells = directChildren(row, "tc");
    if (cells.length < 7) return null;

    const map = cells.length >= 8
      ? { q: 1, a: 2, b: 3, c: 4, d: 5, ans: 6, exp: 7 }
      : { q: 0, a: 1, b: 2, c: 3, d: 4, ans: 5, exp: 6 };
    if (Math.max(...Object.values(map)) >= cells.length) return null;

    const questionCell = cells[map.q];
    const optionCells = [cells[map.a], cells[map.b], cells[map.c], cells[map.d]];
    const answerCell = cells[map.ans];
    const explanationCell = cells[map.exp];

    const question = cellText(questionCell);
    const rawAnswer = cellText(answerCell);
    const answer = normalizeAnswer(rawAnswer);
    const explanation = cellText(explanationCell);
    const serialRaw = map.q === 1 ? cellText(cells[0]) : "";
    const serial = /^\d+[.)।]?$/.test(serialRaw) ? serialRaw.replace(/[.)।]$/, "") : "";

    // Header row should never be treated as a question row.
    if (/^(question|প্রশ্ন)$/i.test(question) && /^(answer|উত্তর|সঠিক উত্তর)$/i.test(rawAnswer)) return null;

    const questionHasContent = cellHasSubstantiveContent(questionCell);
    const optionHasContent = optionCells.map(cellHasSubstantiveContent);
    const answerHasContent = cellHasSubstantiveContent(answerCell);
    const explanationHasContent = cellHasSubstantiveContent(explanationCell);
    const anyData = Boolean(serial || questionHasContent || optionHasContent.some(Boolean) || answerHasContent || explanationHasContent);
    if (!anyData) return null;

    const missingOptions = ["A", "B", "C", "D"].filter((_, index) => !optionHasContent[index]);
    const issue = {
      serial,
      rowNumber,
      question: question || "[Question সেল খালি]",
      answer: rawAnswer,
      missingQuestion: !questionHasContent,
      missingOptions,
      invalidAnswer: !answer
    };

    if (issue.missingQuestion || issue.missingOptions.length || issue.invalidAnswer) {
      return { serial, issue };
    }

    const sourceNodes = [questionCell, ...optionCells, explanationCell];
    const imageRelIds = collectImageRelationshipIds(sourceNodes, relationships);
    const hasMath = sourceNodes.some((cell) =>
      cell.getElementsByTagNameNS(M_NS, "oMath").length > 0 ||
      cell.getElementsByTagNameNS(M_NS, "oMathPara").length > 0
    );
    const drawingNodes = sourceNodes.flatMap((cell) => [
      ...Array.from(cell.getElementsByTagNameNS(W_NS, "drawing")),
      ...Array.from(cell.getElementsByTagNameNS(W_NS, "pict"))
    ]);
    const hasUnsupportedDrawing = drawingNodes.some((drawing) =>
      collectImageRelationshipIds([drawing], relationships).length === 0
    );

    return { serial, record: {
      questionCell,
      optionCells,
      answer,
      explanationCell,
      questionText: question,
      explanationText: explanation,
      hasMath,
      imageRelIds,
      hasUnsupportedDrawing
    } };
  }

  function cellHasSubstantiveContent(cell) {
    if (!cell) return false;
    if (cleanText(cell.textContent || "")) return true;

    // A cell containing an equation, image, drawing, embedded object, symbol or field
    // counts as non-empty even when it has no plain text.
    const checks = [
      [M_NS, "oMath"], [M_NS, "oMathPara"],
      [W_NS, "drawing"], [W_NS, "pict"], [W_NS, "object"],
      [W_NS, "sym"], [W_NS, "fldSimple"], [W_NS, "instrText"]
    ];
    return checks.some(([ns, name]) => cell.getElementsByTagNameNS(ns, name).length > 0);
  }

  function validateGenerationRequest(requestedCount) {
    const blockers = [];
    if (!sourceSlots.length) {
      blockers.push({ number: "1", missingWholeQuestion: true });
      return blockers;
    }

    const numericSlots = sourceSlots.filter((slot) => Number.isInteger(slot.numericSerial) && slot.numericSerial > 0);
    const useSerialNumbers = numericSlots.length >= Math.max(1, Math.ceil(sourceSlots.length * 0.6));

    if (useSerialNumbers) {
      const bySerial = new Map();
      for (const slot of numericSlots) {
        if (!bySerial.has(slot.numericSerial)) bySerial.set(slot.numericSerial, slot);
      }
      for (let number = 1; number <= requestedCount; number += 1) {
        const slot = bySerial.get(number);
        if (!slot) {
          blockers.push({ number: String(number), missingWholeQuestion: true });
        } else if (slot.issue) {
          blockers.push(slot.issue);
        }
      }
    } else {
      for (let index = 0; index < requestedCount; index += 1) {
        const slot = sourceSlots[index];
        if (!slot) blockers.push({ number: String(index + 1), missingWholeQuestion: true });
        else if (slot.issue) blockers.push(slot.issue);
      }
    }

    if (sourceRecords.length < requestedCount && !blockers.some((item) => item.validCountShort)) {
      blockers.push({
        number: "—",
        validCountShort: true,
        requested: requestedCount,
        available: sourceRecords.length
      });
    }

    return blockers;
  }

  function shuffledOptionRecord(record) {
    const order = secureShuffle([0, 1, 2, 3]);
    const oldCorrectIndex = "ABCD".indexOf(record.answer);
    const newCorrectIndex = order.indexOf(oldCorrectIndex);
    return {
      ...record,
      optionCells: order.map((index) => record.optionCells[index]),
      answer: newCorrectIndex >= 0 ? "ABCD"[newCorrectIndex] : record.answer
    };
  }

  async function generateSets(isAutomatic) {
    if (busy && !isAutomatic) return;
    if (!sourceDoc || !sourceSlots.length) {
      setStatus("আগে একটি valid Admin DOCX upload করুন।", "error");
      return;
    }

    const manageBusy = !isAutomatic;
    if (manageBusy) setBusy(true);
    resetResults();

    try {
      ensureJSZip();

      const setCount = clamp(positiveInt($("setCount").value, DEFAULT_SET_COUNT), 1, 26);
      const questionsPerSet = positiveInt($("questionsPerSet").value, 1);
      $("setCount").value = String(setCount);
      $("questionsPerSet").value = String(questionsPerSet);

      const blockers = validateGenerationRequest(questionsPerSet);
      if (blockers.length) {
        showValidationIssues(blockers, questionsPerSet);
        setStatus(`Generate বন্ধ করা হয়েছে। ${questionsPerSet}টি প্রশ্নের জন্য Admin file-এ প্রয়োজনীয় Question এবং চারটি Option সম্পূর্ণ করুন।`, "error");
        return;
      }

      validationIssues.hidden = true;
      validationIssues.innerHTML = "";
      if (issueDialog.open) issueDialog.close();

      await loadTemplate();

      const exam = cleanText($("examName").value) || deriveExamName(selectedFile ? selectedFile.name : "") || "Exam";
      const adminFileName = deriveExamName(selectedFile ? selectedFile.name : "") || exam;
      const subject = cleanText($("subjectName").value);
      const shuffleOptions = $("shuffleOptions").value === "yes";
      const warningCount = sourceRecords.filter((record) => record.hasUnsupportedDrawing).length;
      const imageCount = sourceRecords.filter((record) => record.imageRelIds && record.imageRelIds.length).length;

      lastGenerationLabel = sanitizeFilename([exam, subject, `${setCount}-Sets`].filter(Boolean).join("-"));
      setStatus(`${setCount}টি Set তৈরি হচ্ছে · প্রতি Set-এ ${questionsPerSet}টি প্রশ্ন · Option shuffle: ${shuffleOptions ? "Yes" : "No"}...`);

      generatedFiles = [];
      setGrid.innerHTML = "";
      const sets = [];

      for (let index = 0; index < setCount; index += 1) {
        const setLabel = indexToSetLabel(index);
        const selected = secureShuffle(sourceRecords.slice()).slice(0, questionsPerSet);
        const prepared = shuffleOptions ? selected.map(shuffledOptionRecord) : selected;
        const base = makeSetFilename(exam, subject, setLabel).replace(/\.docx$/i, "");
        const bannerExamName = `${adminFileName} (${setLabel}) Set`;
        const outputs = {};
        for (const [kind, suffix] of [["question", "Questions"], ["answer", "Answers"], ["merged", "Combined"]]) {
          outputs[kind] = { name: `${base} - ${suffix}.docx`, blob: await buildSetDocx(setLabel, prepared, kind, bannerExamName) };
          generatedFiles.push(outputs[kind]);
        }
        const set = { setLabel, count: prepared.length, outputs, preview: prepared.map((record) => ({
          question: record.questionText, options: record.optionCells.map(cellText),
          answer: record.answer, explanation: record.explanationText,
          hasMedia: !!(record.hasMath || record.imageRelIds.length)
        })) };
        sets.push(set);
        renderSetCard(set);

        setStatus(`Set ${setLabel} তৈরি হয়েছে (${index + 1}/${setCount})...`);
        await nextFrame();
      }

      resultsNote.textContent = `${sourceRecords.length}টি সম্পূর্ণ valid প্রশ্ন থেকে ${setCount}টি shuffled Set তৈরি হয়েছে · প্রতি Set-এ ${questionsPerSet}টি প্রশ্ন · Option shuffle: ${shuffleOptions ? "Yes" : "No"}।`;
      resultsEl.classList.add("visible");
      downloadAllBtn.disabled = false;
      try {
        await saveHistory({ id: crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`, createdAt: Date.now(),
          exam, subject, sourceName: selectedFile.name, sourceFile: selectedFile, sourceCount: sourceRecords.length,
          questionsPerSet, shuffleOptions, sets });
        await renderHistory();
      } catch (error) {
        console.warn("History save failed", error);
        setStatus("সেট তৈরি হয়েছে, তবে ব্রাউজারে হিস্ট্রি সংরক্ষণ করা যায়নি। জায়গা খালি করে আবার চেষ্টা করুন।", "warning");
        return;
      }

      if (warningCount) {
        setStatus(`✓ সম্পন্ন। Embedded image copy করা হয়েছে। তবে ${warningCount}টি source row-এ non-image drawing/shape/chart আছে; সেগুলো পুরোপুরি copy নাও হতে পারে।`, "warning");
      } else if (imageCount) {
        setStatus(`✓ সম্পন্ন। ${imageCount}টি source row-এর embedded image generated DOCX-এ copy করা হয়েছে। Option shuffle: ${shuffleOptions ? "Yes" : "No"}।`);
      } else {
        setStatus(`✓ সম্পন্ন। ${setCount}টি DOCX Set প্রস্তুত। Option shuffle: ${shuffleOptions ? "Yes" : "No"}।`);
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

  async function buildSetDocx(setLabel, records, kind = "merged", adminFileName = "") {
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

    // The banner text is part of the template image, so burn the Admin DOCX filename
    // plus the current set label (for example, "T-08 Meghna (A) Set") into that image.
    // This keeps the existing template design unchanged while making the Exam Name dynamic.
    await stampAdminFileNameOnBanner(zip, doc, relsDoc, adminFileName);

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
    if (kind === "question") {
      // Keep the section-break paragraph that ends the template's 2-column
      // question section. Removing it makes question-only DOCX files fall
      // back to the final 1-column section settings.
      let questionSectionEnd = questionTable.nextSibling;
      while (
        questionSectionEnd &&
        questionSectionEnd !== body.lastChild &&
        !paragraphHasSectionProperties(questionSectionEnd)
      ) {
        questionSectionEnd = questionSectionEnd.nextSibling;
      }

      let node = questionSectionEnd && questionSectionEnd !== body.lastChild
        ? questionSectionEnd.nextSibling
        : questionTable.nextSibling;
      while (node && node !== body.lastChild) {
        const next = node.nextSibling;
        body.removeChild(node);
        node = next;
      }
    } else if (kind === "answer") {
      let node = body.firstChild;
      while (node && node !== answerTable.previousSibling) {
        const next = node.nextSibling;
        body.removeChild(node);
        node = next;
      }
    }

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

  async function stampAdminFileNameOnBanner(zip, doc, relsDoc, adminFileName) {
    const name = cleanText(adminFileName);
    if (!name) return;

    const relationshipId = firstEmbeddedImageRelationshipId(doc);
    if (!relationshipId) return;

    const relationships = relationshipMapFromDoc(relsDoc);
    const relationship = relationships.get(relationshipId);
    if (!isImageRelationship(relationship) || /^external$/i.test(relationship.targetMode || "")) return;

    const partPath = resolvePackageTarget("word", relationship.target);
    const entry = zip.file(partPath);
    if (!entry) return;

    const contentType = fallbackImageContentType(partPath);
    if (!/^image\/(png|jpeg|jpg)$/i.test(contentType)) return;

    const originalBytes = await entry.async("uint8array");
    const image = await loadBrowserImage(new Blob([originalBytes], { type: contentType }));
    const width = image.naturalWidth || image.width;
    const height = image.naturalHeight || image.height;
    if (!width || !height) return;

    const canvas = document.createElement("canvas");
    canvas.width = width;
    canvas.height = height;
    const context = canvas.getContext("2d");
    if (!context) return;

    context.drawImage(image, 0, 0, width, height);

    // Coordinates are proportional to the supplied banner artwork. The generated Exam Name
    // (Admin filename + current Set) is placed immediately after the printed "Exam Name:"
    // label and auto-shrinks so long names do not collide with the right-side ornament/name plate.
    const x = width * 0.222;
    const baselineY = height * 0.393;
    const maxWidth = width * 0.302;
    let fontSize = Math.max(24, Math.round(height * 0.078));
    const minFontSize = Math.max(18, Math.round(height * 0.043));
    const fontFamily = '"Noto Sans Bengali", "Noto Serif Bengali", Georgia, "Times New Roman", serif';

    context.textBaseline = "alphabetic";
    context.fillStyle = "#3b103f";
    context.font = `700 ${fontSize}px ${fontFamily}`;
    while (fontSize > minFontSize && context.measureText(name).width > maxWidth) {
      fontSize -= 1;
      context.font = `700 ${fontSize}px ${fontFamily}`;
    }

    context.fillText(name, x, baselineY, maxWidth);

    const outputType = /^image\/jpe?g$/i.test(contentType) ? "image/jpeg" : "image/png";
    const outputBlob = await canvasToBlob(canvas, outputType);
    zip.file(partPath, new Uint8Array(await outputBlob.arrayBuffer()));
  }

  function firstEmbeddedImageRelationshipId(doc) {
    const drawings = Array.from(doc.getElementsByTagNameNS(W_NS, "drawing"));
    for (const drawing of drawings) {
      const elements = [drawing, ...Array.from(drawing.getElementsByTagName("*"))];
      for (const element of elements) {
        for (const attr of Array.from(element.attributes || [])) {
          if (attr.namespaceURI === R_NS && attr.localName === "embed" && attr.value) return attr.value;
        }
      }
    }
    return "";
  }

  function loadBrowserImage(blob) {
    return new Promise((resolve, reject) => {
      const url = URL.createObjectURL(blob);
      const image = new Image();
      image.onload = () => {
        URL.revokeObjectURL(url);
        resolve(image);
      };
      image.onerror = () => {
        URL.revokeObjectURL(url);
        reject(new Error("Template banner image load করা যায়নি।"));
      };
      image.src = url;
    });
  }

  function canvasToBlob(canvas, type) {
    return new Promise((resolve, reject) => {
      if (canvas.toBlob) {
        canvas.toBlob((blob) => {
          if (blob) resolve(blob);
          else reject(new Error("Exam Name banner তৈরি করা যায়নি।"));
        }, type, 0.95);
        return;
      }

      try {
        const dataUrl = canvas.toDataURL(type, 0.95);
        const binary = atob(dataUrl.split(",")[1] || "");
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i += 1) bytes[i] = binary.charCodeAt(i);
        resolve(new Blob([bytes], { type }));
      } catch (error) {
        reject(error);
      }
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
    setCellText(cells[1], record.answer);
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

  function renderSetCard(set) {
    const article = document.createElement("article");
    article.className = "set-card";
    article.innerHTML = `
      <div class="set-badge">SET ${escapeHtml(set.setLabel)}</div>
      <h3>সেট ${escapeHtml(set.setLabel)}</h3>
      <p>${set.count}টি প্রশ্ন</p>
      <div class="set-actions"><button type="button" data-action="preview">◉ প্রশ্ন দেখুন</button>
      <button type="button" data-action="question">⬇ প্রশ্নপত্র</button>
      <button type="button" data-action="answer">⬇ উত্তরপত্র</button>
      <button type="button" data-action="merged">⬇ একত্রিত ফাইল</button></div>
    `;
    article.addEventListener("click", (event) => {
      const action = event.target.closest("button")?.dataset.action;
      if (action === "preview") showPreview(set);
      else if (set.outputs[action]) triggerDownload(set.outputs[action].blob, set.outputs[action].name);
    });
    setGrid.appendChild(article);
  }

  function showPreview(set) {
    $("previewTitle").textContent = `সেট ${set.setLabel} · ${set.count}টি প্রশ্ন`;
    previewContent.innerHTML = `<p class="preview-note">এটি লেখার প্রিভিউ। ছবি ও Word equation দেখতে DOCX ডাউনলোড করুন।</p>` +
      set.preview.map((item, index) => `<article class="preview-question"><h3>${index + 1}. ${escapeHtml(item.question || "[ছবি বা সমীকরণসহ প্রশ্ন]")}</h3>
        <ol type="A">${item.options.map((option) => `<li>${escapeHtml(option || "[ছবি বা সমীকরণ]")}</li>`).join("")}</ol>
        <p><strong>উত্তর: ${escapeHtml(item.answer)}</strong></p><p>${escapeHtml(item.explanation || "")}</p></article>`).join("");
    previewDialog.showModal();
  }

  function openHistoryDb() {
    return new Promise((resolve, reject) => {
      if (!window.indexedDB) return reject(new Error("IndexedDB unavailable"));
      const request = indexedDB.open(HISTORY_DB, 1);
      request.onupgradeneeded = () => request.result.createObjectStore("generations", { keyPath: "id" });
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error);
    });
  }

  async function historyOperation(mode, work) {
    const db = await openHistoryDb();
    try {
      return await new Promise((resolve, reject) => {
        const tx = db.transaction("generations", mode);
        const request = work(tx.objectStore("generations"));
        request.onsuccess = () => resolve(request.result);
        request.onerror = () => reject(request.error);
        tx.onabort = () => reject(tx.error || new Error("History transaction failed"));
      });
    } finally { db.close(); }
  }

  function saveHistory(entry) { return historyOperation("readwrite", (store) => store.put(entry)); }
  function listHistory() { return historyOperation("readonly", (store) => store.getAll()); }
  function removeHistory(id) { return historyOperation("readwrite", (store) => store.delete(id)); }

  async function renderHistory() {
    historyList.textContent = "হিস্ট্রি লোড হচ্ছে...";
    try {
      const entries = (await listHistory()).sort((a, b) => b.createdAt - a.createdAt);
      if (!entries.length) { historyList.textContent = "এখনও কোনো প্রশ্নপত্র তৈরি হয়নি।"; return; }
      historyList.innerHTML = "";
      for (const entry of entries) {
        const item = document.createElement("details");
        item.className = "history-item";
        item.innerHTML = `<summary><span><strong>${escapeHtml(entry.sourceName)}</strong><small>${escapeHtml(entry.exam)}${entry.subject ? ` · ${escapeHtml(entry.subject)}` : ""} · ${entry.sets.length} সেট · প্রতি সেটে ${entry.questionsPerSet} প্রশ্ন · Option shuffle: ${entry.shuffleOptions ? "Yes" : "No"}<br>${new Date(entry.createdAt).toLocaleString("bn-BD")}</small></span><span class="history-open">দেখুন ⌄</span></summary><div class="history-body"><div class="history-tools"><button type="button" data-action="source">⬇ Admin মূল ফাইল</button><button type="button" data-action="zip">⬇ সব আউটপুট ZIP</button><button type="button" data-action="recreate">↻ আবার তৈরি</button><button type="button" data-action="edit">✎ সেটিং বদলান</button><button type="button" data-action="delete" class="danger">মুছুন</button></div><div class="history-sets"></div></div>`;
        for (const set of entry.sets) {
          const card = document.createElement("div");
          card.className = "history-set";
          card.innerHTML = `<strong>সেট ${escapeHtml(set.setLabel)}</strong><div class="history-tools"><button type="button" data-action="preview">◉ দেখুন</button><button type="button" data-action="question">⬇ প্রশ্ন</button><button type="button" data-action="answer">⬇ উত্তর</button><button type="button" data-action="merged">⬇ একত্রিত</button></div>`;
          card.addEventListener("click", (event) => {
            const action = event.target.closest("button")?.dataset.action;
            if (action === "preview") showPreview(set);
            else if (set.outputs[action]) triggerDownload(set.outputs[action].blob, set.outputs[action].name);
          });
          item.querySelector(".history-sets").appendChild(card);
        }
        item.querySelector(".history-tools").addEventListener("click", async (event) => {
          const action = event.target.closest("button")?.dataset.action;
          if (action === "source") triggerDownload(entry.sourceFile, entry.sourceName);
          if ((action === "recreate" || action === "edit") && !busy) {
            await selectFile(entry.sourceFile, true, entry, action === "recreate");
          }
          if (action === "zip") {
            const zip = new window.JSZip();
            zip.file(entry.sourceName, entry.sourceFile);
            entry.sets.forEach((set) => Object.values(set.outputs).forEach((output) => zip.file(output.name, output.blob)));
            triggerDownload(await zip.generateAsync({ type: "blob", compression: "DEFLATE" }), `${sanitizeFilename(entry.exam)}-${entry.id}.zip`);
          }
          if (action === "delete" && confirm("এই হিস্ট্রি এবং সংরক্ষিত ফাইলগুলি মুছে ফেলবেন?")) {
            await removeHistory(entry.id);
            await renderHistory();
          }
        });
        historyList.appendChild(item);
      }
    } catch (error) {
      console.warn("History unavailable", error);
      historyList.textContent = "এই ব্রাউজারে হিস্ট্রি খোলা যাচ্ছে না। ব্রাউজারের storage অনুমতি পরীক্ষা করুন।";
    }
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
    $("shuffleOptions").value = "no";
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
    generateBtn.disabled = value || !sourceSlots.length;
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

  function paragraphHasSectionProperties(node) {
    if (!node || node.nodeType !== 1 || node.namespaceURI !== W_NS || node.localName !== "p") return false;
    const pPr = directChildren(node, "pPr")[0];
    return Boolean(pPr && directChildren(pPr, "sectPr").length);
  }

  function cellText(cell) {
    return cleanText(cell.textContent || "");
  }

  function normalizeAnswer(value) {
    const text = cleanText(value).toUpperCase();
    return /^[ABCD]$/.test(text) ? text : "";
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
