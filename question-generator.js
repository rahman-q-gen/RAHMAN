(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);
  const fileInput = $("docxFile");
  const dropZone = $("dropZone");
  const generateBtn = $("generateBtn");
  const resetBtn = $("resetBtn");
  const statusEl = $("status");
  const selectedFileEl = $("selectedFile");
  const resultsEl = $("results");
  const resultsNote = $("resultsNote");
  const setGrid = $("setGrid");
  const downloadAllBtn = $("downloadAllBtn");
  const debugEl = $("debug");

  let selectedFile = null;
  let generatedFiles = [];
  let lastGenerationLabel = "Question-Generator";

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

  generateBtn.addEventListener("click", generate);
  resetBtn.addEventListener("click", resetAll);
  downloadAllBtn.addEventListener("click", downloadAllAsZip);

  function selectFile(file, fromDrop = false) {
    if (!file) {
      selectedFile = null;
      updateFileState();
      return;
    }

    if (!/\.docx$/i.test(file.name)) {
      selectedFile = null;
      fileInput.value = "";
      updateFileState();
      setStatus("শুধু .docx ফাইল ব্যবহার করুন।", "error");
      return;
    }

    selectedFile = file;
    setStatus("");

    if (fromDrop) {
      try {
        const transfer = new DataTransfer();
        transfer.items.add(file);
        fileInput.files = transfer.files;
      } catch (_) {
        // Some browsers do not allow assigning FileList. The selectedFile variable is enough.
      }
    }

    updateFileState();
  }

  function updateFileState() {
    if (selectedFile) {
      selectedFileEl.textContent = `✓ ${selectedFile.name} (${formatBytes(selectedFile.size)})`;
      selectedFileEl.classList.add("ready");
      generateBtn.disabled = false;
    } else {
      selectedFileEl.textContent = "কোনো ফাইল নির্বাচন করা হয়নি";
      selectedFileEl.classList.remove("ready");
      generateBtn.disabled = true;
    }
  }

  function setStatus(message, type = "") {
    statusEl.textContent = message;
    statusEl.className = `status${type ? ` ${type}` : ""}`;
  }

  function librariesReady() {
    const missing = [];
    if (!window.mammoth) missing.push("Mammoth.js");
    if (!window.docx) missing.push("docx.js");
    if (!window.JSZip) missing.push("JSZip");
    if (missing.length) {
      throw new Error(`${missing.join(", ")} load হয়নি। Internet connection পরীক্ষা করে page reload করুন।`);
    }
  }

  async function generate() {
    if (!selectedFile) return;

    generateBtn.disabled = true;
    downloadAllBtn.disabled = true;
    resultsEl.classList.remove("visible");
    setGrid.innerHTML = "";
    resultsNote.textContent = "";
    generatedFiles = [];
    debugEl.style.display = "none";
    debugEl.textContent = "";

    try {
      librariesReady();
      setStatus("DOCX পড়া হচ্ছে...");

      const arrayBuffer = await selectedFile.arrayBuffer();
      const result = await window.mammoth.extractRawText({ arrayBuffer });
      const raw = normalizeText(result.value || "");

      if (!raw.trim()) {
        throw new Error("DOCX থেকে কোনো text পাওয়া যায়নি।");
      }

      const parsed = parseQuestions(raw);
      if (!parsed.length) {
        showDebug(raw);
        throw new Error("A-D / ক-ঘ format-এর সম্পূর্ণ MCQ শনাক্ত করা যায়নি। নিচে extracted text দেখানো হয়েছে।");
      }

      const requestedCount = positiveInt($("questionCount").value, parsed.length);
      const wanted = Math.min(requestedCount, parsed.length);
      const selectedQuestions = parsed.slice(0, wanted);
      const perSetRequested = positiveInt($("questionsPerSet").value, selectedQuestions.length);
      const perSet = Math.min(perSetRequested, selectedQuestions.length);
      const questions = selectedQuestions.slice(0, perSet);

      if (!questions.length) {
        throw new Error("প্রতি সেটে অন্তত ১টি প্রশ্ন দিতে হবে।");
      }

      const exam = cleanText($("examName").value) || "T-01";
      const subject = cleanText($("subjectName").value);
      const sets = buildSets(questions);
      const answerCount = questions.filter((q) => q.answer).length;

      lastGenerationLabel = [exam, subject, "Sets-A-D"].filter(Boolean).join("-");
      setStatus(`${parsed.length}টি valid MCQ পাওয়া গেছে। Set A-D তৈরি হচ্ছে...`);

      for (let index = 0; index < sets.length; index += 1) {
        const setLetter = String.fromCharCode(65 + index);
        const setQuestions = sets[index];
        const baseTitle = subject ? `${exam} (${subject})` : exam;

        const questionBlob = await makeDocx(`${baseTitle} — Set ${setLetter} — Question`, setQuestions, "question", setLetter);
        const answerBlob = await makeDocx(`${baseTitle} — Set ${setLetter} — Answer`, setQuestions, "answer", setLetter);
        const combinedBlob = await makeDocx(`${baseTitle} — Set ${setLetter} — Combined`, setQuestions, "combined", setLetter);

        const files = [
          { blob: questionBlob, name: makeName(exam, subject, setLetter, "Question"), kind: "question" },
          { blob: answerBlob, name: makeName(exam, subject, setLetter, "Answer"), kind: "answer" },
          { blob: combinedBlob, name: makeName(exam, subject, setLetter, "Combined"), kind: "combined" }
        ];

        generatedFiles.push(...files);
        renderSet(setLetter, setQuestions, files);
      }

      resultsNote.textContent = `${questions.length}টি প্রশ্ন × 4 সেট = ${generatedFiles.length}টি DOCX ফাইল। উৎসে ${answerCount}/${questions.length}টি প্রশ্নের answer marker পাওয়া গেছে।`;
      resultsEl.classList.add("visible");
      downloadAllBtn.disabled = false;

      if (answerCount < questions.length) {
        setStatus(`✓ ফাইল তৈরি হয়েছে। ${questions.length - answerCount}টি প্রশ্নে উৎস DOCX-এ answer marker পাওয়া যায়নি; answer file-এ তা উল্লেখ করা হয়েছে।`, "warning");
      } else {
        setStatus(`✓ সম্পন্ন। 4টি সেট এবং ${generatedFiles.length}টি DOCX ফাইল তৈরি হয়েছে।`);
      }

      resultsEl.scrollIntoView({ behavior: "smooth", block: "start" });
    } catch (error) {
      console.error(error);
      setStatus(error && error.message ? error.message : "ফাইল তৈরির সময় সমস্যা হয়েছে।", "error");
    } finally {
      generateBtn.disabled = !selectedFile;
    }
  }

  function parseQuestions(text) {
    const expanded = expandInlineMarkers(text);
    const lines = expanded.split(/\n+/).map((line) => line.trim()).filter(Boolean);
    const questions = [];
    let current = null;
    let lastOptionIndex = null;

    const questionRe = /^(?:(?:প্রশ্ন|question|q)\s*)?([0-9০-৯]{1,4})\s*[.\)\-:।]\s*(.+)$/i;
    const optionRe = /^\(?\s*([A-Da-dকখগঘ])\s*\)?\s*[.\)\-:।]\s*(.+)$/;
    const answerRe = /^(?:সঠিক\s*উত্তর|উত্তর|correct\s*answer|answer|ans)\s*[.:\-–—]?\s*\(?\s*([A-Da-dকখগঘ])\s*\)?(?:\s.*)?$/i;

    for (const line of lines) {
      const questionMatch = line.match(questionRe);
      if (questionMatch) {
        if (current) finishQuestion(current, questions);
        current = {
          number: parseQuestionNumber(questionMatch[1]),
          question: cleanText(questionMatch[2]),
          options: ["", "", "", ""],
          answer: null
        };
        lastOptionIndex = null;
        continue;
      }

      if (!current) continue;

      const answerMatch = line.match(answerRe);
      if (answerMatch) {
        current.answer = convertAnswer(answerMatch[1]);
        continue;
      }

      const optionMatch = line.match(optionRe);
      if (optionMatch) {
        const optionIndex = optionLabelToIndex(optionMatch[1]);
        if (optionIndex !== null) {
          current.options[optionIndex] = cleanText(optionMatch[2]);
          lastOptionIndex = optionIndex;
          continue;
        }
      }

      if (lastOptionIndex === null) {
        current.question = cleanText(`${current.question} ${line}`);
      } else {
        current.options[lastOptionIndex] = cleanText(`${current.options[lastOptionIndex]} ${line}`);
      }
    }

    if (current) finishQuestion(current, questions);
    return questions;
  }

  function expandInlineMarkers(text) {
    let value = text;
    value = value.replace(/([^\n])\s+(?=\(?\s*[A-Da-dকখগঘ]\s*\)?\s*[.\)\-:।]\s+)/g, "$1\n");
    value = value.replace(/([^\n])\s+(?=(?:সঠিক\s*উত্তর|উত্তর|correct\s*answer|answer|ans)\s*[.:\-–—])/gi, "$1\n");
    return value;
  }

  function finishQuestion(question, target) {
    question.question = cleanText(question.question);
    question.options = question.options.map(cleanText);
    const completeOptions = question.options.filter(Boolean).length;
    if (question.question && completeOptions === 4) {
      target.push(question);
    }
  }

  function parseQuestionNumber(value) {
    const bengaliToEnglish = {
      "০": "0", "১": "1", "২": "2", "৩": "3", "৪": "4",
      "৫": "5", "৬": "6", "৭": "7", "৮": "8", "৯": "9"
    };
    const normalized = String(value).replace(/[০-৯]/g, (digit) => bengaliToEnglish[digit]);
    return Number(normalized) || 0;
  }

  function optionLabelToIndex(label) {
    const map = { A: 0, B: 1, C: 2, D: 3, ক: 0, খ: 1, গ: 2, ঘ: 3 };
    const normalized = /^[a-d]$/i.test(label) ? label.toUpperCase() : label;
    return Object.prototype.hasOwnProperty.call(map, normalized) ? map[normalized] : null;
  }

  function convertAnswer(answer) {
    const index = optionLabelToIndex(answer);
    return index === null ? null : ["A", "B", "C", "D"][index];
  }

  function buildSets(questions) {
    const sets = [];
    for (let setIndex = 0; setIndex < 4; setIndex += 1) {
      sets.push(questions.map((question, questionIndex) => {
        const shift = (setIndex + questionIndex) % 4;
        return {
          number: questionIndex + 1,
          question: question.question,
          options: rotate(question.options, shift),
          answer: question.answer ? rotateAnswer(question.answer, shift) : null
        };
      }));
    }
    return sets;
  }

  function rotate(items, count) {
    const copy = items.slice();
    for (let i = 0; i < count; i += 1) copy.push(copy.shift());
    return copy;
  }

  function rotateAnswer(answer, shift) {
    const sourceIndex = { A: 0, B: 1, C: 2, D: 3 }[answer];
    if (sourceIndex === undefined) return null;
    return ["A", "B", "C", "D"][(sourceIndex - shift + 4) % 4];
  }

  async function makeDocx(title, questions, mode, setLetter) {
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType } = window.docx;
    const children = [
      new Paragraph({
        text: title,
        heading: HeadingLevel.HEADING_1,
        alignment: AlignmentType.CENTER
      }),
      new Paragraph({
        text: `MD. ABDUR RAHMAN — Question Generator — Set ${setLetter}`,
        alignment: AlignmentType.CENTER
      }),
      new Paragraph({ text: "" })
    ];

    questions.forEach((question, index) => {
      if (mode === "answer") {
        const answerText = question.answer
          ? `${question.answer} — ${getAnswerOptionText(question)}`
          : "উৎস ফাইলে উত্তর নির্ধারিত নেই";
        children.push(new Paragraph({
          children: [
            new TextRun({ text: `${index + 1}. `, bold: true }),
            new TextRun({ text: answerText, bold: Boolean(question.answer) })
          ]
        }));
        return;
      }

      children.push(new Paragraph({
        children: [new TextRun({ text: `${index + 1}. ${question.question}`, bold: true })]
      }));

      question.options.forEach((option, optionIndex) => {
        children.push(new Paragraph({
          text: `${["A", "B", "C", "D"][optionIndex]}. ${option}`,
          indent: { left: 720 }
        }));
      });

      if (mode === "combined") {
        const answerText = question.answer
          ? `${question.answer} — ${getAnswerOptionText(question)}`
          : "উৎস ফাইলে উত্তর নির্ধারিত নেই";
        children.push(new Paragraph({
          children: [new TextRun({ text: `উত্তর: ${answerText}`, bold: true })]
        }));
      }

      children.push(new Paragraph({ text: "" }));
    });

    const document = new Document({
      sections: [{ properties: {}, children }]
    });

    return Packer.toBlob(document);
  }

  function getAnswerOptionText(question) {
    const answerIndex = { A: 0, B: 1, C: 2, D: 3 }[question.answer];
    return answerIndex === undefined ? "" : question.options[answerIndex];
  }

  function renderSet(setLetter, questions, files) {
    const card = document.createElement("article");
    card.className = "set-card";
    card.innerHTML = `
      <div class="set-card-head">
        <h3>Set ${setLetter}</h3>
        <div class="meta">${questions.length}টি প্রশ্ন<br>3টি DOCX</div>
      </div>
      <div class="download-links">
        <a class="q-link" href="#">প্রশ্নপত্র ডাউনলোড <span>⬇</span></a>
        <a class="a-link" href="#">উত্তরপত্র ডাউনলোড <span>⬇</span></a>
        <a class="m-link" href="#">একীভূত ফাইল ডাউনলোড <span>⬇</span></a>
      </div>`;

    card.querySelectorAll("a").forEach((link, index) => {
      link.addEventListener("click", (event) => {
        event.preventDefault();
        triggerDownload(files[index].blob, files[index].name);
      });
    });

    setGrid.appendChild(card);
  }

  async function downloadAllAsZip() {
    if (!generatedFiles.length) return;

    try {
      if (!window.JSZip) throw new Error("ZIP library load হয়নি। Page reload করুন।");
      downloadAllBtn.disabled = true;
      setStatus("ZIP তৈরি হচ্ছে...");

      const zip = new window.JSZip();
      generatedFiles.forEach((file) => zip.file(file.name, file.blob));
      const blob = await zip.generateAsync({ type: "blob" });
      triggerDownload(blob, `${sanitizeFilename(lastGenerationLabel)}.zip`);
      setStatus("✓ সব generated file একটি ZIP-এ ডাউনলোড হয়েছে।");
    } catch (error) {
      console.error(error);
      setStatus(error && error.message ? error.message : "ZIP তৈরির সময় সমস্যা হয়েছে।", "error");
    } finally {
      downloadAllBtn.disabled = generatedFiles.length === 0;
    }
  }

  function makeName(exam, subject, setLetter, type) {
    const subjectPart = subject ? ` (${subject})` : "";
    return sanitizeFilename(`${exam}${subjectPart} Set-${setLetter} ${type}.docx`);
  }

  function sanitizeFilename(name) {
    return String(name).replace(/[\\/:*?"<>|]/g, "-").replace(/\s+/g, " ").trim();
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

  function normalizeText(text) {
    return String(text)
      .replace(/\u00a0/g, " ")
      .replace(/\r\n?/g, "\n")
      .split("\n")
      .map((line) => line.replace(/[ \t]+/g, " ").trim())
      .filter(Boolean)
      .join("\n");
  }

  function cleanText(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
  }

  function positiveInt(value, fallback) {
    const number = Number.parseInt(value, 10);
    return Number.isFinite(number) && number > 0 ? number : fallback;
  }

  function formatBytes(bytes) {
    if (!bytes) return "0 B";
    const units = ["B", "KB", "MB", "GB"];
    const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
    const value = bytes / (1024 ** index);
    return `${value.toFixed(index ? 1 : 0)} ${units[index]}`;
  }

  function showDebug(raw) {
    debugEl.textContent = raw.slice(0, 8000);
    debugEl.style.display = "block";
  }

  function resetAll() {
    selectedFile = null;
    fileInput.value = "";
    $("examName").value = "T-01";
    $("subjectName").value = "Meghna";
    $("questionCount").value = "20";
    $("questionsPerSet").value = "20";
    generatedFiles = [];
    lastGenerationLabel = "Question-Generator";
    setGrid.innerHTML = "";
    resultsNote.textContent = "";
    resultsEl.classList.remove("visible");
    downloadAllBtn.disabled = true;
    debugEl.style.display = "none";
    debugEl.textContent = "";
    updateFileState();
    setStatus("");
  }

  updateFileState();
})();
