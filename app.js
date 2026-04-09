"use strict";

const DEFAULT_KG_RULES = {
  levelTokens: { "базовое высшее": "БВО", "бакалавриат": "БАК", "магистратура": "МАГ", "специалитет": "СПЕЦ" },
  formTokens: { "очная": "О", "очно-заочная": "ОЗ", "заочная": "З" },
  financeTokens: { "бюджет": "Б", "договор": "Д" },
  quotaTokens: { target: "Ц", special: "К", separate: "ОК", foreign: "ИН", contract: "Д" },
  yesValues: ["да", "yes", "true", "1"],
  noValues: ["нет", "no", "false", "0"],
  separateQuotaRule: {
    specialFeatureColumn: "Особенности приема",
    specialFeatureValue: "Отдельная квота",
    specialRightColumn: "Особое право",
    specialRightExpected: "Да"
  }
};

const DEFAULT_BENEFITS_RULES = {
  replacements: [
    {
      from: "Дети граждан, проходящих (проходивших) военную службу, в т.ч. по мобилизации, заключивших контракт, при условии их участия в СВО, в составе ВС, Нар.милиции ЛНР, ДНР с 11 мая 2014",
      to: "Дети граждан, проходящих (проходивших) военную службу, в т.ч. по мобилизации, заключивших контракт, при условии их участия в СВО, в составе ВС, Нар.ми"
    },
    {
      from: "Принадлежность к детям граждан, лиц, военнослужащих погибших/получивших увечье, принимавших уч-е в боевых действиях в других госудр., в составе ВС, Нар.милиции ЛНР, ДНР с 11 мая 2014",
      to: "Принадлежность к детям граждан, лиц, военнослужащих погибших/получивших увечье, принимавших уч-е в боевых действиях в других госудр., в составе ВС, На"
    },
    {
      from: "Принадлежность к детям граждан, лиц, военнослужащих погибших/получивших увечье в ходе СВО либо удостоены звания Героя Российской Федерации или награждены тремя орденами Мужества",
      to: "Принадлежность к детям граждан, лиц, военнослужащих погибших/получивших увечье в ходе СВО либо удостоены звания Героя Российской Федерации или награжд"
    }
  ]
};

const DEFAULT_MATCH_RULES = {
  keyFieldPn: "Конкурсная группа",
  keyFieldVi: "Конкурсная группа",
  fieldPairs: [
    { pn: "Институт", vi: "Институт" },
    { pn: "Код", vi: "Код" },
    { pn: "Направление", vi: "Направление" },
    { pn: "Профиль", vi: "Профиль" },
    { pn: "Тип стандарта", vi: "Тип стандарта" },
    { pn: "Форма обучения", vi: "Форма обучения" },
    { pn: "Основание поступление", vi: "Основание поступление" },
    { pn: "Квалификация", vi: "Квалификация" }
  ]
};

const DEFAULT_VI_RULES = {
  examFormTypes: { "егэ": "ege", "экзамен": "exam", "ид": "id" },
  maxScores: { ege: 100, exam: 100, id: 10, idTarget: 5 },
  targetColumn: "Целевая детализированная квота",
  testFormColumn: "Форма испытания",
  testTypeColumn: "Тип вступительного испытания",
  specialMarkColumn: "Особая отметка",
  subjectColumn: "Предмет",
  replaceSubjectColumn: "Заменяемый предмет",
  minScoreColumn: "Минимальный балл",
  maxScoreColumn: "Максимальный балл",
  creativeTypeValue: "Дополнительное испытание творческой и (или) профессиональной направленности"
};

/** Защита от зависаний на «случайных» мегабайтных ячейках Excel (нормализация и .includes по ним линейно дороги). */
const MAX_NORMALIZE_CHARS = 65536;

/** Реже, чем каждые 200 строк — меньше накладных расходов await на больших файлах. */
const CHECK_LOOP_YIELD_EVERY = 800;

const inputs = {
  pnTemplate: document.getElementById("pnTemplate"),
  viTemplate: document.getElementById("viTemplate"),
  pnFile: document.getElementById("pnFile"),
  viFile: document.getElementById("viFile"),
  minBallFile: document.getElementById("minBallFile"),
  spoMinBallFile: document.getElementById("spoMinBallFile"),
  rulesFile: document.getElementById("rulesFile")
};

const runCheckButton = document.getElementById("runCheck");
const criteriaResultsNode = document.getElementById("criteriaResults");
const globalMessageNode = document.getElementById("globalMessage");
const progressSectionNode = document.getElementById("progressSection");
const progressLabelNode = document.getElementById("progressLabel");
const progressPercentNode = document.getElementById("progressPercent");
const progressFillNode = document.getElementById("progressFill");
const progressTrackNode = progressSectionNode.querySelector(".progress-track");
const benchmarkSectionNode = document.getElementById("benchmarkSection");
const benchmarkListNode = document.getElementById("benchmarkList");
const tabUploadButton = document.getElementById("tabUpload");
const tabCriteriaButton = document.getElementById("tabCriteria");
const tabFixedButton = document.getElementById("tabFixed");
const uploadTabPanel = document.getElementById("uploadTabPanel");
const criteriaTabPanel = document.getElementById("criteriaTabPanel");
const fixedTabPanel = document.getElementById("fixedTabPanel");
const fixedTablesMount = document.getElementById("fixedTablesMount");

const checkedIssues = new Set();

Object.values(inputs).forEach((input) => input.addEventListener("change", updateRunButtonState));
runCheckButton.addEventListener("click", onRunCheck);
tabUploadButton.addEventListener("click", () => setActiveTab("upload"));
tabCriteriaButton.addEventListener("click", () => setActiveTab("criteria"));
if (tabFixedButton) tabFixedButton.addEventListener("click", () => setActiveTab("fixed"));
updateRunButtonState();
setActiveTab("upload");

function updateRunButtonState() {
  const requiredInputs = [inputs.pnTemplate, inputs.viTemplate, inputs.pnFile, inputs.viFile, inputs.minBallFile, inputs.spoMinBallFile];
  const allChosen = requiredInputs.every((input) => input.files && input.files[0]);
  runCheckButton.disabled = !allChosen;
  if (!allChosen) {
    criteriaResultsNode.classList.add("hidden");
    criteriaResultsNode.innerHTML = "";
    clearFixedPreview();
  }
}

async function onRunCheck() {
  clearGlobalMessage();
  criteriaResultsNode.classList.add("hidden");
  criteriaResultsNode.innerHTML = "";
  clearFixedPreview();
  checkedIssues.clear();
  clearBenchmark();
  setUiBusy(true);
  setProgress(0, "Подготовка проверки...");
  const perf = createPerfTracker();

  try {
    perf.start("rules", "Чтение rules.json");
    const combinedRules = await readCombinedRules(inputs.rulesFile.files[0]);
    const kgRules = validateKgRules(combinedRules.kgRules);
    const benefitsRules = validateBenefitsRules(combinedRules.benefitsRules);
    const matchRules = validateMatchRules(combinedRules.matchRules);
    const viRules = validateViRules(combinedRules.viRules);
    perf.end("rules");

    setProgress(8, "Чтение шаблона ПН...");
    perf.start("read_pn_template", "Чтение шаблона ПН");
    const pnTemplateHeaders = await readHeadersFromFile(inputs.pnTemplate.files[0], "Шаблон ПН");
    perf.end("read_pn_template");
    setProgress(16, "Чтение шаблона ВИ...");
    perf.start("read_vi_template", "Чтение шаблона ВИ");
    const viTemplateHeaders = await readHeadersFromFile(inputs.viTemplate.files[0], "Шаблон ВИ");
    perf.end("read_vi_template");
    setProgress(30, "Чтение файла ПН...");
    perf.start("read_pn", "Чтение файла ПН");
    const pnWorkbookData = await readWorkbookDataFromFile(inputs.pnFile.files[0], "Файл ПН");
    perf.end("read_pn");
    setProgress(42, "Чтение файла ВИ...");
    perf.start("read_vi", "Чтение файла ВИ");
    const viWorkbookData = await readWorkbookDataFromFile(inputs.viFile.files[0], "Файл ВИ");
    perf.end("read_vi");
    setProgress(52, "Чтение файла min_ball.txt...");
    perf.start("read_min_ball", "Чтение min_ball.txt");
    const minBallData = await readMinBallFromFile(inputs.minBallFile.files[0]);
    perf.end("read_min_ball");
    setProgress(56, "Чтение файла spo_min_ball.txt...");
    perf.start("read_spo_min_ball", "Чтение spo_min_ball.txt");
    const spoMinBallData = await readSpoMinBallFromFile(inputs.spoMinBallFile.files[0]);
    perf.end("read_spo_min_ball");

    setProgress(62, "Проверка соответствия столбцов...");
    perf.start("headers", "Сверка заголовков");
    const pnReport = compareHeaders(pnTemplateHeaders.headers, pnWorkbookData.headers);
    const viReport = compareHeaders(viTemplateHeaders.headers, viWorkbookData.headers);
    const headerIssues = buildHeaderCriterionIssues(pnReport, viReport);
    perf.end("headers");

    const progress = createProgressReporter();
    progress.setPhase(68, 78, "Проверка критериев ПН...");
    perf.start("pn_criteria", "Проверка критериев ПН");
    const pnCriteria = await runPnCriteriaChecks(
      pnWorkbookData,
      kgRules,
      benefitsRules,
      progress.step,
      pnTemplateHeaders.headers
    );
    perf.end("pn_criteria");
    progress.setPhase(78, 82, "Проверка пробелов ПН...");
    perf.start("pn_spaces", "Проверка пробелов ПН");
    const pnWhitespaceCriterion = await checkWhitespaceIssues(pnWorkbookData, "ПН", progress.step);
    perf.end("pn_spaces");
    progress.setPhase(82, 85, "Проверка квалификации ПН...");
    perf.start("pn_qualification", "Проверка квалификации ПН");
    const pnQualificationCriterion = await checkQualificationByLevel(
      pnWorkbookData,
      "ПН",
      progress.step,
      pnTemplateHeaders.headers
    );
    perf.end("pn_qualification");

    progress.setPhase(85, 92, "Проверка критериев ВИ...");
    perf.start("vi_criteria", "Проверка критериев ВИ");
    const viCriteria = await checkViCriteria(
      viWorkbookData,
      minBallData,
      spoMinBallData,
      benefitsRules,
      kgRules,
      viRules,
      progress.step,
      viTemplateHeaders.headers
    );
    perf.end("vi_criteria");
    progress.setPhase(92, 95, "Проверка пробелов ВИ...");
    perf.start("vi_spaces", "Проверка пробелов ВИ");
    const viWhitespaceCriterion = await checkWhitespaceIssues(viWorkbookData, "ВИ", progress.step);
    perf.end("vi_spaces");
    progress.setPhase(95, 97, "Проверка квалификации ВИ...");
    perf.start("vi_qualification", "Проверка квалификации ВИ");
    const viQualificationCriterion = await checkQualificationByLevel(
      viWorkbookData,
      "ВИ",
      progress.step,
      viTemplateHeaders.headers
    );
    perf.end("vi_qualification");

    progress.setPhase(97, 99, "Сверка ПН и ВИ...");
    perf.start("pn_vi_match", "Сверка ПН и ВИ");
    const pnViMatchCriterion = await checkPnViMatch(pnWorkbookData, viWorkbookData, matchRules, progress.step);
    perf.end("pn_vi_match");

    const criteria = [
      {
        sectionId: "pn-file",
        sectionTitle: "Файл План набора",
        blocks: [
          { id: "columns-pn", title: "Соответствие столбцам", issues: headerIssues.pnIssues },
          { id: "kg-name-rules", title: "Корректность названий КГ", issues: pnCriteria.kgNameIssues },
          { id: "code-in-kg", title: "Код направления в названии КГ", issues: pnCriteria.codeIssues },
          { id: "level-in-kg", title: "Уровень образования в названии КГ", issues: pnCriteria.levelIssues },
          { id: "foreign-rules", title: "Правила иностранных КГ", issues: pnCriteria.foreignIssues },
          { id: "rf-only", title: "Проверка столбца Только для граждан РФ", issues: pnCriteria.rfIssues },
          { id: "benefits", title: "Проверка названий льгот", issues: pnCriteria.benefitIssues },
          { id: "target-detailed", title: "Целевая детализированная квота", issues: pnCriteria.targetDetailedIssues },
          { id: "pn-whitespace", title: "Пробелы в значениях ячеек", issues: ensureIssues(pnWhitespaceCriterion) },
          { id: "pn-qualification", title: "Квалификация по уровню образования", issues: ensureIssues(pnQualificationCriterion) },
          { id: "pn-vi-cross-check", title: "Сверка ПН и ВИ", issues: ensureIssues(pnViMatchCriterion) }
        ]
      },
      {
        sectionId: "vi-file",
        sectionTitle: "Файл Вступительные испытания",
        blocks: [
          { id: "columns-vi", title: "Соответствие столбцам", issues: headerIssues.viIssues },
          { id: "max-ball-vi", title: "Максимальные баллы ВИ", issues: viCriteria.maxIssues },
          { id: "min-ball-vi", title: "Минимальные баллы ВИ", issues: viCriteria.minIssues },
          { id: "rule-21-vi", title: "Правило 21 балла", issues: viCriteria.rule21Issues },
          { id: "rule-41-vi", title: "Правило 41 балла (СПО)", issues: viCriteria.rule41Issues },
          { id: "spo-column-vi", title: "Предметы СПО только в «Заменяемый предмет»", issues: viCriteria.spoColumnIssues },
          { id: "special-mark-vi", title: "Формулировки особой отметки", issues: viCriteria.specialMarkIssues },
          { id: "vi-whitespace", title: "Пробелы в значениях ячеек", issues: ensureIssues(viWhitespaceCriterion) },
          { id: "vi-qualification", title: "Квалификация по уровню образования", issues: ensureIssues(viQualificationCriterion) }
        ]
      }
    ];

    setProgress(97, "Формирование отчета по критериям...");
    perf.start("render", "Рендер отчета");
    renderCriteriaReports(criteria);
    setProgress(97, "Подготовка данных исправлений...");
    const pnCorrected = buildCorrectedPnData(pnWorkbookData, kgRules, benefitsRules, pnTemplateHeaders.headers);
    const viCorrected = buildCorrectedViData(
      viWorkbookData,
      minBallData,
      spoMinBallData,
      benefitsRules,
      kgRules,
      viRules,
      viTemplateHeaders.headers
    );
    renderFixedPreview(pnWorkbookData, pnCorrected, viWorkbookData, viCorrected);
    perf.end("render");
    perf.render();
    if (fixedTabPanel && tabFixedButton) setActiveTab("fixed");
    else setActiveTab("criteria");

    const warnings = [
      ...pnTemplateHeaders.warnings,
      ...viTemplateHeaders.warnings,
      ...pnWorkbookData.warnings,
      ...viWorkbookData.warnings,
      ...minBallData.warnings,
      ...spoMinBallData.warnings
    ];
    setProgress(100, "Проверка завершена");
    if (warnings.length > 0) {
      setGlobalMessage(`Обнаружены предупреждения: ${warnings.join(" | ")}`, "warn");
    } else {
      setGlobalMessage("Проверка завершена.", "ok");
    }
  } catch (error) {
    clearFixedPreview();
    setProgress(100, "Проверка завершена с ошибкой");
    setGlobalMessage(error.message || "Не удалось обработать файлы.", "error");
  } finally {
    setUiBusy(false);
    updateRunButtonState();
  }
}

function createPerfTracker() {
  const marks = new Map();
  const results = [];
  return {
    start(id, label) {
      marks.set(id, { label, start: performance.now() });
    },
    end(id) {
      const m = marks.get(id);
      if (!m) return;
      results.push({ label: m.label, ms: performance.now() - m.start });
      marks.delete(id);
    },
    render() {
      benchmarkListNode.innerHTML = "";
      const total = results.reduce((sum, r) => sum + r.ms, 0);
      for (const r of results) {
        const li = document.createElement("li");
        li.textContent = `${r.label}: ${r.ms.toFixed(0)} ms`;
        benchmarkListNode.appendChild(li);
      }
      const totalLi = document.createElement("li");
      totalLi.textContent = `Общее время: ${total.toFixed(0)} ms`;
      benchmarkListNode.appendChild(totalLi);
      benchmarkSectionNode.classList.remove("hidden");
    }
  };
}

function clearBenchmark() {
  benchmarkListNode.innerHTML = "";
  benchmarkSectionNode.classList.add("hidden");
}

function validateKgRules(rules) {
  if (!rules || !rules.levelTokens || !rules.formTokens || !rules.financeTokens || !rules.quotaTokens) {
    throw new Error("kg_rules.json: некорректная структура правил.");
  }
  return rules;
}

function validateBenefitsRules(rules) {
  if (!rules || !Array.isArray(rules.replacements)) {
    throw new Error("benefits_rules.json: некорректная структура правил.");
  }
  return rules;
}

function validateMatchRules(rules) {
  if (!rules || !rules.keyFieldPn || !rules.keyFieldVi || !Array.isArray(rules.fieldPairs)) {
    throw new Error("pn_vi_match_rules.json: некорректная структура правил.");
  }
  return rules;
}

function validateViRules(rules) {
  const source = rules || DEFAULT_VI_RULES;
  if (!source.examFormTypes || !source.maxScores) {
    throw new Error("rules.json: секция viRules заполнена некорректно.");
  }
  return source;
}

function createProgressReporter() {
  let phaseStart = 0;
  let phaseEnd = 100;
  let lastTickMs = 0;
  return {
    setPhase(start, end, label) {
      phaseStart = start;
      phaseEnd = end;
      setProgress(start, label);
    },
    step(done, total, label) {
      // Подпись с номером строки обновляем всегда — иначе при троттлинге 40 мс кажется, что зависло на «5538/10451».
      if (progressLabelNode) progressLabelNode.textContent = label;
      const now = Date.now();
      if (now - lastTickMs < 40 && done < total) return;
      lastTickMs = now;
      const ratio = total > 0 ? Math.max(0, Math.min(1, done / total)) : 1;
      setProgress(phaseStart + (phaseEnd - phaseStart) * ratio, label);
    }
  };
}

async function readCombinedRules(file) {
  if (!file) {
    throw new Error("Не выбран файл: rules.json.");
  }
  try {
    const text = await file.text();
    const parsed = JSON.parse(text);
    if (!parsed || typeof parsed !== "object") {
      throw new Error("bad-object");
    }
    return {
      kgRules: parsed.kgRules,
      benefitsRules: parsed.benefitsRules,
      matchRules: parsed.matchRules,
      viRules: parsed.viRules
    };
  } catch (_e) {
    throw new Error("rules.json: не удалось прочитать JSON-конфиг.");
  }
}

async function readHeadersFromFile(file, sourceLabel) {
  const wb = await readWorkbookDataFromFile(file, sourceLabel);
  return { headers: wb.headers, warnings: wb.warnings };
}

async function readWorkbookDataFromFile(file, sourceLabel) {
  if (!file) throw new Error(`Не выбран файл: ${sourceLabel}.`);
  const arrayBuffer = await file.arrayBuffer();
  return readWorkbookDataFromArrayBuffer(arrayBuffer, sourceLabel);
}

function readWorkbookDataFromArrayBuffer(arrayBuffer, sourceLabel) {
  let workbook;
  try {
    workbook = XLSX.read(arrayBuffer, { type: "array" });
  } catch (_e) {
    throw new Error(`${sourceLabel}: неподдерживаемый или поврежденный Excel-файл.`);
  }
  if (!workbook.SheetNames || workbook.SheetNames.length === 0) throw new Error(`${sourceLabel}: в файле нет листов.`);
  const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
  const sheetRows = XLSX.utils.sheet_to_json(firstSheet, { header: 1, raw: false, blankrows: false });
  if (!sheetRows || sheetRows.length === 0) throw new Error(`${sourceLabel}: на первом листе нет данных.`);

  let row0 = Array.isArray(sheetRows[0]) ? sheetRows[0] : [];
  let row1 = sheetRows.length > 1 && Array.isArray(sheetRows[1]) ? sheetRows[1] : [];
  [row0, row1] = expandHeaderRowsFromSheetMerges(firstSheet, row0, row1);

  const row1HasContent = row1.some((c) => String(c || "").trim() !== "");
  const mergedAcrossTwoHeaderRows = sheetHasMergedHeaderRows(firstSheet);
  const secondRowIsProbablyData = sheetSecondRowLooksLikeDataNotSubHeader(row1);
  const useTwoRowHeader =
    !secondRowIsProbablyData && (mergedAcrossTwoHeaderRows || row1HasContent);

  let built;
  let firstDataRowIndex;
  if (useTwoRowHeader) {
    built = buildCompositeHeadersFromRows(row0, row1);
    firstDataRowIndex = 2;
  } else {
    built = buildLegacySingleRowHeaders(row0);
    firstDataRowIndex = 1;
  }

  const { headers, headerColumns, headerRawRows, maxCol } = built;
  if (headers.length === 0) throw new Error(`${sourceLabel}: не найдено ни одного валидного заголовка.`);
  // Строк данных может не быть (например шаблон ПН только с шапкой) — это нормально; проверка содержимого — по основному файлу.

  const rows = [];
  for (let rowIndex = firstDataRowIndex; rowIndex < sheetRows.length; rowIndex += 1) {
    const row = Array.isArray(sheetRows[rowIndex]) ? sheetRows[rowIndex] : [];
    const rowObj = {};
    headers.forEach((header, idx) => {
      const col = headerColumns[idx];
      rowObj[header] = row[col] == null ? "" : String(row[col]);
    });
    rows.push(rowObj);
  }
  const headerRowCount = useTwoRowHeader ? 2 : 1;
  return { headers, rows, warnings: [], headerRawRows, headerColumns, maxCol, headerRowCount };
}

function normalizeText(value) {
  let s = String(value || "");
  if (s.length > MAX_NORMALIZE_CHARS) s = s.slice(0, MAX_NORMALIZE_CHARS);
  return s.replace(/\s+/g, " ").trim().toLowerCase();
}

function padRowToLength(row, len) {
  const arr = Array.isArray(row) ? row : [];
  const out = [];
  for (let i = 0; i < len; i += 1) out.push(i < arr.length ? arr[i] : undefined);
  return out;
}

function getSheetCellDisplay(sheet, r, c) {
  const addr = XLSX.utils.encode_cell({ r, c });
  const cell = sheet[addr];
  if (!cell || cell.t === "z") return "";
  if (cell.w != null) return String(cell.w).trim();
  if (cell.v != null) return String(cell.v).trim();
  return "";
}

/** Подставляет текст из master-ячейки во все клетки объединения (иначе вторая строка шапки часто пустая). */
function expandHeaderRowsFromSheetMerges(sheet, row0, row1) {
  let refEndCol = 0;
  const ref = sheet["!ref"];
  if (ref) {
    try {
      refEndCol = XLSX.utils.decode_range(ref).e.c + 1;
    } catch (_e) {
      refEndCol = 0;
    }
  }
  const len = Math.max(row0.length, row1.length, refEndCol, 1);
  const p0 = padRowToLength(row0, len);
  const p1 = padRowToLength(row1, len);
  const merges = sheet["!merges"] || [];
  merges.forEach((m) => {
    const { s, e } = m;
    if (s.r > 1) return;
    const txt = getSheetCellDisplay(sheet, s.r, s.c);
    if (!txt) return;
    const endRow = Math.min(e.r, 1);
    for (let r = Math.max(0, s.r); r <= endRow; r += 1) {
      for (let c = s.c; c <= e.c; c += 1) {
        if (r === 0) {
          const cur = p0[c] == null ? "" : String(p0[c]).trim();
          if (!cur) p0[c] = txt;
        }
        if (r === 1) {
          const cur = p1[c] == null ? "" : String(p1[c]).trim();
          if (!cur) p1[c] = txt;
        }
      }
    }
  });
  return [p0, p1];
}

function sheetHasMergedHeaderRows(sheet) {
  const merges = sheet["!merges"] || [];
  return merges.some((m) => m.s.r <= 1 && m.e.r >= 1);
}

/**
 * Вторая строка листа при однострочной шапке часто — уже данные (КГ с «|», длинные поля).
 * Раньше любая непустая row1 включала режим двух строк шапки и ломала столбцы/проверки.
 */
function sheetSecondRowLooksLikeDataNotSubHeader(row1) {
  if (!Array.isArray(row1) || !row1.length) return false;
  for (const cell of row1) {
    const s = String(cell ?? "").trim();
    if (!s) continue;
    if (s.length > 180) return true;
    const pipes = (s.match(/\|/g) || []).length;
    if (pipes >= 2 && s.length > 20) return true;
  }
  return false;
}

/** Для объединённых по горизонтали ячеек в 1-й строке шапки (например «Срок обучения» над V–W). */
function forwardFillTopRow(row0) {
  let last = "";
  return row0.map((cell) => {
    const v = cell == null ? "" : String(cell).trim();
    if (v) last = v;
    return v || last;
  });
}

/** Составной заголовок «родитель / дочерний»; иначе одна строка (в т.ч. «… (2)» после uniquify). */
function splitCompositeHeaderLabel(full) {
  const s = String(full || "").trim();
  if (!s) return { kind: "empty" };
  const m = s.match(/^(.+?)\s+\/\s+(.+)$/);
  if (m) return { kind: "split", top: m[1].trim(), bottom: m[2].trim() };
  return { kind: "single", text: s };
}

function buildHeaderRowsFromUniquifiedLabels(headers) {
  const row1 = [];
  const row2 = [];
  for (const lab of headers) {
    const sp = splitCompositeHeaderLabel(lab);
    if (sp.kind === "split") {
      row1.push(sp.top);
      row2.push(sp.bottom);
    } else if (sp.kind === "single") {
      row1.push(sp.text);
      row2.push("");
    } else {
      row1.push("");
      row2.push("");
    }
  }
  return [row1, row2];
}

function buildCompositeHeadersFromRows(row0, row1) {
  const maxCol = Math.max(row0.length, row1.length, 0);
  const p0 = padRowToLength(row0, maxCol);
  const p1 = padRowToLength(row1, maxCol);
  const topFilled = forwardFillTopRow(p0);
  const headers = [];
  const headerColumns = [];

  for (let i = 0; i < maxCol; i += 1) {
    const t = String(topFilled[i] || "").trim();
    const b = p1[i] == null ? "" : String(p1[i]).trim();
    let name = "";
    if (t && b) {
      name = normalizeText(t) === normalizeText(b) ? t : `${t} / ${b}`;
    } else if (t && !b) {
      name = t;
    } else if (!t && b) {
      name = b;
    } else {
      continue;
    }
    const trimmed = String(name).trim();
    if (!trimmed) continue;
    headers.push(trimmed);
    headerColumns.push(i);
  }

  return { headers, headerColumns, headerRawRows: [p0, p1], maxCol };
}

function buildLegacySingleRowHeaders(row0) {
  const headers = [];
  const headerColumns = [];
  for (let i = 0; i < row0.length; i += 1) {
    const value = row0[i] == null ? "" : String(row0[i]).trim();
    if (!value) continue;
    headers.push(value);
    headerColumns.push(i);
  }
  const maxCol = Math.max(row0.length, 1);
  return { headers, headerColumns, headerRawRows: null, maxCol };
}

/** Последний сегмент после « / » — как в шаблоне 1С, когда верхняя строка дублирует группу, а имя поля во второй. */
function lastHeaderSegmentNormalized(header) {
  const parts = String(header || "")
    .split(/\s*\/\s*/)
    .map((p) => p.trim())
    .filter(Boolean);
  if (!parts.length) return "";
  return normalizeText(parts[parts.length - 1]);
}

/**
 * Шаблон с простыми именами столбцов vs файл с составными «Группа / Поле».
 * Совпадение: полная строка или совпадение последних сегментов / простое имя с последним сегментом файла.
 */
function headerNamesCorrespondForComparison(templateHeader, fileHeader) {
  const nt = normalizeText(templateHeader);
  const nf = normalizeText(fileHeader);
  if (nt && nf && nt === nf) return true;
  const lastT = lastHeaderSegmentNormalized(templateHeader);
  const lastF = lastHeaderSegmentNormalized(fileHeader);
  if (lastT && lastF && lastT === lastF) return true;
  if (nt && lastF && nt === lastF) return true;
  if (nf && lastT && nf === lastT) return true;
  return false;
}

function compareHeaders(templateHeaders, fileHeaders) {
  const missing = [];
  const extra = [];
  for (const th of templateHeaders) {
    if (!normalizeText(th)) continue;
    const ok = fileHeaders.some((fh) => headerNamesCorrespondForComparison(th, fh));
    if (!ok) missing.push(th);
  }
  for (const fh of fileHeaders) {
    if (!normalizeText(fh)) continue;
    const ok = templateHeaders.some((th) => headerNamesCorrespondForComparison(th, fh));
    if (!ok) extra.push(fh);
  }
  return { missing, extra, matched: [] };
}

async function readMinBallFromFile(file) {
  if (!file) throw new Error("Не выбран файл: min_ball.txt.");
  const lines = (await file.text()).split(/\r?\n/);
  const map = new Map();
  const warnings = [];
  lines.forEach((line, idx) => {
    const t = line.trim();
    if (!t) return;
    const parts = t.split(";").map((p) => p.trim());
    if (parts.length < 2) {
      warnings.push(`min_ball.txt: строка ${idx + 1} пропущена (формат "Предмет;Минимальный балл").`);
      return;
    }
    const subject = parts.slice(0, -1).join(";").trim();
    const scoreRaw = parts[parts.length - 1];
    const score = Number(scoreRaw.trim().replace(",", "."));
    if (!subject || Number.isNaN(score)) {
      warnings.push(`min_ball.txt: строка ${idx + 1} пропущена (некорректные данные).`);
      return;
    }
    const key = normalizeText(subject);
    map.set(key, { subject, score });
  });
  if (map.size === 0) throw new Error("min_ball.txt: не найдено валидных строк с минимальными баллами.");
  return { map, warnings };
}

async function readSpoMinBallFromFile(file) {
  if (!file) throw new Error("Не выбран файл: spo_min_ball.txt.");
  const lines = (await file.text()).split(/\r?\n/);
  const spoMap = new Map();
  const warnings = [];
  lines.forEach((line, idx) => {
    const t = line.trim();
    if (!t) return;
    const parts = t.split(";").map((p) => p.trim());
    if (parts.length < 2) {
      warnings.push(`spo_min_ball.txt: строка ${idx + 1} пропущена (формат "Предмет;Минимальный балл").`);
      return;
    }
    const subject = parts.slice(0, -1).join(";").trim();
    const score = Number(parts[parts.length - 1].trim().replace(",", "."));
    if (!subject || Number.isNaN(score)) {
      warnings.push(`spo_min_ball.txt: строка ${idx + 1} пропущена (некорректные данные).`);
      return;
    }
    spoMap.set(normalizeSpoSubjectKey(subject), { subject, score });
  });
  if (spoMap.size === 0) throw new Error("spo_min_ball.txt: не найдено валидных строк с минимальными баллами.");
  return { spoMap, warnings };
}

async function runPnCriteriaChecks(pnWorkbookData, kgRules, benefitsRules, onProgress, templateHeaders = []) {
  const headers = pnWorkbookData.headers;
  const rows = pnWorkbookData.rows;
  const h = {
    kg: findHeader(headers, "Конкурсная группа", templateHeaders),
    level: findHeader(headers, "Уровень образования", templateHeaders),
    form: findHeader(headers, "Форма обучения", templateHeaders),
    finance: findHeader(headers, "Источник финансирования", templateHeaders),
    code: findHeader(headers, "Код", templateHeaders),
    foreignOnly: findHeader(headers, "Только для иностранных граждан", templateHeaders),
    rfOnly: findHeader(headers, "Только для граждан РФ", templateHeaders),
    targetDetailed: findHeader(headers, "Целевая детализированная квота", templateHeaders),
    benefit: findHeader(headers, "Название льготы", templateHeaders),
    specialMark: findHeader(headers, "Особая отметка", templateHeaders),
    specialFeature: findHeader(headers, kgRules.separateQuotaRule.specialFeatureColumn, templateHeaders),
    specialRight: findHeader(headers, kgRules.separateQuotaRule.specialRightColumn, templateHeaders)
  };

  const kgNameIssues = [];
  const codeIssues = [];
  const levelIssues = [];
  const foreignIssues = [];
  const rfIssues = [];
  const benefitIssues = [];
  const targetDetailedIssues = [];
  const benefitPatterns = (benefitsRules.replacements || []).map((item) => ({
    fromRaw: item.from,
    fromNorm: normalizeBenefitText(item.from),
    to: item.to
  }));

  for (let idx = 0; idx < rows.length; idx += 1) {
    const row = rows[idx];
    const rowRef = `Строка ${idx + 2}`;
    const kg = getVal(row, h.kg);
    const level = getVal(row, h.level);
    const form = getVal(row, h.form);
    const finance = getVal(row, h.finance);
    const code = getVal(row, h.code);
    const foreignOnly = getVal(row, h.foreignOnly);
    const rfOnly = getVal(row, h.rfOnly);
    const targetDetailed = getVal(row, h.targetDetailed);
    const benefit = getVal(row, h.benefit);
    const specialMark = getVal(row, h.specialMark);
    const feature = getVal(row, h.specialFeature);
    const specialRight = getVal(row, h.specialRight);
    const hasTokenSpacingIssue = hasPipeTokenSpacingIssue(kg);

    if (hasTokenSpacingIssue) {
      kgNameIssues.push(
        issue(
          "kg-token-spacing",
          rowRef,
          `В названии КГ "${kg}" вокруг разделителя "|" должны быть пробелы с двух сторон (например, "БАК | З | Ц | ...").`
        )
      );
    }

    if (kg && code && !kg.includes(code)) {
      codeIssues.push(issue("code-missing-in-kg", rowRef, `Код "${code}" отсутствует в названии КГ "${kg}".`));
    }

    const levelToken = tokenByValue(level, kgRules.levelTokens);
    if (levelToken && !containsToken(kg, levelToken)) {
      levelIssues.push(issue("level-token", rowRef, `Для уровня "${level}" в КГ должно быть "${levelToken}".`));
    }

    const formToken = tokenByValue(form, kgRules.formTokens);
    if (formToken && !containsToken(kg, formToken)) {
      kgNameIssues.push(issue("form-token", rowRef, `Форма обучения "${form}" не отражена в КГ (ожидали "${formToken}").`));
    }

    const financeToken = tokenByValue(finance, kgRules.financeTokens);
    if (financeToken && !containsToken(kg, financeToken)) {
      kgNameIssues.push(issue("finance-token", rowRef, `Источник финансирования "${finance}" не отражен в КГ (ожидали "${financeToken}").`));
    }

    if (containsToken(kg, kgRules.quotaTokens.foreign)) {
      if (!isYes(foreignOnly, kgRules)) {
        foreignIssues.push(issue("foreign-only", rowRef, `КГ с "ИН": поле "Только для иностранных граждан" должно быть "Да".`));
      }
      if (!isNo(rfOnly, kgRules)) {
        foreignIssues.push(issue("foreign-rf", rowRef, `КГ с "ИН": поле "Только для граждан РФ" должно быть "Нет".`));
      }
    }

    if (rfOnly && !isNo(rfOnly, kgRules)) {
      rfIssues.push(issue("rf-no", rowRef, `Поле "Только для граждан РФ" должно быть "Нет", сейчас "${rfOnly}".`));
    }

    // Токен «Ц» ищем всегда по строке КГ; пробелы у «|» — отдельная ошибка kg-token-spacing и не должны обнулять признак целевой квоты.
    const hasTargetToken = containsToken(kg, kgRules.quotaTokens.target);
    if (targetDetailed) {
      const should = hasTargetToken ? "Да" : "Нет";
      if (normalizeText(targetDetailed) !== normalizeText(should)) {
        targetDetailedIssues.push(issue("target-detailed", rowRef, `Для КГ "${kg}" поле "Целевая детализированная квота" должно быть "${should}".`));
      }
    }

    const separateExpected = normalizeText(feature) === normalizeText(kgRules.separateQuotaRule.specialFeatureValue);
    if (separateExpected) {
      if (!containsToken(kg, kgRules.quotaTokens.separate)) {
        kgNameIssues.push(issue("separate-token", rowRef, `Для "Отдельная квота" в КГ должен быть токен "${kgRules.quotaTokens.separate}".`));
      }
      if (normalizeText(specialRight) !== normalizeText(kgRules.separateQuotaRule.specialRightExpected)) {
        kgNameIssues.push(issue("special-right", rowRef, `Для "Отдельная квота" поле "Особое право" должно быть "${kgRules.separateQuotaRule.specialRightExpected}".`));
      }
    }

    const benefitColumns = [
      { header: h.benefit, value: benefit },
      { header: h.specialMark, value: specialMark }
    ].filter((entry) => entry.header && entry.value);
    for (const column of benefitColumns) {
      const expected = applyBenefitToValue(column.value, benefitPatterns);
      if (expected !== null) {
        benefitIssues.push(
          issue(
            "benefit-replace",
            rowRef,
            `Некорректная формулировка в столбце "${column.header}". Найдено: "${String(column.value).trim()}". Значение должно быть заменено на "${expected}".`
          )
        );
      }
    }
    if (onProgress) onProgress(idx + 1, rows.length, `Проверка критериев ПН... ${idx + 1}/${rows.length}`);
    if ((idx + 1) % CHECK_LOOP_YIELD_EVERY === 0) await yieldToUi();
  }

  return { kgNameIssues, codeIssues, levelIssues, foreignIssues, rfIssues, benefitIssues, targetDetailedIssues };
}

async function checkViCriteria(viWorkbookData, minBallData, spoMinBallData, benefitsRules, kgRules, viRules, onProgress, templateHeaders = []) {
  const subjectHeader = findHeader(viWorkbookData.headers, viRules.subjectColumn, templateHeaders);
  const replaceSubjectHeader = findHeader(viWorkbookData.headers, viRules.replaceSubjectColumn, templateHeaders);
  const minScoreHeader = findHeader(viWorkbookData.headers, viRules.minScoreColumn, templateHeaders);
  const maxScoreHeader = findHeader(viWorkbookData.headers, viRules.maxScoreColumn, templateHeaders);
  const testFormHeader = findHeader(viWorkbookData.headers, viRules.testFormColumn, templateHeaders);
  const testTypeHeader = findHeader(viWorkbookData.headers, viRules.testTypeColumn, templateHeaders);
  const targetHeader = findHeader(viWorkbookData.headers, viRules.targetColumn, templateHeaders);
  const specialMarkHeader = findHeader(viWorkbookData.headers, viRules.specialMarkColumn, templateHeaders);

  const maxIssues = [];
  const minIssues = [];
  const rule21Issues = [];
  const rule41Issues = [];
  const spoColumnIssues = [];
  const specialMarkIssues = [];
  const benefitPatterns = (benefitsRules.replacements || []).map((item) => ({
    fromRaw: item.from,
    fromNorm: normalizeBenefitText(item.from),
    to: item.to
  }));

  if (!subjectHeader || !replaceSubjectHeader || !minScoreHeader || !maxScoreHeader || !testFormHeader) {
    return {
      maxIssues: [issue("missing-required-columns", "ВИ", "Не найдены обязательные столбцы для критериев ВИ.")],
      minIssues: [],
      rule21Issues: [],
      rule41Issues: [],
      spoColumnIssues: [],
      specialMarkIssues: []
    };
  }

  for (let index = 0; index < viWorkbookData.rows.length; index += 1) {
    const row = viWorkbookData.rows[index];
    const rowRef = `Строка ${index + 2}`;
    const baseSubject = getVal(row, subjectHeader);
    const altSubject = getVal(row, replaceSubjectHeader);
    const chosenSubject = altSubject || baseSubject;
    const minRaw = getVal(row, minScoreHeader);
    const maxRaw = getVal(row, maxScoreHeader);
    const form = getVal(row, testFormHeader);
    const testType = getVal(row, testTypeHeader);
    const targetFlag = getVal(row, targetHeader);
    const specialMark = getVal(row, specialMarkHeader);

    const formType = detectViFormType(form, viRules);
    const minScore = Number(String(minRaw).replace(",", "."));
    const maxScore = Number(String(maxRaw).replace(",", "."));
    const isTarget = isYes(targetFlag, kgRules);

    if (baseSubject && spoMinBallData.spoMap.has(normalizeSpoSubjectKey(baseSubject))) {
      spoColumnIssues.push(
        issue(
          "spo-subject-in-base",
          rowRef,
          `Предмет СПО «${baseSubject}» не должен быть в столбце «Предмет». Укажите его только в столбце «Заменяемый предмет» (список СПО — spo_min_ball.txt).`
        )
      );
    }
    const spoSubjectInBaseColumn = Boolean(
      baseSubject && spoMinBallData.spoMap.has(normalizeSpoSubjectKey(baseSubject))
    );

    if (!Number.isNaN(maxScore)) {
      const expectedMax = getExpectedMax(formType, isTarget, viRules);
      if (expectedMax !== null && maxScore !== expectedMax) {
        maxIssues.push(issue("vi-max", rowRef, `Некорректный максимальный балл: ожидается ${expectedMax}, указано ${maxScore}.`));
      }
    }

    if (Number.isNaN(minScore)) {
      if (minRaw) minIssues.push(issue("invalid-score", rowRef, `Некорректное значение минимального балла "${minRaw}".`));
    } else {
      if (formType === "id") {
        if (minScore !== 0) {
          minIssues.push(issue("vi-min-id", rowRef, `Для ИД минимальный балл должен быть 0, указано ${minScore}.`));
        }
      } else if (formType === "ege" || formType === "exam") {
        if (!chosenSubject) {
          minIssues.push(issue("empty-subjects", rowRef, "Не заполнены поля \"Предмет\" и \"Заменяемый предмет\"."));
        } else if (!spoSubjectInBaseColumn) {
          const resolved = resolveViMinScoreRef(chosenSubject, altSubject, minBallData, spoMinBallData);
          if (resolved.kind === "missing") {
            minIssues.push(
              issue(
                "missing-reference",
                rowRef,
                `Некорректный минимальный балл по предмету «${chosenSubject}»: нет эталона в min_ball.txt и в spo_min_ball.txt.`
              )
            );
          } else if (resolved.kind === "spo-requires-replace") {
            minIssues.push(issue("spo-requires-replace-column", rowRef, resolved.message));
          } else if (resolved.score !== minScore) {
            minIssues.push(
              issue(
                "score-mismatch",
                rowRef,
                `Некорректный минимальный балл по предмету «${chosenSubject}»: в файле ${minScore}, в ${resolved.sourceName} ${resolved.score}.`
              )
            );
          }
        }
      }

      if (minScore === 21) {
        if (normalizeText(testType) !== normalizeText(viRules.creativeTypeValue)) {
          rule21Issues.push(issue("vi-rule-21-type", rowRef, `При минимальном балле 21 тип испытания должен быть "${viRules.creativeTypeValue}".`));
        }
        if (altSubject) {
          rule21Issues.push(issue("vi-rule-21-replace", rowRef, "При минимальном балле 21 поле \"Заменяемый предмет\" должно быть пустым."));
        }
      }

      if (minScore === 41 && !spoSubjectInBaseColumn) {
        const spoKey = normalizeSpoSubjectKey(chosenSubject);
        const spoRef = spoMinBallData.spoMap.get(spoKey);
        if (!spoRef) {
          rule41Issues.push(
            issue(
              "vi-rule-41-spo",
              rowRef,
              `Минимальный балл 41 допустим только для предметов СПО. Предмет «${chosenSubject}» не найден в spo_min_ball.txt.`
            )
          );
        }
      }
    }

    if (specialMark) {
      const expected = applyBenefitToValue(specialMark, benefitPatterns);
      if (expected !== null) {
        specialMarkIssues.push(
          issue(
            "benefit-replace",
            rowRef,
            `Некорректная формулировка в "Особая отметка". Найдено: "${String(specialMark).trim()}". Значение должно быть заменено на "${expected}".`
          )
        );
      }
    }
    if (onProgress) onProgress(index + 1, viWorkbookData.rows.length, `Проверка критериев ВИ... ${index + 1}/${viWorkbookData.rows.length}`);
    if ((index + 1) % CHECK_LOOP_YIELD_EVERY === 0) await yieldToUi();
  }

  return { maxIssues, minIssues, rule21Issues, rule41Issues, spoColumnIssues, specialMarkIssues };
}

async function checkWhitespaceIssues(workbookData, fileLabel, onProgress) {
  const issues = [];
  for (let rowIndex = 0; rowIndex < workbookData.rows.length; rowIndex += 1) {
    const row = workbookData.rows[rowIndex];
    const rowRef = `Строка ${rowIndex + 2}`;
    for (const header of workbookData.headers) {
      const rawValue = row[header] == null ? "" : String(row[header]);
      if (!rawValue) continue;
      if (rawValue !== rawValue.trim()) {
        issues.push(issue(
          "whitespace-trim",
          rowRef,
          `[${fileLabel}] В столбце "${header}" есть пробел в начале или в конце значения.`
        ));
      }
      if (rawValue.includes("  ")) {
        issues.push(issue(
          "whitespace-double",
          rowRef,
          `[${fileLabel}] В столбце "${header}" есть двойной пробел внутри значения.`
        ));
      }
    }
    if (onProgress) onProgress(rowIndex + 1, workbookData.rows.length, `Проверка пробелов ${fileLabel}... ${rowIndex + 1}/${workbookData.rows.length}`);
    if ((rowIndex + 1) % CHECK_LOOP_YIELD_EVERY === 0) await yieldToUi();
  }
  return { issues };
}

async function checkQualificationByLevel(workbookData, fileLabel, onProgress, templateHeaders = []) {
  const issues = [];
  const levelHeader =
    findHeader(workbookData.headers, "Уровень образования", templateHeaders) ||
    findHeader(workbookData.headers, "Уровень подготовки", templateHeaders);
  const qualificationHeader = findHeader(workbookData.headers, "Квалификация", templateHeaders);
  if (!levelHeader || !qualificationHeader) {
    return {
      issues: [
        issue(
          "qualification-missing-columns",
          fileLabel,
          `[${fileLabel}] Не найдены столбцы "Уровень образования" (или "Уровень подготовки") и/или "Квалификация".`
        )
      ]
    };
  }

  for (let rowIndex = 0; rowIndex < workbookData.rows.length; rowIndex += 1) {
    const row = workbookData.rows[rowIndex];
    const rowRef = `Строка ${rowIndex + 2}`;
    const level = getVal(row, levelHeader);
    const qualification = getVal(row, qualificationHeader);
    const levelNorm = normalizeText(level);

    let expected = "-";
    if (levelNorm.includes("бакалавриат")) expected = "Бакалавр";
    else if (levelNorm.includes("магистратура")) expected = "Магистр";

    if (!qualification) {
      issues.push(issue("qualification-empty", rowRef, `[${fileLabel}] Поле "Квалификация" не должно быть пустым.`));
    } else if (normalizeText(qualification) !== normalizeText(expected)) {
      issues.push(
        issue(
          "qualification-mismatch",
          rowRef,
          `[${fileLabel}] Некорректная квалификация: при уровне "${level || "-"}" ожидается "${expected}", указано "${qualification}".`
        )
      );
    }

    if (onProgress) onProgress(rowIndex + 1, workbookData.rows.length, `Проверка квалификации ${fileLabel}... ${rowIndex + 1}/${workbookData.rows.length}`);
    if ((rowIndex + 1) % CHECK_LOOP_YIELD_EVERY === 0) await yieldToUi();
  }

  return { issues };
}

function detectViFormType(formValue, viRules) {
  const n = normalizeText(formValue);
  for (const [key, type] of Object.entries(viRules.examFormTypes || {})) {
    if (n.includes(normalizeText(key))) return type;
  }
  return null;
}

function getExpectedMax(formType, isTarget, viRules) {
  if (!formType) return null;
  if (formType === "id") return isTarget ? Number(viRules.maxScores.idTarget) : Number(viRules.maxScores.id);
  if (formType === "ege") return Number(viRules.maxScores.ege);
  if (formType === "exam") return Number(viRules.maxScores.exam);
  return null;
}

function normalizeSpoSubjectKey(subject) {
  let raw = String(subject || "").trim();
  if (raw.length > MAX_NORMALIZE_CHARS) raw = raw.slice(0, MAX_NORMALIZE_CHARS);
  const withoutPrefix = raw.replace(/^\s*спо\s*[:;,-]?\s*/i, "");
  return normalizeText(withoutPrefix);
}

/**
 * Эталон минимума: обычные предметы — min_ball.txt; только СПО — spo_min_ball.txt и строка должна опираться на «Заменяемый предмет».
 */
function resolveViMinScoreRef(chosenSubject, altSubject, minBallData, spoMinBallData) {
  const chosenNorm = normalizeText(chosenSubject);
  const chosenSpoKey = normalizeSpoSubjectKey(chosenSubject);
  const refMin = minBallData.map.get(chosenNorm);
  const refSpo = spoMinBallData.spoMap.get(chosenSpoKey);

  if (refMin && refSpo) {
    return { kind: "ok", score: refMin.score, sourceName: "min_ball.txt" };
  }
  if (refMin && !refSpo) {
    return { kind: "ok", score: refMin.score, sourceName: "min_ball.txt" };
  }
  if (!refMin && refSpo) {
    const altTrim = altSubject ? String(altSubject).trim() : "";
    if (!altTrim || normalizeSpoSubjectKey(altSubject) !== chosenSpoKey) {
      return {
        kind: "spo-requires-replace",
        message: `Предмет СПО «${chosenSubject}» должен быть указан только в столбце «Заменяемый предмет» (эталон — spo_min_ball.txt).`
      };
    }
    return { kind: "ok", score: refSpo.score, sourceName: "spo_min_ball.txt" };
  }
  return { kind: "missing" };
}

function normalizeBenefitText(value) {
  let s = String(value || "");
  if (s.length > MAX_NORMALIZE_CHARS) s = s.slice(0, MAX_NORMALIZE_CHARS);
  return s
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function findBenefitReplacement(cellValue, patterns) {
  const normalized = normalizeBenefitText(cellValue);
  if (!normalized) return null;
  for (const pattern of patterns) {
    if (!pattern || !pattern.fromNorm) continue;
    if (normalized === pattern.fromNorm || normalized.includes(pattern.fromNorm)) {
      return pattern;
    }
  }
  return null;
}

async function checkPnViMatch(pnWorkbookData, viWorkbookData, rules, onProgress) {
  const issues = [];
  const pnKgHeader = findHeader(pnWorkbookData.headers, rules.keyFieldPn);
  const viKgHeader = findHeader(viWorkbookData.headers, rules.keyFieldVi);
  if (!pnKgHeader || !viKgHeader) {
    return { issues: [issue("missing-key", "ПН/ВИ", "Не найден столбец ключа Конкурсная группа для сверки ПН и ВИ.")] };
  }

  const viByKg = new Map();
  const resolvedPairs = [];
  for (const pair of rules.fieldPairs || []) {
    const pnHeader = findHeader(pnWorkbookData.headers, pair.pn);
    const viHeader = findHeader(viWorkbookData.headers, pair.vi);
    if (pnHeader && viHeader) {
      resolvedPairs.push({ pn: pair.pn, pnHeader, viHeader });
    }
  }
  for (const viRow of viWorkbookData.rows) {
    const kg = normalizeKgKey(getVal(viRow, viKgHeader));
    if (!kg || viByKg.has(kg)) continue;
    viByKg.set(kg, viRow);
  }

  for (let index = 0; index < pnWorkbookData.rows.length; index += 1) {
    const pnRow = pnWorkbookData.rows[index];
    const rowRef = `Строка ${index + 2}`;
    const kgValue = getVal(pnRow, pnKgHeader);
    const key = normalizeKgKey(kgValue);
    if (!key) continue;
    const viRow = viByKg.get(key);
    if (!viRow) {
      issues.push(issue("missing-kg-in-vi", rowRef, `КГ "${kgValue}" есть в ПН, но отсутствует в ВИ.`));
      continue;
    }
    for (const pair of resolvedPairs) {
      const pnValue = getVal(pnRow, pair.pnHeader);
      const viValue = getVal(viRow, pair.viHeader);
      if (normalizeText(pnValue) !== normalizeText(viValue)) {
        issues.push(issue("pn-vi-mismatch", rowRef, `Несовпадение ПН↔ВИ: "${pair.pn}" в ПН="${pnValue}", в ВИ="${viValue}" (КГ "${kgValue}").`));
      }
    }
    if (onProgress) onProgress(index + 1, pnWorkbookData.rows.length, `Сверка ПН и ВИ... ${index + 1}/${pnWorkbookData.rows.length}`);
    if ((index + 1) % CHECK_LOOP_YIELD_EVERY === 0) await yieldToUi();
  }

  return { issues };
}

function buildHeaderCriterionIssues(pnReport, viReport) {
  return {
    pnIssues: [
      ...pnReport.missing.map((f) => issue("missing-column", "ПН", `В файле ПН отсутствует поле "${f}".`)),
      ...pnReport.extra.map((f) => issue("extra-column", "ПН", `В файле ПН лишнее поле "${f}".`))
    ],
    viIssues: [
      ...viReport.missing.map((f) => issue("missing-column", "ВИ", `В файле ВИ отсутствует поле "${f}".`)),
      ...viReport.extra.map((f) => issue("extra-column", "ВИ", `В файле ВИ лишнее поле "${f}".`))
    ]
  };
}

function issue(type, rowRef, message) {
  return { type, rowRef, message };
}

function ensureIssues(result) {
  if (result && Array.isArray(result.issues)) return result.issues;
  return [];
}

function normalizeHeaderLookupName(value) {
  return normalizeText(value);
}

function findHeader(headers, expectedName, templateHeaders = []) {
  const target = normalizeHeaderLookupName(expectedName);
  if (!target) return null;

  for (const h of headers) {
    if (normalizeHeaderLookupName(h) === target) return h;
  }

  const byLastSegment = headers.filter((h) => {
    const parts = String(h)
      .split(/\s*\/\s*/)
      .map((p) => normalizeHeaderLookupName(p.trim()))
      .filter(Boolean);
    return parts.length > 0 && parts[parts.length - 1] === target;
  });
  if (byLastSegment.length > 0) return byLastSegment[0];

  const byIncludes = headers.filter((h) => normalizeHeaderLookupName(h).includes(target));
  if (byIncludes.length > 0) return byIncludes[0];

  if (Array.isArray(templateHeaders) && templateHeaders.length > 0) {
    const templateCandidate = templateHeaders.find((th) =>
      headerNamesCorrespondForComparison(expectedName, th)
    );
    if (templateCandidate) {
      const resolved = headers.find((h) =>
        headerNamesCorrespondForComparison(templateCandidate, h)
      );
      if (resolved) return resolved;
    }
  }

  return null;
}

function getVal(row, header) {
  return header ? String(row[header] || "").trim() : "";
}

function tokenByValue(value, map) {
  const n = normalizeText(value);
  const entries = Object.entries(map).sort((a, b) => b[0].length - a[0].length);
  for (const [key, token] of entries) {
    if (n.includes(key)) return token;
  }
  return "";
}

/**
 * Excel/Word часто подставляют не ASCII U+007C, а полноширинную черту U+FF5C и др.
 * Без замены на ASCII | не срабатывает удаление пробелов у разделителей.
 */
function normalizeKgVerticalBarsToAscii(value) {
  let s = String(value || "");
  if (s.length > MAX_NORMALIZE_CHARS) s = s.slice(0, MAX_NORMALIZE_CHARS);
  return s.replace(/\uFF5C/g, "|").replace(/\u2223/g, "|").replace(/\u2502/g, "|");
}

function containsToken(kgName, token) {
  if (!kgName || !token) return false;
  const text = fixPipeSpacingAroundBars(kgName).toUpperCase();
  const t = String(token).toUpperCase();
  const parts = text
    .split("|")
    .map((p) => p.trim())
    .filter(Boolean);
  if (parts.includes(t)) return true;
  return text.includes(` ${t} `) || text.startsWith(`${t} `) || text.endsWith(` ${t}`) || text === t;
}

function hasPipeTokenSpacingIssue(kgName) {
  if (!kgName) return false;
  const text = normalizeKgVerticalBarsToAscii(kgName);
  return /(^|[^ ])\|/.test(text) || /\|([^ ]|$)/.test(text);
}

function isYes(value, rules) {
  return rules.yesValues.includes(normalizeText(value));
}

function isNo(value, rules) {
  return rules.noValues.includes(normalizeText(value));
}

function clearFixedPreview() {
  if (!fixedTablesMount) return;
  fixedTablesMount.innerHTML = "";
  const wrap = document.createElement("div");
  wrap.className = "fixed-empty-placeholder";
  const p = document.createElement("p");
  p.innerHTML =
    "<strong>Сначала выполните проверку.</strong> На вкладке «Загрузка» выберите все файлы и нажмите «Проверить» — затем здесь появятся таблицы ПН и ВИ с подсветкой изменённых ячеек.";
  wrap.appendChild(p);
  fixedTablesMount.appendChild(wrap);
}

function fixWhitespaceInCell(raw) {
  return String(raw || "")
    .replace(/\s+/g, " ")
    .trim();
}

function fixPipeSpacingAroundBars(kgName) {
  if (kgName == null || kgName === "") return "";
  return normalizeKgVerticalBarsToAscii(kgName)
    .replace(/\s*\|\s*/g, " | ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeKgKey(value) {
  return normalizeText(fixPipeSpacingAroundBars(value));
}

function buildBenefitPatternsFromRules(benefitsRules) {
  return (benefitsRules.replacements || []).map((item) => ({
    fromNorm: normalizeBenefitText(item.from),
    to: item.to
  }));
}

/**
 * Замены по rules.json для одной ячейки. Если в тексте несколько отметок через «;»,
 * каждая часть сопоставляется отдельно, результат снова склеивается через «; ».
 */
function applyBenefitToValue(value, patterns) {
  const trimmed = String(value ?? "").trim();
  if (!trimmed) return null;

  if (!trimmed.includes(";")) {
    const matched = findBenefitReplacement(trimmed, patterns);
    return matched ? matched.to : null;
  }

  const parts = trimmed
    .split(";")
    .map((p) => p.trim())
    .filter((p) => p.length > 0);
  if (parts.length === 0) return null;

  let any = false;
  const out = parts.map((part) => {
    const matched = findBenefitReplacement(part, patterns);
    if (matched) {
      any = true;
      return String(matched.to).trim();
    }
    return part;
  });
  return any ? out.join("; ") : null;
}

function expectedQualificationForLevel(level) {
  const levelNorm = normalizeText(level);
  if (levelNorm.includes("бакалавриат")) return "Бакалавр";
  if (levelNorm.includes("магистратура")) return "Магистр";
  return "-";
}

function buildCorrectedPnData(pnWorkbookData, kgRules, benefitsRules, templateHeaders = []) {
  const headers = pnWorkbookData.headers;
  const patterns = buildBenefitPatternsFromRules(benefitsRules);
  const h = {
    kg: findHeader(headers, "Конкурсная группа", templateHeaders),
    level:
      findHeader(headers, "Уровень образования", templateHeaders) ||
      findHeader(headers, "Уровень подготовки", templateHeaders),
    benefit: findHeader(headers, "Название льготы", templateHeaders),
    specialMark: findHeader(headers, "Особая отметка", templateHeaders),
    targetDetailed: findHeader(headers, "Целевая детализированная квота", templateHeaders),
    rfOnly: findHeader(headers, "Только для граждан РФ", templateHeaders),
    foreignOnly: findHeader(headers, "Только для иностранных граждан", templateHeaders),
    specialFeature: findHeader(headers, kgRules.separateQuotaRule.specialFeatureColumn, templateHeaders),
    specialRight: findHeader(headers, kgRules.separateQuotaRule.specialRightColumn, templateHeaders)
  };
  const qualH = findHeader(headers, "Квалификация", templateHeaders);

  const rows = pnWorkbookData.rows.map((row) => {
    const next = {};
    for (const header of headers) {
      next[header] = fixWhitespaceInCell(row[header]);
    }
    if (h.kg && next[h.kg]) {
      next[h.kg] = fixPipeSpacingAroundBars(next[h.kg]);
    }
    if (h.benefit && next[h.benefit]) {
      const rep = applyBenefitToValue(next[h.benefit], patterns);
      if (rep !== null) next[h.benefit] = rep;
    }
    if (h.specialMark && next[h.specialMark]) {
      const rep = applyBenefitToValue(next[h.specialMark], patterns);
      if (rep !== null) next[h.specialMark] = rep;
    }
    if (h.level && qualH) {
      next[qualH] = expectedQualificationForLevel(getVal(next, h.level));
    }
    if (h.kg && h.targetDetailed && next[h.kg] && getVal(row, h.targetDetailed)) {
      const should = containsToken(next[h.kg], kgRules.quotaTokens.target) ? "Да" : "Нет";
      next[h.targetDetailed] = should;
    }
    if (h.rfOnly && getVal(row, h.rfOnly) && !isNo(getVal(next, h.rfOnly), kgRules)) {
      next[h.rfOnly] = "Нет";
    }
    if (h.kg && h.foreignOnly && next[h.kg] && containsToken(next[h.kg], kgRules.quotaTokens.foreign)) {
      if (!isYes(getVal(next, h.foreignOnly), kgRules)) {
        next[h.foreignOnly] = "Да";
      }
    }
    if (
      h.specialFeature &&
      h.specialRight &&
      normalizeText(getVal(next, h.specialFeature)) === normalizeText(kgRules.separateQuotaRule.specialFeatureValue)
    ) {
      next[h.specialRight] = kgRules.separateQuotaRule.specialRightExpected;
    }
    return next;
  });

  return {
    headers,
    rows,
    headerRawRows: pnWorkbookData.headerRawRows,
    headerColumns: pnWorkbookData.headerColumns,
    maxCol: pnWorkbookData.maxCol,
    headerRowCount: pnWorkbookData.headerRowCount
  };
}

function buildCorrectedViData(viWorkbookData, minBallData, spoMinBallData, benefitsRules, kgRules, viRules, templateHeaders = []) {
  const headers = viWorkbookData.headers;
  const patterns = buildBenefitPatternsFromRules(benefitsRules);
  const kgHeader = findHeader(headers, "Конкурсная группа", templateHeaders);
  const subjectHeader = findHeader(headers, viRules.subjectColumn, templateHeaders);
  const replaceHeader = findHeader(headers, viRules.replaceSubjectColumn, templateHeaders);
  const minScoreHeader = findHeader(headers, viRules.minScoreColumn, templateHeaders);
  const maxScoreHeader = findHeader(headers, viRules.maxScoreColumn, templateHeaders);
  const testFormHeader = findHeader(headers, viRules.testFormColumn, templateHeaders);
  const targetHeader = findHeader(headers, viRules.targetColumn, templateHeaders);
  const specialMarkHeader = findHeader(headers, viRules.specialMarkColumn, templateHeaders);
  const levelHeader =
    findHeader(headers, "Уровень образования", templateHeaders) ||
    findHeader(headers, "Уровень подготовки", templateHeaders);
  const qualHeader = findHeader(headers, "Квалификация", templateHeaders);

  const rows = viWorkbookData.rows.map((row) => {
    const next = {};
    for (const header of headers) {
      next[header] = fixWhitespaceInCell(row[header]);
    }
    if (kgHeader && next[kgHeader]) {
      next[kgHeader] = fixPipeSpacingAroundBars(next[kgHeader]);
    }

    if (subjectHeader && replaceHeader) {
      const base = next[subjectHeader];
      const alt = next[replaceHeader];
      if (base && spoMinBallData.spoMap.has(normalizeSpoSubjectKey(base)) && !String(alt || "").trim()) {
        next[replaceHeader] = base;
        next[subjectHeader] = "";
      }
    }

    if (specialMarkHeader && next[specialMarkHeader]) {
      const rep = applyBenefitToValue(next[specialMarkHeader], patterns);
      if (rep !== null) next[specialMarkHeader] = rep;
    }

    if (levelHeader && qualHeader) {
      next[qualHeader] = expectedQualificationForLevel(getVal(next, levelHeader));
    }

    const form = getVal(next, testFormHeader);
    const formType = detectViFormType(form, viRules);
    const targetFlag = getVal(next, targetHeader);
    const isTarget = isYes(targetFlag, kgRules);

    const maxRaw = maxScoreHeader ? next[maxScoreHeader] : "";
    const maxScore = Number(String(maxRaw || "").replace(",", "."));

    if (maxScoreHeader && !Number.isNaN(maxScore) && formType) {
      const expectedMax = getExpectedMax(formType, isTarget, viRules);
      if (expectedMax !== null) next[maxScoreHeader] = String(expectedMax);
    }

    if (minScoreHeader && formType === "id") {
      next[minScoreHeader] = "0";
    }

    const chosenAfter = getVal(next, replaceHeader) || getVal(next, subjectHeader);
    const altAfter = getVal(next, replaceHeader);

    if (minScoreHeader && (formType === "ege" || formType === "exam") && chosenAfter) {
      const resolved = resolveViMinScoreRef(chosenAfter, altAfter, minBallData, spoMinBallData);
      if (resolved.kind === "ok") {
        next[minScoreHeader] = String(resolved.score);
      }
    }

    const minScore = Number(String(getVal(next, minScoreHeader) || "").replace(",", "."));
    if (minScoreHeader && replaceHeader && minScore === 21) {
      next[replaceHeader] = "";
    }

    return next;
  });

  return {
    headers,
    rows,
    headerRawRows: viWorkbookData.headerRawRows,
    headerColumns: viWorkbookData.headerColumns,
    maxCol: viWorkbookData.maxCol,
    headerRowCount: viWorkbookData.headerRowCount
  };
}

function workbookDataHasTwoHeaderRows(workbookData) {
  if (!workbookData) return false;
  if (workbookData.headerRowCount === 2) return true;
  return Boolean(
    workbookData.headerRawRows &&
      workbookData.headerRawRows.length >= 2 &&
      workbookData.headerRawRows[1].some((c) => String(c || "").trim() !== "")
  );
}

function workbookDataToSheet(workbookData) {
  const { headers, rows } = workbookData;
  const aoa = [];
  if (workbookDataHasTwoHeaderRows(workbookData)) {
    const [r1, r2] = buildHeaderRowsFromUniquifiedLabels(headers);
    aoa.push(r1, r2);
  } else {
    aoa.push(headers);
  }
  for (const r of rows) {
    aoa.push(headers.map((h) => (r[h] == null ? "" : String(r[h]))));
  }
  return XLSX.utils.aoa_to_sheet(aoa);
}

function downloadWorkbookDataXlsx(workbookData, filename, sheetName) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, workbookDataToSheet(workbookData), sheetName || "Лист1");
  XLSX.writeFile(wb, filename);
}

function renderFixedPreview(pnOrig, pnFixed, viOrig, viFixed) {
  if (!fixedTablesMount) return;
  fixedTablesMount.innerHTML = "";

  setProgress(98, "Сборка таблицы исправлений (ПН)...");
  fixedTablesMount.appendChild(
    buildFixedTableSection("План набора (исправлено)", pnOrig, pnFixed, "pn_corrected.xlsx", "ПН")
  );
  setProgress(99, "Сборка таблицы исправлений (ВИ)...");
  fixedTablesMount.appendChild(
    buildFixedTableSection("Вступительные испытания (исправлено)", viOrig, viFixed, "vi_corrected.xlsx", "ВИ")
  );
}

function buildFixedTableSection(title, origWb, fixedWb, downloadFilename, sheetLabel) {
  const section = document.createElement("section");
  section.className = "fixed-table-section";

  const h3 = document.createElement("h3");
  h3.textContent = title;

  const toolbar = document.createElement("div");
  toolbar.className = "fixed-table-toolbar";
  const dl = document.createElement("button");
  dl.type = "button";
  dl.textContent = "Скачать .xlsx";
  dl.addEventListener("click", () => {
    downloadWorkbookDataXlsx(fixedWb, downloadFilename, sheetLabel);
  });
  toolbar.appendChild(dl);

  const scroll = document.createElement("div");
  scroll.className = "fixed-table-scroll";
  scroll.appendChild(buildFixedHtmlTable(origWb, fixedWb));

  section.appendChild(h3);
  section.appendChild(toolbar);
  section.appendChild(scroll);
  return section;
}

function buildFixedHtmlTable(origWb, fixedWb) {
  const twoLine = workbookDataHasTwoHeaderRows(origWb) && origWb.headerColumns && origWb.headerColumns.length;
  const dataRowStart1Based = twoLine ? 3 : 2;

  const table = document.createElement("table");
  table.className = "fixed-data-table";
  const thead = document.createElement("thead");
  if (twoLine) {
    const labels = origWb.headers;
    const tr1 = document.createElement("tr");
    const corner1 = document.createElement("th");
    corner1.rowSpan = 2;
    corner1.textContent = "Стр.";
    tr1.appendChild(corner1);
    const tr2 = document.createElement("tr");
    for (let i = 0; i < labels.length; i += 1) {
      const sp = splitCompositeHeaderLabel(labels[i]);
      if (sp.kind === "split") {
        const thTop = document.createElement("th");
        thTop.textContent = sp.top;
        tr1.appendChild(thTop);
        const thBot = document.createElement("th");
        thBot.textContent = sp.bottom;
        tr2.appendChild(thBot);
      } else {
        const th = document.createElement("th");
        th.rowSpan = 2;
        th.textContent = sp.kind === "single" ? sp.text : "";
        tr1.appendChild(th);
      }
    }
    thead.appendChild(tr1);
    thead.appendChild(tr2);
  } else {
    const trh = document.createElement("tr");
    const corner = document.createElement("th");
    corner.textContent = "Стр.";
    trh.appendChild(corner);
    for (const h of origWb.headers) {
      const th = document.createElement("th");
      th.textContent = h;
      trh.appendChild(th);
    }
    thead.appendChild(trh);
  }
  table.appendChild(thead);

  const tbody = document.createElement("tbody");
  const n = Math.max(origWb.rows.length, fixedWb.rows.length);
  for (let i = 0; i < n; i += 1) {
    const origRow = origWb.rows[i] || {};
    const fixedRow = fixedWb.rows[i] || {};
    const tr = document.createElement("tr");
    const rowHead = document.createElement("th");
    rowHead.scope = "row";
    rowHead.textContent = String(i + dataRowStart1Based);
    tr.appendChild(rowHead);
    for (const h of origWb.headers) {
      const td = document.createElement("td");
      const o = origRow[h] == null ? "" : String(origRow[h]);
      const f = fixedRow[h] == null ? "" : String(fixedRow[h]);
      td.textContent = f;
      if (o !== f) {
        td.classList.add("cell-corrected");
        td.title = `Было: ${o}`;
      }
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }
  table.appendChild(tbody);
  return table;
}

function renderCriteriaReports(criteria) {
  criteriaResultsNode.innerHTML = "";
  criteriaResultsNode.classList.remove("hidden");
  for (const section of criteria) criteriaResultsNode.appendChild(renderCriteriaSection(section));
}

function renderCriteriaSection(section) {
  const wrap = document.createElement("section");
  wrap.className = "criteria-section";
  if (section.sectionId === "pn-file") {
    wrap.classList.add("criteria-section-pn");
  } else if (section.sectionId === "vi-file") {
    wrap.classList.add("criteria-section-vi");
  }
  const title = document.createElement("h3");
  title.className = "criteria-section-title";
  title.textContent = section.sectionTitle;
  wrap.appendChild(title);
  for (const block of section.blocks || []) {
    if (!block) continue;
    wrap.appendChild(renderCriterionBlock(block));
  }
  return wrap;
}

function renderCriterionBlock(criterion) {
  const issues = Array.isArray(criterion && criterion.issues) ? criterion.issues : [];
  const details = document.createElement("details");
  details.className = "criteria-block";
  details.open = issues.length > 0;
  const summary = document.createElement("summary");
  summary.textContent = `${criterion.title || "Критерий"} (${issues.length} ошибок)`;
  details.appendChild(summary);

  const content = document.createElement("div");
  content.className = "criteria-content";
  if (!issues.length) {
    const ok = document.createElement("p");
    ok.className = "ok";
    ok.textContent = "Ошибок по данному критерию не найдено.";
    content.appendChild(ok);
    details.appendChild(content);
    return details;
  }

  const groups = groupIssuesByType(issues);
  let index = 0;
  for (const [groupType, groupedIssues] of groups.entries()) {
    const groupWrap = document.createElement("section");
    groupWrap.className = "issue-group";
    const groupTitle = document.createElement("h4");
    groupTitle.className = "issue-group-title";
    groupTitle.textContent = `${getIssueTypeLabel(groupType)} (${groupedIssues.length})`;
    groupWrap.appendChild(groupTitle);
    const list = document.createElement("ol");
    list.className = "issue-list";
    for (const item of groupedIssues) {
      list.appendChild(renderIssueItem(`${criterion.id}-${index}`, item));
      index += 1;
    }
    groupWrap.appendChild(list);
    content.appendChild(groupWrap);
  }
  details.appendChild(content);
  return details;
}

function groupIssuesByType(issues) {
  const groups = new Map();
  for (const i of issues) {
    const t = i.type || "other";
    if (!groups.has(t)) groups.set(t, []);
    groups.get(t).push(i);
  }
  return groups;
}

function getIssueTypeLabel(type) {
  const map = {
    "missing-column": "Отсутствующие столбцы",
    "extra-column": "Лишние столбцы",
    "form-token": "Форма обучения в названии КГ",
    "finance-token": "Вид финансирования в названии КГ",
    "kg-token-spacing": "Пробелы вокруг токенов в названии КГ",
    "separate-token": "Признак отдельной квоты",
    "special-right": "Особое право для отдельной квоты",
    "code-missing-in-kg": "Код направления в названии КГ",
    "level-token": "Уровень образования в названии КГ",
    "foreign-only": "Только для иностранных граждан",
    "foreign-rf": "Только для граждан РФ для ИН",
    "rf-no": "Только для граждан РФ",
    "target-detailed": "Целевая детализированная квота",
    "missing-required-columns": "Отсутствуют обязательные столбцы",
    "invalid-score": "Некорректные значения баллов",
    "empty-subjects": "Пустые предметы",
    "score-mismatch": "Некорректный минимальный балл",
    "missing-reference": "Некорректный минимальный балл (нет эталона в min_ball.txt / spo_min_ball.txt)",
    "spo-subject-in-base": "Предмет СПО в столбце «Предмет»",
    "spo-requires-replace-column": "Предмет СПО не в столбце «Заменяемый предмет»",
    "vi-max": "Некорректный максимальный балл",
    "vi-min-id": "Минимальный балл для ИД",
    "vi-rule-21-type": "Правило 21 балла: тип испытания",
    "vi-rule-21-replace": "Правило 21 балла: заменяемый предмет",
    "vi-rule-41-spo": "Правило 41 балла (СПО)",
    "whitespace-trim": "Пробелы в начале/конце",
    "whitespace-double": "Двойные пробелы",
    "qualification-missing-columns": "Отсутствуют столбцы квалификации",
    "qualification-empty": "Пустая квалификация",
    "qualification-mismatch": "Несоответствие квалификации уровню образования",
    "missing-kg-in-vi": "КГ отсутствует в файле ВИ",
    "pn-vi-mismatch": "Несовпадения ПН и ВИ",
    "benefit-replace": "Некорректные названия льгот"
  };
  return map[type] || "Прочие ошибки";
}

function renderIssueItem(issueId, issueObj) {
  const li = document.createElement("li");
  li.className = "issue-item";
  if (checkedIssues.has(issueId)) li.classList.add("is-done");
  const label = document.createElement("label");
  label.className = "issue-label";
  const checkbox = document.createElement("input");
  checkbox.type = "checkbox";
  checkbox.checked = checkedIssues.has(issueId);
  checkbox.addEventListener("change", () => {
    if (checkbox.checked) {
      checkedIssues.add(issueId);
      li.classList.add("is-done");
    } else {
      checkedIssues.delete(issueId);
      li.classList.remove("is-done");
    }
  });
  const textWrap = document.createElement("div");
  const text = document.createElement("div");
  text.className = "issue-text";
  text.textContent = issueObj.message;
  const meta = document.createElement("div");
  meta.className = "issue-meta";
  meta.textContent = issueObj.rowRef;
  textWrap.appendChild(text);
  textWrap.appendChild(meta);
  label.appendChild(checkbox);
  label.appendChild(textWrap);
  li.appendChild(label);
  return li;
}

function clearGlobalMessage() {
  globalMessageNode.textContent = "";
  globalMessageNode.className = "summary";
}

function setGlobalMessage(text, typeClass) {
  globalMessageNode.textContent = text;
  globalMessageNode.className = `summary ${typeClass}`;
}

function setProgress(percent, label) {
  const clamped = Math.max(0, Math.min(100, percent));
  progressSectionNode.classList.remove("hidden");
  progressLabelNode.textContent = label;
  // Не округлять до целого: иначе длинные шаги (98–99.5) визуально «застывают» на одном проценте.
  const displayPct = clamped >= 99.95 ? "100%" : `${clamped.toFixed(1)}%`;
  progressPercentNode.textContent = displayPct;
  progressFillNode.style.width = `${clamped}%`;
  progressTrackNode.setAttribute("aria-valuenow", String(Math.round(clamped * 10) / 10));
}

function setUiBusy(isBusy) {
  Object.values(inputs).forEach((input) => { input.disabled = isBusy; });
  runCheckButton.disabled = isBusy;
  // Вкладки не отключаем: иначе во время проверки переключение не работает (disabled не получает click).
}

/** Лёгкая отдача потока; двойной rAF здесь многократно замедлял проверку на больших файлах. */
async function yieldToUi() {
  await new Promise((resolve) => setTimeout(resolve, 0));
}

function setActiveTab(tab) {
  let active = tab;
  if (active === "fixed" && !fixedTabPanel) active = "criteria";

  if (tabUploadButton) {
    tabUploadButton.classList.toggle("is-active", active === "upload");
    tabUploadButton.setAttribute("aria-selected", active === "upload" ? "true" : "false");
  }
  if (tabCriteriaButton) {
    tabCriteriaButton.classList.toggle("is-active", active === "criteria");
    tabCriteriaButton.setAttribute("aria-selected", active === "criteria" ? "true" : "false");
  }
  if (tabFixedButton) {
    tabFixedButton.classList.toggle("is-active", active === "fixed");
    tabFixedButton.setAttribute("aria-selected", active === "fixed" ? "true" : "false");
  }
  if (uploadTabPanel) uploadTabPanel.classList.toggle("is-active", active === "upload");
  if (criteriaTabPanel) criteriaTabPanel.classList.toggle("is-active", active === "criteria");
  if (fixedTabPanel) fixedTabPanel.classList.toggle("is-active", active === "fixed");

  const panelEl =
    active === "upload" ? uploadTabPanel : active === "criteria" ? criteriaTabPanel : fixedTabPanel;
  if (panelEl && typeof panelEl.scrollIntoView === "function") {
    panelEl.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }
}
