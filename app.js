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
  creativeTypeValue: "Дополнительное испытание творческой и (или) профессиональной направленности",
  russianForeignSubject: "Русский язык для иностранных граждан",
  russianForeignMinScore: 41
};

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
const uploadTabPanel = document.getElementById("uploadTabPanel");
const criteriaTabPanel = document.getElementById("criteriaTabPanel");

const checkedIssues = new Set();

Object.values(inputs).forEach((input) => input.addEventListener("change", updateRunButtonState));
runCheckButton.addEventListener("click", onRunCheck);
tabUploadButton.addEventListener("click", () => setActiveTab("upload"));
tabCriteriaButton.addEventListener("click", () => setActiveTab("criteria"));
updateRunButtonState();
setActiveTab("upload");

function updateRunButtonState() {
  const requiredInputs = [inputs.pnTemplate, inputs.viTemplate, inputs.pnFile, inputs.viFile, inputs.minBallFile, inputs.spoMinBallFile];
  const allChosen = requiredInputs.every((input) => input.files && input.files[0]);
  runCheckButton.disabled = !allChosen;
  if (!allChosen) {
    criteriaResultsNode.classList.add("hidden");
    criteriaResultsNode.innerHTML = "";
  }
}

async function onRunCheck() {
  clearGlobalMessage();
  criteriaResultsNode.classList.add("hidden");
  criteriaResultsNode.innerHTML = "";
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
    const pnCriteria = await runPnCriteriaChecks(pnWorkbookData, kgRules, benefitsRules, progress.step);
    perf.end("pn_criteria");
    progress.setPhase(78, 82, "Проверка пробелов ПН...");
    perf.start("pn_spaces", "Проверка пробелов ПН");
    const pnWhitespaceCriterion = await checkWhitespaceIssues(pnWorkbookData, "ПН", progress.step);
    perf.end("pn_spaces");
    progress.setPhase(82, 85, "Проверка квалификации ПН...");
    perf.start("pn_qualification", "Проверка квалификации ПН");
    const pnQualificationCriterion = await checkQualificationByLevel(pnWorkbookData, "ПН", progress.step);
    perf.end("pn_qualification");

    progress.setPhase(85, 92, "Проверка критериев ВИ...");
    perf.start("vi_criteria", "Проверка критериев ВИ");
    const viCriteria = await checkViCriteria(viWorkbookData, minBallData, spoMinBallData, benefitsRules, kgRules, viRules, progress.step);
    perf.end("vi_criteria");
    progress.setPhase(92, 95, "Проверка пробелов ВИ...");
    perf.start("vi_spaces", "Проверка пробелов ВИ");
    const viWhitespaceCriterion = await checkWhitespaceIssues(viWorkbookData, "ВИ", progress.step);
    perf.end("vi_spaces");
    progress.setPhase(95, 97, "Проверка квалификации ВИ...");
    perf.start("vi_qualification", "Проверка квалификации ВИ");
    const viQualificationCriterion = await checkQualificationByLevel(viWorkbookData, "ВИ", progress.step);
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
          { id: "ru-foreign-vi", title: "Русский язык для иностранных граждан", issues: viCriteria.ruForeignIssues },
          { id: "special-mark-vi", title: "Формулировки особой отметки", issues: viCriteria.specialMarkIssues },
          { id: "vi-whitespace", title: "Пробелы в значениях ячеек", issues: ensureIssues(viWhitespaceCriterion) },
          { id: "vi-qualification", title: "Квалификация по уровню образования", issues: ensureIssues(viQualificationCriterion) }
        ]
      }
    ];

    setProgress(97, "Формирование отчета по критериям...");
    perf.start("render", "Рендер отчета");
    renderCriteriaReports(criteria);
    perf.end("render");
    perf.render();
    setActiveTab("criteria");

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

  const firstRow = Array.isArray(sheetRows[0]) ? sheetRows[0] : [];
  const headers = [];
  const headerColumns = [];
  for (let i = 0; i < firstRow.length; i += 1) {
    const value = firstRow[i] == null ? "" : String(firstRow[i]).trim();
    if (!value) continue;
    headers.push(value);
    headerColumns.push(i);
  }
  if (headers.length === 0) throw new Error(`${sourceLabel}: не найдено ни одного валидного заголовка.`);

  const rows = [];
  for (let rowIndex = 1; rowIndex < sheetRows.length; rowIndex += 1) {
    const row = Array.isArray(sheetRows[rowIndex]) ? sheetRows[rowIndex] : [];
    const rowObj = {};
    headers.forEach((header, idx) => {
      const col = headerColumns[idx];
      rowObj[header] = row[col] == null ? "" : String(row[col]);
    });
    rows.push(rowObj);
  }
  return { headers, rows, warnings: [] };
}

function normalizeText(value) {
  return String(value || "").replace(/\s+/g, " ").trim().toLowerCase();
}

function compareHeaders(templateHeaders, fileHeaders) {
  const templateMap = createNormalizedMap(templateHeaders);
  const fileMap = createNormalizedMap(fileHeaders);
  const missing = [];
  const extra = [];
  for (const [n, original] of templateMap.entries()) if (!fileMap.has(n)) missing.push(original);
  for (const [n, original] of fileMap.entries()) if (!templateMap.has(n)) extra.push(original);
  return { missing, extra, matched: [] };
}

function createNormalizedMap(headers) {
  const map = new Map();
  for (const h of headers) {
    const n = normalizeText(h);
    if (n && !map.has(n)) map.set(n, h);
  }
  return map;
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

async function runPnCriteriaChecks(pnWorkbookData, kgRules, benefitsRules, onProgress) {
  const headers = pnWorkbookData.headers;
  const rows = pnWorkbookData.rows;
  const h = {
    kg: findHeader(headers, "Конкурсная группа"),
    level: findHeader(headers, "Уровень образования"),
    form: findHeader(headers, "Форма обучения"),
    finance: findHeader(headers, "Источник финансирования"),
    code: findHeader(headers, "Код"),
    foreignOnly: findHeader(headers, "Только для иностранных граждан"),
    rfOnly: findHeader(headers, "Только для граждан РФ"),
    targetDetailed: findHeader(headers, "Целевая детализированная квота"),
    benefit: findHeader(headers, "Название льготы"),
    specialMark: findHeader(headers, "Особая отметка"),
    specialFeature: findHeader(headers, kgRules.separateQuotaRule.specialFeatureColumn),
    specialRight: findHeader(headers, kgRules.separateQuotaRule.specialRightColumn)
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
      const matched = findBenefitReplacement(column.value, benefitPatterns);
      if (matched) {
        benefitIssues.push(
          issue(
            "benefit-replace",
            rowRef,
            `Некорректная формулировка в столбце "${column.header}". Найдено: "${String(column.value).trim()}". Значение должно быть заменено на "${matched.to}".`
          )
        );
      }
    }
    if (onProgress) onProgress(idx + 1, rows.length, `Проверка критериев ПН... ${idx + 1}/${rows.length}`);
    if ((idx + 1) % 200 === 0) await yieldToUi();
  }

  return { kgNameIssues, codeIssues, levelIssues, foreignIssues, rfIssues, benefitIssues, targetDetailedIssues };
}

async function checkViCriteria(viWorkbookData, minBallData, spoMinBallData, benefitsRules, kgRules, viRules, onProgress) {
  const subjectHeader = findHeader(viWorkbookData.headers, viRules.subjectColumn);
  const replaceSubjectHeader = findHeader(viWorkbookData.headers, viRules.replaceSubjectColumn);
  const minScoreHeader = findHeader(viWorkbookData.headers, viRules.minScoreColumn);
  const maxScoreHeader = findHeader(viWorkbookData.headers, viRules.maxScoreColumn);
  const testFormHeader = findHeader(viWorkbookData.headers, viRules.testFormColumn);
  const testTypeHeader = findHeader(viWorkbookData.headers, viRules.testTypeColumn);
  const targetHeader = findHeader(viWorkbookData.headers, viRules.targetColumn);
  const specialMarkHeader = findHeader(viWorkbookData.headers, viRules.specialMarkColumn);

  const maxIssues = [];
  const minIssues = [];
  const rule21Issues = [];
  const rule41Issues = [];
  const ruForeignIssues = [];
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
      ruForeignIssues: [],
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
        } else {
          const ref = minBallData.map.get(normalizeText(chosenSubject));
          if (!ref) {
            minIssues.push(issue("missing-reference", rowRef, `Некорректный минимальный балл по предмету "${chosenSubject}": отсутствует значение в min_ball.txt.`));
          } else if (ref.score !== minScore) {
            minIssues.push(issue("score-mismatch", rowRef, `Некорректный минимальный балл по предмету "${chosenSubject}": в файле ${minScore}, в min_ball.txt ${ref.score}.`));
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

      if (minScore === 41) {
        const spoKey = normalizeSpoSubjectKey(chosenSubject);
        const spoRef = spoMinBallData.spoMap.get(spoKey);
        if (!spoRef) {
          rule41Issues.push(issue("vi-rule-41-spo", rowRef, `Минимальный балл 41 допустим только для предметов СПО. Предмет "${chosenSubject}" не найден в СПО-разделе min_ball.txt.`));
        }
      }

      if (normalizeText(baseSubject) === normalizeText(viRules.russianForeignSubject) && minScore !== Number(viRules.russianForeignMinScore)) {
        ruForeignIssues.push(issue("vi-ru-foreign", rowRef, `Для "${viRules.russianForeignSubject}" минимальный балл должен быть ${viRules.russianForeignMinScore}.`));
      }
    }

    if (specialMark) {
      const matched = findBenefitReplacement(specialMark, benefitPatterns);
      if (matched) {
        specialMarkIssues.push(
          issue(
            "benefit-replace",
            rowRef,
            `Некорректная формулировка в "Особая отметка". Найдено: "${String(specialMark).trim()}". Значение должно быть заменено на "${matched.to}".`
          )
        );
      }
    }
    if (onProgress) onProgress(index + 1, viWorkbookData.rows.length, `Проверка критериев ВИ... ${index + 1}/${viWorkbookData.rows.length}`);
    if ((index + 1) % 200 === 0) await yieldToUi();
  }

  return { maxIssues, minIssues, rule21Issues, rule41Issues, ruForeignIssues, specialMarkIssues };
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
    if ((rowIndex + 1) % 250 === 0) await yieldToUi();
  }
  return { issues };
}

async function checkQualificationByLevel(workbookData, fileLabel, onProgress) {
  const issues = [];
  const levelHeader = findHeader(workbookData.headers, "Уровень образования");
  const qualificationHeader = findHeader(workbookData.headers, "Квалификация");
  if (!levelHeader || !qualificationHeader) {
    return { issues: [issue("qualification-missing-columns", fileLabel, `[${fileLabel}] Не найдены столбцы "Уровень образования" и/или "Квалификация".`)] };
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
    if ((rowIndex + 1) % 250 === 0) await yieldToUi();
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
  const raw = String(subject || "").trim();
  const withoutPrefix = raw.replace(/^\s*спо\s*[:;,-]?\s*/i, "");
  return normalizeText(withoutPrefix);
}

function normalizeBenefitText(value) {
  return String(value || "")
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
    const kg = normalizeText(getVal(viRow, viKgHeader));
    if (!kg || viByKg.has(kg)) continue;
    viByKg.set(kg, viRow);
  }

  for (let index = 0; index < pnWorkbookData.rows.length; index += 1) {
    const pnRow = pnWorkbookData.rows[index];
    const rowRef = `Строка ${index + 2}`;
    const kgValue = getVal(pnRow, pnKgHeader);
    const key = normalizeText(kgValue);
    if (!key) return;
    const viRow = viByKg.get(key);
    if (!viRow) {
      issues.push(issue("missing-kg-in-vi", rowRef, `КГ "${kgValue}" есть в ПН, но отсутствует в ВИ.`));
      return;
    }
    for (const pair of resolvedPairs) {
      const pnValue = getVal(pnRow, pair.pnHeader);
      const viValue = getVal(viRow, pair.viHeader);
      if (normalizeText(pnValue) !== normalizeText(viValue)) {
        issues.push(issue("pn-vi-mismatch", rowRef, `Несовпадение ПН↔ВИ: "${pair.pn}" в ПН="${pnValue}", в ВИ="${viValue}" (КГ "${kgValue}").`));
      }
    }
    if (onProgress) onProgress(index + 1, pnWorkbookData.rows.length, `Сверка ПН и ВИ... ${index + 1}/${pnWorkbookData.rows.length}`);
    if ((index + 1) % 200 === 0) await yieldToUi();
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

function findHeader(headers, expectedName) {
  const target = normalizeText(expectedName);
  return headers.find((h) => normalizeText(h) === target) || null;
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

function containsToken(kgName, token) {
  if (!kgName || !token) return false;
  const text = String(kgName).toUpperCase();
  const t = String(token).toUpperCase();
  return text.includes(`|${t}|`) || text.includes(` ${t} `) || text.startsWith(`${t} `) || text.endsWith(` ${t}`) || text === t;
}

function isYes(value, rules) {
  return rules.yesValues.includes(normalizeText(value));
}

function isNo(value, rules) {
  return rules.noValues.includes(normalizeText(value));
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
    "missing-reference": "Некорректный минимальный балл (нет значения в min_ball.txt)",
    "vi-max": "Некорректный максимальный балл",
    "vi-min-id": "Минимальный балл для ИД",
    "vi-rule-21-type": "Правило 21 балла: тип испытания",
    "vi-rule-21-replace": "Правило 21 балла: заменяемый предмет",
    "vi-rule-41-spo": "Правило 41 балла (СПО)",
    "vi-ru-foreign": "Русский язык для иностранных граждан",
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
  const safe = Math.max(0, Math.min(100, Math.round(percent)));
  progressSectionNode.classList.remove("hidden");
  progressLabelNode.textContent = label;
  progressPercentNode.textContent = `${safe}%`;
  progressFillNode.style.width = `${safe}%`;
  progressTrackNode.setAttribute("aria-valuenow", String(safe));
}

function setUiBusy(isBusy) {
  Object.values(inputs).forEach((input) => { input.disabled = isBusy; });
  runCheckButton.disabled = isBusy;
  tabUploadButton.disabled = isBusy;
  tabCriteriaButton.disabled = isBusy;
}

async function yieldToUi() {
  await new Promise((resolve) => setTimeout(resolve, 0));
}

function setActiveTab(tab) {
  const uploadActive = tab === "upload";
  tabUploadButton.classList.toggle("is-active", uploadActive);
  tabCriteriaButton.classList.toggle("is-active", !uploadActive);
  uploadTabPanel.classList.toggle("is-active", uploadActive);
  criteriaTabPanel.classList.toggle("is-active", !uploadActive);
}
