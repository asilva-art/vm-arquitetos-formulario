var EXEC_SHEET_NAME = "⚙️ EXECUÇÃO";
var EAP_SHEET_NAME = "📅 EAP - LINHA DE BASE";
var CONTROL_SHEET_NAME = "📌 CONTROLE_EAP_ATUAL";
var EXEC_DATA_START_ROW = 5;
var EAP_DATA_START_ROW = 4;
var CONTROL_DATA_START_ROW = 2;
var DEFAULT_TZ = "America/Sao_Paulo";

var CONTROL_HEADERS = [
  "KEY",
  "PROJETO",
  "ID CT",
  "REF EAP",
  "FASE EAP",
  "TAREFA EAP",
  "PRAZO DU",
  "INICIO PLANEJADO",
  "FIM PLANEJADO",
  "STATUS ATUAL",
  "RESPONSAVEL",
  "INICIO REAL",
  "FIM REAL",
  "% REAL",
  "DIAS DESVIO",
  "BLOQUEIO",
  "DESC BLOQUEIO",
  "ATUALIZADO EM",
  "DATA ULT. REGISTRO",
  "ID ULT. ENVIO",
  "OBS ULT. REGISTRO",
  "ORIGEM"
];

var STATUS_CONTROL = {
  NAO_INICIADA: "NAO INICIADA",
  EM_ANDAMENTO: "EM ANDAMENTO",
  CONCLUIDA: "CONCLUIDA",
  BLOQUEADA: "BLOQUEADA"
};

var PROJECT_MAP = {
  "CT-117": { projectName: "PH Top Green", contractId: "CT-117", sectionName: "Secao CT-117" },
  "CT-119": { projectName: "KT Costa Laguna", contractId: "CT-119", sectionName: "Secao CT-119" },
  "CT-120": { projectName: "Alliance Sede NL", contractId: "CT-120", sectionName: "Secao CT-120" },
  "CT-121": { projectName: "VC Hermes Sabara", contractId: "CT-121", sectionName: "Secao CT-121" },
  "CT-122": { projectName: "IT Bernardo Guimaraes", contractId: "CT-122", sectionName: "Secao CT-122" },
  "ADMIN-INTERNO": { projectName: "Administrativo/Interno", contractId: "ADMIN-INTERNO", sectionName: "Secao Administrativo/Interno" }
};

var STATUS_LABEL_BY_CODE = {
  CONCLUIDO: "Concluido hoje",
  EM_ANDAMENTO: "Em andamento - continua amanha",
  PAUSADO: "Pausado - aguardando"
};

var BLOCK_LABEL_BY_CODE = {
  NAO: "Nao - tudo ok",
  SIM_APROVACAO: "Sim - aguardando aprovacao do cliente",
  SIM_ACESSO: "Sim - aguardando acesso ao local",
  SIM_FORNECEDOR: "Sim - aguardando fornecedor",
  SIM_OUTRO: "Sim - outro motivo"
};

var PHASE_BY_CATEGORY = {
  "Estudos Preliminares": "Estudos Preliminares",
  Anteprojeto: "Anteprojeto",
  "Pre-Executivo": "Projeto Pre-Executivo",
  Executivo: "Projeto Executivo",
  Detalhamentos: "Detalhamentos",
  Acompanhamentos: "Acompanhamentos",
  "Aprovacoes e Marcos": "Aprovacoes e Marcos",
  Comunicacao: "Comunicacao",
  Outros: "Outros"
};

var PHASE_CODE = {
  "Estudos Preliminares": "EP",
  Anteprojeto: "AP",
  "Projeto Pre-Executivo": "PE",
  "Projeto Executivo": "EX",
  Detalhamentos: "DT",
  Acompanhamentos: "AC",
  "Aprovacoes e Marcos": "AM",
  Comunicacao: "CM",
  Outros: "OT"
};

function doGet(e) {
  var action = "health";
  if (e && e.parameter && e.parameter.action) {
    action = e.parameter.action;
  }

  if (action === "health") {
    return json_({
      status: "ok",
      message: "Backend ativo.",
      controlSheet: CONTROL_SHEET_NAME
    });
  }

  if (action === "config") {
    var projects = [];
    var code;
    for (code in PROJECT_MAP) {
      if (Object.prototype.hasOwnProperty.call(PROJECT_MAP, code)) {
        projects.push({
          code: code,
          label: PROJECT_MAP[code].projectName
        });
      }
    }

    return json_({
      status: "ok",
      projects: projects,
      statuses: STATUS_CONTROL
    });
  }

  if (action === "resumo_controle") {
    return json_(resumoControleEapAtual());
  }

  return json_({
    status: "error",
    message: "Acao invalida."
  });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  var lockAcquired = false;

  try {
    lock.waitLock(30000);
    lockAcquired = true;

    var payload = parsePayload_(e);
    var input = normalizePayload_(payload);
    var executionSheet = getExecutionSheet_();

    enforceNoDuplicate_(executionSheet, input.professional, input.dateKey);

    var controlContext = prepareControlContext_();
    var formId = generateFormId_(executionSheet, input.dateObj);
    var records = buildSubmissionRecords_(input, controlContext.rows);
    var executionRows = buildExecutionRows_(records, input, formId);
    var executionStartRow = Math.max(executionSheet.getLastRow() + 1, EXEC_DATA_START_ROW);

    executionSheet.getRange(executionStartRow, 1, executionRows.length, 17).setValues(executionRows);

    syncControlFromRecords_(controlContext, records, {
      dateBr: input.dateBr,
      professional: input.professional,
      formId: formId,
      updatedAtBr: Utilities.formatDate(new Date(), getTz_(), "dd/MM/yyyy")
    });

    return json_({
      status: "ok",
      message: "Registro salvo com sucesso.",
      formId: formId,
      rowsCreated: executionRows.length,
      controlUpdated: true
    });
  } catch (error) {
    return json_({
      status: "error",
      message: error && error.message ? error.message : "Erro desconhecido no envio."
    });
  } finally {
    if (lockAcquired) {
      lock.releaseLock();
    }
  }
}

function parsePayload_(e) {
  var raw = "";
  if (e && e.postData && e.postData.contents) {
    raw = String(e.postData.contents).trim();
  }

  if (raw) {
    try {
      return JSON.parse(raw);
    } catch (err) {
      if (raw.indexOf("payload=") === 0) {
        return JSON.parse(decodeURIComponent(raw.substring(8)));
      }
      throw new Error("Formato invalido: nao foi possivel ler JSON.");
    }
  }

  if (e && e.parameter && e.parameter.payload) {
    return JSON.parse(e.parameter.payload);
  }

  throw new Error("Requisicao vazia.");
}

function normalizePayload_(payload) {
  if (!payload || typeof payload !== "object") {
    throw new Error("Payload invalido.");
  }

  var professional = clean_(payload.professional);
  if (!professional) {
    throw new Error("Campo 1 obrigatorio: nome do profissional.");
  }

  var dateObj = parseDate_(payload.date);
  if (!dateObj) {
    throw new Error("Campo 2 obrigatorio: data invalida.");
  }

  var projects = payload.projects;
  if (!projects || !projects.length) {
    throw new Error("Campo 3 obrigatorio: marque ao menos um projeto.");
  }

  var normalizedProjects = [];
  var i;
  for (i = 0; i < projects.length; i += 1) {
    normalizedProjects.push(normalizeProject_(projects[i], i + 1));
  }

  return {
    professional: professional,
    dateObj: dateObj,
    dateKey: Utilities.formatDate(dateObj, getTz_(), "yyyy-MM-dd"),
    dateBr: Utilities.formatDate(dateObj, getTz_(), "dd/MM/yyyy"),
    projects: normalizedProjects
  };
}

function normalizeProject_(item, sectionNumber) {
  item = item || {};

  var projectCode = clean_(item.projectCode);
  var projectLabel = clean_(item.projectLabel);
  var tasks = [];
  var sourceTasks = item.tasks || [];
  var i;

  for (i = 0; i < sourceTasks.length; i += 1) {
    var task = clean_(sourceTasks[i]);
    if (task) {
      tasks.push(task);
    }
  }

  if (!projectCode) {
    throw new Error("Projeto invalido na secao " + sectionNumber + ".");
  }
  if (!tasks.length) {
    throw new Error("Campo 4 obrigatorio na secao " + sectionNumber + ".");
  }

  var statusCode = clean_(item.statusCode);
  var statusLabel = clean_(item.statusLabel) || STATUS_LABEL_BY_CODE[statusCode] || statusCode;
  if (!statusLabel) {
    throw new Error("Campo 5 obrigatorio na secao " + sectionNumber + ".");
  }

  var blockCode = clean_(item.blockCode);
  var blockLabel = clean_(item.blockLabel) || BLOCK_LABEL_BY_CODE[blockCode] || blockCode;
  if (!blockCode && !blockLabel) {
    throw new Error("Campo 6 obrigatorio na secao " + sectionNumber + ".");
  }

  var isBlocked = blockCode ? blockCode !== "NAO" : !/^nao/i.test(blockLabel);
  var blockDescription = clean_(item.blockDescription);
  if (isBlocked && !blockDescription) {
    throw new Error("Campo 7 obrigatorio quando ha bloqueio na secao " + sectionNumber + ".");
  }

  return {
    projectCode: projectCode,
    projectLabel: projectLabel || projectCode,
    tasks: tasks,
    statusText: statusLabel,
    isBlocked: isBlocked,
    blockReason: isBlocked ? (blockDescription || blockLabel) : "",
    observation: clean_(item.observation)
  };
}

function getExecutionSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXEC_SHEET_NAME);
  if (!sheet) {
    throw new Error("Aba " + EXEC_SHEET_NAME + " nao encontrada.");
  }
  return sheet;
}

function getEapSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EAP_SHEET_NAME);
  if (!sheet) {
    throw new Error("Aba " + EAP_SHEET_NAME + " nao encontrada.");
  }
  return sheet;
}

function getOrCreateControlSheet_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(CONTROL_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(CONTROL_SHEET_NAME);
  }

  ensureControlHeader_(sheet);
  return sheet;
}

function ensureControlHeader_(sheet) {
  var current = sheet.getRange(1, 1, 1, CONTROL_HEADERS.length).getValues()[0];
  var needsUpdate = false;
  var i;

  for (i = 0; i < CONTROL_HEADERS.length; i += 1) {
    if (clean_(current[i]) !== CONTROL_HEADERS[i]) {
      needsUpdate = true;
      break;
    }
  }

  if (needsUpdate) {
    sheet.getRange(1, 1, 1, CONTROL_HEADERS.length).setValues([CONTROL_HEADERS]);
    sheet.getRange(1, 1, 1, CONTROL_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
}

function readControlRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < CONTROL_DATA_START_ROW) {
    return [];
  }

  return sheet
    .getRange(CONTROL_DATA_START_ROW, 1, lastRow - CONTROL_DATA_START_ROW + 1, CONTROL_HEADERS.length)
    .getValues();
}

function writeControlRows_(sheet, rows) {
  var lastRow = sheet.getLastRow();
  if (lastRow >= CONTROL_DATA_START_ROW) {
    sheet.getRange(CONTROL_DATA_START_ROW, 1, lastRow - CONTROL_DATA_START_ROW + 1, CONTROL_HEADERS.length).clearContent();
  }

  if (rows && rows.length) {
    sheet.getRange(CONTROL_DATA_START_ROW, 1, rows.length, CONTROL_HEADERS.length).setValues(rows);
  }
}

function prepareControlContext_() {
  var controlSheet = getOrCreateControlSheet_();
  var rows = readControlRows_(controlSheet);

  if (!rows.length) {
    rows = seedControlRowsFromEap_();
    writeControlRows_(controlSheet, rows);
  }

  return {
    sheet: controlSheet,
    rows: rows
  };
}

function seedControlRowsFromEap_() {
  var eapSheet = getEapSheet_();
  var lastRow = eapSheet.getLastRow();
  if (lastRow < EAP_DATA_START_ROW) {
    return [];
  }

  var values = eapSheet.getRange(EAP_DATA_START_ROW, 1, lastRow - EAP_DATA_START_ROW + 1, 12).getValues();
  var counters = {};
  var controlRows = [];
  var i;

  for (i = 0; i < values.length; i += 1) {
    var row = values[i];
    var projectName = clean_(row[0]);
    var contractId = clean_(row[1]);
    var phase = clean_(row[2]);
    var task = clean_(row[4]);
    var prazoDu = clean_(row[5]);

    if (!isValidEapTask_(projectName, contractId, phase, task)) {
      continue;
    }

    var phaseCode = PHASE_CODE[phase] || "OT";
    var counterKey = contractId + "|" + phaseCode;
    counters[counterKey] = (counters[counterKey] || 0) + 1;

    var ref = compactContractId_(contractId) + "-" + phaseCode + "-" + pad_(counters[counterKey], 2);
    controlRows.push(buildControlRow_({
      projectName: projectName,
      contractId: contractId,
      refEap: ref,
      phase: phase,
      task: task,
      prazoDu: prazoDu,
      status: STATUS_CONTROL.NAO_INICIADA,
      origem: "EAP_BASE"
    }));
  }

  return controlRows;
}

function isValidEapTask_(projectName, contractId, phase, task) {
  if (!projectName || !contractId || !phase || !task) {
    return false;
  }

  if (projectName.indexOf("▶") === 0) {
    return false;
  }

  if (projectName.indexOf("VM +") === 0) {
    return false;
  }

  if (contractId.indexOf("CT-") !== 0 && contractId.indexOf("ADMIN") !== 0) {
    return false;
  }

  return true;
}

function buildControlRow_(data) {
  return [
    makeControlKey_(data.contractId, data.refEap),
    clean_(data.projectName),
    clean_(data.contractId),
    clean_(data.refEap),
    clean_(data.phase),
    clean_(data.task),
    clean_(data.prazoDu),
    clean_(data.startPlanned),
    clean_(data.endPlanned),
    clean_(data.status || STATUS_CONTROL.NAO_INICIADA),
    clean_(data.responsible),
    clean_(data.startReal),
    clean_(data.endReal),
    toNumberOrBlank_(data.percentReal, 0),
    clean_(data.delayDays),
    clean_(data.blocked || "Nao"),
    clean_(data.blockReason),
    clean_(data.updatedAt),
    clean_(data.lastRecordDate),
    clean_(data.lastFormId),
    clean_(data.lastObservation),
    clean_(data.origem || "FORM_DAILY")
  ];
}

function buildSubmissionRecords_(input, controlRows) {
  var records = [];
  var i;

  for (i = 0; i < input.projects.length; i += 1) {
    var item = input.projects[i];
    var projectInfo = resolveProjectInfo_(item.projectCode, item.projectLabel);
    var phase = inferPhase_(item.tasks);
    var taskText = item.tasks.join(" | ");
    var refEap = findBestRefEap_(controlRows, projectInfo.contractId, phase, taskText);

    if (!refEap) {
      refEap = buildRefEap_(projectInfo.contractId, phase);
    }

    records.push({
      projectName: projectInfo.projectName,
      contractId: projectInfo.contractId,
      sectionName: projectInfo.sectionName,
      phase: phase,
      taskText: taskText,
      refEap: refEap,
      statusText: item.statusText,
      statusClass: classifyControlStatus_(item.statusText, item.isBlocked),
      isBlocked: item.isBlocked,
      blockReason: item.blockReason,
      observation: item.observation,
      responsible: input.professional
    });
  }

  return records;
}

function buildExecutionRows_(records, input, formId) {
  var todayBr = Utilities.formatDate(new Date(), getTz_(), "dd/MM/yyyy");
  var rows = [];
  var i;

  for (i = 0; i < records.length; i += 1) {
    var rec = records[i];
    rows.push([
      rec.projectName,
      rec.contractId,
      rec.phase,
      rec.taskText,
      rec.sectionName,
      rec.refEap,
      rec.responsible,
      "",
      rec.statusText,
      "",
      rec.isBlocked ? "Sim" : "Nao",
      rec.isBlocked ? rec.blockReason : "",
      "",
      input.dateBr,
      rec.observation,
      todayBr,
      formId
    ]);
  }

  return rows;
}

function syncControlFromRecords_(controlContext, records, meta) {
  var rows = controlContext.rows || [];
  var i;

  for (i = 0; i < records.length; i += 1) {
    var record = records[i];
    var rowIndex = findControlRowIndexForRecord_(rows, record);

    if (rowIndex < 0) {
      rows.push(createAdHocControlRow_(record, meta));
      rowIndex = rows.length - 1;
    }

    updateControlRowFromRecord_(rows[rowIndex], record, meta);
  }

  writeControlRows_(controlContext.sheet, rows);
}

function findBestRefEap_(rows, contractId, phase, taskText) {
  var bestRef = "";
  var bestScore = -1;
  var fallbackRef = "";
  var normalizedTask = normalizeText_(taskText);
  var i;

  for (i = 0; i < rows.length; i += 1) {
    var row = rows[i];
    if (clean_(row[2]) !== clean_(contractId)) {
      continue;
    }
    if (clean_(row[4]) !== clean_(phase)) {
      continue;
    }

    var status = clean_(row[9]);
    if (!fallbackRef && status !== STATUS_CONTROL.CONCLUIDA) {
      fallbackRef = clean_(row[3]);
    }

    var rowTask = normalizeText_(row[5]);
    var score = 0;

    if (status !== STATUS_CONTROL.CONCLUIDA) {
      score += 3;
    }
    if (status === STATUS_CONTROL.EM_ANDAMENTO) {
      score += 2;
    }
    if (normalizedTask && rowTask) {
      if (normalizedTask.indexOf(rowTask) >= 0 || rowTask.indexOf(normalizedTask) >= 0) {
        score += 6;
      }
      if (hasOverlapWords_(normalizedTask, rowTask)) {
        score += 4;
      }
    }

    if (score > bestScore) {
      bestScore = score;
      bestRef = clean_(row[3]);
    }
  }

  if (bestRef && bestScore >= 3) {
    return bestRef;
  }
  return fallbackRef || bestRef;
}

function findControlRowIndexForRecord_(rows, record) {
  var i;

  for (i = 0; i < rows.length; i += 1) {
    if (clean_(rows[i][3]) === clean_(record.refEap) && clean_(rows[i][2]) === clean_(record.contractId)) {
      return i;
    }
  }

  var bestIndex = -1;
  var bestScore = -1;
  var normalizedTask = normalizeText_(record.taskText);

  for (i = 0; i < rows.length; i += 1) {
    var row = rows[i];
    if (clean_(row[2]) !== clean_(record.contractId)) {
      continue;
    }
    if (clean_(row[4]) !== clean_(record.phase)) {
      continue;
    }

    var score = 0;
    var rowStatus = clean_(row[9]);
    if (rowStatus === STATUS_CONTROL.NAO_INICIADA) {
      score += 4;
    } else if (rowStatus === STATUS_CONTROL.EM_ANDAMENTO || rowStatus === STATUS_CONTROL.BLOQUEADA) {
      score += 3;
    } else {
      score += 1;
    }

    var rowTask = normalizeText_(row[5]);
    if (normalizedTask && rowTask) {
      if (normalizedTask.indexOf(rowTask) >= 0 || rowTask.indexOf(normalizedTask) >= 0) {
        score += 5;
      }
      if (hasOverlapWords_(normalizedTask, rowTask)) {
        score += 3;
      }
    }

    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }

  return bestIndex;
}

function createAdHocControlRow_(record, meta) {
  return buildControlRow_({
    projectName: record.projectName,
    contractId: record.contractId,
    refEap: record.refEap,
    phase: record.phase,
    task: record.taskText,
    status: record.statusClass,
    responsible: meta.professional,
    startReal: meta.dateBr,
    endReal: record.statusClass === STATUS_CONTROL.CONCLUIDA ? meta.dateBr : "",
    percentReal: record.statusClass === STATUS_CONTROL.CONCLUIDA ? 100 : 50,
    blocked: record.isBlocked ? "Sim" : "Nao",
    blockReason: record.blockReason,
    updatedAt: meta.updatedAtBr,
    lastRecordDate: meta.dateBr,
    lastFormId: meta.formId,
    lastObservation: record.observation,
    origem: "FORM_ADHOC"
  });
}

function updateControlRowFromRecord_(row, record, meta) {
  row[0] = makeControlKey_(record.contractId, record.refEap);
  row[1] = clean_(row[1]) || record.projectName;
  row[2] = clean_(row[2]) || record.contractId;
  row[3] = clean_(row[3]) || record.refEap;
  row[4] = clean_(row[4]) || record.phase;
  row[5] = clean_(row[5]) || record.taskText;
  row[9] = record.statusClass;
  row[10] = meta.professional;

  if (!clean_(row[11])) {
    row[11] = meta.dateBr;
  }

  if (record.statusClass === STATUS_CONTROL.CONCLUIDA) {
    row[12] = meta.dateBr;
    row[13] = 100;
  } else if (record.statusClass === STATUS_CONTROL.EM_ANDAMENTO) {
    row[13] = Math.max(toNumber_(row[13], 0), 50);
    row[12] = clean_(row[12]);
  } else if (record.statusClass === STATUS_CONTROL.BLOQUEADA) {
    row[13] = Math.max(toNumber_(row[13], 0), 25);
  }

  row[14] = calculateDelayDays_(row[8], record.statusClass === STATUS_CONTROL.CONCLUIDA ? row[12] : meta.dateBr);
  row[15] = record.isBlocked ? "Sim" : "Nao";
  row[16] = record.isBlocked ? record.blockReason : "";
  row[17] = meta.updatedAtBr;
  row[18] = meta.dateBr;
  row[19] = meta.formId;
  row[20] = record.observation;
  row[21] = clean_(row[21]) || "FORM_DAILY";
}

function classifyControlStatus_(statusText, isBlocked) {
  if (isBlocked) {
    return STATUS_CONTROL.BLOQUEADA;
  }

  var text = normalizeText_(statusText);
  if (text.indexOf("conclu") >= 0) {
    return STATUS_CONTROL.CONCLUIDA;
  }
  if (text.indexOf("paus") >= 0 || text.indexOf("aguard") >= 0) {
    return STATUS_CONTROL.BLOQUEADA;
  }
  if (text.indexOf("andamento") >= 0 || text.indexOf("continua") >= 0) {
    return STATUS_CONTROL.EM_ANDAMENTO;
  }

  return STATUS_CONTROL.EM_ANDAMENTO;
}

function calculateDelayDays_(plannedEndValue, referenceValue) {
  var planned = parseDate_(plannedEndValue);
  var reference = parseDate_(referenceValue);

  if (!planned || !reference) {
    return "";
  }

  var millis = reference.getTime() - planned.getTime();
  var diffDays = Math.floor(millis / (24 * 60 * 60 * 1000));
  return diffDays > 0 ? diffDays : 0;
}

function resolveProjectInfo_(projectCode, projectLabel) {
  var mapValue = PROJECT_MAP[projectCode];
  if (mapValue) {
    return mapValue;
  }

  return {
    projectName: projectLabel || projectCode,
    contractId: projectCode,
    sectionName: "Secao " + projectCode
  };
}

function inferPhase_(tasks) {
  var i;
  for (i = 0; i < tasks.length; i += 1) {
    if (PHASE_BY_CATEGORY[tasks[i]]) {
      return PHASE_BY_CATEGORY[tasks[i]];
    }
  }
  return "Outros";
}

function buildRefEap_(contractId, phase) {
  var contractKey = compactContractId_(contractId);
  var phaseKey = PHASE_CODE[phase] || "OT";

  if (!contractKey) {
    return "";
  }

  return contractKey + "-" + phaseKey + "-PEND";
}

function compactContractId_(contractId) {
  return String(contractId || "")
    .replace(/[^A-Za-z0-9]/g, "")
    .toUpperCase();
}

function makeControlKey_(contractId, refEap) {
  return compactContractId_(contractId) + "|" + clean_(refEap);
}

function enforceNoDuplicate_(sheet, professional, dateKey) {
  var lastRow = sheet.getLastRow();
  if (lastRow < EXEC_DATA_START_ROW) {
    return;
  }

  var rows = sheet.getRange(EXEC_DATA_START_ROW, 7, lastRow - EXEC_DATA_START_ROW + 1, 8).getValues();
  var targetProfessional = normalizeName_(professional);
  var i;

  for (i = 0; i < rows.length; i += 1) {
    var row = rows[i];
    var existingProfessional = normalizeName_(row[0]);
    var existingDateKey = toDateKey_(row[7]);

    if (!existingProfessional || !existingDateKey) {
      continue;
    }
    if (existingProfessional === targetProfessional && existingDateKey === dateKey) {
      throw new Error("Ja existe envio desse profissional nesta data. Regra: maximo 1 envio por dia.");
    }
  }
}

function generateFormId_(sheet, dateObj) {
  var dateKey = Utilities.formatDate(dateObj, getTz_(), "yyyyMMdd");
  var prefix = "FORM-" + dateKey + "-";
  var idRegex = new RegExp("^FORM-" + dateKey + "-(\\d{3})$");
  var lastRow = sheet.getLastRow();

  if (lastRow < EXEC_DATA_START_ROW) {
    return prefix + "001";
  }

  var values = sheet.getRange(EXEC_DATA_START_ROW, 17, lastRow - EXEC_DATA_START_ROW + 1, 1).getValues();
  var maxCounter = 0;
  var i;

  for (i = 0; i < values.length; i += 1) {
    var id = clean_(values[i][0]);
    if (!id) {
      continue;
    }
    var match = id.match(idRegex);
    if (!match) {
      continue;
    }
    var num = Number(match[1]);
    if (num > maxCounter) {
      maxCounter = num;
    }
  }

  return prefix + ("000" + (maxCounter + 1)).slice(-3);
}

function parseDate_(value) {
  if (!value) {
    return null;
  }

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value)) {
    return value;
  }

  var text = String(value).trim();
  if (!text) {
    return null;
  }

  var isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    return new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]));
  }

  var brMatch = text.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (brMatch) {
    return new Date(Number(brMatch[3]), Number(brMatch[2]) - 1, Number(brMatch[1]));
  }

  var parsed = new Date(text);
  return isNaN(parsed) ? null : parsed;
}

function toDateKey_(value) {
  var parsed = parseDate_(value);
  if (!parsed) {
    return "";
  }
  return Utilities.formatDate(parsed, getTz_(), "yyyy-MM-dd");
}

function normalizeName_(value) {
  return clean_(value).replace(/\s+/g, " ").trim().toUpperCase();
}

function normalizeText_(value) {
  return clean_(value)
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function hasOverlapWords_(a, b) {
  if (!a || !b) {
    return false;
  }

  var wordsA = a.split(" ");
  var wordsB = b.split(" ");
  var bag = {};
  var i;

  for (i = 0; i < wordsA.length; i += 1) {
    if (wordsA[i].length >= 4) {
      bag[wordsA[i]] = true;
    }
  }

  for (i = 0; i < wordsB.length; i += 1) {
    if (wordsB[i].length >= 4 && bag[wordsB[i]]) {
      return true;
    }
  }

  return false;
}

function toNumber_(value, fallback) {
  var num = Number(value);
  if (isNaN(num)) {
    return fallback;
  }
  return num;
}

function toNumberOrBlank_(value, fallback) {
  var num = Number(value);
  if (isNaN(num)) {
    return fallback === undefined ? "" : fallback;
  }
  return num;
}

function pad_(value, size) {
  var text = String(value);
  while (text.length < size) {
    text = "0" + text;
  }
  return text;
}

function clean_(value) {
  if (value === null || value === undefined) {
    return "";
  }
  return String(value).trim();
}

function getTz_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || DEFAULT_TZ;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// -------------------------
// Funcoes manuais (rodar no editor Apps Script)
// -------------------------

function setupControleEapAtual() {
  var context = prepareControlContext_();
  return {
    status: "ok",
    sheet: CONTROL_SHEET_NAME,
    rows: context.rows.length,
    message: "Controle pronto para uso."
  };
}

function sincronizarControleComExecucaoHistorica() {
  var context = prepareControlContext_();
  var executionSheet = getExecutionSheet_();
  var lastRow = executionSheet.getLastRow();

  if (lastRow < EXEC_DATA_START_ROW) {
    return {
      status: "ok",
      message: "Sem dados historicos na EXECUCAO.",
      rowsProcessed: 0
    };
  }

  var values = executionSheet.getRange(EXEC_DATA_START_ROW, 1, lastRow - EXEC_DATA_START_ROW + 1, 17).getValues();
  var rows = context.rows;
  var updatedAtBr = Utilities.formatDate(new Date(), getTz_(), "dd/MM/yyyy");
  var processed = 0;
  var i;

  for (i = 0; i < values.length; i += 1) {
    var row = values[i];
    var contractId = clean_(row[1]);
    var projectName = clean_(row[0]);
    var phase = clean_(row[2]);
    var taskText = clean_(row[3]);

    if (!contractId || !projectName || !phase || !taskText) {
      continue;
    }

    var record = {
      projectName: projectName,
      contractId: contractId,
      phase: phase,
      taskText: taskText,
      refEap: clean_(row[5]) || buildRefEap_(contractId, phase),
      statusText: clean_(row[8]),
      statusClass: classifyControlStatus_(clean_(row[8]), /^sim/i.test(clean_(row[10]))),
      isBlocked: /^sim/i.test(clean_(row[10])),
      blockReason: clean_(row[11]),
      observation: clean_(row[14]),
      responsible: clean_(row[6]),
      sectionName: ""
    };

    var meta = {
      dateBr: clean_(row[13]) || updatedAtBr,
      professional: clean_(row[6]),
      formId: clean_(row[16]),
      updatedAtBr: updatedAtBr
    };

    var rowIndex = findControlRowIndexForRecord_(rows, record);
    if (rowIndex < 0) {
      rows.push(createAdHocControlRow_(record, meta));
      rowIndex = rows.length - 1;
    }

    updateControlRowFromRecord_(rows[rowIndex], record, meta);
    processed += 1;
  }

  writeControlRows_(context.sheet, rows);

  return {
    status: "ok",
    message: "Sincronizacao concluida.",
    rowsProcessed: processed,
    controlRows: rows.length
  };
}

function resumoControleEapAtual() {
  var sheet = getOrCreateControlSheet_();
  var rows = readControlRows_(sheet);
  var totals = {
    total: rows.length,
    naoIniciada: 0,
    emAndamento: 0,
    concluida: 0,
    bloqueada: 0
  };

  var i;
  for (i = 0; i < rows.length; i += 1) {
    var status = clean_(rows[i][9]);
    if (status === STATUS_CONTROL.NAO_INICIADA) {
      totals.naoIniciada += 1;
    } else if (status === STATUS_CONTROL.EM_ANDAMENTO) {
      totals.emAndamento += 1;
    } else if (status === STATUS_CONTROL.CONCLUIDA) {
      totals.concluida += 1;
    } else if (status === STATUS_CONTROL.BLOQUEADA) {
      totals.bloqueada += 1;
    }
  }

  return {
    status: "ok",
    sheet: CONTROL_SHEET_NAME,
    totals: totals
  };
}
