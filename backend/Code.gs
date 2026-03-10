var EXEC_SHEET_NAME = "⚙️ EXECUÇÃO";
var EAP_SHEET_NAME = "📅 EAP - LINHA DE BASE";
var CONTROL_SHEET_NAME = "📌 CONTROLE_EAP_ATUAL";
var EVENTS_SHEET_NAME = "📍 EVENTOS_COORDENACAO";
var EXEC_DATA_START_ROW = 5;
var EXEC_COLUMNS = 17;
var EAP_DATA_START_ROW = 4;
var CONTROL_DATA_START_ROW = 2;
var EVENTS_DATA_START_ROW = 2;
var DEFAULT_TZ = "America/Sao_Paulo";
var RECESS_START_MONTH_DAY = "12-25";
var RECESS_END_MONTH_DAY = "01-01";

var DEFAULT_CONTRACT_DEADLINE_SUMMARY = "";

var CONTRACT_DEADLINE_LEGEND = [
  { sigla: "F1/F2/F3", descricao: "Fases I, II e III de Estudos Preliminares + Anteprojeto" },
  { sigla: "RAF1/2/3", descricao: "Revisao de Anteprojeto nas Fases I, II e III" },
  { sigla: "PRE", descricao: "Projeto Pre-Executivo" },
  { sigla: "RFA", descricao: "Revisao Final do Anteprojeto" },
  { sigla: "EXE", descricao: "Projeto Executivo" },
  { sigla: "DET", descricao: "Detalhamento" },
  { sigla: "RDE", descricao: "Revisao de Detalhamento e Executivo" },
  { sigla: "DU", descricao: "Dias uteis" }
];

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
  "ORIGEM",
  "PREDECESSORA REF",
  "RELACAO DEP",
  "LAG DU"
];

var EAP_HEADERS = [
  "PROJETO",
  "ID CT",
  "FASE",
  "SUBFASE",
  "TAREFA",
  "PRAZO DU",
  "DEPENDENCIA",
  "% CONTRATUAL",
  "TIPO",
  "MARCO",
  "OBS",
  "ORIGEM"
];

var EVENTS_HEADERS = [
  "ID EVENTO",
  "ID CT",
  "DATA EVENTO",
  "TIPO",
  "TITULO",
  "RESPONSAVEL",
  "STATUS",
  "OBS",
  "ATUALIZADO EM",
  "ATUALIZADO POR"
];

var STATUS_CONTROL = {
  NAO_INICIADA: "NAO INICIADA",
  EM_ANDAMENTO: "EM ANDAMENTO",
  CONCLUIDA: "CONCLUIDA",
  BLOQUEADA: "BLOQUEADA"
};

var PROJECT_MAP = {
  "CT-117": {
    projectName: "CT 20251212-058-117 - Paulo Henrique",
    contractId: "CT-117",
    sectionName: "Secao CT-117",
    signatureDate: "2025-12-12"
  },
  "CT-119": {
    projectName: "CT 20251216-824-119 - Karina e Toshi",
    contractId: "CT-119",
    sectionName: "Secao CT-119",
    signatureDate: "2025-12-16"
  },
  "CT-120": {
    projectName: "CT 20251222-367-120 - Alliance Health",
    contractId: "CT-120",
    sectionName: "Secao CT-120",
    signatureDate: "2025-12-17"
  },
  "CT-121": {
    projectName: "CT 20260120-570-121 - Valdete e Carlos",
    contractId: "CT-121",
    sectionName: "Secao CT-121",
    signatureDate: "2025-12-18"
  },
  "CT-122": {
    projectName: "CT 20260212-089-122 - Iara e Tales",
    contractId: "CT-122",
    sectionName: "Secao CT-122",
    signatureDate: "2026-02-12"
  },
  "CT-123": {
    projectName: "CT 20250108-525-097 - Katia e Luis",
    contractId: "CT-123",
    sectionName: "Secao CT-123",
    signatureDate: "2025-01-08"
  },
  "CT-124": {
    projectName: "CT 20250127-254-098 - Monica e Oswaldo",
    contractId: "CT-124",
    sectionName: "Secao CT-124",
    signatureDate: "2025-01-27"
  },
  "CT-125": {
    projectName: "CT 20250212-851-099 - Carina e Eider",
    contractId: "CT-125",
    sectionName: "Secao CT-125",
    signatureDate: "2025-02-12"
  },
  "CT-126": {
    projectName: "CT 20250214-101-100 - Rachel e Matheus",
    contractId: "CT-126",
    sectionName: "Secao CT-126",
    signatureDate: "2025-02-14"
  },
  "CT-127": {
    projectName: "CT 20250312-400-101 - Natalia e Igor",
    contractId: "CT-127",
    sectionName: "Secao CT-127",
    signatureDate: "2025-03-12"
  },
  "CT-128": {
    projectName: "CT 20250404-857-102 - Priscila e Tiago",
    contractId: "CT-128",
    sectionName: "Secao CT-128",
    signatureDate: "2025-04-04"
  },
  "CT-129": {
    projectName: "CT 20250415-042-103 - Mariana e Igor",
    contractId: "CT-129",
    sectionName: "Secao CT-129",
    signatureDate: "2025-04-15"
  },
  "CT-130": {
    projectName: "CT 20250515-104 - Paula e Renato",
    contractId: "CT-130",
    sectionName: "Secao CT-130",
    signatureDate: "2025-05-15"
  },
  "CT-131": {
    projectName: "CT 20250515-083-105 - Priscila e Ricardo",
    contractId: "CT-131",
    sectionName: "Secao CT-131",
    signatureDate: "2025-05-15"
  },
  "CT-132": {
    projectName: "CT 20250523-065-106 - Fernanda e Henrique",
    contractId: "CT-132",
    sectionName: "Secao CT-132",
    signatureDate: "2025-05-23"
  },
  "CT-133": {
    projectName: "CT 20250902-054-112 - Daniela e Rafael",
    contractId: "CT-133",
    sectionName: "Secao CT-133",
    signatureDate: "2025-09-02"
  },
  "CT-134": {
    projectName: "CT 20250626-990-109 - Patricia e Alessandro",
    contractId: "CT-134",
    sectionName: "Secao CT-134",
    signatureDate: "2025-06-26"
  },
  "CT-135": {
    projectName: "CT 20250604-000-108 - Vanessa e Breno",
    contractId: "CT-135",
    sectionName: "Secao CT-135",
    signatureDate: "2025-06-04"
  },
  "CT-136": {
    projectName: "CT 20251126-086-115 - Maira e Eduardo",
    contractId: "CT-136",
    sectionName: "Secao CT-136",
    signatureDate: "2025-11-26"
  },
  "ADMIN-INTERNO": { projectName: "Administrativo/Interno", contractId: "ADMIN-INTERNO", sectionName: "Secao Administrativo/Interno" }
};
var CONTRACTS_SHEET_NAME = "CONTRATOS";
var PROJECT_MAP_CACHE = null;

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

var EAP_TEMPLATE_ITEMS = [
  { wbs: "1.1.1", phase: "Estudos Preliminares", subphase: "Coleta de Dados", type: "TAREFA", task: "Elaboracao do Briefing Documentado" },
  { wbs: "1.1.2", phase: "Estudos Preliminares", subphase: "Coleta de Dados", type: "TAREFA", task: "Geracao da Analise do Terreno" },
  { wbs: "1.1.3", phase: "Estudos Preliminares", subphase: "Coleta de Dados", type: "TAREFA", task: "Levantamento da Analise Legal" },
  { wbs: "2.1.1", phase: "Anteprojeto - Etapa I (Implantacao e Layout)", subphase: "Concepcao Inicial", type: "TAREFA", task: "Estudo de Implantacao" },
  { wbs: "2.1.2", phase: "Anteprojeto - Etapa I (Implantacao e Layout)", subphase: "Concepcao Inicial", type: "TAREFA", task: "Planta de Layout" },
  { wbs: "2.2.1", phase: "Anteprojeto - Etapa I (Implantacao e Layout)", subphase: "Revisao e Aprovacao", type: "TAREFA", task: "Execucao de Revisoes (ate 2 rodadas)" },
  { wbs: "2.2.2", phase: "Anteprojeto - Etapa I (Implantacao e Layout)", subphase: "Revisao e Aprovacao", type: "MARCO", task: "Aprovacao do Layout" },
  { wbs: "3.1.1", phase: "Anteprojeto - Etapa II (Volumetria e Materiais Externos)", subphase: "Desenvolvimento Volumetrico", type: "TAREFA", task: "Modelagem de Imagens 3D Externas" },
  { wbs: "3.1.2", phase: "Anteprojeto - Etapa II (Volumetria e Materiais Externos)", subphase: "Desenvolvimento Volumetrico", type: "TAREFA", task: "Pre-selecao de Materiais Externos" },
  { wbs: "3.1.3", phase: "Anteprojeto - Etapa II (Volumetria e Materiais Externos)", subphase: "Desenvolvimento Volumetrico", type: "TAREFA", task: "Estudo de Alternativa de Volumetria (se solicitada)" },
  { wbs: "3.2.1", phase: "Anteprojeto - Etapa II (Volumetria e Materiais Externos)", subphase: "Revisao e Aprovacao", type: "TAREFA", task: "Execucao de Revisoes (ate 2 rodadas)" },
  { wbs: "3.2.2", phase: "Anteprojeto - Etapa II (Volumetria e Materiais Externos)", subphase: "Revisao e Aprovacao", type: "MARCO", task: "Aprovacao da Volumetria e Materiais Externos" },
  { wbs: "4.1.1", phase: "Projeto Legal", subphase: "Preparacao para Aprovacao", type: "TAREFA", task: "Elaboracao do Projeto para Aprovacao na Prefeitura" },
  { wbs: "4.2.1", phase: "Projeto Legal", subphase: "Protocolo e Acompanhamento", type: "TAREFA", task: "Acompanhamento do Protocolo na Prefeitura" },
  { wbs: "4.2.2", phase: "Projeto Legal", subphase: "Protocolo e Acompanhamento", type: "MARCO", task: "Projeto Legal Aprovado" },
  { wbs: "5.1.1", phase: "Anteprojeto - Etapa III (Interiores)", subphase: "Concepcao de Interiores", type: "TAREFA", task: "Modelagem de Imagens 3D Internas" },
  { wbs: "5.1.2", phase: "Anteprojeto - Etapa III (Interiores)", subphase: "Concepcao de Interiores", type: "TAREFA", task: "Pre-selecao de Materiais Internos" },
  { wbs: "5.1.3", phase: "Anteprojeto - Etapa III (Interiores)", subphase: "Concepcao de Interiores", type: "TAREFA", task: "Estudo de Alternativa por Ambiente (se solicitada)" },
  { wbs: "5.2.1", phase: "Anteprojeto - Etapa III (Interiores)", subphase: "Revisao e Aprovacao", type: "TAREFA", task: "Execucao de Revisoes (ate 2 rodadas)" },
  { wbs: "5.2.2", phase: "Anteprojeto - Etapa III (Interiores)", subphase: "Revisao e Aprovacao", type: "MARCO", task: "Aprovacao dos Interiores" },
  { wbs: "6.1.1", phase: "Projeto Pre-Executivo", subphase: "Detalhamento Basico", type: "TAREFA", task: "Planta Layout Cotada" },
  { wbs: "6.1.2", phase: "Projeto Pre-Executivo", subphase: "Detalhamento Basico", type: "TAREFA", task: "Detalhamento de Pedras" },
  { wbs: "6.1.3", phase: "Projeto Pre-Executivo", subphase: "Detalhamento Basico", type: "TAREFA", task: "Quantitativo de Revestimentos" },
  { wbs: "6.1.4", phase: "Projeto Pre-Executivo", subphase: "Detalhamento Basico", type: "TAREFA", task: "Lista de Loucas e Metais" },
  { wbs: "7.1.1", phase: "Definicao de Materiais e Revisao 3D", subphase: "Selecao Final de Materiais", type: "TAREFA", task: "Acompanhamento em Lojas (ate 7 visitas)" },
  { wbs: "7.2.1", phase: "Definicao de Materiais e Revisao 3D", subphase: "Visualizacao Final", type: "TAREFA", task: "Execucao da Revisao Final 3D" },
  { wbs: "7.2.2", phase: "Definicao de Materiais e Revisao 3D", subphase: "Visualizacao Final", type: "MARCO", task: "Aprovacao Final 3D" },
  { wbs: "8.1.1", phase: "Projeto Executivo", subphase: "Desenhos Arquitetonicos", type: "TAREFA", task: "Planta de Situacao" },
  { wbs: "8.1.2", phase: "Projeto Executivo", subphase: "Desenhos Arquitetonicos", type: "TAREFA", task: "Planta de Implantacao" },
  { wbs: "8.1.3", phase: "Projeto Executivo", subphase: "Desenhos Arquitetonicos", type: "TAREFA", task: "Planta de Cobertura" },
  { wbs: "8.1.4", phase: "Projeto Executivo", subphase: "Desenhos Arquitetonicos", type: "TAREFA", task: "Cortes" },
  { wbs: "8.1.5", phase: "Projeto Executivo", subphase: "Desenhos Arquitetonicos", type: "TAREFA", task: "Fachadas" },
  { wbs: "8.2.1", phase: "Projeto Executivo", subphase: "Desenhos Complementares", type: "TAREFA", task: "Pontos Eletricos" },
  { wbs: "8.2.2", phase: "Projeto Executivo", subphase: "Desenhos Complementares", type: "TAREFA", task: "Pontos Hidraulicos" },
  { wbs: "8.2.3", phase: "Projeto Executivo", subphase: "Desenhos Complementares", type: "TAREFA", task: "Planta de Forro" },
  { wbs: "8.2.4", phase: "Projeto Executivo", subphase: "Desenhos Complementares", type: "TAREFA", task: "Planta de Piso" },
  { wbs: "8.2.5", phase: "Projeto Executivo", subphase: "Desenhos Complementares", type: "TAREFA", task: "Projeto Luminotecnico" },
  { wbs: "8.3.1", phase: "Projeto Executivo", subphase: "Detalhamentos Construtivos", type: "TAREFA", task: "Detalhamento de Escadas e Rampas" },
  { wbs: "8.3.2", phase: "Projeto Executivo", subphase: "Detalhamentos Construtivos", type: "TAREFA", task: "Detalhamento de Esquadrias" },
  { wbs: "8.3.3", phase: "Projeto Executivo", subphase: "Detalhamentos Construtivos", type: "TAREFA", task: "Detalhes Construtivos" },
  { wbs: "8.3.4", phase: "Projeto Executivo", subphase: "Detalhamentos Construtivos", type: "TAREFA", task: "Detalhamento de Gradil" },
  { wbs: "8.3.5", phase: "Projeto Executivo", subphase: "Detalhamentos Construtivos", type: "TAREFA", task: "Detalhamento de Areas Molhadas" },
  { wbs: "8.3.6", phase: "Projeto Executivo", subphase: "Detalhamentos Construtivos", type: "TAREFA", task: "Detalhamento de Guarda-corpo" },
  { wbs: "9.1.1", phase: "Detalhamentos Complementares", subphase: "Detalhamentos Especificos", type: "TAREFA", task: "Detalhamento de Vidros e Espelhos" },
  { wbs: "9.1.2", phase: "Detalhamentos Complementares", subphase: "Detalhamentos Especificos", type: "TAREFA", task: "Design e Detalhamento de Moveis" },
  { wbs: "9.1.3", phase: "Detalhamentos Complementares", subphase: "Detalhamentos Especificos", type: "TAREFA", task: "Detalhamentos Complementares Especificos" },
  { wbs: "10.1.1", phase: "Acompanhamento de Obra e Producao", subphase: "Acompanhamento de Obra", type: "TAREFA", task: "Realizacao de Visitas Tecnicas em Obra (2 visitas)" },
  { wbs: "10.2.1", phase: "Acompanhamento de Obra e Producao", subphase: "Acompanhamento de Producao", type: "TAREFA", task: "Realizacao de Acompanhamento de Producao (1 visita)" }
];

var EAP_TEMPLATE_NETWORK = {
  "1.1.1": { du: 3, deps: [] },
  "1.1.2": { du: 3, deps: [] },
  "1.1.3": { du: 2, deps: [] },
  "2.1.1": { du: 4, deps: ["1.1.1", "1.1.2", "1.1.3"] },
  "2.1.2": { du: 3, deps: ["2.1.1"] },
  "2.2.1": { du: 2, deps: ["2.1.2"] },
  "2.2.2": { du: 1, deps: ["2.2.1"] },
  "3.1.1": { du: 4, deps: ["2.2.2"] },
  "3.1.2": { du: 2, deps: ["2.2.2"] },
  "3.1.3": { du: 2, deps: ["2.2.2"] },
  "3.2.1": { du: 2, deps: ["3.1.1", "3.1.2", "3.1.3"] },
  "3.2.2": { du: 1, deps: ["3.2.1"] },
  "4.1.1": { du: 5, deps: ["3.2.2"] },
  "4.2.1": { du: 3, deps: ["4.1.1"] },
  "4.2.2": { du: 1, deps: ["4.2.1"] },
  "5.1.1": { du: 5, deps: ["3.2.2"] },
  "5.1.2": { du: 3, deps: ["3.2.2"] },
  "5.1.3": { du: 3, deps: ["3.2.2"] },
  "5.2.1": { du: 2, deps: ["5.1.1", "5.1.2", "5.1.3"] },
  "5.2.2": { du: 1, deps: ["5.2.1"] },
  "6.1.1": { du: 3, deps: ["5.2.2"] },
  "6.1.2": { du: 2, deps: ["5.2.2"] },
  "6.1.3": { du: 2, deps: ["5.2.2"] },
  "6.1.4": { du: 2, deps: ["5.2.2"] },
  "7.1.1": { du: 5, deps: ["6.1.1", "6.1.2", "6.1.3", "6.1.4"] },
  "7.2.1": { du: 3, deps: ["7.1.1"] },
  "7.2.2": { du: 1, deps: ["7.2.1"] },
  "8.1.1": { du: 2, deps: ["7.2.2"] },
  "8.1.2": { du: 2, deps: ["7.2.2"] },
  "8.1.3": { du: 2, deps: ["7.2.2"] },
  "8.1.4": { du: 3, deps: ["7.2.2"] },
  "8.1.5": { du: 3, deps: ["7.2.2"] },
  "8.2.1": { du: 2, deps: ["7.2.2"] },
  "8.2.2": { du: 2, deps: ["7.2.2"] },
  "8.2.3": { du: 2, deps: ["7.2.2"] },
  "8.2.4": { du: 2, deps: ["7.2.2"] },
  "8.2.5": { du: 2, deps: ["7.2.2"] },
  "8.3.1": { du: 2, deps: ["8.1.4"] },
  "8.3.2": { du: 2, deps: ["8.1.5"] },
  "8.3.3": { du: 3, deps: ["8.1.4", "8.1.5"] },
  "8.3.4": { du: 2, deps: ["8.3.3"] },
  "8.3.5": { du: 2, deps: ["8.2.2", "8.2.4"] },
  "8.3.6": { du: 2, deps: ["8.3.3"] },
  "9.1.1": { du: 2, deps: ["8.3.2"] },
  "9.1.2": { du: 4, deps: ["7.2.2"] },
  "9.1.3": { du: 2, deps: ["8.3.3"] },
  "10.1.1": { du: 6, deps: ["8.1.2", "8.2.1", "8.2.2", "8.2.3", "8.2.4", "8.2.5"] },
  "10.2.1": { du: 3, deps: ["10.1.1", "9.1.2"] }
};

function doGet(e) {
  PROJECT_MAP_CACHE = null;
  var action = "health";
  if (e && e.parameter && e.parameter.action) {
    action = e.parameter.action;
  }
  var callback = e && e.parameter ? e.parameter.callback : "";
  var response;

  if (action === "health") {
    response = {
      status: "ok",
      message: "Backend ativo.",
      controlSheet: CONTROL_SHEET_NAME
    };
    return jsonOrJsonp_(response, callback);
  }

  if (action === "config") {
    response = buildConfigPayload_();
    return jsonOrJsonp_(response, callback);
  }

  if (action === "ppm_snapshot") {
    response = buildPpmSnapshotPayload_();
    return jsonOrJsonp_(response, callback);
  }

  if (action === "coord_project") {
    response = buildCoordProjectPayload_(e && e.parameter ? e.parameter.contractId : "");
    return jsonOrJsonp_(response, callback);
  }

  if (action === "history") {
    response = buildHistoryPayload_(
      e && e.parameter ? e.parameter.professional : "",
      e && e.parameter ? e.parameter.limit : ""
    );
    return jsonOrJsonp_(response, callback);
  }

  if (action === "resumo_controle") {
    response = resumoControleEapAtual();
    return jsonOrJsonp_(response, callback);
  }

  response = {
    status: "error",
    message: "Acao invalida."
  };
  return jsonOrJsonp_(response, callback);
}

function doPost(e) {
  PROJECT_MAP_CACHE = null;
  var lock = LockService.getScriptLock();
  var lockAcquired = false;

  try {
    lock.waitLock(30000);
    lockAcquired = true;

    var payload = parsePayload_(e);
    if (clean_(payload.action) === "coord_update_project") {
      return handleCoordProjectUpdate_(payload);
    }

    var input = normalizePayload_(payload);
    var editFormId = clean_(payload.editFormId);
    var executionSheet = getExecutionSheet_();
    var nowBr = Utilities.formatDate(new Date(), getTz_(), "dd/MM/yyyy");
    var controlContext = prepareControlContext_();
    var formId = editFormId || "";

    if (!editFormId) {
      enforceNoDuplicate_(executionSheet, input.professional, input.dateKey);
      formId = generateFormId_(executionSheet, input.dateObj);
    } else {
      ensureExistingFormId_(executionSheet, editFormId);
    }

    var records = buildSubmissionRecords_(input, controlContext.rows);
    var executionRows = buildExecutionRows_(records, input, formId);

    if (editFormId) {
      var preservedControlRows = cloneControlRows_(controlContext.rows);
      replaceExecutionRowsByFormId_(executionSheet, editFormId, executionRows);
      controlContext.rows = rebuildControlFromExecution_(executionSheet, {
        preserveRows: preservedControlRows
      });
    } else {
      var executionStartRow = Math.max(executionSheet.getLastRow() + 1, EXEC_DATA_START_ROW);
      executionSheet.getRange(executionStartRow, 1, executionRows.length, EXEC_COLUMNS).setValues(executionRows);

      var syncMeta = {
        dateBr: input.dateBr,
        professional: input.professional,
        formId: formId,
        updatedAtBr: nowBr
      };
      syncControlFromRecords_(controlContext, records, syncMeta);
    }

    var dailySummary = buildDailySubmissionSummary_(input, records, controlContext.rows);

    return json_({
      status: "ok",
      message: editFormId ? "Edicao salva com sucesso." : "Registro salvo com sucesso.",
      formId: formId,
      rowsCreated: executionRows.length,
      edited: !!editFormId,
      controlUpdated: true,
      dailySummary: dailySummary
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
  var refEap = clean_(item.refEap);
  var refLabel = clean_(item.refLabel);
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
  if (!refEap) {
    throw new Error("Campo REF EAP obrigatorio na secao " + sectionNumber + ".");
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

  var taskFreeText = clean_(item.taskFreeText);
  if (taskFreeText.length > 30) {
    throw new Error("Campo 4B complementar excede 30 caracteres na secao " + sectionNumber + ".");
  }

  var observation = clean_(item.observation);
  if (observation.length > 500) {
    throw new Error("Campo 8 excede 500 caracteres na secao " + sectionNumber + ".");
  }

  return {
    projectCode: projectCode,
    projectLabel: projectLabel || projectCode,
    refEap: refEap,
    refLabel: refLabel,
    tasks: tasks,
    statusText: statusLabel,
    isBlocked: isBlocked,
    blockReason: isBlocked ? (blockDescription || blockLabel) : "",
    observation: observation
  };
}

function getExecutionSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXEC_SHEET_NAME);
  if (!sheet) {
    throw new Error("Aba " + EXEC_SHEET_NAME + " nao encontrada.");
  }
  ensureExecutionSheetReady_(sheet);
  return sheet;
}

function ensureExecutionSheetReady_(sheet) {
  if (!sheet) {
    throw new Error("Aba de execucao nao encontrada.");
  }
  if (sheet.getLastColumn() < EXEC_COLUMNS) {
    throw new Error("Aba " + EXEC_SHEET_NAME + " com estrutura invalida: minimo de " + EXEC_COLUMNS + " colunas.");
  }
}

function getEapSheet_() {
  return getOrCreateEapSheet_();
}

function getOrCreateEapSheet_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(EAP_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(EAP_SHEET_NAME);
  }
  ensureEapHeader_(sheet);
  return sheet;
}

function ensureEapHeader_(sheet) {
  var headerRow = EAP_DATA_START_ROW - 1;
  var current = sheet.getRange(headerRow, 1, 1, EAP_HEADERS.length).getValues()[0];
  var needsUpdate = false;
  var i;

  for (i = 0; i < EAP_HEADERS.length; i += 1) {
    if (clean_(current[i]) !== EAP_HEADERS[i]) {
      needsUpdate = true;
      break;
    }
  }

  if (needsUpdate) {
    sheet.getRange(headerRow, 1, 1, EAP_HEADERS.length).setValues([EAP_HEADERS]);
    sheet.getRange(headerRow, 1, 1, EAP_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(headerRow);
  }
}

function writeEapRows_(sheet, rows) {
  var lastRow = sheet.getLastRow();
  if (lastRow >= EAP_DATA_START_ROW) {
    sheet.getRange(EAP_DATA_START_ROW, 1, lastRow - EAP_DATA_START_ROW + 1, EAP_HEADERS.length).clearContent();
  }

  if (rows && rows.length) {
    sheet.getRange(EAP_DATA_START_ROW, 1, rows.length, EAP_HEADERS.length).setValues(rows);
  }
}

function clearExecutionData_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < EXEC_DATA_START_ROW) {
    return 0;
  }
  var rowsToClear = lastRow - EXEC_DATA_START_ROW + 1;
  sheet.getRange(EXEC_DATA_START_ROW, 1, rowsToClear, EXEC_COLUMNS).clearContent();
  return rowsToClear;
}

function clearEventsData_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < EVENTS_DATA_START_ROW) {
    return 0;
  }
  var rowsToClear = lastRow - EVENTS_DATA_START_ROW + 1;
  sheet.getRange(EVENTS_DATA_START_ROW, 1, rowsToClear, EVENTS_HEADERS.length).clearContent();
  return rowsToClear;
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

function cloneControlRows_(rows) {
  var source = Array.isArray(rows) ? rows : [];
  var result = [];
  var i;
  for (i = 0; i < source.length; i += 1) {
    result.push(source[i].slice(0, CONTROL_HEADERS.length));
  }
  return result;
}

function mergePreservedControlRows_(baseRows, preservedRows) {
  var merged = cloneControlRows_(baseRows);
  var preserved = cloneControlRows_(preservedRows);
  var byKey = {};
  var i;

  for (i = 0; i < merged.length; i += 1) {
    byKey[makeControlKey_(merged[i][2], merged[i][3])] = i;
  }

  for (i = 0; i < preserved.length; i += 1) {
    var row = preserved[i];
    var key = makeControlKey_(row[2], row[3]);
    if (!clean_(row[2]) || !clean_(row[3])) {
      continue;
    }

    if (byKey[key] >= 0) {
      merged[byKey[key]] = row.slice(0, CONTROL_HEADERS.length);
      merged[byKey[key]][0] = key;
    } else {
      row[0] = key;
      merged.push(row.slice(0, CONTROL_HEADERS.length));
    }
  }

  return merged;
}

function preserveCoordinatorDataForUntouchedRows_(rows, preservedRows, touchedByExecution) {
  var source = cloneControlRows_(preservedRows);
  var byKey = {};
  var i;

  for (i = 0; i < source.length; i += 1) {
    var row = source[i];
    byKey[makeControlKey_(row[2], row[3])] = row;
  }

  for (i = 0; i < rows.length; i += 1) {
    var current = rows[i];
    var key = makeControlKey_(current[2], current[3]);
    if (touchedByExecution[key]) {
      continue;
    }
    var preserved = byKey[key];
    if (!preserved) {
      continue;
    }

    // Mantem configuracoes de coordenacao em tarefas nao tocadas por este replay.
    current[6] = clean_(preserved[6]);  // PRAZO DU
    current[7] = clean_(preserved[7]);  // INICIO PLANEJADO
    current[8] = clean_(preserved[8]);  // FIM PLANEJADO
    current[9] = clean_(preserved[9]);  // STATUS
    current[10] = clean_(preserved[10]); // RESPONSAVEL
    current[11] = clean_(preserved[11]); // INICIO REAL
    current[12] = clean_(preserved[12]); // FIM REAL
    current[13] = toNumberOrBlank_(preserved[13], 0); // % REAL
    current[14] = clean_(preserved[14]); // DIAS DESVIO
    current[15] = clean_(preserved[15]); // BLOQUEIO
    current[16] = clean_(preserved[16]); // DESC BLOQUEIO
    current[17] = clean_(preserved[17]); // ATUALIZADO EM
    current[18] = clean_(preserved[18]); // DATA ULT. REGISTRO
    current[19] = clean_(preserved[19]); // ID ULT. ENVIO
    current[20] = clean_(preserved[20]); // OBS ULT. REGISTRO
    current[21] = clean_(preserved[21]); // ORIGEM
    current[22] = clean_(preserved[22]); // PREDECESSORA REF
    current[23] = clean_(preserved[23]) || "FS"; // RELACAO DEP
    current[24] = toNumberOrBlank_(preserved[24], 0); // LAG DU
  }
}

function applyCanonicalProjectNamesToControlRows_(rows, map) {
  var projectMap = map || getProjectMap_();
  var i;
  for (i = 0; i < rows.length; i += 1) {
    var contractId = clean_(rows[i][2]);
    rows[i][1] = resolveCanonicalProjectName_(contractId, rows[i][1], projectMap);
    rows[i][0] = makeControlKey_(contractId, rows[i][3]);
  }
}

function getOrCreateEventsSheet_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(EVENTS_SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(EVENTS_SHEET_NAME);
  }

  ensureEventsHeader_(sheet);
  return sheet;
}

function ensureEventsHeader_(sheet) {
  var current = sheet.getRange(1, 1, 1, EVENTS_HEADERS.length).getValues()[0];
  var needsUpdate = false;
  var i;

  for (i = 0; i < EVENTS_HEADERS.length; i += 1) {
    if (clean_(current[i]) !== EVENTS_HEADERS[i]) {
      needsUpdate = true;
      break;
    }
  }

  if (needsUpdate) {
    sheet.getRange(1, 1, 1, EVENTS_HEADERS.length).setValues([EVENTS_HEADERS]);
    sheet.getRange(1, 1, 1, EVENTS_HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
}

function readEventsByContract_(sheet, contractId) {
  var target = clean_(contractId);
  var values = [];
  var result = [];
  var lastRow = sheet.getLastRow();
  var i;

  if (!target || lastRow < EVENTS_DATA_START_ROW) {
    return result;
  }

  values = sheet.getRange(EVENTS_DATA_START_ROW, 1, lastRow - EVENTS_DATA_START_ROW + 1, EVENTS_HEADERS.length).getValues();

  for (i = 0; i < values.length; i += 1) {
    var row = values[i];
    if (clean_(row[1]) !== target) {
      continue;
    }
    result.push({
      eventDate: toIsoDateOrBlank_(row[2]),
      eventType: clean_(row[3]),
      title: clean_(row[4]),
      responsible: clean_(row[5]),
      status: clean_(row[6]) || "Planejado",
      observation: clean_(row[7])
    });
  }

  result.sort(function (a, b) {
    var left = clean_(a.eventDate);
    var right = clean_(b.eventDate);
    if (left && right) {
      return left.localeCompare(right);
    }
    if (left) {
      return -1;
    }
    if (right) {
      return 1;
    }
    return clean_(a.title).localeCompare(clean_(b.title));
  });

  return result;
}

function replaceEventsByContract_(sheet, contractId, events, editor) {
  var target = clean_(contractId);
  var safeEditor = clean_(editor) || "COORDENACAO";
  var nowBr = Utilities.formatDate(new Date(), getTz_(), "dd/MM/yyyy");
  var lastRow = sheet.getLastRow();
  var toDelete = [];
  var i;

  if (!target) {
    return 0;
  }

  if (lastRow >= EVENTS_DATA_START_ROW) {
    var values = sheet
      .getRange(EVENTS_DATA_START_ROW, 1, lastRow - EVENTS_DATA_START_ROW + 1, EVENTS_HEADERS.length)
      .getValues();
    for (i = 0; i < values.length; i += 1) {
      if (clean_(values[i][1]) === target) {
        toDelete.push(EVENTS_DATA_START_ROW + i);
      }
    }
  }

  toDelete.sort(function (a, b) {
    return b - a;
  });
  for (i = 0; i < toDelete.length; i += 1) {
    sheet.deleteRow(toDelete[i]);
  }

  if (!events || !events.length) {
    return 0;
  }

  var rows = [];
  for (i = 0; i < events.length; i += 1) {
    var item = events[i];
    rows.push([
      target + "-EV-" + pad_(i + 1, 3),
      target,
      isoToBr_(item.eventDate),
      clean_(item.eventType),
      clean_(item.title),
      clean_(item.responsible),
      clean_(item.status) || "Planejado",
      clean_(item.observation),
      nowBr,
      safeEditor
    ]);
  }

  var startRow = Math.max(sheet.getLastRow() + 1, EVENTS_DATA_START_ROW);
  sheet.getRange(startRow, 1, rows.length, EVENTS_HEADERS.length).setValues(rows);
  return rows.length;
}

function prepareControlContext_() {
  var controlSheet = getOrCreateControlSheet_();
  var rows = readControlRows_(controlSheet);

  if (!rows.length) {
    rows = seedControlRowsFromStandardTemplate_();
    writeControlRows_(controlSheet, rows);
  }

  return {
    sheet: controlSheet,
    rows: rows
  };
}

function seedControlRowsFromEap_() {
  // Compatibilidade: o controle agora nasce da EAP padrao em codigo.
  return seedControlRowsFromStandardTemplate_();
}

function seedControlRowsFromStandardTemplate_() {
  return buildControlRowsFromStandardTemplate_();
}

function buildBaselineRowsFromStandardTemplate_() {
  var rows = [];
  var projects = getActiveContractProjects_();
  var i;
  var j;

  for (i = 0; i < projects.length; i += 1) {
    var project = projects[i];
    for (j = 0; j < EAP_TEMPLATE_ITEMS.length; j += 1) {
      var item = EAP_TEMPLATE_ITEMS[j];
      rows.push([
        project.projectName,
        project.contractId,
        item.phase,
        item.subphase,
        item.task,
        "",
        "",
        "",
        item.type,
        item.type === "MARCO" ? "SIM" : "NAO",
        "WBS " + item.wbs,
        "EAP_PADRAO_V1"
      ]);
    }
  }

  // Mantem opcao de trabalho administrativo/interno.
  rows.push([
    "Administrativo/Interno",
    "ADMIN-INTERNO",
    "Administrativo/Interno",
    "Rotinas Internas",
    "Atividades administrativas internas",
    "",
    "",
    "",
    "TAREFA",
    "NAO",
    "WBS INT.1",
    "EAP_PADRAO_V1"
  ]);

  return rows;
}

function buildControlRowsFromStandardTemplate_() {
  var rows = [];
  var projects = getActiveContractProjects_();
  var i;
  var j;

  for (i = 0; i < projects.length; i += 1) {
    var project = projects[i];
    for (j = 0; j < EAP_TEMPLATE_ITEMS.length; j += 1) {
      var item = EAP_TEMPLATE_ITEMS[j];
      var network = EAP_TEMPLATE_NETWORK[item.wbs] || { du: 1, deps: [] };
      rows.push(buildControlRow_({
        projectName: project.projectName,
        contractId: project.contractId,
        refEap: buildRefFromWbs_(project.contractId, item.wbs),
        phase: item.phase,
        task: item.task,
        prazoDu: Math.max(1, toNumber_(network.du, 1)),
        status: STATUS_CONTROL.NAO_INICIADA,
        predecessorRef: buildDependencyRefTextFromWbs_(project.contractId, network.deps),
        relationType: "FS",
        lagDu: 0,
        origem: "EAP_PADRAO_V1"
      }));
    }
  }

  rows.push(buildControlRow_({
    projectName: "Administrativo/Interno",
    contractId: "ADMIN-INTERNO",
    refEap: "ADMININTERNO-WBS-INT_1",
    phase: "Administrativo/Interno",
    task: "Atividades administrativas internas",
    prazoDu: 1,
    status: STATUS_CONTROL.NAO_INICIADA,
    predecessorRef: "",
    relationType: "FS",
    lagDu: 0,
    origem: "EAP_PADRAO_V1"
  }));

  return rows;
}

function buildDependencyRefTextFromWbs_(contractId, deps) {
  var list = Array.isArray(deps) ? deps : [];
  var refs = [];
  var i;
  for (i = 0; i < list.length; i += 1) {
    refs.push(buildRefFromWbs_(contractId, list[i]));
  }
  return refs.join("; ");
}

function getActiveContractProjects_() {
  var map = getProjectMap_();
  var list = [];
  var code;

  for (code in map) {
    if (!Object.prototype.hasOwnProperty.call(map, code)) {
      continue;
    }
    if (String(code).indexOf("CT-") !== 0) {
      continue;
    }
    list.push(map[code]);
  }

  list.sort(function (a, b) {
    return compareContractId_(a.contractId, b.contractId);
  });
  return list;
}

function getProjectsForConfig_() {
  var map = getProjectMap_();
  var projects = [];
  var code;

  for (code in map) {
    if (!Object.prototype.hasOwnProperty.call(map, code)) {
      continue;
    }
    projects.push({
      code: map[code].contractId,
      label: map[code].projectName,
      signatureDate: clean_(map[code].signatureDate),
      deadlineSummary: clean_(map[code].deadlineSummary)
    });
  }

  projects.sort(function (a, b) {
    return compareContractId_(a.code, b.code);
  });

  return projects;
}

function getProjectMap_() {
  if (PROJECT_MAP_CACHE) {
    return PROJECT_MAP_CACHE;
  }

  var map = {};
  var code;
  var i;
  var discovered = readProjectsFromContractsSheet_();
  var hints = readProjectHintsFromOperationalSheets_();

  for (code in PROJECT_MAP) {
    if (!Object.prototype.hasOwnProperty.call(PROJECT_MAP, code)) {
      continue;
    }
    map[code] = {
      projectName: clean_(PROJECT_MAP[code].projectName),
      contractId: clean_(PROJECT_MAP[code].contractId),
      sectionName: clean_(PROJECT_MAP[code].sectionName),
      signatureDate: clean_(PROJECT_MAP[code].signatureDate),
      deadlineSummary: clean_(PROJECT_MAP[code].deadlineSummary)
    };
  }

  for (i = 0; i < discovered.length; i += 1) {
    var item = discovered[i];
    var existing = map[item.contractId] || {};
    map[item.contractId] = {
      projectName: resolveCanonicalProjectName_(item.contractId, item.projectName, map),
      contractId: item.contractId,
      sectionName: "Secao " + item.contractId,
      signatureDate: clean_(item.signatureDate) || clean_(existing.signatureDate),
      deadlineSummary: clean_(item.deadlineSummary) || clean_(existing.deadlineSummary)
    };
  }

  for (i = 0; i < hints.length; i += 1) {
    var hint = hints[i];
    var prev = map[hint.contractId] || {};
    map[hint.contractId] = {
      projectName: resolveCanonicalProjectName_(hint.contractId, hint.projectName || prev.projectName, map),
      contractId: hint.contractId,
      sectionName: "Secao " + hint.contractId,
      signatureDate: clean_(prev.signatureDate),
      deadlineSummary: clean_(prev.deadlineSummary)
    };
  }

  for (code in map) {
    if (!Object.prototype.hasOwnProperty.call(map, code)) {
      continue;
    }
    map[code].projectName = resolveCanonicalProjectName_(code, map[code].projectName, map);
  }

  if (!map["ADMIN-INTERNO"]) {
    map["ADMIN-INTERNO"] = {
      projectName: "Administrativo/Interno",
      contractId: "ADMIN-INTERNO",
      sectionName: "Secao Administrativo/Interno"
    };
  }

  PROJECT_MAP_CACHE = map;
  return map;
}

function readProjectHintsFromOperationalSheets_() {
  var hints = {};
  var spreadsheets = SpreadsheetApp.getActiveSpreadsheet();
  var controlSheet = spreadsheets.getSheetByName(CONTROL_SHEET_NAME);

  collectProjectHintsFromSheet_(controlSheet, CONTROL_DATA_START_ROW, 2, 1, hints);

  var list = [];
  var code;
  for (code in hints) {
    if (Object.prototype.hasOwnProperty.call(hints, code)) {
      list.push({
        contractId: code,
        projectName: hints[code]
      });
    }
  }
  list.sort(function (a, b) {
    return compareContractId_(a.contractId, b.contractId);
  });
  return list;
}

function collectProjectHintsFromSheet_(sheet, startRow, contractColumn, projectColumn, store) {
  if (!sheet || sheet.getLastRow() < startRow) {
    return;
  }

  var width = Math.max(contractColumn, projectColumn) + 1;
  var values = sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, width).getValues();
  var i;

  for (i = 0; i < values.length; i += 1) {
    var row = values[i];
    var contractId = extractContractId_(row[contractColumn]);
    if (!contractId || contractId === "ADMIN-INTERNO") {
      continue;
    }

    var projectName = normalizeProjectDisplayName_(row[projectColumn]);
    if (!clean_(store[contractId])) {
      store[contractId] = projectName || contractId;
      continue;
    }

    if (isContractDisplayName_(projectName) && !isContractDisplayName_(store[contractId])) {
      store[contractId] = projectName;
    }
  }
}

function readProjectsFromContractsSheet_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONTRACTS_SHEET_NAME);
  if (!sheet) {
    return [];
  }

  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) {
    return [];
  }

  var headers = values[0].map(normalizeHeaderText_);
  var contractIdIndex = findHeaderIndex_(headers, [
    "id ct",
    "id contrato",
    "contrato",
    "codigo contrato",
    "codigo ct",
    "ct"
  ]);
  var projectNameIndex = findHeaderIndexByPriority_(headers, [
    "nome exibicao",
    "nome de exibicao",
    "nome completo",
    "nome contrato",
    "nome do contrato",
    "titulo contrato",
    "titulo do contrato",
    "cliente",
    "clientes",
    "nome cliente",
    "nome clientes",
    "arquivo contrato",
    "arquivo do contrato",
    "apelido",
    "nome projeto",
    "projeto",
    "nome do projeto",
    "nome"
  ]);
  if (projectNameIndex < 0) {
    projectNameIndex = findHeaderIndex_(headers, [
      "apelido",
      "nome projeto",
      "projeto",
      "nome do projeto",
      "nome"
    ]);
  }
  var statusIndex = findHeaderIndex_(headers, ["status", "situacao", "situação", "ativo"]);
  var deadlineSummaryIndex = findHeaderIndexByPriority_(headers, [
    "clausula prazos texto completo",
    "clausula de prazos texto completo",
    "texto completo clausula prazos",
    "texto prazos contratuais",
    "prazos contratuais completo",
    "prazos contratuais completos",
    "prazos detalhados",
    "clausula prazos",
    "clausula de prazos",
    "prazos contratuais",
    "prazos resumo",
    "resumo prazos",
    "prazo resumo",
    "resumo clausula prazos",
    "resumo clausula de prazos",
    "prazos"
  ]);
  if (deadlineSummaryIndex < 0) {
    deadlineSummaryIndex = findHeaderIndex_(headers, [
      "prazos resumo",
      "resumo prazos",
      "prazo resumo",
      "prazos contratuais",
      "resumo clausula prazos",
      "resumo clausula de prazos",
      "clausula prazos"
    ]);
  }
  var signatureDateIndex = findHeaderIndex_(headers, [
    "data assinatura",
    "data de assinatura",
    "assinatura",
    "dt assinatura",
    "data contrato",
    "data do contrato",
    "assinado em"
  ]);
  var map = {};
  var i;

  for (i = 1; i < values.length; i += 1) {
    var row = values[i];
    if (isBlankRow_(row)) {
      continue;
    }

    var contractId = contractIdIndex >= 0 ? extractContractId_(row[contractIdIndex]) : "";
    if (!contractId) {
      contractId = findContractIdInRow_(row);
    }
    if (!contractId) {
      continue;
    }
    if (statusIndex >= 0 && shouldSkipContractByStatus_(row[statusIndex])) {
      continue;
    }

    var projectName = projectNameIndex >= 0 ? clean_(row[projectNameIndex]) : "";
    var contractCell = contractIdIndex >= 0 ? row[contractIdIndex] : "";
    var contractDisplayName = extractContractDisplayName_(contractCell);
    if (!projectName && contractIdIndex >= 0 && contractIdIndex + 1 < row.length) {
      projectName = clean_(row[contractIdIndex + 1]);
    }
    if (!projectName && contractDisplayName) {
      projectName = contractDisplayName;
    }
    if (!projectName) {
      projectName = findContractDisplayNameInRow_(row);
    }
    projectName = choosePreferredProjectName_(contractId, projectName, contractDisplayName);

    map[contractId] = {
      projectName: projectName,
      contractId: contractId,
      sectionName: "Secao " + contractId,
      signatureDate: signatureDateIndex >= 0 ? toIsoDateOrBlank_(row[signatureDateIndex]) : "",
      deadlineSummary: deadlineSummaryIndex >= 0 ? clean_(row[deadlineSummaryIndex]) : ""
    };
  }

  var list = [];
  var code;
  for (code in map) {
    if (Object.prototype.hasOwnProperty.call(map, code)) {
      list.push(map[code]);
    }
  }

  list.sort(function (a, b) {
    return compareContractId_(a.contractId, b.contractId);
  });
  return list;
}

function compareContractId_(a, b) {
  var left = clean_(a).toUpperCase();
  var right = clean_(b).toUpperCase();

  if (left === "ADMIN-INTERNO") {
    return right === "ADMIN-INTERNO" ? 0 : 1;
  }
  if (right === "ADMIN-INTERNO") {
    return -1;
  }

  return left.localeCompare(right);
}

function normalizeHeaderText_(value) {
  return normalizeText_(value).replace(/\s+/g, " ").trim();
}

function findHeaderIndex_(headers, aliases) {
  var i;
  var j;
  for (i = 0; i < headers.length; i += 1) {
    for (j = 0; j < aliases.length; j += 1) {
      var alias = normalizeHeaderText_(aliases[j]);
      if (headers[i] === alias || headers[i].indexOf(alias) >= 0) {
        return i;
      }
    }
  }
  return -1;
}

function findHeaderIndexByPriority_(headers, aliases) {
  var i;
  var j;
  for (j = 0; j < aliases.length; j += 1) {
    var alias = normalizeHeaderText_(aliases[j]);
    for (i = 0; i < headers.length; i += 1) {
      if (headers[i] === alias || headers[i].indexOf(alias) >= 0) {
        return i;
      }
    }
  }
  return -1;
}

function isBlankRow_(row) {
  var i;
  for (i = 0; i < row.length; i += 1) {
    if (clean_(row[i])) {
      return false;
    }
  }
  return true;
}

function findContractIdInRow_(row) {
  var i;
  for (i = 0; i < row.length; i += 1) {
    var id = extractContractId_(row[i]);
    if (id) {
      return id;
    }
  }
  return "";
}

function findContractDisplayNameInRow_(row) {
  var i;
  for (i = 0; i < row.length; i += 1) {
    var display = extractContractDisplayName_(row[i]);
    if (display) {
      return display;
    }
  }
  return "";
}

function extractContractDisplayName_(value) {
  var text = clean_(value);
  if (!text) {
    return "";
  }

  text = text.replace(/\\/g, "/");
  if (text.indexOf("/") >= 0) {
    text = text.substring(text.lastIndexOf("/") + 1);
  }

  text = text
    .replace(/^Complete_with_Docusign_/i, "")
    .replace(/^Complete with Docusign\s*/i, "")
    .replace(/\.(pdf|docx?|xlsx?)$/i, "")
    .replace(/[_]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  text = text.replace(/\s*-\s*(clicksign|d4sign|docusign|docusing)\s*$/i, "").trim();
  text = text.replace(/\s*-\s*pdf\s*$/i, "").trim();

  var structured = text.match(
    /CT[\s\-_]*([0-9]{8})\s*[-_/]\s*([0-9]{3})(?:\s*[-_/]\s*([0-9]{3}))?\s*[-_/]?\s*(.*)$/i
  );
  if (structured) {
    var codePart = structured[1] + "-" + structured[2] + (structured[3] ? "-" + structured[3] : "");
    var suffix = clean_(structured[4]).replace(/^[-_\s]+/, "").replace(/[-_\s]+$/, "");
    suffix = suffix.replace(/\s+/g, " ");
    if (suffix) {
      return normalizeProjectDisplayName_("CT " + codePart + " - " + suffix);
    }
    return normalizeProjectDisplayName_("CT " + codePart);
  }

  if (!/CT[\s\-_]*[0-9]{2,6}/i.test(text)) {
    return "";
  }

  return normalizeProjectDisplayName_(text);
}

function normalizeProjectDisplayName_(value) {
  var text = clean_(value);
  if (!text) {
    return "";
  }

  text = text
    .replace(/\.(pdf|docx?|xlsx?)$/i, "")
    .replace(/\s*-\s*(clicksign|d4sign|docusign|docusing)\s*$/i, "")
    .replace(/\s*-\s*pdf\s*$/i, "")
    .replace(/\s+/g, " ")
    .trim();

  return text;
}

function isContractDisplayName_(value) {
  var text = clean_(value).toUpperCase();
  if (!text) {
    return false;
  }
  return /^CT[\s\-_]+[0-9]{8}[-_/][0-9]{3}(?:[-_/][0-9]{3})?\b/.test(text);
}

function resolveCanonicalProjectName_(contractId, fallbackName, map) {
  var target = clean_(contractId);
  var current = clean_(fallbackName);
  var projectMap = map || getProjectMap_();
  var fromMap = projectMap[target] ? clean_(projectMap[target].projectName) : "";

  if (isContractDisplayName_(fromMap)) {
    return fromMap;
  }
  if (isContractDisplayName_(current)) {
    return current;
  }
  if (fromMap) {
    return fromMap;
  }
  if (current) {
    return current;
  }
  return target;
}

function choosePreferredProjectName_(contractId, primaryName, contractDisplayName) {
  var primary = normalizeProjectDisplayName_(primaryName);
  var display = normalizeProjectDisplayName_(contractDisplayName);
  var fallbackFromMap = PROJECT_MAP[contractId] ? normalizeProjectDisplayName_(PROJECT_MAP[contractId].projectName) : "";

  if (isContractDisplayName_(display)) {
    return display;
  }
  if (isContractDisplayName_(primary)) {
    return primary;
  }
  if (isContractDisplayName_(fallbackFromMap)) {
    return fallbackFromMap;
  }
  return primary || display || fallbackFromMap || clean_(contractId);
}

function extractContractId_(value) {
  var text = clean_(value).toUpperCase();
  if (!text) {
    return "";
  }

  var longMatch = text.match(/CT[\s\-_]*([0-9]{8})(?:[\s\-_\/]*([0-9]{3}))(?:[\s\-_\/]*([0-9]{3}))?/);
  if (longMatch) {
    var longNumber = longMatch[3] || longMatch[2];
    var normalizedLong = String(longNumber).replace(/^0+/, "");
    if (!normalizedLong) {
      normalizedLong = "0";
    }
    return "CT-" + normalizedLong;
  }

  var match = text.match(/CT[\s\-_/]*([0-9]{2,6})\b/);
  if (!match) {
    return "";
  }

  var numberPart = String(match[1]).replace(/^0+/, "");
  if (!numberPart) {
    numberPart = "0";
  }

  return "CT-" + numberPart;
}

function shouldSkipContractByStatus_(value) {
  var status = normalizeText_(value);
  if (!status) {
    return false;
  }

  if (
    status.indexOf("cancel") >= 0 ||
    status.indexOf("distrat") >= 0 ||
    status.indexOf("arquiv") >= 0 ||
    status.indexOf("inativ") >= 0
  ) {
    return true;
  }

  return false;
}

function buildRefFromWbs_(contractId, wbs) {
  var compact = compactContractId_(contractId);
  var wbsKey = clean_(wbs).replace(/[^A-Za-z0-9]+/g, "_");
  return compact + "-WBS-" + wbsKey;
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
    clean_(data.origem || "FORM_DAILY"),
    clean_(data.predecessorRef),
    clean_(data.relationType || "FS"),
    toNumberOrBlank_(data.lagDu, 0)
  ];
}

function buildSubmissionRecords_(input, controlRows) {
  var records = [];
  var i;

  for (i = 0; i < input.projects.length; i += 1) {
    var item = input.projects[i];
    var projectInfo = resolveProjectInfo_(item.projectCode, item.projectLabel);
    var phase = resolvePhaseForSubmission_(controlRows, projectInfo.contractId, item.refEap, item.refLabel, item.tasks);
    var taskText = item.tasks.join(" | ");
    var refEap = clean_(item.refEap);

    if (!refEap) {
      refEap = findBestRefEap_(controlRows, projectInfo.contractId, phase, taskText);
      if (!refEap) {
        refEap = buildRefEap_(projectInfo.contractId, phase);
      }
    }

    records.push({
      projectCode: item.projectCode,
      projectLabel: item.projectLabel || projectInfo.projectName,
      projectName: projectInfo.projectName,
      contractId: projectInfo.contractId,
      sectionName: projectInfo.sectionName,
      phase: phase,
      taskText: taskText,
      refEap: refEap,
      refLabel: clean_(item.refLabel),
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

function resolvePhaseForSubmission_(controlRows, contractId, refEap, refLabel, tasks) {
  var phase = findPhaseByRef_(controlRows, contractId, refEap);
  if (phase) {
    return phase;
  }

  phase = parsePhaseFromRefLabel_(refLabel);
  if (phase) {
    return phase;
  }

  return inferPhase_(tasks);
}

function findPhaseByRef_(rows, contractId, refEap) {
  var targetContract = clean_(contractId);
  var targetRef = clean_(refEap);
  var i;

  if (!targetContract || !targetRef) {
    return "";
  }

  for (i = 0; i < rows.length; i += 1) {
    if (clean_(rows[i][2]) !== targetContract) {
      continue;
    }
    if (clean_(rows[i][3]) !== targetRef) {
      continue;
    }
    return clean_(rows[i][4]);
  }

  return "";
}

function parsePhaseFromRefLabel_(refLabel) {
  var text = clean_(refLabel);
  if (!text) {
    return "";
  }

  var parts = text.split("|");
  if (parts.length < 2) {
    return "";
  }

  return clean_(parts[1]);
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

  applyCanonicalProjectNamesToControlRows_(rows, getProjectMap_());
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
    projectName: resolveCanonicalProjectName_(record.contractId, record.projectName),
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
  row[1] = resolveCanonicalProjectName_(record.contractId, record.projectName || row[1]);
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

function calculateDaysSinceDate_(dateValue, referenceDate) {
  var start = parseDate_(dateValue);
  if (!start) {
    return "";
  }

  var end = parseDate_(referenceDate) || new Date();
  var millis = end.getTime() - start.getTime();
  if (millis < 0) {
    return 0;
  }

  return Math.floor(millis / (24 * 60 * 60 * 1000));
}

function resolveProjectInfo_(projectCode, projectLabel) {
  var mapValue = getProjectMap_()[projectCode];
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

function ensureExistingFormId_(sheet, formId) {
  var indexes = findExecutionRowIndexesByFormId_(sheet, formId);
  if (!indexes.length) {
    throw new Error("Nao foi encontrado envio com ID " + formId + " para edicao.");
  }
}

function findExecutionRowIndexesByFormId_(sheet, formId) {
  var target = clean_(formId);
  var indexes = [];
  var lastRow = sheet.getLastRow();
  var i;

  if (!target || lastRow < EXEC_DATA_START_ROW) {
    return indexes;
  }

  var values = sheet.getRange(EXEC_DATA_START_ROW, 17, lastRow - EXEC_DATA_START_ROW + 1, 1).getValues();
  for (i = 0; i < values.length; i += 1) {
    if (clean_(values[i][0]) === target) {
      indexes.push(EXEC_DATA_START_ROW + i);
    }
  }
  return indexes;
}

function replaceExecutionRowsByFormId_(sheet, formId, newRows) {
  var indexes = findExecutionRowIndexesByFormId_(sheet, formId);
  var i;

  if (!indexes.length) {
    throw new Error("Nao foi encontrado envio com ID " + formId + " para edicao.");
  }

  indexes.sort(function (a, b) {
    return b - a;
  });

  for (i = 0; i < indexes.length; i += 1) {
    sheet.deleteRow(indexes[i]);
  }

  if (newRows && newRows.length) {
    var startRow = Math.max(sheet.getLastRow() + 1, EXEC_DATA_START_ROW);
    sheet.getRange(startRow, 1, newRows.length, EXEC_COLUMNS).setValues(newRows);
  }
}

function rebuildControlFromExecution_(executionSheet, options) {
  options = options || {};
  var controlSheet = getOrCreateControlSheet_();
  var preservedRows = Array.isArray(options.preserveRows) ? options.preserveRows : [];
  var projectMap = getProjectMap_();
  var rows = mergePreservedControlRows_(buildControlRowsFromStandardTemplate_(), preservedRows);
  applyCanonicalProjectNamesToControlRows_(rows, projectMap);
  var touchedByExecution = {};
  var updatedAtBr = Utilities.formatDate(new Date(), getTz_(), "dd/MM/yyyy");
  var lastRow = executionSheet.getLastRow();
  var values = [];
  var i;

  if (lastRow >= EXEC_DATA_START_ROW) {
    values = executionSheet.getRange(EXEC_DATA_START_ROW, 1, lastRow - EXEC_DATA_START_ROW + 1, EXEC_COLUMNS).getValues();
  }

  for (i = 0; i < values.length; i += 1) {
    var executionRow = values[i];
    var parsed = parseExecutionRowForControl_(executionRow, updatedAtBr);
    if (!parsed) {
      continue;
    }

    var rowIndex = findControlRowIndexForRecord_(rows, parsed.record);
    if (rowIndex < 0) {
      rows.push(createAdHocControlRow_(parsed.record, parsed.meta));
      rowIndex = rows.length - 1;
    }

    updateControlRowFromRecord_(rows[rowIndex], parsed.record, parsed.meta);
    touchedByExecution[makeControlKey_(rows[rowIndex][2], rows[rowIndex][3])] = true;
  }

  if (preservedRows.length) {
    preserveCoordinatorDataForUntouchedRows_(rows, preservedRows, touchedByExecution);
  }

  applyCanonicalProjectNamesToControlRows_(rows, projectMap);
  writeControlRows_(controlSheet, rows);
  return rows;
}

function parseExecutionRowForControl_(row, fallbackUpdatedAtBr) {
  var contractId = clean_(row[1]);
  var projectName = clean_(row[0]);
  var phase = clean_(row[2]);
  var taskText = clean_(row[3]);

  if (!contractId || !projectName || !phase || !taskText) {
    return null;
  }

  var isBlocked = /^sim/i.test(clean_(row[10]));
  var statusText = clean_(row[8]);
  var record = {
    projectName: projectName,
    contractId: contractId,
    phase: phase,
    taskText: taskText,
    refEap: clean_(row[5]) || buildRefEap_(contractId, phase),
    statusText: statusText,
    statusClass: classifyControlStatus_(statusText, isBlocked),
    isBlocked: isBlocked,
    blockReason: clean_(row[11]),
    observation: clean_(row[14]),
    responsible: clean_(row[6]),
    sectionName: clean_(row[4])
  };

  var meta = {
    dateBr: clean_(row[13]) || fallbackUpdatedAtBr,
    professional: clean_(row[6]),
    formId: clean_(row[16]),
    updatedAtBr: clean_(row[15]) || fallbackUpdatedAtBr
  };

  return { record: record, meta: meta };
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

function cloneDate_(value) {
  var base = value instanceof Date ? value : new Date();
  return new Date(base.getFullYear(), base.getMonth(), base.getDate());
}

function isBusinessDay_(date) {
  var day = date.getDay();
  if (day === 0 || day === 6) {
    return false;
  }
  return !isRecessoDay_(date);
}

function isRecessoDay_(date) {
  if (!(date instanceof Date)) {
    return false;
  }

  var md = Utilities.formatDate(date, getTz_(), "MM-dd");
  if (RECESS_START_MONTH_DAY > RECESS_END_MONTH_DAY) {
    return md >= RECESS_START_MONTH_DAY || md <= RECESS_END_MONTH_DAY;
  }
  return md >= RECESS_START_MONTH_DAY && md <= RECESS_END_MONTH_DAY;
}

function nextBusinessDay_(date) {
  var cursor = cloneDate_(date);
  while (!isBusinessDay_(cursor)) {
    cursor.setDate(cursor.getDate() + 1);
  }
  return cursor;
}

function addBusinessDays_(date, businessDays) {
  var total = Number(businessDays || 0);
  var cursor = nextBusinessDay_(date);
  var step = total >= 0 ? 1 : -1;
  var remaining = Math.abs(total);

  while (remaining > 0) {
    cursor.setDate(cursor.getDate() + step);
    if (isBusinessDay_(cursor)) {
      remaining -= 1;
    }
  }

  return cursor;
}

function subtractBusinessDays_(date, businessDays) {
  return addBusinessDays_(date, -Math.abs(Number(businessDays || 0)));
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

function buildConfigPayload_() {
  var projects = getProjectsForConfig_();
  var signatureDates = {};
  var i;

  for (i = 0; i < projects.length; i += 1) {
    if (!projects[i].code) {
      continue;
    }
    signatureDates[projects[i].code] = clean_(projects[i].signatureDate);
  }

  return {
    status: "ok",
    projects: projects,
    statuses: STATUS_CONTROL,
    refOptions: buildRefOptionsByProject_(),
    signatureDates: signatureDates,
    deadlineSummaryDefault: DEFAULT_CONTRACT_DEADLINE_SUMMARY,
    deadlineLegend: CONTRACT_DEADLINE_LEGEND
  };
}

function buildPpmSnapshotPayload_() {
  var sheet = getOrCreateControlSheet_();
  var rows = readControlRows_(sheet);
  var projectMap = getProjectMap_();
  var signatureDates = {};
  var byContract = {};
  var order = [];
  var i;

  for (var code in projectMap) {
    if (!Object.prototype.hasOwnProperty.call(projectMap, code)) {
      continue;
    }
    signatureDates[code] = clean_(projectMap[code].signatureDate);
  }

  for (i = 0; i < rows.length; i += 1) {
    var row = rows[i];
    var contractId = clean_(row[2]);
    var refEap = clean_(row[3]);

    if (!contractId || !refEap || contractId === "ADMIN-INTERNO") {
      continue;
    }

    if (!byContract[contractId]) {
      var projectInfo = projectMap[contractId] || {};
      byContract[contractId] = {
        contractId: contractId,
        projectName: resolveCanonicalProjectName_(contractId, clean_(row[1]) || clean_(projectInfo.projectName), projectMap),
        signatureDate: clean_(projectInfo.signatureDate),
        tasks: [],
        totals: {
          total: 0,
          naoIniciada: 0,
          emAndamento: 0,
          concluida: 0,
          bloqueada: 0
        }
      };
      order.push(contractId);
    }

    var status = normalizeControlStatus_(row[9]);
    var task = {
      refEap: refEap,
      phase: clean_(row[4]),
      task: clean_(row[5]),
      status: status,
      plannedStart: toIsoDateOrBlank_(row[7]),
      plannedEnd: toIsoDateOrBlank_(row[8]),
      realStart: toIsoDateOrBlank_(row[11]),
      realEnd: toIsoDateOrBlank_(row[12]),
      lastRecordDate: toIsoDateOrBlank_(row[18]),
      updatedAt: toIsoDateOrBlank_(row[17]),
      responsible: clean_(row[10]),
      blocked: /^sim/i.test(clean_(row[15])),
      blockReason: clean_(row[16]),
      percentReal: toNumberOrBlank_(row[13], "")
    };

    byContract[contractId].tasks.push(task);
    byContract[contractId].totals.total += 1;

    if (status === STATUS_CONTROL.CONCLUIDA) {
      byContract[contractId].totals.concluida += 1;
    } else if (status === STATUS_CONTROL.EM_ANDAMENTO) {
      byContract[contractId].totals.emAndamento += 1;
    } else if (status === STATUS_CONTROL.BLOQUEADA) {
      byContract[contractId].totals.bloqueada += 1;
    } else {
      byContract[contractId].totals.naoIniciada += 1;
    }
  }

  var projects = [];
  for (i = 0; i < order.length; i += 1) {
    var contractKey = order[i];
    byContract[contractKey].tasks.sort(function (a, b) {
      return String(a.refEap).localeCompare(String(b.refEap));
    });
    projects.push(byContract[contractKey]);
  }

  projects.sort(function (a, b) {
    return compareContractId_(a.contractId, b.contractId);
  });

  return {
    status: "ok",
    generatedAt: Utilities.formatDate(new Date(), getTz_(), "yyyy-MM-dd'T'HH:mm:ss"),
    projects: projects,
    signatureDates: signatureDates
  };
}

function buildCoordProjectPayload_(contractId) {
  var target = clean_(contractId);
  if (!target) {
    return {
      status: "error",
      message: "Informe o contractId para abrir a coordenacao."
    };
  }

  var controlSheet = getOrCreateControlSheet_();
  var controlRows = readControlRows_(controlSheet);
  var eventsSheet = getOrCreateEventsSheet_();
  var fullProjectMap = getProjectMap_();
  var projectInfo = fullProjectMap[target] || {
    projectName: target,
    contractId: target,
    sectionName: "Secao " + target
  };
  var tasks = [];
  var summary = {
    total: 0,
    concluida: 0,
    emAndamento: 0,
    bloqueada: 0,
    naoIniciada: 0
  };
  var i;

  for (i = 0; i < controlRows.length; i += 1) {
    var row = controlRows[i];
    if (clean_(row[2]) !== target) {
      continue;
    }

    var status = normalizeControlStatus_(row[9]);
    tasks.push({
      refEap: clean_(row[3]),
      phase: clean_(row[4]),
      task: clean_(row[5]),
      status: status,
      responsible: clean_(row[10]),
      plannedStart: toIsoDateOrBlank_(row[7]),
      plannedEnd: toIsoDateOrBlank_(row[8]),
      prazoDu: Math.max(1, toNumber_(row[6], 1)),
      predecessorRef: clean_(row[22]),
      relationType: clean_(row[23]) || "FS",
      lagDu: toNumberOrBlank_(row[24], 0),
      percentReal: toNumberOrBlank_(row[13], 0),
      blocked: /^sim/i.test(clean_(row[15])) || status === STATUS_CONTROL.BLOQUEADA,
      blockReason: clean_(row[16]),
      updatedAt: toIsoDateOrBlank_(row[17]),
      lastRecordDate: toIsoDateOrBlank_(row[18])
    });

    summary.total += 1;
    if (status === STATUS_CONTROL.CONCLUIDA) {
      summary.concluida += 1;
    } else if (status === STATUS_CONTROL.EM_ANDAMENTO) {
      summary.emAndamento += 1;
    } else if (status === STATUS_CONTROL.BLOQUEADA) {
      summary.bloqueada += 1;
    } else {
      summary.naoIniciada += 1;
    }
  }

  tasks.sort(function (a, b) {
    return clean_(a.refEap).localeCompare(clean_(b.refEap));
  });

  var signatureDate = clean_(projectInfo.signatureDate);
  var daysSinceSignature = calculateDaysSinceDate_(signatureDate, new Date());
  var contractDeadlineSummary = clean_(projectInfo.deadlineSummary);

  return {
    status: "ok",
    project: {
      contractId: target,
      projectName: resolveCanonicalProjectName_(target, clean_(projectInfo.projectName), fullProjectMap),
      signatureDate: signatureDate,
      daysSinceSignature: daysSinceSignature,
      contractDeadlineSummary: contractDeadlineSummary,
      summary: summary,
      tasks: tasks,
      events: readEventsByContract_(eventsSheet, target),
      warnings: buildCoordWarningsFromTasks_(tasks)
    }
  };
}

function handleCoordProjectUpdate_(payload) {
  var target = clean_(payload.contractId);
  var editor = clean_(payload.editor) || "COORDENACAO";
  if (!target) {
    throw new Error("contractId obrigatorio para salvar coordenacao.");
  }

  var updates = Array.isArray(payload.updates) ? payload.updates : [];
  var shouldRecalculate = clean_(payload.recalculate).toLowerCase() !== "false";
  var events = normalizeCoordEvents_(payload.events || []);
  var controlSheet = getOrCreateControlSheet_();
  var rows = readControlRows_(controlSheet);
  var nowBr = Utilities.formatDate(new Date(), getTz_(), "dd/MM/yyyy");
  var updatedTasks = 0;
  var i;

  for (i = 0; i < updates.length; i += 1) {
    var update = normalizeCoordTaskUpdate_(updates[i]);
    if (!update || !update.refEap) {
      continue;
    }

    var rowIndex = findControlRowByContractAndRef_(rows, target, update.refEap);
    if (rowIndex < 0) {
      continue;
    }

    var row = rows[rowIndex];
    row[9] = update.status;
    row[10] = update.responsible || row[10];
    row[6] = toNumberOrBlank_(update.prazoDu, 1);
    row[7] = isoToBr_(update.plannedStart);
    row[8] = isoToBr_(update.plannedEnd);
    row[22] = clean_(update.predecessorRef);
    row[23] = clean_(update.relationType) || "FS";
    row[24] = toNumberOrBlank_(update.lagDu, 0);
    row[13] = toNumberOrBlank_(update.percentReal, 0);
    row[15] = update.blocked ? "Sim" : "Nao";
    row[16] = update.blocked ? update.blockReason : "";
    row[17] = nowBr;
    row[18] = clean_(row[18]);
    row[21] = "COORD_CARD";

    if (update.status === STATUS_CONTROL.EM_ANDAMENTO && !clean_(row[11])) {
      row[11] = nowBr;
    }
    if (update.status === STATUS_CONTROL.CONCLUIDA) {
      if (!clean_(row[11])) {
        row[11] = nowBr;
      }
      row[12] = nowBr;
      row[13] = 100;
      row[15] = "Nao";
      row[16] = "";
    }
    if (update.status !== STATUS_CONTROL.CONCLUIDA) {
      row[12] = clean_(row[12]);
    }

    row[14] = calculateDelayDays_(row[8], update.status === STATUS_CONTROL.CONCLUIDA ? row[12] : nowBr);
    updatedTasks += 1;
  }

  var recalcResult = {
    recalculatedDates: 0,
    warnings: []
  };
  if (shouldRecalculate) {
    recalcResult = recalculateProjectScheduleRows_(rows, target);
  }

  applyCanonicalProjectNamesToControlRows_(rows, getProjectMap_());
  writeControlRows_(controlSheet, rows);

  var eventsSheet = getOrCreateEventsSheet_();
  var savedEvents = replaceEventsByContract_(eventsSheet, target, events, editor);

  return json_({
    status: "ok",
    message: "Configuracoes de coordenacao salvas.",
    contractId: target,
    updatedTasks: updatedTasks,
    savedEvents: savedEvents,
    recalculatedDates: recalcResult.recalculatedDates,
    warnings: recalcResult.warnings
  });
}

function normalizeCoordTaskUpdate_(item) {
  var raw = item || {};
  var status = normalizeControlStatus_(raw.status);
  var blocked = clean_(raw.blocked) === "Sim" || clean_(raw.blocked).toUpperCase() === "TRUE";

  return {
    refEap: clean_(raw.refEap),
    status: status,
    responsible: clean_(raw.responsible),
    prazoDu: Math.max(1, toNumber_(raw.prazoDu, 1)),
    predecessorRef: clean_(raw.predecessorRef),
    relationType: normalizeRelationType_(raw.relationType),
    lagDu: toNumber_(raw.lagDu, 0),
    plannedStart: clean_(raw.plannedStart),
    plannedEnd: clean_(raw.plannedEnd),
    percentReal: Math.max(0, Math.min(100, toNumber_(raw.percentReal, 0))),
    blocked: blocked || status === STATUS_CONTROL.BLOQUEADA,
    blockReason: clean_(raw.blockReason)
  };
}

function normalizeRelationType_(value) {
  var relation = clean_(value).toUpperCase();
  if (relation === "SS" || relation === "FF" || relation === "SF") {
    return relation;
  }
  return "FS";
}

function splitDependencyRefs_(value) {
  var text = clean_(value);
  if (!text) {
    return [];
  }
  return text
    .split(/[;,]+/)
    .map(function (item) {
      return clean_(item);
    })
    .filter(function (item) {
      return !!item;
    });
}

function recalculateProjectScheduleRows_(rows, contractId) {
  var target = clean_(contractId);
  var taskMap = {};
  var refs = [];
  var recalculatedDates = 0;
  var warnings = [];
  var i;

  for (i = 0; i < rows.length; i += 1) {
    var row = rows[i];
    if (clean_(row[2]) !== target) {
      continue;
    }
    var ref = clean_(row[3]);
    if (!ref) {
      continue;
    }
    taskMap[ref] = row;
    refs.push(ref);
  }

  if (!refs.length) {
    return {
      recalculatedDates: 0,
      warnings: []
    };
  }

  var orderResult = buildTopologicalOrder_(refs, taskMap);
  var orderedRefs = orderResult.order;
  warnings = warnings.concat(orderResult.warnings);
  var projectStart = inferProjectStartDate_(orderedRefs, taskMap);

  for (i = 0; i < orderedRefs.length; i += 1) {
    var currentRef = orderedRefs[i];
    var currentRow = taskMap[currentRef];
    if (!currentRow) {
      continue;
    }
    if (normalizeControlStatus_(currentRow[9]) === STATUS_CONTROL.CONCLUIDA) {
      continue;
    }

    var duration = Math.max(1, toNumber_(currentRow[6], 1));
    var relationType = normalizeRelationType_(currentRow[23]);
    var lagDu = toNumber_(currentRow[24], 0);
    var dependencies = splitDependencyRefs_(currentRow[22]);
    var currentStart = parseDate_(currentRow[7]) || cloneDate_(projectStart);
    var computedStart = cloneDate_(currentStart);
    var depIndex;

    for (depIndex = 0; depIndex < dependencies.length; depIndex += 1) {
      var depRef = dependencies[depIndex];
      var predecessorRow = taskMap[depRef];
      if (!predecessorRow) {
        warnings.push(currentRef + ": predecessora nao encontrada (" + depRef + ").");
        continue;
      }

      var predecessorStart = parseDate_(predecessorRow[7]);
      var predecessorEnd = parseDate_(predecessorRow[8]);
      var candidateStart = computeStartByRelation_(predecessorStart, predecessorEnd, relationType, lagDu, duration);

      if (!candidateStart) {
        warnings.push(currentRef + ": datas insuficientes na predecessora " + depRef + ".");
        continue;
      }

      if (candidateStart.getTime() > computedStart.getTime()) {
        computedStart = cloneDate_(candidateStart);
      }
    }

    computedStart = nextBusinessDay_(computedStart);
    var computedEnd = addBusinessDays_(computedStart, duration - 1);
    var previousStart = toDateKey_(currentRow[7]);
    var previousEnd = toDateKey_(currentRow[8]);
    var nextStart = toDateKey_(computedStart);
    var nextEnd = toDateKey_(computedEnd);

    if (previousStart !== nextStart || previousEnd !== nextEnd) {
      recalculatedDates += 1;
    }

    currentRow[7] = Utilities.formatDate(computedStart, getTz_(), "dd/MM/yyyy");
    currentRow[8] = Utilities.formatDate(computedEnd, getTz_(), "dd/MM/yyyy");
  }

  warnings = warnings.concat(buildCoordWarningsFromRows_(rows, target));
  warnings = dedupeTextList_(warnings).slice(0, 20);

  return {
    recalculatedDates: recalculatedDates,
    warnings: warnings
  };
}

function buildTopologicalOrder_(refs, taskMap) {
  var adjacency = {};
  var indegree = {};
  var queue = [];
  var order = [];
  var warnings = [];
  var i;

  for (i = 0; i < refs.length; i += 1) {
    adjacency[refs[i]] = [];
    indegree[refs[i]] = 0;
  }

  for (i = 0; i < refs.length; i += 1) {
    var ref = refs[i];
    var row = taskMap[ref];
    var deps = splitDependencyRefs_(row[22]);
    var j;
    for (j = 0; j < deps.length; j += 1) {
      var depRef = deps[j];
      if (!taskMap[depRef]) {
        continue;
      }
      adjacency[depRef].push(ref);
      indegree[ref] += 1;
    }
  }

  for (i = 0; i < refs.length; i += 1) {
    if (indegree[refs[i]] === 0) {
      queue.push(refs[i]);
    }
  }
  queue.sort();

  while (queue.length) {
    var current = queue.shift();
    order.push(current);
    var edges = adjacency[current] || [];
    var e;
    for (e = 0; e < edges.length; e += 1) {
      var next = edges[e];
      indegree[next] -= 1;
      if (indegree[next] === 0) {
        queue.push(next);
      }
    }
    queue.sort();
  }

  if (order.length < refs.length) {
    warnings.push("Dependencias com ciclo detectado. Foi aplicada ordenacao por REF EAP como fallback.");
    refs
      .slice()
      .sort()
      .forEach(function (ref) {
        if (order.indexOf(ref) < 0) {
          order.push(ref);
        }
      });
  }

  return {
    order: order,
    warnings: warnings
  };
}

function inferProjectStartDate_(orderedRefs, taskMap) {
  var minDate = null;
  var i;

  for (i = 0; i < orderedRefs.length; i += 1) {
    var row = taskMap[orderedRefs[i]];
    if (!row) {
      continue;
    }
    var start = parseDate_(row[7]);
    if (!start) {
      continue;
    }
    if (!minDate || start.getTime() < minDate.getTime()) {
      minDate = cloneDate_(start);
    }
  }

  if (!minDate) {
    minDate = new Date();
  }
  minDate.setHours(0, 0, 0, 0);
  return minDate;
}

function computeStartByRelation_(predStart, predEnd, relationType, lagDu, duration) {
  var lag = toNumber_(lagDu, 0);
  var durationDays = Math.max(1, toNumber_(duration, 1));
  var rel = normalizeRelationType_(relationType);
  var base;
  var end;

  if (rel === "SS") {
    if (!predStart) {
      return null;
    }
    return addBusinessDays_(predStart, lag);
  }

  if (rel === "FF") {
    if (!predEnd) {
      return null;
    }
    end = addBusinessDays_(predEnd, lag);
    return subtractBusinessDays_(end, durationDays - 1);
  }

  if (rel === "SF") {
    if (!predStart) {
      return null;
    }
    end = addBusinessDays_(predStart, lag);
    return subtractBusinessDays_(end, durationDays - 1);
  }

  base = predEnd || predStart;
  if (!base) {
    return null;
  }
  return addBusinessDays_(base, lag + 1);
}

function buildCoordWarningsFromRows_(rows, contractId) {
  var target = clean_(contractId);
  var tasks = [];
  var i;

  for (i = 0; i < rows.length; i += 1) {
    var row = rows[i];
    if (clean_(row[2]) !== target) {
      continue;
    }
    tasks.push({
      refEap: clean_(row[3]),
      status: clean_(row[9]),
      plannedStart: toIsoDateOrBlank_(row[7]),
      plannedEnd: toIsoDateOrBlank_(row[8]),
      predecessorRef: clean_(row[22]),
      relationType: clean_(row[23]) || "FS",
      lagDu: toNumberOrBlank_(row[24], 0),
      prazoDu: toNumberOrBlank_(row[6], 1)
    });
  }

  return buildCoordWarningsFromTasks_(tasks);
}

function buildCoordWarningsFromTasks_(tasks) {
  var list = Array.isArray(tasks) ? tasks : [];
  var map = {};
  var warnings = [];
  var i;

  for (i = 0; i < list.length; i += 1) {
    var task = list[i];
    map[clean_(task.refEap)] = task;
  }

  for (i = 0; i < list.length; i += 1) {
    var current = list[i];
    var ref = clean_(current.refEap);
    if (!ref) {
      continue;
    }

    var deps = splitDependencyRefs_(current.predecessorRef);
    var d;
    for (d = 0; d < deps.length; d += 1) {
      var depRef = deps[d];
      if (depRef === ref) {
        warnings.push(ref + ": dependencia circular (ela mesma).");
        continue;
      }
      if (!map[depRef]) {
        warnings.push(ref + ": predecessora nao encontrada (" + depRef + ").");
        continue;
      }

      var pred = map[depRef];
      var predStart = parseDate_(pred.plannedStart);
      var predEnd = parseDate_(pred.plannedEnd);
      var currentStart = parseDate_(current.plannedStart);
      var requiredStart = computeStartByRelation_(
        predStart,
        predEnd,
        current.relationType,
        current.lagDu,
        current.prazoDu
      );

      if (requiredStart && currentStart && currentStart.getTime() < requiredStart.getTime()) {
        warnings.push(
          ref +
            ": inicio planejado (" +
            Utilities.formatDate(currentStart, getTz_(), "dd/MM/yyyy") +
            ") viola dependencia " +
            depRef +
            " (" +
            normalizeRelationType_(current.relationType) +
            ")."
        );
      }
    }
  }

  return dedupeTextList_(warnings);
}

function dedupeTextList_(items) {
  var source = Array.isArray(items) ? items : [];
  var map = {};
  var result = [];
  var i;

  for (i = 0; i < source.length; i += 1) {
    var text = clean_(source[i]);
    if (!text || map[text]) {
      continue;
    }
    map[text] = true;
    result.push(text);
  }

  return result;
}

function normalizeCoordEvents_(events) {
  var source = Array.isArray(events) ? events : [];
  var result = [];
  var i;

  for (i = 0; i < source.length; i += 1) {
    var item = source[i] || {};
    var normalized = {
      eventDate: clean_(item.eventDate),
      eventType: clean_(item.eventType),
      title: clean_(item.title),
      responsible: clean_(item.responsible),
      status: clean_(item.status) || "Planejado",
      observation: clean_(item.observation)
    };

    if (!normalized.eventDate && !normalized.eventType && !normalized.title && !normalized.observation) {
      continue;
    }
    result.push(normalized);
  }

  result.sort(function (a, b) {
    var left = clean_(a.eventDate);
    var right = clean_(b.eventDate);
    if (left && right) {
      return left.localeCompare(right);
    }
    if (left) {
      return -1;
    }
    if (right) {
      return 1;
    }
    return clean_(a.title).localeCompare(clean_(b.title));
  });

  return result;
}

function findControlRowByContractAndRef_(rows, contractId, refEap) {
  var targetContract = clean_(contractId);
  var targetRef = clean_(refEap);
  var i;

  for (i = 0; i < rows.length; i += 1) {
    if (clean_(rows[i][2]) !== targetContract) {
      continue;
    }
    if (clean_(rows[i][3]) === targetRef) {
      return i;
    }
  }
  return -1;
}

function buildHistoryPayload_(professional, limit) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXEC_SHEET_NAME);
  var maxItems = Number(limit);
  var targetProfessional = normalizeName_(professional);
  var grouped = {};
  var order = [];
  var history = [];
  var i;

  if (isNaN(maxItems) || maxItems <= 0) {
    maxItems = 10;
  }
  maxItems = Math.min(maxItems, 30);

  if (!sheet || sheet.getLastRow() < EXEC_DATA_START_ROW) {
    return { status: "ok", history: [] };
  }
  ensureExecutionSheetReady_(sheet);

  var values = sheet.getRange(EXEC_DATA_START_ROW, 1, sheet.getLastRow() - EXEC_DATA_START_ROW + 1, EXEC_COLUMNS).getValues();

  for (i = values.length - 1; i >= 0; i -= 1) {
    var row = values[i];
    var rowProfessional = clean_(row[6]);
    if (targetProfessional && normalizeName_(rowProfessional) !== targetProfessional) {
      continue;
    }

    var formId = clean_(row[16]);
    if (!formId) {
      continue;
    }

    if (!grouped[formId]) {
      grouped[formId] = {
        formId: formId,
        professional: rowProfessional,
        dateBr: toBrDateLabel_(row[13]),
        submittedAtBr: toBrDateLabel_(row[15]),
        projects: []
      };
      order.push(formId);
    }

    grouped[formId].projects.unshift({
      projectCode: clean_(row[1]),
      projectName: clean_(row[0]),
      refEap: clean_(row[5]),
      taskText: clean_(row[3]),
      statusText: clean_(row[8])
    });
  }

  for (i = 0; i < order.length && history.length < maxItems; i += 1) {
    history.push(grouped[order[i]]);
  }

  return {
    status: "ok",
    history: history
  };
}

function normalizeControlStatus_(value) {
  var text = normalizeText_(value);
  if (text.indexOf("conclu") >= 0) {
    return STATUS_CONTROL.CONCLUIDA;
  }
  if (text.indexOf("andamento") >= 0 || text.indexOf("continua") >= 0) {
    return STATUS_CONTROL.EM_ANDAMENTO;
  }
  if (text.indexOf("bloque") >= 0 || text.indexOf("paus") >= 0 || text.indexOf("aguard") >= 0) {
    return STATUS_CONTROL.BLOQUEADA;
  }
  return STATUS_CONTROL.NAO_INICIADA;
}

function toIsoDateOrBlank_(value) {
  var parsed = parseDate_(value);
  if (!parsed) {
    return "";
  }
  return Utilities.formatDate(parsed, getTz_(), "yyyy-MM-dd");
}

function isoToBr_(value) {
  var parsed = parseDate_(value);
  if (!parsed) {
    return "";
  }
  return Utilities.formatDate(parsed, getTz_(), "dd/MM/yyyy");
}

function toBrDateLabel_(value) {
  var parsed = parseDate_(value);
  if (!parsed) {
    return clean_(value);
  }
  return Utilities.formatDate(parsed, getTz_(), "dd/MM/yyyy");
}

function buildRefOptionsByProject_() {
  var optionsByProject = {};
  var projectMap = getProjectMap_();
  var code;

  for (code in projectMap) {
    if (Object.prototype.hasOwnProperty.call(projectMap, code)) {
      optionsByProject[projectMap[code].contractId] = [];
    }
  }

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONTROL_SHEET_NAME);
  if (!sheet) {
    optionsByProject["ADMIN-INTERNO"] = optionsByProject["ADMIN-INTERNO"] || [];
    optionsByProject["ADMIN-INTERNO"].push({
      value: "ADMININTERNO-WBS-INT_1",
      label: "ADMININTERNO-WBS-INT_1 | Administrativo/Interno | Atividades administrativas internas",
      status: STATUS_CONTROL.EM_ANDAMENTO
    });
    return optionsByProject;
  }

  var rows = readControlRows_(sheet);
  var i;
  for (i = 0; i < rows.length; i += 1) {
    var row = rows[i];
    var contractId = clean_(row[2]);
    var refEap = clean_(row[3]);
    var phase = clean_(row[4]);
    var task = clean_(row[5]);
    var status = clean_(row[9]);

    if (!contractId || !refEap) {
      continue;
    }
    if (status === STATUS_CONTROL.CONCLUIDA) {
      continue;
    }

    if (!optionsByProject[contractId]) {
      optionsByProject[contractId] = [];
    }

    optionsByProject[contractId].push({
      value: refEap,
      label: refEap + " | " + phase + " | " + task,
      status: status
    });
  }

  for (code in optionsByProject) {
    if (Object.prototype.hasOwnProperty.call(optionsByProject, code)) {
      optionsByProject[code].sort(function (a, b) {
        var aOrder = controlStatusOrder_(a.status);
        var bOrder = controlStatusOrder_(b.status);
        if (aOrder !== bOrder) {
          return aOrder - bOrder;
        }
        return String(a.label).localeCompare(String(b.label));
      });
    }
  }

  return optionsByProject;
}

function controlStatusOrder_(status) {
  if (status === STATUS_CONTROL.EM_ANDAMENTO) {
    return 1;
  }
  if (status === STATUS_CONTROL.BLOQUEADA) {
    return 2;
  }
  if (status === STATUS_CONTROL.NAO_INICIADA) {
    return 3;
  }
  if (status === STATUS_CONTROL.CONCLUIDA) {
    return 4;
  }
  return 5;
}

function buildDailySubmissionSummary_(input, records, controlRows) {
  var byContractId = {};
  var contractOrder = [];
  var i;

  for (i = 0; i < records.length; i += 1) {
    var record = records[i];
    var contractId = clean_(record.contractId);

    if (!byContractId[contractId]) {
      byContractId[contractId] = {
        projectCode: clean_(record.projectCode) || contractId,
        contractId: contractId,
        projectName: clean_(record.projectName) || contractId,
        todayItems: [],
        snapshot: {
          total: 0,
          naoIniciada: 0,
          emAndamento: 0,
          concluida: 0,
          bloqueada: 0,
          pendentes: 0
        },
        nextTasks: []
      };
      contractOrder.push(contractId);
    }

    byContractId[contractId].todayItems.push({
      refEap: clean_(record.refEap),
      taskText: clean_(record.taskText),
      statusText: clean_(record.statusText),
      isBlocked: !!record.isBlocked,
      blockReason: clean_(record.blockReason)
    });
  }

  for (i = 0; i < controlRows.length; i += 1) {
    var row = controlRows[i];
    var rowContractId = clean_(row[2]);
    var projectSummary = byContractId[rowContractId];
    if (!projectSummary) {
      continue;
    }

    var status = clean_(row[9]);
    projectSummary.snapshot.total += 1;
    if (status === STATUS_CONTROL.NAO_INICIADA) {
      projectSummary.snapshot.naoIniciada += 1;
    } else if (status === STATUS_CONTROL.EM_ANDAMENTO) {
      projectSummary.snapshot.emAndamento += 1;
    } else if (status === STATUS_CONTROL.CONCLUIDA) {
      projectSummary.snapshot.concluida += 1;
    } else if (status === STATUS_CONTROL.BLOQUEADA) {
      projectSummary.snapshot.bloqueada += 1;
    }

    if (status !== STATUS_CONTROL.CONCLUIDA) {
      projectSummary.snapshot.pendentes += 1;
      projectSummary.nextTasks.push({
        refEap: clean_(row[3]),
        phase: clean_(row[4]),
        task: clean_(row[5]),
        status: status
      });
    }
  }

  var projects = [];
  for (i = 0; i < contractOrder.length; i += 1) {
    var key = contractOrder[i];
    var item = byContractId[key];

    item.nextTasks.sort(function (a, b) {
      var aOrder = nextTaskPriority_(a.status);
      var bOrder = nextTaskPriority_(b.status);
      if (aOrder !== bOrder) {
        return aOrder - bOrder;
      }
      return String(a.refEap).localeCompare(String(b.refEap));
    });
    item.nextTasks = item.nextTasks.slice(0, 3);

    projects.push(item);
  }

  return {
    professional: input.professional,
    dateBr: input.dateBr,
    generatedAtBr: Utilities.formatDate(new Date(), getTz_(), "dd/MM/yyyy HH:mm"),
    projects: projects
  };
}

function nextTaskPriority_(status) {
  if (status === STATUS_CONTROL.BLOQUEADA) {
    return 1;
  }
  if (status === STATUS_CONTROL.EM_ANDAMENTO) {
    return 2;
  }
  if (status === STATUS_CONTROL.NAO_INICIADA) {
    return 3;
  }
  if (status === STATUS_CONTROL.CONCLUIDA) {
    return 4;
  }
  return 5;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function jsonOrJsonp_(obj, callback) {
  var cb = clean_(callback);
  if (cb && /^[A-Za-z0-9_\\.]+$/.test(cb)) {
    return ContentService.createTextOutput(cb + "(" + JSON.stringify(obj) + ");").setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return json_(obj);
}

// -------------------------
// Funcoes manuais (rodar no editor Apps Script)
// -------------------------

function setupControleEapAtual() {
  var context = prepareControlContext_();
  getOrCreateEventsSheet_();
  return {
    status: "ok",
    sheet: CONTROL_SHEET_NAME,
    rows: context.rows.length,
    message: "Controle pronto para uso."
  };
}

function resetBancoDoZeroEapPadrao() {
  var executionSheet = getExecutionSheet_();
  var eapSheet = getOrCreateEapSheet_();
  var controlSheet = getOrCreateControlSheet_();
  var eventsSheet = getOrCreateEventsSheet_();

  var clearedExecutionRows = clearExecutionData_(executionSheet);
  var clearedEventsRows = clearEventsData_(eventsSheet);
  var eapRows = buildBaselineRowsFromStandardTemplate_();
  writeEapRows_(eapSheet, eapRows);

  var controlRows = buildControlRowsFromStandardTemplate_();
  writeControlRows_(controlSheet, controlRows);

  return {
    status: "ok",
    message: "Reset completo concluido com a EAP padrao.",
    templateVersion: "EAP_PADRAO_V1",
    executionRowsCleared: clearedExecutionRows,
    eventsRowsCleared: clearedEventsRows,
    eapRowsCreated: eapRows.length,
    controlRowsCreated: controlRows.length
  };
}

function sincronizarControleComExecucaoHistorica() {
  var executionSheet = getExecutionSheet_();
  var lastRow = executionSheet.getLastRow();

  if (lastRow < EXEC_DATA_START_ROW) {
    return {
      status: "ok",
      message: "Sem dados historicos na EXECUCAO.",
      rowsProcessed: 0
    };
  }

  var controlRows = rebuildControlFromExecution_(executionSheet);
  var processed = lastRow - EXEC_DATA_START_ROW + 1;

  return {
    status: "ok",
    message: "Sincronizacao concluida.",
    rowsProcessed: processed,
    controlRows: controlRows.length
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

function reconciliarControleComEapPadrao() {
  var controlSheet = getOrCreateControlSheet_();
  var currentRows = readControlRows_(controlSheet);
  var baselineRows = buildControlRowsFromStandardTemplate_();
  var mergedRows = mergePreservedControlRows_(baselineRows, currentRows);
  var projectMap = getProjectMap_();
  applyCanonicalProjectNamesToControlRows_(mergedRows, projectMap);
  writeControlRows_(controlSheet, mergedRows);

  var currentByContract = summarizeRowsByContract_(currentRows);
  var mergedByContract = summarizeRowsByContract_(mergedRows);
  var contracts = Object.keys(mergedByContract).sort(compareContractId_);
  var coverage = [];
  var i;

  for (i = 0; i < contracts.length; i += 1) {
    var contractId = contracts[i];
    var before = currentByContract[contractId] || 0;
    var after = mergedByContract[contractId] || 0;
    coverage.push({
      contractId: contractId,
      beforeRows: before,
      afterRows: after,
      addedRows: Math.max(after - before, 0)
    });
  }

  return {
    status: "ok",
    message: "Reconciliação concluída sem perda de dados existentes.",
    beforeRows: currentRows.length,
    afterRows: mergedRows.length,
    addedRows: Math.max(mergedRows.length - currentRows.length, 0),
    coverage: coverage
  };
}

function auditoriaCoberturaControle() {
  var rows = readControlRows_(getOrCreateControlSheet_());
  var baselineRows = buildControlRowsFromStandardTemplate_();
  var currentByContract = summarizeRowsByContract_(rows);
  var expectedByContract = summarizeRowsByContract_(baselineRows);
  var contracts = Object.keys(expectedByContract).sort(compareContractId_);
  var coverage = [];
  var i;

  for (i = 0; i < contracts.length; i += 1) {
    var contractId = contracts[i];
    var expected = expectedByContract[contractId] || 0;
    var current = currentByContract[contractId] || 0;
    coverage.push({
      contractId: contractId,
      currentRows: current,
      expectedRows: expected,
      missingRows: Math.max(expected - current, 0)
    });
  }

  return {
    status: "ok",
    coverage: coverage
  };
}

function summarizeRowsByContract_(rows) {
  var source = Array.isArray(rows) ? rows : [];
  var summary = {};
  var i;
  for (i = 0; i < source.length; i += 1) {
    var contractId = clean_(source[i][2]);
    if (!contractId) {
      continue;
    }
    summary[contractId] = (summary[contractId] || 0) + 1;
  }
  return summary;
}
