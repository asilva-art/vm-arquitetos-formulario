(function () {
  "use strict";

  var cfg = window.VM_FORM_CONFIG || {};
  var localProjectsFromConfig = Array.isArray(cfg.projetos) ? cfg.projetos : [];
  var formEl = document.getElementById("daily-form");
  var professionalEl = document.getElementById("professional");
  var dateEl = document.getElementById("work-date");
  var projectListEl = document.getElementById("project-list");
  var projectSectionsEl = document.getElementById("project-sections");
  var feedbackEl = document.getElementById("feedback");
  var portfolioFilterEl = document.getElementById("portfolio-filter");
  var portfolioCutoffEl = document.getElementById("portfolio-cutoff");
  var portfolioRefreshEl = document.getElementById("portfolio-refresh");
  var portfolioToggleEl = document.getElementById("portfolio-toggle");
  var portfolioContentEl = document.getElementById("portfolio-content");
  var portfolioSummaryEl = document.getElementById("portfolio-summary");
  var portfolioCardsEl = document.getElementById("portfolio-cards");
  var portfolioBlocksEl = document.getElementById("portfolio-blocks");
  var portfolioMeetingEl = document.getElementById("portfolio-meeting");
  var meetingAgendaEl = document.getElementById("meeting-agenda");
  var meetingCopyEl = document.getElementById("meeting-copy");
  var guideToggleEl = document.getElementById("guide-toggle");
  var guideStepsEl = document.getElementById("guide-steps");
  var summaryToggleEl = document.getElementById("summary-toggle");
  var dailySummaryEl = document.getElementById("daily-summary");
  var submitButtonEl = document.getElementById("submit-button");
  var lastSubmittedSummary = null;
  var lastMeetingSummaryText = "";
  var portfolioDataset = [];
  var ppmSnapshotMap = {};

  var projectMap = {};
  var projectList = [];
  var refOptionsByProject = normalizeRefOptions(cfg.refOptions || {});
  rebuildProjectCache_(localProjectsFromConfig);

  var LOCAL_EAP_TEMPLATE = [
    { wbs: "1.1.1", phase: "Estudos Preliminares", task: "Elaboracao do Briefing Documentado" },
    { wbs: "1.1.2", phase: "Estudos Preliminares", task: "Geracao da Analise do Terreno" },
    { wbs: "1.1.3", phase: "Estudos Preliminares", task: "Levantamento da Analise Legal" },
    { wbs: "2.1.1", phase: "Anteprojeto - Etapa I (Implantacao e Layout)", task: "Estudo de Implantacao" },
    { wbs: "2.1.2", phase: "Anteprojeto - Etapa I (Implantacao e Layout)", task: "Planta de Layout" },
    { wbs: "2.2.1", phase: "Anteprojeto - Etapa I (Implantacao e Layout)", task: "Execucao de Revisoes (ate 2 rodadas)" },
    { wbs: "2.2.2", phase: "Anteprojeto - Etapa I (Implantacao e Layout)", task: "Aprovacao do Layout" },
    { wbs: "3.1.1", phase: "Anteprojeto - Etapa II (Volumetria e Materiais Externos)", task: "Modelagem de Imagens 3D Externas" },
    { wbs: "3.1.2", phase: "Anteprojeto - Etapa II (Volumetria e Materiais Externos)", task: "Pre-selecao de Materiais Externos" },
    { wbs: "3.1.3", phase: "Anteprojeto - Etapa II (Volumetria e Materiais Externos)", task: "Estudo de Alternativa de Volumetria (se solicitada)" },
    { wbs: "3.2.1", phase: "Anteprojeto - Etapa II (Volumetria e Materiais Externos)", task: "Execucao de Revisoes (ate 2 rodadas)" },
    { wbs: "3.2.2", phase: "Anteprojeto - Etapa II (Volumetria e Materiais Externos)", task: "Aprovacao da Volumetria e Materiais Externos" },
    { wbs: "4.1.1", phase: "Projeto Legal", task: "Elaboracao do Projeto para Aprovacao na Prefeitura" },
    { wbs: "4.2.1", phase: "Projeto Legal", task: "Acompanhamento do Protocolo na Prefeitura" },
    { wbs: "4.2.2", phase: "Projeto Legal", task: "Projeto Legal Aprovado" },
    { wbs: "5.1.1", phase: "Anteprojeto - Etapa III (Interiores)", task: "Modelagem de Imagens 3D Internas" },
    { wbs: "5.1.2", phase: "Anteprojeto - Etapa III (Interiores)", task: "Pre-selecao de Materiais Internos" },
    { wbs: "5.1.3", phase: "Anteprojeto - Etapa III (Interiores)", task: "Estudo de Alternativa por Ambiente (se solicitada)" },
    { wbs: "5.2.1", phase: "Anteprojeto - Etapa III (Interiores)", task: "Execucao de Revisoes (ate 2 rodadas)" },
    { wbs: "5.2.2", phase: "Anteprojeto - Etapa III (Interiores)", task: "Aprovacao dos Interiores" },
    { wbs: "6.1.1", phase: "Projeto Pre-Executivo", task: "Planta Layout Cotada" },
    { wbs: "6.1.2", phase: "Projeto Pre-Executivo", task: "Detalhamento de Pedras" },
    { wbs: "6.1.3", phase: "Projeto Pre-Executivo", task: "Quantitativo de Revestimentos" },
    { wbs: "6.1.4", phase: "Projeto Pre-Executivo", task: "Lista de Loucas e Metais" },
    { wbs: "7.1.1", phase: "Definicao de Materiais e Revisao 3D", task: "Acompanhamento em Lojas (ate 7 visitas)" },
    { wbs: "7.2.1", phase: "Definicao de Materiais e Revisao 3D", task: "Execucao da Revisao Final 3D" },
    { wbs: "7.2.2", phase: "Definicao de Materiais e Revisao 3D", task: "Aprovacao Final 3D" },
    { wbs: "8.1.1", phase: "Projeto Executivo", task: "Planta de Situacao" },
    { wbs: "8.1.2", phase: "Projeto Executivo", task: "Planta de Implantacao" },
    { wbs: "8.1.3", phase: "Projeto Executivo", task: "Planta de Cobertura" },
    { wbs: "8.1.4", phase: "Projeto Executivo", task: "Cortes" },
    { wbs: "8.1.5", phase: "Projeto Executivo", task: "Fachadas" },
    { wbs: "8.2.1", phase: "Projeto Executivo", task: "Pontos Eletricos" },
    { wbs: "8.2.2", phase: "Projeto Executivo", task: "Pontos Hidraulicos" },
    { wbs: "8.2.3", phase: "Projeto Executivo", task: "Planta de Forro" },
    { wbs: "8.2.4", phase: "Projeto Executivo", task: "Planta de Piso" },
    { wbs: "8.2.5", phase: "Projeto Executivo", task: "Projeto Luminotecnico" },
    { wbs: "8.3.1", phase: "Projeto Executivo", task: "Detalhamento de Escadas e Rampas" },
    { wbs: "8.3.2", phase: "Projeto Executivo", task: "Detalhamento de Esquadrias" },
    { wbs: "8.3.3", phase: "Projeto Executivo", task: "Detalhes Construtivos" },
    { wbs: "8.3.4", phase: "Projeto Executivo", task: "Detalhamento de Gradil" },
    { wbs: "8.3.5", phase: "Projeto Executivo", task: "Detalhamento de Areas Molhadas" },
    { wbs: "8.3.6", phase: "Projeto Executivo", task: "Detalhamento de Guarda-corpo" },
    { wbs: "9.1.1", phase: "Detalhamentos Complementares", task: "Detalhamento de Vidros e Espelhos" },
    { wbs: "9.1.2", phase: "Detalhamentos Complementares", task: "Design e Detalhamento de Moveis" },
    { wbs: "9.1.3", phase: "Detalhamentos Complementares", task: "Detalhamentos Complementares Especificos" },
    { wbs: "10.1.1", phase: "Acompanhamento de Obra e Producao", task: "Realizacao de Visitas Tecnicas em Obra (2 visitas)" },
    { wbs: "10.2.1", phase: "Acompanhamento de Obra e Producao", task: "Realizacao de Acompanhamento de Producao (1 visita)" }
  ];

  var LOCAL_EAP_NETWORK = {
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

  var DEFAULT_SIGNATURE_DATE = "2025-12-12";
  var PROJECT_SIGNATURE_DATE = Object.assign(
    {
      "CT-117": "2025-12-12",
      "CT-119": "2025-12-16",
      "CT-120": "2025-12-17",
      "CT-121": "2025-12-18",
      "CT-122": "2026-02-12",
      "CT-123": "2025-01-08",
      "CT-124": "2025-01-27",
      "CT-125": "2025-02-12",
      "CT-126": "2025-02-14",
      "CT-127": "2025-03-12",
      "CT-128": "2025-04-04",
      "CT-129": "2025-04-15",
      "CT-130": "2025-05-15",
      "CT-131": "2025-05-15",
      "CT-132": "2025-05-23",
      "CT-133": "2025-09-02",
      "CT-134": "2025-06-26",
      "CT-135": "2025-06-04",
      "CT-136": "2025-11-26"
    },
    cfg.datasAssinatura || {}
  );

  init();

  function init() {
    if (!formEl) {
      return;
    }

    populateProfessionals();
    populateProjects();
    setDefaultDate();
    setDefaultPortfolioCutoffDate_();
    renderProjectSections([]);

    projectListEl.addEventListener("change", handleProjectSelectionChange);
    formEl.addEventListener("submit", handleSubmit);
    if (portfolioFilterEl) {
      portfolioFilterEl.addEventListener("change", renderPortfolioOverview);
    }
    if (portfolioCutoffEl) {
      portfolioCutoffEl.addEventListener("change", renderPortfolioOverview);
    }
    if (portfolioRefreshEl) {
      portfolioRefreshEl.addEventListener("click", function () {
        loadRefOptionsFromServer();
        loadPpmSnapshotFromServer();
      });
    }
    if (meetingCopyEl) {
      meetingCopyEl.addEventListener("click", handleMeetingCopy);
      meetingCopyEl.disabled = true;
    }
    if (portfolioToggleEl && portfolioContentEl) {
      portfolioToggleEl.addEventListener("click", handlePortfolioToggle);
    }
    if (guideToggleEl && guideStepsEl) {
      guideToggleEl.addEventListener("click", handleGuideToggle);
    }
    if (summaryToggleEl) {
      summaryToggleEl.addEventListener("click", handleSummaryToggle);
    }

    loadRefOptionsFromServer();
    loadPpmSnapshotFromServer();
    portfolioDataset = buildPortfolioDataset_();
    renderPortfolioOverview();
    clearSummaryUI();
  }

  function handleGuideToggle() {
    if (!guideToggleEl || !guideStepsEl) {
      return;
    }

    var isHidden = guideStepsEl.classList.toggle("hidden");
    guideToggleEl.setAttribute("aria-expanded", String(!isHidden));
    guideToggleEl.textContent = isHidden ? "Expandir explicacoes" : "Recolher explicacoes";
  }

  function handlePortfolioToggle() {
    if (!portfolioToggleEl || !portfolioContentEl) {
      return;
    }

    var isHidden = portfolioContentEl.classList.toggle("hidden");
    portfolioToggleEl.setAttribute("aria-expanded", String(!isHidden));
    portfolioToggleEl.textContent = isHidden ? "Expandir status visual" : "Recolher status visual";
  }

  function handleMeetingCopy() {
    if (!lastMeetingSummaryText) {
      showFeedback("warn", "Nao ha pauta de reuniao para copiar.");
      return;
    }

    copyTextToClipboard_(lastMeetingSummaryText)
      .then(function () {
        showFeedback("ok", "Resumo da reuniao copiado. Cole em COMUNICACOES.");
      })
      .catch(function () {
        showFeedback("error", "Nao foi possivel copiar o resumo automaticamente.");
      });
  }

  function populateProfessionals() {
    var list = Array.isArray(cfg.profissionais) && cfg.profissionais.length
      ? cfg.profissionais
      : ["Alencar", "Breno", "Ester", "Fran", "Lucas", "Lud", "Luiza", "Uly", "Vivi"];

    // Garante que a lista final nao duplica opcoes entre fallback HTML e JS.
    professionalEl.innerHTML = '<option value="">Selecione...</option>';

    list.forEach(function (name) {
      var option = document.createElement("option");
      option.value = name;
      option.textContent = name;
      professionalEl.appendChild(option);
    });
  }

  function populateProjects() {
    var html = (projectList || [])
      .map(function (project) {
        var safeId = "project-" + sanitizeKey(project.code);
        return [
          '<label class="check-item" for="' + safeId + '">',
          '<input id="' + safeId + '" type="checkbox" name="projects" value="' + escapeAttr(project.code) + '" />',
          "<span>" + escapeHtml(project.label) + "</span>",
          "</label>"
        ].join("");
      })
      .join("");

    projectListEl.innerHTML = html || "<p>Nenhum projeto configurado.</p>";
  }

  function rebuildProjectCache_(projects) {
    var list = Array.isArray(projects) ? projects : [];
    var normalized = [];
    var map = {};

    list.forEach(function (item) {
      var code = cleanText(item && item.code);
      if (!code) {
        return;
      }
      var label = formatProjectLabel_(code, cleanText(item && item.label));
      var project = { code: code, label: label };
      normalized.push(project);
      map[code] = project;
    });

    projectList = normalized;
    projectMap = map;
  }

  function formatProjectLabel_(code, label) {
    if (code === "ADMIN-INTERNO") {
      return label || "Administrativo/Interno";
    }
    if (!label) {
      return code;
    }
    if (label.indexOf(code) === 0) {
      return label;
    }
    return code + " | " + label;
  }

  function handleProjectSelectionChange() {
    renderProjectSections(getSelectedProjectCodes());
    renderPortfolioOverview();
  }

  function renderProjectSections(selectedCodes) {
    projectSectionsEl.innerHTML = "";

    selectedCodes.forEach(function (projectCode) {
      var project = projectMap[projectCode] || { code: projectCode, label: projectCode };
      var key = sanitizeKey(project.code);

      var sectionEl = document.createElement("section");
      sectionEl.className = "project-card";
      sectionEl.dataset.projectCode = project.code;
      sectionEl.innerHTML =
        "<h3>Secao do projeto: " + escapeHtml(project.label) + "</h3>" +
        '<div class="field">' +
        '<label for="ref-' + key + '">Campo 4A - Tarefa EAP (REF EAP)</label>' +
        '<select id="ref-' + key + '" required>' + buildRefOptionsHtml(project.code) + "</select>" +
        '<p class="hint">Obrigatorio para rastreabilidade.</p>' +
        "</div>" +
        '<div class="field">' +
        '<label for="task-ref-' + key + '">Campo 4B - Tarefa executada hoje (EAP)</label>' +
        '<select id="task-ref-' + key + '">' + buildRefOptionsHtml(project.code) + "</select>" +
        '<p class="hint">Opcional se preencher tarefa livre abaixo.</p>' +
        "</div>" +
        '<div class="field">' +
        '<label for="task-free-' + key + '">Campo 4B complementar - Tarefa livre (max. 30)</label>' +
        '<input id="task-free-' + key + '" type="text" maxlength="30" placeholder="Ex: ajuste de prancha" />' +
        '<p class="hint">Preencha se a tarefa do dia nao estiver na EAP.</p>' +
        "</div>" +
        '<div class="field">' +
        '<p class="field-title">Campo 5 - Status</p>' +
        '<div class="option-list">' + buildStatusOptionsHtml(key) + "</div>" +
        "</div>" +
        '<div class="field">' +
        '<p class="field-title">Campo 6 - Ha bloqueio neste projeto?</p>' +
        '<div class="option-list">' + buildBlockOptionsHtml(key) + "</div>" +
        "</div>" +
        '<div class="field hidden" id="block-wrap-' + key + '">' +
        '<label for="block-desc-' + key + '">Campo 7 - Descricao do bloqueio</label>' +
        '<input id="block-desc-' + key + '" type="text" maxlength="180" placeholder="Descreva em 1 ou 2 linhas" />' +
        "</div>" +
        '<div class="field">' +
        '<label for="obs-' + key + '">Campo 8 - Observacao livre (opcional)</label>' +
        '<textarea id="obs-' + key + '" maxlength="500" placeholder="Maximo de 500 caracteres"></textarea>' +
        '<p class="hint">Opcional. Maximo de 500 caracteres.</p>' +
        "</div>";

      projectSectionsEl.appendChild(sectionEl);
      wireBlockToggle(key);
    });
  }

  function buildRefOptionsHtml(projectCode) {
    var options = getRefOptionsForProject(projectCode);
    var html = ['<option value="">Selecione...</option>'];
    var i;

    for (i = 0; i < options.length; i += 1) {
      html.push(
        '<option value="' + escapeAttr(options[i].value) + '">' + escapeHtml(options[i].label) + "</option>"
      );
    }

    if (!options.length) {
      html.push(
        '<option value="ADHOC-' +
          escapeAttr(sanitizeKey(projectCode)) +
          '">ADHOC - tarefa nao listada</option>'
      );
    }

    return html.join("");
  }

  function getRefOptionsForProject(projectCode) {
    var project = projectMap[projectCode];
    var key = project ? project.code : projectCode;
    var serverOptions = refOptionsByProject[key] || [];
    if (serverOptions.length) {
      return serverOptions;
    }

    return buildDefaultRefOptionsForProject_(key);
  }

  function buildDefaultRefOptionsForProject_(projectCode) {
    if (projectCode === "ADMIN-INTERNO") {
      return [
        {
          value: "ADMININTERNO-WBS-INT_1",
          label: "ADMININTERNO-WBS-INT_1 | Administrativo/Interno | Atividades administrativas internas",
          status: "NAO INICIADA"
        }
      ];
    }

    if (!/^CT-\d+$/i.test(projectCode)) {
      return [];
    }

    var compact = String(projectCode).replace(/[^A-Za-z0-9]/g, "").toUpperCase();
    return LOCAL_EAP_TEMPLATE.map(function (item) {
      var value = compact + "-WBS-" + String(item.wbs).replace(/[^A-Za-z0-9]+/g, "_");
      return {
        value: value,
        label: value + " | " + item.phase + " | " + item.task,
        status: "NAO INICIADA"
      };
    });
  }

  function mergeProjects_(localProjects, serverProjects) {
    var mergedMap = {};
    var merged = [];

    function addItem(item, isServer) {
      var code = cleanText(item && item.code);
      if (!code || mergedMap[code]) {
        return;
      }
      var label = formatProjectLabel_(code, cleanText(item && item.label));
      mergedMap[code] = true;
      merged.push({
        code: code,
        label: isServer && label ? label : label
      });
    }

    (localProjects || []).forEach(function (item) {
      addItem(item, false);
    });

    (serverProjects || []).forEach(function (item) {
      addItem(item, true);
    });

    return merged;
  }

  function buildStatusOptionsHtml(key) {
    return (cfg.opcoesStatus || [])
      .map(function (item, index) {
        var inputId = "status-" + key + "-" + index;
        return [
          '<label class="radio-item" for="' + inputId + '">',
          '<input id="' + inputId + '" type="radio" name="status-' + key + '" value="' + escapeAttr(item.code) + '" data-label="' + escapeAttr(item.label) + '" />',
          "<span>" + escapeHtml(item.label) + "</span>",
          "</label>"
        ].join("");
      })
      .join("");
  }

  function buildBlockOptionsHtml(key) {
    return (cfg.opcoesBloqueio || [])
      .map(function (item, index) {
        var inputId = "block-" + key + "-" + index;
        return [
          '<label class="radio-item" for="' + inputId + '">',
          '<input id="' + inputId + '" type="radio" name="block-' + key + '" value="' + escapeAttr(item.code) + '" data-label="' + escapeAttr(item.label) + '" />',
          "<span>" + escapeHtml(item.label) + "</span>",
          "</label>"
        ].join("");
      })
      .join("");
  }

  function wireBlockToggle(key) {
    var blockInputs = document.querySelectorAll('input[name="block-' + key + '"]');
    var blockWrap = document.getElementById("block-wrap-" + key);
    var blockDesc = document.getElementById("block-desc-" + key);

    for (var i = 0; i < blockInputs.length; i += 1) {
      var input = blockInputs[i];
      input.addEventListener("change", function () {
        if (this.value === "NAO") {
          blockWrap.classList.add("hidden");
          blockDesc.value = "";
        } else {
          blockWrap.classList.remove("hidden");
        }
      });
    }
  }

  function handleSubmit(event) {
    event.preventDefault();
    clearFeedback();

    var appsScriptUrl = (cfg.appsScriptUrl || "").trim();
    if (!appsScriptUrl || appsScriptUrl.indexOf("COLE_AQUI") >= 0) {
      showFeedback("error", "Configuracao pendente: cole a URL do Apps Script em web/config.js.");
      return;
    }

    var professional = (professionalEl.value || "").trim();
    var workDate = (dateEl.value || "").trim();
    var selectedCodes = getSelectedProjectCodes();

    var errors = [];
    if (!professional) {
      errors.push("Campo 1 obrigatorio: selecione o profissional.");
    }
    if (!workDate) {
      errors.push("Campo 2 obrigatorio: selecione a data.");
    }
    if (!selectedCodes.length) {
      errors.push("Campo 3 obrigatorio: marque ao menos um projeto.");
    }

    var sections = [];
    selectedCodes.forEach(function (code) {
      var project = projectMap[code] || { code: code, label: code };
      var key = sanitizeKey(code);
      var refEl = document.getElementById("ref-" + key);
      var refEap = refEl ? cleanText(refEl.value) : "";
      var refLabel = "";
      if (refEl && refEl.selectedIndex >= 0) {
        refLabel = cleanText(refEl.options[refEl.selectedIndex].text);
      }
      var taskRefEl = document.getElementById("task-ref-" + key);
      var taskRefValue = taskRefEl ? cleanText(taskRefEl.value) : "";
      var taskRefLabel = "";
      if (taskRefEl && taskRefEl.selectedIndex >= 0) {
        taskRefLabel = cleanText(taskRefEl.options[taskRefEl.selectedIndex].text);
      }
      var taskFree = ((document.getElementById("task-free-" + key) || {}).value || "").trim();
      var taskItems = [];
      if (taskRefLabel) {
        taskItems.push(extractTaskNameFromRefLabel(taskRefLabel));
      }
      if (taskFree) {
        taskItems.push("Livre: " + taskFree);
      }
      var statusEl = document.querySelector('input[name="status-' + key + '"]:checked');
      var blockEl = document.querySelector('input[name="block-' + key + '"]:checked');
      var blockDesc = ((document.getElementById("block-desc-" + key) || {}).value || "").trim();
      var obs = ((document.getElementById("obs-" + key) || {}).value || "").trim();

      if (!refEap) {
        errors.push("Campo REF EAP obrigatorio na secao " + project.label + ".");
      }
      if (!taskItems.length) {
        errors.push("Campo 4B obrigatorio na secao " + project.label + " (EAP ou tarefa livre).");
      }
      if (!statusEl) {
        errors.push("Campo 5 obrigatorio na secao " + project.label + ".");
      }
      if (!blockEl) {
        errors.push("Campo 6 obrigatorio na secao " + project.label + ".");
      }
      if (blockEl && blockEl.value !== "NAO" && !blockDesc) {
        errors.push("Campo 7 obrigatorio quando ha bloqueio na secao " + project.label + ".");
      }

      sections.push({
        projectCode: code,
        projectLabel: project.label,
        refEap: refEap,
        refLabel: refLabel,
        tasks: taskItems,
        taskRefEap: taskRefValue,
        taskRefLabel: taskRefLabel,
        taskFreeText: taskFree,
        statusCode: statusEl ? statusEl.value : "",
        statusLabel: statusEl ? statusEl.dataset.label : "",
        blockCode: blockEl ? blockEl.value : "",
        blockLabel: blockEl ? blockEl.dataset.label : "",
        blockDescription: blockDesc,
        observation: obs
      });
    });

    if (errors.length) {
      showFeedback("error", errors.join(" "));
      return;
    }

    var payload = {
      professional: professional,
      date: workDate,
      projects: sections,
      submittedAt: new Date().toISOString(),
      source: "vm-arquitetos-web-v1"
    };

    submitButtonEl.disabled = true;
    submitButtonEl.textContent = "Enviando...";

    sendPayload(appsScriptUrl, payload)
      .then(function (result) {
        if (result.mode === "cors" && result.data && result.data.status === "ok") {
          var msg = "Registro enviado com sucesso.";
          if (result.data.formId) {
            msg += " ID: " + result.data.formId + ".";
          }
          if (result.data.rowsCreated) {
            msg += " Linhas criadas: " + result.data.rowsCreated + ".";
          }
          showFeedback("ok", msg);
          setSummaryForSubmission(result.data.dailySummary || buildFallbackDailySummary(payload));
          resetForm();
          loadRefOptionsFromServer();
          loadPpmSnapshotFromServer();
          return;
        }

        if (result.mode === "cors" && result.data && result.data.status === "error") {
          showFeedback("error", result.data.message || "Falha ao salvar o registro.");
          return;
        }

        showFeedback(
          "warn",
          "Envio realizado no modo de compatibilidade. Confirme a entrada na aba EXECUCAO da planilha."
        );
        resetForm();
        loadRefOptionsFromServer();
        loadPpmSnapshotFromServer();
      })
      .catch(function (err) {
        showFeedback("error", err.message || "Nao foi possivel enviar. Tente novamente.");
      })
      .then(function () {
        submitButtonEl.disabled = false;
        submitButtonEl.textContent = "Enviar registro diario";
      });
  }

  function sendPayload(url, payload) {
    var body = JSON.stringify(payload);

    return fetch(url, {
      method: "POST",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: body
    })
      .then(function (response) {
        return response.text().then(function (text) {
          var data = null;
          try {
            data = text ? JSON.parse(text) : null;
          } catch (err) {
            data = null;
          }

          if (!response.ok) {
            throw new Error((data && data.message) || "Falha de comunicacao com o backend.");
          }

          return { mode: "cors", data: data };
        });
      })
      .catch(function () {
        return fetch(url, {
          method: "POST",
          mode: "no-cors",
          headers: { "Content-Type": "text/plain;charset=utf-8" },
          body: body
        }).then(function () {
          return { mode: "no-cors", data: null };
        });
      });
  }

  function loadRefOptionsFromServer() {
    var appsScriptUrl = cleanText(cfg.appsScriptUrl);
    if (!appsScriptUrl || appsScriptUrl.indexOf("COLE_AQUI") >= 0) {
      portfolioDataset = buildPortfolioDataset_();
      renderPortfolioOverview();
      return;
    }

    var callbackName = "vmFormConfigCb_" + String(Date.now());
    var script = document.createElement("script");
    var settled = false;

    function finish() {
      if (settled) {
        return;
      }
      settled = true;
      try {
        delete window[callbackName];
      } catch (err) {
        window[callbackName] = undefined;
      }
      if (script && script.parentNode) {
        script.parentNode.removeChild(script);
      }
    }

    window[callbackName] = function (data) {
      try {
        if (data && data.status === "ok") {
          var previouslySelected = getSelectedProjectCodes();

          if (Array.isArray(data.projects) && data.projects.length) {
            rebuildProjectCache_(mergeProjects_(localProjectsFromConfig, data.projects));
            populateProjects();
            selectProjectCodes_(previouslySelected);
          }

          if (data.refOptions) {
            refOptionsByProject = normalizeRefOptions(data.refOptions);
          }

          renderProjectSections(getSelectedProjectCodes());
          portfolioDataset = buildPortfolioDataset_();
          renderPortfolioOverview();
        }
      } finally {
        finish();
      }
    };

    script.async = true;
    script.src = buildUrlWithParams(appsScriptUrl, {
      action: "config",
      callback: callbackName,
      _: Date.now()
    });
    script.onerror = finish;
    document.body.appendChild(script);

    setTimeout(finish, 12000);
  }

  function loadPpmSnapshotFromServer() {
    var appsScriptUrl = cleanText(cfg.appsScriptUrl);
    if (!appsScriptUrl || appsScriptUrl.indexOf("COLE_AQUI") >= 0) {
      return;
    }

    var callbackName = "vmPpmCb_" + String(Date.now());
    var script = document.createElement("script");
    var settled = false;

    function finish() {
      if (settled) {
        return;
      }
      settled = true;
      try {
        delete window[callbackName];
      } catch (err) {
        window[callbackName] = undefined;
      }
      if (script && script.parentNode) {
        script.parentNode.removeChild(script);
      }
    }

    window[callbackName] = function (data) {
      try {
        if (data && data.status === "ok") {
          ppmSnapshotMap = normalizePpmSnapshot_(data);
          renderPortfolioOverview();
        }
      } finally {
        finish();
      }
    };

    script.async = true;
    script.src = buildUrlWithParams(appsScriptUrl, {
      action: "ppm_snapshot",
      callback: callbackName,
      _: Date.now()
    });
    script.onerror = finish;
    document.body.appendChild(script);
    setTimeout(finish, 12000);
  }

  function normalizePpmSnapshot_(payload) {
    var map = {};
    var projects = (payload && Array.isArray(payload.projects)) ? payload.projects : [];

    projects.forEach(function (project) {
      var contractId = cleanText(project && (project.contractId || project.projectCode));
      if (!contractId) {
        return;
      }

      var tasks = Array.isArray(project.tasks) ? project.tasks : [];
      map[contractId] = {
        contractId: contractId,
        projectName: cleanText(project.projectName),
        totals: project.totals || {},
        tasks: tasks.map(function (task) {
          return {
            refEap: cleanText(task && task.refEap),
            phase: cleanText(task && task.phase),
            task: cleanText(task && task.task),
            status: normalizePortfolioStatus_(task && task.status),
            plannedStart: parseDateFlexible_(task && task.plannedStart),
            plannedEnd: parseDateFlexible_(task && task.plannedEnd),
            realStart: parseDateFlexible_(task && task.realStart),
            realEnd: parseDateFlexible_(task && task.realEnd),
            lastRecordDate: parseDateFlexible_(task && task.lastRecordDate),
            updatedAt: parseDateFlexible_(task && task.updatedAt),
            blocked: !!(task && task.blocked),
            blockReason: cleanText(task && task.blockReason)
          };
        })
      };
    });

    return map;
  }

  function renderPortfolioOverview() {
    if (!portfolioSummaryEl || !portfolioCardsEl || !portfolioBlocksEl) {
      return;
    }

    portfolioDataset = buildPortfolioDataset_();
    var dataset = Array.isArray(portfolioDataset) ? portfolioDataset : [];
    var filtered = applyPortfolioFilter_(dataset);

    renderPortfolioSummary_(filtered);
    renderPortfolioCards_(filtered);
    renderPortfolioBlocks_(filtered);
    renderMeetingAgenda_(filtered);
  }

  function applyPortfolioFilter_(dataset) {
    var mode = cleanText(portfolioFilterEl && portfolioFilterEl.value);
    if (mode === "monday") {
      var mondayItems = dataset.filter(function (item) {
        return hasWeeklyAttention_(item);
      });

      if (!mondayItems.length) {
        mondayItems = dataset.filter(function (item) {
          return Number(item.emAndamento || 0) > 0 || (item.nextTasks && item.nextTasks.length > 0);
        });
      }

      return mondayItems.sort(compareMeetingPriority_);
    }

    if (mode !== "selected") {
      return dataset.slice();
    }

    var selected = getSelectedProjectCodes();
    if (!selected.length) {
      return [];
    }

    return dataset.filter(function (item) {
      return selected.indexOf(item.code) >= 0;
    });
  }

  function hasWeeklyAttention_(item) {
    if (!item) {
      return false;
    }
    if (Number(item.atrasadasAteCorte || 0) > 0) {
      return true;
    }
    if (Number(item.bloqueada || 0) > 0) {
      return true;
    }
    if (String(item.semaforo) !== "green") {
      return true;
    }
    return Number(item.emAndamento || 0) > 0;
  }

  function compareMeetingPriority_(a, b) {
    var diff = scoreMeetingPriority_(a) - scoreMeetingPriority_(b);
    if (diff !== 0) {
      return diff;
    }
    return String(a.code || "").localeCompare(String(b.code || ""));
  }

  function scoreMeetingPriority_(item) {
    var semaforoScore = 3;
    if (item && item.semaforo === "red") {
      semaforoScore = 0;
    } else if (item && item.semaforo === "yellow") {
      semaforoScore = 1;
    } else if (item && item.semaforo === "green") {
      semaforoScore = 2;
    }

    var blockedPenalty = Number(item && item.bloqueada ? item.bloqueada : 0) > 0 ? -2 : 0;
    var delayedPenalty = -Math.min(Number(item && item.atrasadasAteCorte ? item.atrasadasAteCorte : 0), 9);

    return semaforoScore * 10 + blockedPenalty + delayedPenalty;
  }

  function buildPortfolioDataset_() {
    var data = [];
    var cutoffDate = getCutoffDate_();

    projectList.forEach(function (project) {
      var code = cleanText(project.code);
      if (!code || code === "ADMIN-INTERNO") {
        return;
      }

      var optionList = getRefOptionsForProject(code);
      var snapshotProject = ppmSnapshotMap[code] || null;
      var schedule = buildProjectSchedule_(code, optionList, snapshotProject);
      var metrics = buildPortfolioMetricsFromSchedule_(schedule, cutoffDate);
      var nextTasks = buildPortfolioNextTasks_(schedule);
      var blockedItems = buildPortfolioBlockedItems_(schedule);

      data.push({
        code: code,
        label: cleanText(project.label) || code,
        total: metrics.total,
        concluida: metrics.concluida,
        emAndamento: metrics.emAndamento,
        bloqueada: metrics.bloqueada,
        naoIniciada: metrics.naoIniciada,
        pendentes: metrics.pendentes,
        percentual: metrics.percentual,
        planejadasAteCorte: metrics.planejadasAteCorte,
        realizadasAteCorte: metrics.realizadasAteCorte,
        atrasadasAteCorte: metrics.atrasadasAteCorte,
        adiantadasAteCorte: metrics.adiantadasAteCorte,
        semaforo: classifyPortfolioSemaforo_(metrics),
        nextTasks: nextTasks,
        blockedItems: blockedItems
      });
    });

    data.sort(function (a, b) {
      return String(a.code).localeCompare(String(b.code));
    });
    return data;
  }

  function getCutoffDate_() {
    if (portfolioCutoffEl && cleanText(portfolioCutoffEl.value)) {
      var fromInput = parseIsoDate_(portfolioCutoffEl.value);
      if (fromInput) {
        return fromInput;
      }
    }
    var now = new Date();
    now.setHours(0, 0, 0, 0);
    return now;
  }

  function buildProjectSchedule_(projectCode, optionList, snapshotProject) {
    var snapshotTasks = snapshotProject && Array.isArray(snapshotProject.tasks) ? snapshotProject.tasks : [];
    if (snapshotTasks.length) {
      return buildProjectScheduleFromSnapshot_(projectCode, snapshotTasks);
    }

    var signatureDate = parseIsoDate_(PROJECT_SIGNATURE_DATE[projectCode] || DEFAULT_SIGNATURE_DATE);
    var statusByRef = buildStatusMapFromOptions_(optionList);
    var scheduleByWbs = {};
    var tasks = [];

    LOCAL_EAP_TEMPLATE.forEach(function (item) {
      var rule = LOCAL_EAP_NETWORK[item.wbs] || { du: 1, deps: [] };
      var refEap = buildRefFromWbsForFront_(projectCode, item.wbs);
      var status = statusByRef[refEap] || "CONCLUIDA";
      var dates = planTaskDates_(signatureDate, rule, scheduleByWbs);

      var task = {
        refEap: refEap,
        wbs: item.wbs,
        phase: item.phase,
        task: item.task,
        status: status,
        plannedStart: dates.start,
        plannedEnd: dates.end
      };

      scheduleByWbs[item.wbs] = task;
      tasks.push(task);
    });

    return tasks;
  }

  function buildStatusMapFromOptions_(optionList) {
    var map = {};
    (optionList || []).forEach(function (item) {
      var key = cleanText(item && item.value);
      if (!key) {
        return;
      }
      map[key] = normalizePortfolioStatus_(item && item.status);
    });
    return map;
  }

  function buildProjectScheduleFromSnapshot_(projectCode, snapshotTasks) {
    var signatureDate = parseIsoDate_(PROJECT_SIGNATURE_DATE[projectCode] || DEFAULT_SIGNATURE_DATE);
    var scheduleByWbs = {};
    var taskByRef = {};
    var tasks = [];

    LOCAL_EAP_TEMPLATE.forEach(function (item) {
      var rule = LOCAL_EAP_NETWORK[item.wbs] || { du: 1, deps: [] };
      var refEap = buildRefFromWbsForFront_(projectCode, item.wbs);
      var dates = planTaskDates_(signatureDate, rule, scheduleByWbs);

      var task = {
        refEap: refEap,
        wbs: item.wbs,
        phase: item.phase,
        task: item.task,
        status: "NAO INICIADA",
        plannedStart: dates.start,
        plannedEnd: dates.end,
        realStart: null,
        realEnd: null,
        lastRecordDate: null,
        updatedAt: null,
        blocked: false,
        blockReason: ""
      };

      scheduleByWbs[item.wbs] = task;
      taskByRef[refEap] = task;
      tasks.push(task);
    });

    snapshotTasks.forEach(function (snapshotTask) {
      var refEap = cleanText(snapshotTask && snapshotTask.refEap);
      if (!refEap) {
        return;
      }

      var target = taskByRef[refEap];
      if (!target) {
        target = {
          refEap: refEap,
          wbs: extractWbsFromRef_(refEap),
          phase: "",
          task: "",
          status: "NAO INICIADA",
          plannedStart: null,
          plannedEnd: null,
          realStart: null,
          realEnd: null,
          lastRecordDate: null,
          updatedAt: null,
          blocked: false,
          blockReason: ""
        };
        taskByRef[refEap] = target;
        tasks.push(target);
      }

      target.phase = cleanText(snapshotTask.phase) || target.phase;
      target.task = cleanText(snapshotTask.task) || target.task;
      target.status = normalizePortfolioStatus_(snapshotTask.status);
      target.plannedStart = parseDateFlexible_(snapshotTask.plannedStart) || target.plannedStart;
      target.plannedEnd = parseDateFlexible_(snapshotTask.plannedEnd) || target.plannedEnd;
      target.realStart = parseDateFlexible_(snapshotTask.realStart);
      target.realEnd = parseDateFlexible_(snapshotTask.realEnd);
      target.lastRecordDate = parseDateFlexible_(snapshotTask.lastRecordDate);
      target.updatedAt = parseDateFlexible_(snapshotTask.updatedAt);
      target.blocked = Boolean(snapshotTask && snapshotTask.blocked);
      target.blockReason = cleanText(snapshotTask && snapshotTask.blockReason);
    });

    tasks.sort(function (a, b) {
      return String(a.refEap).localeCompare(String(b.refEap));
    });
    return tasks;
  }

  function buildRefFromWbsForFront_(projectCode, wbs) {
    var compact = String(projectCode || "").replace(/[^A-Za-z0-9]/g, "").toUpperCase();
    var key = String(wbs || "").replace(/[^A-Za-z0-9]+/g, "_");
    return compact + "-WBS-" + key;
  }

  function extractWbsFromRef_(refEap) {
    var text = cleanText(refEap);
    var match = text.match(/-WBS-([0-9_]+)/i);
    if (!match) {
      return "";
    }
    return String(match[1]).replace(/_/g, ".");
  }

  function planTaskDates_(projectStartDate, rule, scheduleByWbs) {
    var deps = Array.isArray(rule.deps) ? rule.deps : [];
    var duration = Math.max(Number(rule.du || 1), 1);
    var start = cloneDate_(projectStartDate);

    if (deps.length) {
      var maxPredEnd = null;
      deps.forEach(function (depWbs) {
        var depTask = scheduleByWbs[depWbs];
        if (!depTask || !depTask.plannedEnd) {
          return;
        }
        if (!maxPredEnd || depTask.plannedEnd.getTime() > maxPredEnd.getTime()) {
          maxPredEnd = depTask.plannedEnd;
        }
      });
      if (maxPredEnd) {
        start = nextBusinessDay_(addDays_(maxPredEnd, 1));
      }
    }

    start = nextBusinessDay_(start);
    var end = addBusinessDays_(start, duration - 1);
    return { start: start, end: end };
  }

  function buildPortfolioMetricsFromSchedule_(schedule, cutoffDate) {
    var metrics = {
      total: schedule.length,
      concluida: 0,
      emAndamento: 0,
      bloqueada: 0,
      naoIniciada: 0,
      pendentes: 0,
      percentual: 0,
      planejadasAteCorte: 0,
      realizadasAteCorte: 0,
      atrasadasAteCorte: 0,
      adiantadasAteCorte: 0
    };

    schedule.forEach(function (task) {
      var status = task.status;
      var isPlannedByCutoff =
        task.plannedEnd instanceof Date && task.plannedEnd.getTime() <= cutoffDate.getTime();
      var isRealizedByCutoff = isTaskRealizedByCutoff_(task, cutoffDate);

      if (status === "CONCLUIDA") {
        metrics.concluida += 1;
      } else if (status === "EM ANDAMENTO") {
        metrics.emAndamento += 1;
      } else if (status === "BLOQUEADA") {
        metrics.bloqueada += 1;
      } else {
        metrics.naoIniciada += 1;
      }

      if (isPlannedByCutoff) {
        metrics.planejadasAteCorte += 1;
        if (isRealizedByCutoff) {
          metrics.realizadasAteCorte += 1;
        } else {
          metrics.atrasadasAteCorte += 1;
        }
      } else if (isRealizedByCutoff) {
        metrics.adiantadasAteCorte += 1;
      }
    });

    metrics.pendentes = metrics.total - metrics.concluida;
    metrics.percentual = metrics.total > 0 ? Math.round((metrics.concluida / metrics.total) * 100) : 0;
    return metrics;
  }

  function isTaskRealizedByCutoff_(task, cutoffDate) {
    if (!task || task.status !== "CONCLUIDA") {
      return false;
    }

    var dates = [task.realEnd, task.lastRecordDate, task.updatedAt, task.realStart];
    var i;
    for (i = 0; i < dates.length; i += 1) {
      var date = parseDateFlexible_(dates[i]);
      if (date) {
        return date.getTime() <= cutoffDate.getTime();
      }
    }

    // Fallback para tarefas concluidas sem data real preenchida.
    return true;
  }

  function normalizePortfolioStatus_(value) {
    var text = cleanText(value).toUpperCase();
    if (text.indexOf("BLOQUE") >= 0) {
      return "BLOQUEADA";
    }
    if (text.indexOf("ANDAMENTO") >= 0) {
      return "EM ANDAMENTO";
    }
    if (text.indexOf("CONCL") >= 0) {
      return "CONCLUIDA";
    }
    return "NAO INICIADA";
  }

  function classifyPortfolioSemaforo_(metrics) {
    if (metrics.bloqueada > 0 && metrics.atrasadasAteCorte > 0) {
      return "red";
    }
    if (metrics.atrasadasAteCorte > 3) {
      return "red";
    }
    if (metrics.atrasadasAteCorte > 0 || metrics.bloqueada > 0) {
      return "yellow";
    }
    return "green";
  }

  function buildPortfolioNextTasks_(schedule) {
    return schedule
      .filter(function (task) {
        return task.status !== "CONCLUIDA";
      })
      .sort(function (a, b) {
        var byStatus = portfolioTaskPriority_(a.status) - portfolioTaskPriority_(b.status);
        if (byStatus !== 0) {
          return byStatus;
        }
        if (a.plannedEnd && b.plannedEnd) {
          return a.plannedEnd.getTime() - b.plannedEnd.getTime();
        }
        return String(a.refEap).localeCompare(String(b.refEap));
      })
      .slice(0, 3)
      .map(function (task) {
        return {
          refEap: task.refEap,
          task: task.task,
          status: task.status
        };
      });
  }

  function buildPortfolioBlockedItems_(schedule) {
    return schedule
      .filter(function (task) {
        return task.status === "BLOQUEADA";
      })
      .map(function (task) {
        return {
          refEap: task.refEap,
          task: task.task,
          status: task.status,
          blockReason: cleanText(task.blockReason)
        };
      });
  }

  function portfolioTaskPriority_(status) {
    if (status === "EM ANDAMENTO") {
      return 1;
    }
    if (status === "BLOQUEADA") {
      return 2;
    }
    if (status === "NAO INICIADA") {
      return 3;
    }
    return 4;
  }

  function renderPortfolioSummary_(items) {
    if (!items.length) {
      portfolioSummaryEl.innerHTML = '<div class="portfolio-empty">Selecione projetos ou aguarde a atualizacao de dados.</div>';
      return;
    }

    var cutoffDate = getCutoffDate_();
    var totais = {
      projetos: items.length,
      verde: 0,
      amarelo: 0,
      vermelho: 0,
      bloqueios: 0,
      planejadas: 0,
      realizadas: 0,
      atrasadas: 0
    };

    items.forEach(function (item) {
      if (item.semaforo === "green") {
        totais.verde += 1;
      } else if (item.semaforo === "yellow") {
        totais.amarelo += 1;
      } else {
        totais.vermelho += 1;
      }
      totais.bloqueios += item.bloqueada;
      totais.planejadas += item.planejadasAteCorte;
      totais.realizadas += item.realizadasAteCorte;
      totais.atrasadas += item.atrasadasAteCorte;
    });

    portfolioSummaryEl.innerHTML = [
      buildPortfolioStat_("Corte", formatDateBr_(cutoffDate)),
      buildPortfolioStat_("Semaforo Verde", totais.verde),
      buildPortfolioStat_("Semaforo Amarelo", totais.amarelo),
      buildPortfolioStat_("Semaforo Vermelho", totais.vermelho),
      buildPortfolioStat_("Planejado ate corte", totais.planejadas),
      buildPortfolioStat_("Realizado ate corte*", totais.realizadas),
      buildPortfolioStat_("Atrasadas", totais.atrasadas),
      buildPortfolioStat_("Bloqueios", totais.bloqueios),
      buildPortfolioStat_("Projetos", totais.projetos)
    ].join("");
  }

  function buildPortfolioStat_(label, value) {
    return (
      '<article class="portfolio-stat"><strong>' +
      escapeHtml(String(value)) +
      "</strong><span>" +
      escapeHtml(label) +
      "</span></article>"
    );
  }

  function renderPortfolioCards_(items) {
    if (!items.length) {
      portfolioCardsEl.innerHTML = '<div class="portfolio-empty">Nenhum projeto para exibir neste filtro.</div>';
      return;
    }

    portfolioCardsEl.innerHTML = items
      .map(function (item) {
        var aderencia =
          item.planejadasAteCorte > 0
            ? Math.round((item.realizadasAteCorte / item.planejadasAteCorte) * 100)
            : 100;
        var nextTasks = item.nextTasks.length
          ? "<ul>" +
            item.nextTasks
              .map(function (task) {
                return "<li>" + escapeHtml(task.refEap + " - " + task.task) + "</li>";
              })
              .join("") +
            "</ul>"
          : "<p class='portfolio-empty'>Sem proximo foco identificado.</p>";

        return [
          '<article class="portfolio-card">',
          '<div class="portfolio-card-head">',
          "<h3>" + escapeHtml(item.label) + "</h3>",
          '<span class="portfolio-badge ' + escapeAttr(item.semaforo) + '">' + escapeHtml(labelSemaforo_(item.semaforo)) + "</span>",
          "</div>",
          '<div class="portfolio-progress">',
          '<div class="portfolio-progress-meta"><span>Concluido</span><strong>' + escapeHtml(String(item.percentual)) + "%</strong></div>",
          '<div class="portfolio-progress-bar"><div class="portfolio-progress-fill" style="width:' + escapeAttr(String(Math.max(0, Math.min(100, item.percentual)))) + '%;"></div></div>',
          "</div>",
          '<div class="portfolio-metrics">',
          '<div class="portfolio-metric"><strong>' + escapeHtml(String(item.concluida)) + '</strong><span>Concluidas</span></div>',
          '<div class="portfolio-metric"><strong>' + escapeHtml(String(item.emAndamento)) + '</strong><span>Em andamento</span></div>',
          '<div class="portfolio-metric"><strong>' + escapeHtml(String(item.bloqueada)) + '</strong><span>Bloqueadas</span></div>',
          '<div class="portfolio-metric"><strong>' + escapeHtml(String(item.naoIniciada)) + '</strong><span>Nao iniciadas</span></div>',
          '<div class="portfolio-metric"><strong>' + escapeHtml(String(item.realizadasAteCorte)) + "/" + escapeHtml(String(item.planejadasAteCorte)) + '</strong><span>Real x Plan ate corte</span></div>',
          '<div class="portfolio-metric"><strong>' + escapeHtml(String(item.atrasadasAteCorte)) + '</strong><span>Atrasadas no corte</span></div>',
          '<div class="portfolio-metric"><strong>' + escapeHtml(String(item.adiantadasAteCorte)) + '</strong><span>Adiantadas no corte</span></div>',
          '<div class="portfolio-metric"><strong>' + escapeHtml(String(Math.max(0, Math.min(150, aderencia)))) + '%</strong><span>Aderencia no corte</span></div>',
          "</div>",
          '<div class="portfolio-next"><h4>Proximo foco</h4>' + nextTasks + "</div>",
          "</article>"
        ].join("");
      })
      .join("");
  }

  function renderPortfolioBlocks_(items) {
    var blocks = [];
    items.forEach(function (item) {
      item.blockedItems.forEach(function (block) {
        blocks.push({
          project: item.label,
          refEap: block.refEap,
          task: block.task,
          blockReason: block.blockReason
        });
      });
    });

    if (!blocks.length) {
      portfolioBlocksEl.innerHTML = '<div class="portfolio-empty">Nenhum bloqueio ativo no filtro atual.</div>';
      return;
    }

    portfolioBlocksEl.innerHTML = blocks
      .slice(0, 12)
      .map(function (block) {
        var details = cleanText(block.refEap) + " - " + cleanText(block.task);
        if (cleanText(block.blockReason)) {
          details += " | " + cleanText(block.blockReason);
        }
        return (
          '<article class="portfolio-block-item"><strong>' +
          escapeHtml(block.project) +
          "</strong><p>" +
          escapeHtml(details) +
          "</p></article>"
        );
      })
      .join("");
  }

  function renderMeetingAgenda_(items) {
    if (!portfolioMeetingEl || !meetingAgendaEl) {
      return;
    }

    var mode = cleanText(portfolioFilterEl && portfolioFilterEl.value);
    if (mode !== "monday") {
      portfolioMeetingEl.classList.add("hidden");
      meetingAgendaEl.innerHTML = "";
      lastMeetingSummaryText = "";
      if (meetingCopyEl) {
        meetingCopyEl.disabled = true;
      }
      return;
    }

    portfolioMeetingEl.classList.remove("hidden");

    var cutoffDate = getCutoffDate_();
    var ranges = buildMeetingWeekRanges_(cutoffDate);
    var agendaItems = buildMeetingAgendaItems_(items);

    if (!agendaItems.length) {
      meetingAgendaEl.innerHTML =
        '<div class="meeting-empty">Nenhuma pendencia critica para a semana. Use este espaco para confirmar focos de manutencao.</div>';
      lastMeetingSummaryText = buildMeetingSummaryText_([], ranges, cutoffDate);
      if (meetingCopyEl) {
        meetingCopyEl.disabled = false;
      }
      return;
    }

    meetingAgendaEl.innerHTML = agendaItems
      .map(function (item) {
        var focoHtml = item.focusList.length
          ? "<ul>" +
            item.focusList
              .map(function (focus) {
                return "<li>" + escapeHtml(focus) + "</li>";
              })
              .join("") +
            "</ul>"
          : "<p class='portfolio-empty'>Sem foco sugerido.</p>";

        return [
          '<article class="meeting-card">',
          '<div class="meeting-card-head">',
          "<strong>" + escapeHtml(item.label) + "</strong>",
          '<span class="portfolio-badge ' + escapeAttr(item.semaforo) + '">' + escapeHtml(labelSemaforo_(item.semaforo)) + "</span>",
          "</div>",
          "<ul>",
          "<li>Semana anterior: " +
            escapeHtml(String(item.realizadasAteCorte)) +
            "/" +
            escapeHtml(String(item.planejadasAteCorte)) +
            " entregas no corte (" +
            escapeHtml(String(item.aderencia)) +
            "%).</li>",
          "<li>Pendencias abertas: " + escapeHtml(String(item.atrasadasAteCorte)) + ".</li>",
          "<li>Bloqueios ativos: " + escapeHtml(String(item.bloqueada)) + ".</li>",
          "<li>Foco da semana atual:</li>",
          "</ul>",
          focoHtml,
          "</article>"
        ].join("");
      })
      .join("");

    lastMeetingSummaryText = buildMeetingSummaryText_(agendaItems, ranges, cutoffDate);
    if (meetingCopyEl) {
      meetingCopyEl.disabled = false;
    }
  }

  function buildMeetingAgendaItems_(items) {
    return (items || [])
      .map(function (item) {
        var aderencia =
          Number(item.planejadasAteCorte || 0) > 0
            ? Math.round((Number(item.realizadasAteCorte || 0) / Number(item.planejadasAteCorte || 1)) * 100)
            : 100;
        var focusList = (item.nextTasks || []).slice(0, 3).map(function (task) {
          var text = cleanText(task.refEap);
          if (cleanText(task.task)) {
            text += " - " + cleanText(task.task);
          }
          if (cleanText(task.status)) {
            text += " (" + cleanText(task.status) + ")";
          }
          return text;
        });

        return {
          code: cleanText(item.code),
          label: cleanText(item.label),
          semaforo: cleanText(item.semaforo),
          planejadasAteCorte: Number(item.planejadasAteCorte || 0),
          realizadasAteCorte: Number(item.realizadasAteCorte || 0),
          atrasadasAteCorte: Number(item.atrasadasAteCorte || 0),
          bloqueada: Number(item.bloqueada || 0),
          aderencia: Math.max(0, Math.min(150, aderencia)),
          focusList: focusList
        };
      })
      .sort(compareMeetingPriority_);
  }

  function buildMeetingWeekRanges_(cutoffDate) {
    var monday = getMondayFromDate_(cutoffDate);
    var friday = addDays_(monday, 4);
    if (friday.getTime() > cutoffDate.getTime()) {
      friday = cloneDate_(cutoffDate);
    }

    var nextMonday = addDays_(monday, 7);
    var nextFriday = addDays_(nextMonday, 4);

    return {
      previousStart: monday,
      previousEnd: friday,
      currentStart: nextMonday,
      currentEnd: nextFriday
    };
  }

  function getMondayFromDate_(date) {
    var cursor = cloneDate_(date);
    while (cursor.getDay() !== 1) {
      cursor = addDays_(cursor, -1);
    }
    return cursor;
  }

  function buildMeetingSummaryText_(agendaItems, ranges, cutoffDate) {
    var lines = [];
    lines.push("VM + Arquitetos | Resumo da reuniao de coordenacao");
    lines.push("Data de corte: " + formatDateBr_(cutoffDate));
    lines.push(
      "Semana anterior: " +
        formatDateBr_(ranges.previousStart) +
        " a " +
        formatDateBr_(ranges.previousEnd)
    );
    lines.push(
      "Semana atual: " +
        formatDateBr_(ranges.currentStart) +
        " a " +
        formatDateBr_(ranges.currentEnd)
    );
    lines.push("");

    if (!agendaItems.length) {
      lines.push("Sem pendencias criticas nesta virada semanal.");
      lines.push("Ajuste focos de manutencao por projeto.");
      return lines.join("\n");
    }

    agendaItems.forEach(function (item, index) {
      lines.push(String(index + 1) + ". " + cleanText(item.label) + " [" + labelSemaforo_(item.semaforo) + "]");
      lines.push(
        "   - Semana anterior: " +
          item.realizadasAteCorte +
          "/" +
          item.planejadasAteCorte +
          " entregas (" +
          item.aderencia +
          "%)."
      );
      lines.push("   - Pendencias abertas: " + item.atrasadasAteCorte + ".");
      lines.push("   - Bloqueios ativos: " + item.bloqueada + ".");
      if (item.focusList.length) {
        lines.push("   - Foco da semana: " + item.focusList.join("; ") + ".");
      } else {
        lines.push("   - Foco da semana: definir em reuniao.");
      }
      lines.push("");
    });

    return lines.join("\n");
  }

  function labelSemaforo_(value) {
    if (value === "green") {
      return "Verde";
    }
    if (value === "red") {
      return "Vermelho";
    }
    return "Amarelo";
  }

  function buildUrlWithParams(url, params) {
    var query = [];
    var key;
    for (key in params) {
      if (Object.prototype.hasOwnProperty.call(params, key)) {
        query.push(encodeURIComponent(key) + "=" + encodeURIComponent(String(params[key])));
      }
    }
    return url + (url.indexOf("?") >= 0 ? "&" : "?") + query.join("&");
  }

  function selectProjectCodes_(codes) {
    (codes || []).forEach(function (code) {
      var input = document.getElementById("project-" + sanitizeKey(code));
      if (input) {
        input.checked = true;
      }
    });
  }

  function renderDailySummary(summary) {
    if (!dailySummaryEl || !summary || !Array.isArray(summary.projects) || !summary.projects.length) {
      return;
    }

    var header = [
      '<h2>Resumo rapido do dia</h2>',
      '<p class="summary-meta">' +
        escapeHtml(cleanText(summary.professional)) +
        " | " +
        escapeHtml(cleanText(summary.dateBr)) +
        "</p>"
    ].join("");

    var projectsHtml = summary.projects
      .map(function (project) {
        return buildProjectSummaryHtml(project || {});
      })
      .join("");

    dailySummaryEl.innerHTML = header + projectsHtml;
  }

  function setSummaryForSubmission(summary) {
    if (!summary || !Array.isArray(summary.projects) || !summary.projects.length) {
      clearSummaryUI();
      return;
    }

    lastSubmittedSummary = summary;
    renderDailySummary(summary);

    if (dailySummaryEl) {
      dailySummaryEl.classList.add("hidden");
    }
    if (summaryToggleEl) {
      summaryToggleEl.textContent = "Ver resumo deste envio";
      summaryToggleEl.classList.remove("hidden");
    }
  }

  function handleSummaryToggle() {
    if (!dailySummaryEl || !summaryToggleEl || !lastSubmittedSummary) {
      return;
    }

    if (dailySummaryEl.classList.contains("hidden")) {
      dailySummaryEl.classList.remove("hidden");
      summaryToggleEl.textContent = "Ocultar resumo deste envio";
      return;
    }

    dailySummaryEl.classList.add("hidden");
    summaryToggleEl.textContent = "Ver resumo deste envio";
  }

  function clearSummaryUI() {
    lastSubmittedSummary = null;
    if (summaryToggleEl) {
      summaryToggleEl.classList.add("hidden");
      summaryToggleEl.textContent = "Ver resumo deste envio";
    }
    if (dailySummaryEl) {
      dailySummaryEl.classList.add("hidden");
      dailySummaryEl.innerHTML = "";
    }
  }

  function buildProjectSummaryHtml(project) {
    var snapshot = project.snapshot || {};
    var todayItems = Array.isArray(project.todayItems) ? project.todayItems : [];
    var nextTasks = Array.isArray(project.nextTasks) ? project.nextTasks : [];
    var title = cleanText(project.projectName) || cleanText(project.contractId) || "Projeto";

    return [
      '<article class="summary-project">',
      '<h3>' + escapeHtml(title) + "</h3>",
      '<p><strong>Hoje:</strong> ' + escapeHtml(formatTodayItems(todayItems)) + "</p>",
      '<p><strong>Pendencias:</strong> ' + escapeHtml(formatSnapshot(snapshot)) + "</p>",
      '<p><strong>Proximo dia:</strong> ' + escapeHtml(formatNextTasks(nextTasks)) + "</p>",
      "</article>"
    ].join("");
  }

  function formatTodayItems(items) {
    if (!items.length) {
      return "Sem itens registrados.";
    }

    return items
      .map(function (item) {
        var line = cleanText(item.refEap) || "REF";
        if (cleanText(item.taskText)) {
          line += ": " + cleanText(item.taskText);
        }
        if (cleanText(item.statusText)) {
          line += " (" + cleanText(item.statusText) + ")";
        }
        if (item.isBlocked && cleanText(item.blockReason)) {
          line += " - bloqueio: " + cleanText(item.blockReason);
        }
        return line;
      })
      .join("; ");
  }

  function formatSnapshot(snapshot) {
    var pendentes = Number(snapshot.pendentes || 0);
    var andamento = Number(snapshot.emAndamento || 0);
    var bloqueadas = Number(snapshot.bloqueada || 0);
    var naoIniciadas = Number(snapshot.naoIniciada || 0);
    var concluidas = Number(snapshot.concluida || 0);

    if (!pendentes && !andamento && !bloqueadas && !naoIniciadas && !concluidas) {
      return "Sem dados de pendencias para este projeto.";
    }

    return (
      pendentes +
      " pendentes (" +
      andamento +
      " em andamento, " +
      bloqueadas +
      " bloqueadas, " +
      naoIniciadas +
      " nao iniciadas, " +
      concluidas +
      " concluidas)"
    );
  }

  function formatNextTasks(tasks) {
    if (!tasks.length) {
      return "Sem proximo item sugerido.";
    }

    return tasks
      .map(function (task) {
        var parts = [cleanText(task.refEap)];
        if (cleanText(task.status)) {
          parts.push(cleanText(task.status));
        }
        if (cleanText(task.task)) {
          parts.push(cleanText(task.task));
        }
        return parts.filter(Boolean).join(" - ");
      })
      .join("; ");
  }

  function buildFallbackDailySummary(payload) {
    var projects = (payload.projects || []).map(function (project) {
      var optionList = getRefOptionsForProject(project.projectCode);
      return {
        projectCode: cleanText(project.projectCode),
        contractId: cleanText(project.projectCode),
        projectName: cleanText(project.projectLabel) || cleanText(project.projectCode),
        todayItems: [
          {
            refEap: cleanText(project.refEap),
            taskText: Array.isArray(project.tasks) ? project.tasks.join(" | ") : "",
            statusText: cleanText(project.statusLabel),
            isBlocked: cleanText(project.blockCode) && cleanText(project.blockCode) !== "NAO",
            blockReason: cleanText(project.blockDescription)
          }
        ],
        snapshot: buildSnapshotFromOptions(optionList),
        nextTasks: buildNextTasksFromOptions(optionList)
      };
    });

    return {
      professional: cleanText(payload.professional),
      dateBr: formatIsoDateToBr(payload.date),
      generatedAtBr: "",
      projects: projects
    };
  }

  function buildSnapshotFromOptions(optionList) {
    var snapshot = {
      total: 0,
      naoIniciada: 0,
      emAndamento: 0,
      concluida: 0,
      bloqueada: 0,
      pendentes: 0
    };

    var i;
    for (i = 0; i < optionList.length; i += 1) {
      var status = cleanText(optionList[i] && optionList[i].status).toUpperCase();
      snapshot.total += 1;
      snapshot.pendentes += 1;
      if (status === "NAO INICIADA") {
        snapshot.naoIniciada += 1;
      } else if (status === "EM ANDAMENTO") {
        snapshot.emAndamento += 1;
      } else if (status === "BLOQUEADA") {
        snapshot.bloqueada += 1;
      } else if (status === "CONCLUIDA") {
        snapshot.concluida += 1;
      }
    }

    return snapshot;
  }

  function buildNextTasksFromOptions(optionList) {
    return optionList.slice(0, 3).map(function (item) {
      var parsed = parseRefOptionLabel(cleanText(item.label));
      return {
        refEap: parsed.refEap,
        phase: parsed.phase,
        task: parsed.task,
        status: cleanText(item.status)
      };
    });
  }

  function parseRefOptionLabel(label) {
    var parts = cleanText(label).split(" | ");
    return {
      refEap: cleanText(parts[0]),
      phase: cleanText(parts[1]),
      task: cleanText(parts.slice(2).join(" | "))
    };
  }

  function extractTaskNameFromRefLabel(label) {
    var parsed = parseRefOptionLabel(label);
    return cleanText(parsed.task) || cleanText(parsed.refEap) || cleanText(label);
  }

  function formatIsoDateToBr(value) {
    var text = cleanText(value);
    var match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) {
      return text;
    }
    return match[3] + "/" + match[2] + "/" + match[1];
  }

  function resetForm() {
    formEl.reset();
    setDefaultDate();
    renderProjectSections([]);
  }

  function getSelectedProjectCodes() {
    var checked = document.querySelectorAll('input[name="projects"]:checked');
    var values = [];
    var i;
    for (i = 0; i < checked.length; i += 1) {
      values.push(checked[i].value);
    }
    return values;
  }

  function normalizeRefOptions(raw) {
    var normalized = {};
    var projectCode;

    if (!raw || typeof raw !== "object") {
      return normalized;
    }

    for (projectCode in raw) {
      if (!Object.prototype.hasOwnProperty.call(raw, projectCode)) {
        continue;
      }
      if (!Array.isArray(raw[projectCode])) {
        normalized[projectCode] = [];
        continue;
      }

      normalized[projectCode] = raw[projectCode]
        .map(function (item) {
          return {
            value: cleanText(item && item.value),
            label: cleanText(item && item.label),
            status: cleanText(item && item.status)
          };
        })
        .filter(function (item) {
          return item.value && item.label;
        });
    }

    return normalized;
  }

  function parseIsoDate_(value) {
    var text = cleanText(value);
    var match = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (!match) {
      return null;
    }
    var year = Number(match[1]);
    var month = Number(match[2]) - 1;
    var day = Number(match[3]);
    var date = new Date(year, month, day);
    date.setHours(0, 0, 0, 0);
    if (isNaN(date.getTime())) {
      return null;
    }
    return date;
  }

  function parseDateFlexible_(value) {
    if (value instanceof Date && !isNaN(value.getTime())) {
      var fromDate = cloneDate_(value);
      fromDate.setHours(0, 0, 0, 0);
      return fromDate;
    }

    var text = cleanText(value);
    if (!text) {
      return null;
    }

    var iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (iso) {
      return parseIsoDate_(iso[1] + "-" + iso[2] + "-" + iso[3]);
    }

    var br = text.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
    if (br) {
      return parseIsoDate_(br[3] + "-" + br[2] + "-" + br[1]);
    }

    var parsed = new Date(text);
    if (isNaN(parsed.getTime())) {
      return null;
    }
    parsed.setHours(0, 0, 0, 0);
    return parsed;
  }

  function cloneDate_(date) {
    var base = date instanceof Date ? date : new Date();
    return new Date(base.getFullYear(), base.getMonth(), base.getDate());
  }

  function addDays_(date, days) {
    var next = cloneDate_(date);
    next.setDate(next.getDate() + Number(days || 0));
    return next;
  }

  function isBusinessDay_(date) {
    var day = date.getDay();
    return day !== 0 && day !== 6;
  }

  function nextBusinessDay_(date) {
    var cursor = cloneDate_(date);
    while (!isBusinessDay_(cursor)) {
      cursor = addDays_(cursor, 1);
    }
    return cursor;
  }

  function addBusinessDays_(date, businessDays) {
    var total = Math.max(Number(businessDays || 0), 0);
    var cursor = nextBusinessDay_(date);
    var count = 0;
    while (count < total) {
      cursor = addDays_(cursor, 1);
      if (isBusinessDay_(cursor)) {
        count += 1;
      }
    }
    return cursor;
  }

  function formatDateBr_(date) {
    if (!(date instanceof Date) || isNaN(date.getTime())) {
      return "";
    }
    var dd = String(date.getDate()).padStart(2, "0");
    var mm = String(date.getMonth() + 1).padStart(2, "0");
    var yyyy = String(date.getFullYear());
    return dd + "/" + mm + "/" + yyyy;
  }

  function setDefaultDate() {
    var now = new Date();
    now.setMinutes(now.getMinutes() - now.getTimezoneOffset());
    dateEl.value = now.toISOString().slice(0, 10);
  }

  function setDefaultPortfolioCutoffDate_() {
    if (!portfolioCutoffEl) {
      return;
    }
    if (cleanText(portfolioCutoffEl.value)) {
      return;
    }
    var now = new Date();
    now.setMinutes(now.getMinutes() - now.getTimezoneOffset());
    portfolioCutoffEl.value = now.toISOString().slice(0, 10);
  }

  function copyTextToClipboard_(text) {
    if (navigator.clipboard && typeof navigator.clipboard.writeText === "function") {
      return navigator.clipboard.writeText(text);
    }

    return new Promise(function (resolve, reject) {
      try {
        var textArea = document.createElement("textarea");
        textArea.value = text;
        textArea.setAttribute("readonly", "readonly");
        textArea.style.position = "fixed";
        textArea.style.top = "-1000px";
        textArea.style.left = "-1000px";
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        var ok = document.execCommand("copy");
        document.body.removeChild(textArea);
        if (ok) {
          resolve();
          return;
        }
      } catch (error) {
        reject(error);
        return;
      }

      reject(new Error("copy_failed"));
    });
  }

  function showFeedback(type, message) {
    feedbackEl.classList.remove("hidden", "ok", "warn", "error");
    feedbackEl.classList.add(type);
    feedbackEl.textContent = message;
  }

  function clearFeedback() {
    feedbackEl.classList.add("hidden");
    feedbackEl.classList.remove("ok", "warn", "error");
    feedbackEl.textContent = "";
  }

  function sanitizeKey(value) {
    return String(value).replace(/[^a-zA-Z0-9_-]/g, "_");
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  function escapeAttr(value) {
    return escapeHtml(value);
  }

  function cleanText(value) {
    if (value === null || value === undefined) {
      return "";
    }
    return String(value).trim();
  }
})();
