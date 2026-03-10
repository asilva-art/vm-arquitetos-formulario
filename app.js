(function () {
  "use strict";

  var cfg = window.VM_FORM_CONFIG || {};
  var localProjectsFromConfig = Array.isArray(cfg.projetos) ? cfg.projetos : [];
  var formEl = document.getElementById("daily-form");
  var professionalEl = document.getElementById("professional");
  var dateEl = document.getElementById("work-date");
  var projectListEl = document.getElementById("project-list");
  var projectSectionsEl = document.getElementById("project-sections");
  var bulkStatusEl = document.getElementById("bulk-status");
  var bulkBlockEl = document.getElementById("bulk-block");
  var applyBulkEl = document.getElementById("apply-bulk");
  var repeatLastEl = document.getElementById("repeat-last");
  var formChecklistEl = document.getElementById("form-checklist");
  var feedbackEl = document.getElementById("feedback");
  var portfolioFilterEl = document.getElementById("portfolio-filter");
  var portfolioSortEl = document.getElementById("portfolio-sort");
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
  var editLastEl = document.getElementById("edit-last");
  var dailySummaryEl = document.getElementById("daily-summary");
  var dailyFocusEl = document.getElementById("daily-focus");
  var dailyFocusListEl = document.getElementById("daily-focus-list");
  var startFocusEl = document.getElementById("start-focus");
  var offlineQueueEl = document.getElementById("offline-queue");
  var offlineQueueTextEl = document.getElementById("offline-queue-text");
  var retryQueueEl = document.getElementById("retry-queue");
  var historyPanelEl = document.getElementById("history-panel");
  var historyListEl = document.getElementById("history-list");
  var historyRefreshEl = document.getElementById("history-refresh");
  var coordModalEl = document.getElementById("coord-modal");
  var coordOverlayEl = document.getElementById("coord-overlay");
  var coordTitleEl = document.getElementById("coord-title");
  var coordSubtitleEl = document.getElementById("coord-subtitle");
  var coordContractMetaEl = document.getElementById("coord-contract-meta");
  var coordDeadlineMetaEl = document.getElementById("coord-deadline-meta");
  var coordCloseEl = document.getElementById("coord-close");
  var coordFilterEl = document.getElementById("coord-filter");
  var coordSaveEl = document.getElementById("coord-save");
  var coordFeedbackEl = document.getElementById("coord-feedback");
  var coordAlertsEl = document.getElementById("coord-alerts");
  var coordTaskListEl = document.getElementById("coord-task-list");
  var coordAddEventEl = document.getElementById("coord-add-event");
  var coordEventsListEl = document.getElementById("coord-events-list");
  var submitButtonEl = document.getElementById("submit-button");
  var lastSubmittedSummary = null;
  var lastMeetingSummaryText = "";
  var lastFocusItems = [];
  var historyItems = [];
  var pendingEditFormId = "";
  var pendingEditPayload = null;
  var queueRetryRunning = false;
  var draftSaveTimer = null;
  var portfolioDataset = [];
  var ppmSnapshotMap = {};
  var coordState = {
    open: false,
    contractId: "",
    projectName: "",
    signatureDate: "",
    daysSinceSignature: null,
    deadlineSummary: "",
    tasks: [],
    events: [],
    warnings: [],
    originalTasksByRef: {},
    originalEventsKey: ""
  };

  var STORAGE_KEYS = {
    LAST_PROFESSIONAL: "vm_form_last_professional_v1",
    DRAFT: "vm_form_draft_v1",
    LAST_SUBMISSION: "vm_form_last_submission_v1",
    PROJECT_USAGE: "vm_form_project_usage_v1",
    OFFLINE_QUEUE: "vm_form_offline_queue_v1",
    EDIT_CONTEXT: "vm_form_edit_context_v1"
  };
  var JSONP_TIMEOUT_MS = 20000;

  var projectMap = {};
  var projectList = [];
  var refOptionsByProject = normalizeRefOptions(cfg.refOptions || {});

  var DEFAULT_SIGNATURE_DATE = "";
  var PROJECT_SIGNATURE_DATE = Object.assign({}, cfg.datasAssinatura || {});
  var DEFAULT_DEADLINE_SUMMARY = cleanText(cfg.deadlineSummaryDefault) || "";
  var LEGACY_GENERIC_DEADLINE_SUMMARY =
    "F1 10DU | F2 5DU | F3 15DU | RAF1/2/3 5/5/10DU | PRE 10DU | RFA 7DU | EXE 15DU | DET 15DU | RDE 10DU | Recesso Natal/Ano Novo nao contabiliza";
  var DEFAULT_DEADLINE_LEGEND_MAP = {
    F1: "Fase I: ate 10 dias uteis a partir do pagamento da primeira parcela e levantamento.",
    F2: "Fase II: ate 5 dias uteis apos aprovacao do Anteprojeto Fase I.",
    F3: "Fase III: ate 15 dias uteis apos entrega do Anteprojeto Fase II.",
    "RAF1/2/3": "Revisao de Anteprojeto: Fase I ate 5 DU; Fase II ate 5 DU; Fase III ate 10 DU.",
    PRE: "Pre-Executivo: ate 10 dias uteis apos aprovacao do Anteprojeto Fase III.",
    RFA: "Revisao Final do Anteprojeto: ate 7 dias uteis apos aprovacao do Pre-Executivo.",
    EXE: "Projeto Executivo: ate 15 dias uteis apos entrega da revisao final do Anteprojeto.",
    DET: "Detalhamento: ate 15 dias uteis apos entrega do Projeto Executivo.",
    RDE: "Revisao de detalhamento e executivo: ate 10 dias uteis.",
    RECESSO: "Nao sao contabilizados os prazos no periodo de recesso entre Natal e Ano Novo.",
    DU: "DU = dias uteis."
  };
  var deadlineLegendMap = buildDeadlineLegendMap_(cfg.deadlineLegend);
  rebuildProjectCache_(localProjectsFromConfig);

  init();

  function init() {
    if (!formEl) {
      return;
    }

    populateProfessionals();
    restoreLastProfessional_();
    populateProjects();
    populateBulkSelectors_();
    setDefaultDate();
    setDefaultPortfolioCutoffDate_();
    renderProjectSections([]);
    applyDailySimpleMode_();

    professionalEl.addEventListener("change", handleProfessionalChange_);
    projectListEl.addEventListener("change", handleProjectSelectionChange);
    formEl.addEventListener("submit", handleSubmit);
    formEl.addEventListener("change", handleFormInteraction_);
    formEl.addEventListener("input", handleFormInteraction_);
    if (portfolioFilterEl) {
      portfolioFilterEl.addEventListener("change", renderPortfolioOverview);
    }
    if (portfolioSortEl) {
      portfolioSortEl.addEventListener("change", renderPortfolioOverview);
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
    if (applyBulkEl) {
      applyBulkEl.addEventListener("click", handleApplyBulkAction_);
    }
    if (repeatLastEl) {
      repeatLastEl.addEventListener("click", handleRepeatLastSubmission_);
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
    if (editLastEl) {
      editLastEl.addEventListener("click", handleEditLastSubmission_);
    }
    if (startFocusEl) {
      startFocusEl.addEventListener("click", handleStartFocus_);
    }
    if (historyRefreshEl) {
      historyRefreshEl.addEventListener("click", loadHistoryFromServer_);
    }
    if (retryQueueEl) {
      retryQueueEl.addEventListener("click", flushOfflineQueue_);
    }
    if (portfolioCardsEl) {
      portfolioCardsEl.addEventListener("click", handlePortfolioCardClick_);
      portfolioCardsEl.addEventListener("keydown", handlePortfolioCardKeydown_);
    }
    if (coordCloseEl) {
      coordCloseEl.addEventListener("click", closeCoordPanel_);
    }
    if (coordOverlayEl) {
      coordOverlayEl.addEventListener("click", closeCoordPanel_);
    }
    if (coordSaveEl) {
      coordSaveEl.addEventListener("click", handleCoordSave_);
    }
    if (coordFilterEl) {
      coordFilterEl.addEventListener("change", applyCoordFilter_);
    }
    if (coordAddEventEl) {
      coordAddEventEl.addEventListener("click", function () {
        addCoordEventRow_();
      });
    }
    if (coordEventsListEl) {
      coordEventsListEl.addEventListener("click", handleCoordEventRemove_);
    }
    document.addEventListener("keydown", handleCoordModalKeydown_);
    window.addEventListener("online", function () {
      updateOfflineQueueUi_();
      flushOfflineQueue_();
    });

    restoreDraftIfAny_();
    restoreEditContext_();
    loadRefOptionsFromServer();
    loadPpmSnapshotFromServer();
    loadHistoryFromServer_();
    portfolioDataset = buildPortfolioDataset_();
    renderPortfolioOverview();
    renderDailyFocus_();
    updateOfflineQueueUi_();
    flushOfflineQueue_();
    updateFormChecklist_();
    ensureEditButtonState_();
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

  function applyDailySimpleMode_() {
    if (guideStepsEl && guideToggleEl && !guideStepsEl.classList.contains("hidden")) {
      guideStepsEl.classList.add("hidden");
      guideToggleEl.setAttribute("aria-expanded", "false");
      guideToggleEl.textContent = "Expandir explicacoes";
    }

    if (portfolioContentEl && portfolioToggleEl && !portfolioContentEl.classList.contains("hidden")) {
      portfolioContentEl.classList.add("hidden");
      portfolioToggleEl.setAttribute("aria-expanded", "false");
      portfolioToggleEl.textContent = "Expandir status visual";
    }
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

  function handleProfessionalChange_() {
    rememberLastProfessional_(cleanText((professionalEl || {}).value));
    var draft = buildCurrentDraft_();
    refreshProjectsKeepingDraft_(draft);
    loadHistoryFromServer_();
    renderDailyFocus_();
  }

  function handleFormInteraction_(event) {
    var target = event && event.target;
    if (!target) {
      return;
    }

    if (target === professionalEl || target.id === "portfolio-filter" || target.id === "portfolio-sort" || target.id === "portfolio-cutoff") {
      return;
    }

    updateFormChecklist_();
    scheduleDraftSave_();
  }

  function refreshProjectsKeepingDraft_(draft) {
    var snapshot = draft || buildCurrentDraft_();
    var selectedCodes = Array.isArray(snapshot.selectedCodes) ? snapshot.selectedCodes.slice() : [];

    populateProjects();
    selectProjectCodes_(selectedCodes);
    renderProjectSections(selectedCodes);
    applyDraftValuesToSections_(snapshot);
    updateFormChecklist_();
    scheduleDraftSave_();
  }

  function populateBulkSelectors_() {
    if (bulkStatusEl) {
      bulkStatusEl.innerHTML = '<option value="">Nao alterar</option>';
      (cfg.opcoesStatus || []).forEach(function (item) {
        var option = document.createElement("option");
        option.value = cleanText(item && item.code);
        option.textContent = cleanText(item && item.label) || cleanText(item && item.code);
        bulkStatusEl.appendChild(option);
      });
    }

    if (bulkBlockEl) {
      bulkBlockEl.innerHTML = '<option value="">Nao alterar</option>';
      (cfg.opcoesBloqueio || []).forEach(function (item) {
        var option = document.createElement("option");
        option.value = cleanText(item && item.code);
        option.textContent = cleanText(item && item.label) || cleanText(item && item.code);
        bulkBlockEl.appendChild(option);
      });
    }
  }

  function handleApplyBulkAction_() {
    var statusCode = cleanText((bulkStatusEl || {}).value);
    var blockCode = cleanText((bulkBlockEl || {}).value);
    var selectedCodes = getSelectedProjectCodes();

    if (!selectedCodes.length) {
      showFeedback("warn", "Marque ao menos um projeto antes de aplicar em lote.");
      return;
    }

    if (!statusCode && !blockCode) {
      showFeedback("warn", "Selecione status e/ou bloqueio para aplicar em lote.");
      return;
    }

    selectedCodes.forEach(function (code) {
      var key = sanitizeKey(code);
      if (statusCode) {
        var statusInput = document.querySelector(
          'input[name="status-' + key + '"][value="' + statusCode + '"]'
        );
        if (statusInput) {
          statusInput.checked = true;
        }
      }

      if (blockCode) {
        var blockInput = document.querySelector(
          'input[name="block-' + key + '"][value="' + blockCode + '"]'
        );
        if (blockInput) {
          blockInput.checked = true;
          blockInput.dispatchEvent(new Event("change"));
        }
      }
    });

    updateFormChecklist_();
    scheduleDraftSave_();
    showFeedback("ok", "Aplicacao em lote concluida nas secoes selecionadas.");
  }

  function handleRepeatLastSubmission_() {
    clearEditMode_();
    var last = getStorageJson_(STORAGE_KEYS.LAST_SUBMISSION);
    if (!last || !Array.isArray(last.projects) || !last.projects.length) {
      showFeedback("warn", "Nenhum preenchimento anterior encontrado neste navegador.");
      return;
    }

    if (last.professional && professionalEl) {
      professionalEl.value = cleanText(last.professional);
      rememberLastProfessional_(cleanText(last.professional));
    }

    if (!cleanText((dateEl || {}).value)) {
      setDefaultDate();
    }

    var selectedCodes = last.projects
      .map(function (item) {
        return cleanText(item && item.projectCode);
      })
      .filter(Boolean);

    populateProjects();
    selectProjectCodes_(selectedCodes);
    renderProjectSections(selectedCodes);

    last.projects.forEach(function (section) {
      applySectionValues_(cleanText(section && section.projectCode), {
        refEap: cleanText(section && section.refEap),
        taskRefEap: cleanText(section && section.taskRefEap),
        taskFreeText: cleanText(section && section.taskFreeText),
        statusCode: cleanText(section && section.statusCode),
        blockCode: cleanText(section && section.blockCode),
        blockDescription: cleanText(section && section.blockDescription),
        observation: cleanText(section && section.observation)
      });
    });

    updateFormChecklist_();
    renderDailyFocus_();
    scheduleDraftSave_();
    showFeedback("ok", "Ultimo preenchimento replicado. Revise e envie.");
  }

  function saveLastSubmissionTemplate_(payload) {
    if (!payload || !Array.isArray(payload.projects) || !payload.projects.length) {
      return;
    }

    setStorageJson_(STORAGE_KEYS.LAST_SUBMISSION, {
      professional: cleanText(payload.professional),
      date: cleanText(payload.date),
      projects: payload.projects.map(function (item) {
        return {
          projectCode: cleanText(item && item.projectCode),
          refEap: cleanText(item && item.refEap),
          taskRefEap: cleanText(item && item.taskRefEap),
          taskFreeText: cleanText(item && item.taskFreeText),
          statusCode: cleanText(item && item.statusCode),
          blockCode: cleanText(item && item.blockCode),
          blockDescription: cleanText(item && item.blockDescription),
          observation: cleanText(item && item.observation)
        };
      }),
      savedAt: new Date().toISOString()
    });
  }

  function scheduleDraftSave_() {
    if (draftSaveTimer) {
      clearTimeout(draftSaveTimer);
    }
    draftSaveTimer = setTimeout(function () {
      saveDraftNow_();
      draftSaveTimer = null;
    }, 250);
  }

  function saveDraftNow_() {
    var draft = buildCurrentDraft_();
    if (!hasDraftContent_(draft)) {
      clearDraft_();
      return;
    }
    setStorageJson_(STORAGE_KEYS.DRAFT, draft);
  }

  function clearDraft_() {
    removeStorageKey_(STORAGE_KEYS.DRAFT);
  }

  function restoreDraftIfAny_() {
    var draft = getStorageJson_(STORAGE_KEYS.DRAFT);
    if (!draft || typeof draft !== "object") {
      return;
    }

    var savedAt = parseDateFlexible_(draft.savedAt);
    if (savedAt && Math.abs(new Date().getTime() - savedAt.getTime()) > 7 * 24 * 60 * 60 * 1000) {
      clearDraft_();
      return;
    }

    if (cleanText(draft.professional) && professionalEl) {
      professionalEl.value = cleanText(draft.professional);
      rememberLastProfessional_(cleanText(draft.professional));
    }

    if (cleanText(draft.date) && dateEl) {
      dateEl.value = cleanText(draft.date);
    }

    populateProjects();
    var selectedCodes = Array.isArray(draft.selectedCodes) ? draft.selectedCodes : [];
    selectProjectCodes_(selectedCodes);
    renderProjectSections(selectedCodes);
    applyDraftValuesToSections_(draft);
    updateFormChecklist_();

    if (selectedCodes.length) {
      showFeedback("warn", "Rascunho restaurado automaticamente neste navegador.");
    }
  }

  function buildCurrentDraft_() {
    var selectedCodes = getSelectedProjectCodes();
    var sections = {};

    selectedCodes.forEach(function (code) {
      var key = sanitizeKey(code);
      var statusEl = document.querySelector('input[name="status-' + key + '"]:checked');
      var blockEl = document.querySelector('input[name="block-' + key + '"]:checked');

      sections[code] = {
        refEap: cleanText((document.getElementById("ref-" + key) || {}).value),
        taskRefEap: cleanText((document.getElementById("task-ref-" + key) || {}).value),
        taskFreeText: cleanText((document.getElementById("task-free-" + key) || {}).value),
        statusCode: cleanText(statusEl && statusEl.value),
        blockCode: cleanText(blockEl && blockEl.value),
        blockDescription: cleanText((document.getElementById("block-desc-" + key) || {}).value),
        observation: cleanText((document.getElementById("obs-" + key) || {}).value)
      };
    });

    return {
      professional: cleanText((professionalEl || {}).value),
      date: cleanText((dateEl || {}).value),
      selectedCodes: selectedCodes,
      sectionsByCode: sections,
      savedAt: new Date().toISOString()
    };
  }

  function hasDraftContent_(draft) {
    if (!draft) {
      return false;
    }
    if (cleanText(draft.professional) || cleanText(draft.date)) {
      return true;
    }
    return Array.isArray(draft.selectedCodes) && draft.selectedCodes.length > 0;
  }

  function applyDraftValuesToSections_(draft) {
    if (!draft || !draft.sectionsByCode) {
      return;
    }
    var sectionMap = draft.sectionsByCode;
    var selectedCodes = getSelectedProjectCodes();

    selectedCodes.forEach(function (code) {
      applySectionValues_(code, sectionMap[code] || {});
    });
  }

  function applySectionValues_(code, values) {
    var key = sanitizeKey(code);
    var refEl = document.getElementById("ref-" + key);
    var taskRefEl = document.getElementById("task-ref-" + key);
    var taskFreeEl = document.getElementById("task-free-" + key);
    var blockDescEl = document.getElementById("block-desc-" + key);
    var obsEl = document.getElementById("obs-" + key);

    if (refEl && cleanText(values.refEap)) {
      refEl.value = cleanText(values.refEap);
    }
    if (taskRefEl && cleanText(values.taskRefEap)) {
      taskRefEl.value = cleanText(values.taskRefEap);
    }
    if (taskFreeEl && cleanText(values.taskFreeText)) {
      taskFreeEl.value = cleanText(values.taskFreeText);
    }

    if (cleanText(values.statusCode)) {
      var statusInput = document.querySelector(
        'input[name="status-' + key + '"][value="' + cleanText(values.statusCode) + '"]'
      );
      if (statusInput) {
        statusInput.checked = true;
      }
    }

    if (cleanText(values.blockCode)) {
      var blockInput = document.querySelector(
        'input[name="block-' + key + '"][value="' + cleanText(values.blockCode) + '"]'
      );
      if (blockInput) {
        blockInput.checked = true;
        blockInput.dispatchEvent(new Event("change"));
      }
    }

    if (blockDescEl && cleanText(values.blockDescription)) {
      blockDescEl.value = cleanText(values.blockDescription);
    }
    if (obsEl && cleanText(values.observation)) {
      obsEl.value = cleanText(values.observation);
    }
  }

  function rememberLastProfessional_(professional) {
    if (!professional) {
      return;
    }
    setStorageJson_(STORAGE_KEYS.LAST_PROFESSIONAL, professional);
  }

  function restoreLastProfessional_() {
    var saved = cleanText(getStorageJson_(STORAGE_KEYS.LAST_PROFESSIONAL));
    if (!saved || !professionalEl) {
      return;
    }
    professionalEl.value = saved;
  }

  function rememberProjectUsage_(professional, projectCodes) {
    var owner = cleanText(professional).toUpperCase();
    if (!owner || !Array.isArray(projectCodes) || !projectCodes.length) {
      return;
    }

    var map = getStorageJson_(STORAGE_KEYS.PROJECT_USAGE) || {};
    var ownerMap = map[owner] || {};
    var stamp = Date.now();

    projectCodes.forEach(function (code) {
      var key = cleanText(code);
      if (key) {
        ownerMap[key] = stamp;
      }
    });

    map[owner] = ownerMap;
    setStorageJson_(STORAGE_KEYS.PROJECT_USAGE, map);
  }

  function getProjectUsageForProfessional_(professional) {
    var owner = cleanText(professional).toUpperCase();
    if (!owner) {
      return {};
    }
    var map = getStorageJson_(STORAGE_KEYS.PROJECT_USAGE) || {};
    return map[owner] || {};
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
    var currentProfessional = cleanText((professionalEl || {}).value);
    var orderedProjects = getProjectsOrderedForProfessional_(currentProfessional);
    var html = orderedProjects
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

  function getProjectsOrderedForProfessional_(professional) {
    var base = (projectList || []).slice();
    var usage = getProjectUsageForProfessional_(professional);

    base.sort(function (a, b) {
      var scoreA = Number(usage[a.code] || 0);
      var scoreB = Number(usage[b.code] || 0);
      if (scoreA !== scoreB) {
        return scoreB - scoreA;
      }
      return String(a.code).localeCompare(String(b.code));
    });

    return base;
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
      var signatureDate = cleanText(item && item.signatureDate);
      var deadlineSummary = cleanText(item && item.deadlineSummary);
      var project = {
        code: code,
        label: label,
        signatureDate: signatureDate,
        deadlineSummary: deadlineSummary
      };
      normalized.push(project);
      map[code] = project;

      if (signatureDate) {
        PROJECT_SIGNATURE_DATE[code] = signatureDate;
      }
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
    if (isContractDisplayName_(label)) {
      return label;
    }
    if (label.indexOf(code) === 0) {
      return label;
    }
    return code + " | " + label;
  }

  function isContractDisplayName_(label) {
    var text = cleanText(label).toUpperCase();
    if (!text) {
      return false;
    }
    return /^CT[\s_-]+[0-9]{8}[-_/][0-9]{3}(?:[-_/][0-9]{3})?/.test(text);
  }

  function isLegacyGenericDeadlineSummary_(value) {
    return cleanText(value) === LEGACY_GENERIC_DEADLINE_SUMMARY;
  }

  function handleProjectSelectionChange() {
    var draft = buildCurrentDraft_();
    renderProjectSections(getSelectedProjectCodes());
    applyDraftValuesToSections_(draft);
    renderPortfolioOverview();
    renderDailyFocus_();
    updateFormChecklist_();
    scheduleDraftSave_();
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
      wireRefAutoFill_(key);
      setDefaultBlockValue_(key);
    });

    updateFormChecklist_();
  }

  function wireRefAutoFill_(key) {
    var refEl = document.getElementById("ref-" + key);
    var taskRefEl = document.getElementById("task-ref-" + key);
    if (!refEl || !taskRefEl) {
      return;
    }

    refEl.addEventListener("change", function () {
      if (cleanText(refEl.value)) {
        taskRefEl.value = cleanText(refEl.value);
      }
      updateFormChecklist_();
      scheduleDraftSave_();
    });
  }

  function setDefaultBlockValue_(key) {
    var blockInput = document.querySelector('input[name="block-' + key + '"][value="NAO"]');
    if (!blockInput) {
      return;
    }
    blockInput.checked = true;
    blockInput.dispatchEvent(new Event("change"));
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

    return [];
  }

  function mergeProjects_(localProjects, serverProjects) {
    var mergedMap = {};
    var merged = [];

    function getSignatureDateFromProject_(item) {
      return cleanText(item && (item.signatureDate || item.dataAssinatura || item.assinatura));
    }

    function getDeadlineSummaryFromProject_(item) {
      return cleanText(item && (item.deadlineSummary || item.prazoResumo || item.prazosResumo));
    }

    function addOrMerge(item, preferIncoming) {
      var code = cleanText(item && item.code);
      if (!code) {
        return;
      }

      var incomingLabel = cleanText(item && item.label);
      var incomingSignatureDate = getSignatureDateFromProject_(item);
      var incomingDeadlineSummary = getDeadlineSummaryFromProject_(item);
      var existing = mergedMap[code];

      if (!existing) {
        existing = {
          code: code,
          label: formatProjectLabel_(code, incomingLabel),
          signatureDate: incomingSignatureDate,
          deadlineSummary: incomingDeadlineSummary
        };
        mergedMap[code] = existing;
        merged.push(existing);
        return;
      }

      if (preferIncoming && incomingLabel) {
        var formattedIncomingLabel = formatProjectLabel_(code, incomingLabel);
        var hasExistingLabel = !!cleanText(existing.label);
        var incomingIsContractDisplay = isContractDisplayName_(formattedIncomingLabel);
        if (!hasExistingLabel || incomingIsContractDisplay) {
          existing.label = formattedIncomingLabel;
        }
      }
      if (incomingSignatureDate) {
        existing.signatureDate = incomingSignatureDate;
      }
      if (incomingDeadlineSummary) {
        var existingSummary = cleanText(existing.deadlineSummary);
        var incomingIsGeneric = isLegacyGenericDeadlineSummary_(incomingDeadlineSummary);
        var existingIsDetailed = /4\.\d/.test(existingSummary);
        if (!existingSummary || !incomingIsGeneric || !existingIsDetailed) {
          existing.deadlineSummary = incomingDeadlineSummary;
        }
      }
    }

    (localProjects || []).forEach(function (item) {
      addOrMerge(item, false);
    });

    (serverProjects || []).forEach(function (item) {
      addOrMerge(item, true);
    });

    return merged;
  }

  function mergeSignatureDates_(signatureDates) {
    var code;
    for (code in signatureDates) {
      if (!Object.prototype.hasOwnProperty.call(signatureDates, code)) {
        continue;
      }
      var normalizedCode = cleanText(code);
      var normalizedDate = cleanText(signatureDates[code]);
      if (!normalizedCode || !normalizedDate) {
        continue;
      }
      PROJECT_SIGNATURE_DATE[normalizedCode] = normalizedDate;
      if (projectMap[normalizedCode]) {
        projectMap[normalizedCode].signatureDate = normalizedDate;
      }
    }
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
        updateFormChecklist_();
        scheduleDraftSave_();
      });
    }
  }

  function updateFormChecklist_() {
    if (!formChecklistEl) {
      return;
    }

    var selectedCodes = getSelectedProjectCodes();
    formChecklistEl.classList.remove("ok", "warn");

    if (!selectedCodes.length) {
      formChecklistEl.classList.add("warn");
      formChecklistEl.textContent = "Selecione ao menos 1 projeto para iniciar o preenchimento.";
      return;
    }

    var pending = [];
    selectedCodes.forEach(function (code) {
      var project = projectMap[code] || { label: code };
      var missing = getSectionMissingFields_(code);
      if (missing.length) {
        pending.push({
          label: cleanText(project.label) || code,
          missing: missing
        });
      }
    });

    if (!pending.length) {
      formChecklistEl.classList.add("ok");
      formChecklistEl.textContent =
        "Tudo certo para envio: " + String(selectedCodes.length) + " secoes completas.";
      return;
    }

    var first = pending[0];
    formChecklistEl.classList.add("warn");
    formChecklistEl.textContent =
      "Faltam " +
      String(pending.length) +
      " secoes incompletas. Exemplo: " +
      first.label +
      " (" +
      first.missing.join(", ") +
      ").";
  }

  function getSectionMissingFields_(code) {
    var key = sanitizeKey(code);
    var missing = [];
    var refEap = cleanText((document.getElementById("ref-" + key) || {}).value);
    var taskRef = cleanText((document.getElementById("task-ref-" + key) || {}).value);
    var taskFree = cleanText((document.getElementById("task-free-" + key) || {}).value);
    var statusEl = document.querySelector('input[name="status-' + key + '"]:checked');
    var blockEl = document.querySelector('input[name="block-' + key + '"]:checked');
    var blockDesc = cleanText((document.getElementById("block-desc-" + key) || {}).value);

    if (!refEap) {
      missing.push("REF EAP");
    }
    if (!taskRef && !taskFree) {
      missing.push("Tarefa 4B");
    }
    if (!statusEl) {
      missing.push("Status");
    }
    if (!blockEl) {
      missing.push("Bloqueio");
    } else if (cleanText(blockEl.value) !== "NAO" && !blockDesc) {
      missing.push("Desc. bloqueio");
    }

    return missing;
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
    var isEditing = !!cleanText(pendingEditFormId);
    if (isEditing) {
      payload.editFormId = cleanText(pendingEditFormId);
    }

    rememberLastProfessional_(professional);

    submitButtonEl.disabled = true;
    submitButtonEl.textContent = "Enviando...";

    sendPayload(appsScriptUrl, payload)
      .then(function (result) {
        if (result.mode === "cors" && result.data && result.data.status === "ok") {
          var formId = cleanText(result.data.formId) || cleanText(payload.editFormId);
          var msg = isEditing ? "Edicao salva com sucesso." : "Registro enviado com sucesso.";
          if (result.data.formId) {
            msg += " ID: " + result.data.formId + ".";
          }
          if (result.data.rowsCreated) {
            msg += " Linhas criadas: " + result.data.rowsCreated + ".";
          }
          var summary = result.data.dailySummary || buildFallbackDailySummary(payload);
          msg += " " + buildPostSubmitHint_(summary);
          rememberProjectUsage_(professional, selectedCodes);
          saveLastSubmissionTemplate_(payload);
          saveEditContext_(formId, payload);
          clearDraft_();
          showFeedback("ok", msg);
          setSummaryForSubmission(summary);
          clearEditMode_();
          resetForm();
          loadRefOptionsFromServer();
          loadPpmSnapshotFromServer();
          loadHistoryFromServer_();
          renderDailyFocus_();
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
        rememberProjectUsage_(professional, selectedCodes);
        saveLastSubmissionTemplate_(payload);
        if (cleanText(payload.editFormId)) {
          saveEditContext_(cleanText(payload.editFormId), payload);
        }
        clearDraft_();
        clearEditMode_();
        resetForm();
        loadRefOptionsFromServer();
        loadPpmSnapshotFromServer();
        loadHistoryFromServer_();
        renderDailyFocus_();
      })
      .catch(function (err) {
        if (shouldQueueSubmission_(err)) {
          enqueueOfflineSubmission_(payload);
          showFeedback(
            "warn",
            "Sem internet agora. O envio foi guardado e sera reenviado automaticamente quando a conexao voltar."
          );
          clearEditMode_();
          return;
        }
        showFeedback("error", err.message || "Nao foi possivel enviar. Tente novamente.");
      })
      .then(function () {
        submitButtonEl.disabled = false;
        if (!cleanText(pendingEditFormId)) {
          submitButtonEl.textContent = "Enviar registro diario";
        }
      });
  }

  function buildPostSubmitHint_(summary) {
    var projects = summary && Array.isArray(summary.projects) ? summary.projects : [];
    if (!projects.length) {
      return "Faltou algo? Use Editar ultimo envio para ajustar.";
    }

    var totalPendencias = 0;
    var nextFocus = "";
    projects.forEach(function (project) {
      var snapshot = project && project.snapshot ? project.snapshot : {};
      totalPendencias += Number(snapshot.pendentes || 0);
      if (!nextFocus && project && Array.isArray(project.nextTasks) && project.nextTasks.length) {
        nextFocus = cleanText(project.nextTasks[0].refEap) || cleanText(project.nextTasks[0].task);
      }
    });

    var message = "Faltou algo? Use Editar ultimo envio para ajustes.";
    if (totalPendencias > 0) {
      message += " Pendencias abertas: " + totalPendencias + ".";
    }
    if (nextFocus) {
      message += " Proximo foco sugerido: " + nextFocus + ".";
    }
    return message;
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
            var httpError = new Error((data && data.message) || "Falha de comunicacao com o backend.");
            httpError.isHttp = true;
            throw httpError;
          }

          return { mode: "cors", data: data };
        });
      })
      .catch(function (error) {
        if (error && error.isHttp) {
          throw error;
        }
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
          var incomingDeadlineSummaryDefault = cleanText(data.deadlineSummaryDefault);
          if (incomingDeadlineSummaryDefault) {
            DEFAULT_DEADLINE_SUMMARY = incomingDeadlineSummaryDefault;
          }
          if (Array.isArray(data.deadlineLegend) && data.deadlineLegend.length) {
            deadlineLegendMap = buildDeadlineLegendMap_(data.deadlineLegend);
          }

          var previousDraft = buildCurrentDraft_();
          var previouslySelected = Array.isArray(previousDraft.selectedCodes)
            ? previousDraft.selectedCodes
            : [];

          if (data.signatureDates && typeof data.signatureDates === "object") {
            mergeSignatureDates_(data.signatureDates);
          }

          if (Array.isArray(data.projects) && data.projects.length) {
            rebuildProjectCache_(mergeProjects_(localProjectsFromConfig, data.projects));
            populateProjects();
            selectProjectCodes_(previouslySelected);
          }

          if (data.refOptions) {
            refOptionsByProject = normalizeRefOptions(data.refOptions);
          }

          renderProjectSections(getSelectedProjectCodes());
          applyDraftValuesToSections_(previousDraft);
          portfolioDataset = buildPortfolioDataset_();
          renderPortfolioOverview();
          renderDailyFocus_();
          updateFormChecklist_();
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

    setTimeout(finish, JSONP_TIMEOUT_MS);
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
          if (data.signatureDates && typeof data.signatureDates === "object") {
            mergeSignatureDates_(data.signatureDates);
          }
          renderPortfolioOverview();
          renderDailyFocus_();
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
    setTimeout(finish, JSONP_TIMEOUT_MS);
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
        projectName: resolvePreferredProjectName_(contractId, cleanText(project.projectName)),
        signatureDate: cleanText(project && project.signatureDate),
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
            responsible: cleanText(task && task.responsible),
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
    var sorted = applyPortfolioSort_(filtered);

    renderPortfolioSummary_(sorted);
    renderPortfolioCards_(sorted);
    renderPortfolioBlocks_(sorted);
    renderMeetingAgenda_(sorted);
  }

  function renderDailyFocus_() {
    if (!dailyFocusEl || !dailyFocusListEl) {
      return;
    }

    var professional = cleanText((professionalEl || {}).value);
    lastFocusItems = buildDailyFocusItems_(professional);
    ensureEditButtonState_();

    if (!lastFocusItems.length) {
      dailyFocusEl.classList.add("hidden");
      dailyFocusListEl.innerHTML = "";
      return;
    }

    dailyFocusEl.classList.remove("hidden");
    dailyFocusListEl.innerHTML = lastFocusItems
      .map(function (item, index) {
        var statusTagClass = item.status === "BLOQUEADA" ? "blocked" : (item.status === "EM ANDAMENTO" ? "doing" : "todo");
        return [
          '<article class="daily-focus-item' + (index === 0 ? " highlight" : "") + '">',
          '<div class="daily-focus-item-head">',
          '<strong>' + escapeHtml(item.projectLabel) + "</strong>",
          '<span class="daily-focus-tag ' + escapeAttr(statusTagClass) + '">' + escapeHtml(item.status) + "</span>",
          "</div>",
          '<p><span class="daily-focus-ref">' + escapeHtml(item.refEap) + "</span> - " + escapeHtml(item.task || "Sem descricao da tarefa") + "</p>",
          '<p class="daily-focus-meta">Responsavel: ' + escapeHtml(item.responsible || "Nao definido") + "</p>",
          "</article>"
        ].join("");
      })
      .join("");
  }

  function buildDailyFocusItems_(professional) {
    var selectedCodes = getSelectedProjectCodes();
    var selectedMap = {};
    selectedCodes.forEach(function (code) {
      selectedMap[code] = true;
    });

    var owner = cleanText(professional).toUpperCase();
    var today = getCutoffDate_();
    var items = [];
    var code;

    for (code in ppmSnapshotMap) {
      if (!Object.prototype.hasOwnProperty.call(ppmSnapshotMap, code)) {
        continue;
      }

      var project = ppmSnapshotMap[code] || {};
      var projectLabel = formatProjectLabel_(code, cleanText(project.projectName));
      var tasks = Array.isArray(project.tasks) ? project.tasks : [];
      var i;
      for (i = 0; i < tasks.length; i += 1) {
        var task = tasks[i];
        var status = normalizePortfolioStatus_(task && task.status);
        if (status === "CONCLUIDA") {
          continue;
        }

        var responsible = cleanText(task && task.responsible);
        var plannedEnd = parseDateFlexible_(task && task.plannedEnd);
        var overdueDays = 0;
        if (plannedEnd && plannedEnd.getTime() < today.getTime()) {
          overdueDays = Math.floor((today.getTime() - plannedEnd.getTime()) / (24 * 60 * 60 * 1000));
        }

        var score = 0;
        if (status === "EM ANDAMENTO") {
          score -= 20;
        } else if (status === "BLOQUEADA") {
          score -= 15;
        }

        if (owner && cleanText(responsible).toUpperCase() === owner) {
          score -= 25;
        }
        if (selectedMap[code]) {
          score -= 8;
        }
        if (overdueDays > 0) {
          score -= Math.min(overdueDays, 10);
        }

        items.push({
          projectCode: code,
          projectLabel: projectLabel,
          refEap: cleanText(task && task.refEap),
          task: cleanText(task && task.task),
          status: status,
          responsible: responsible,
          plannedEnd: plannedEnd,
          score: score
        });
      }
    }

    items.sort(function (a, b) {
      if (a.score !== b.score) {
        return a.score - b.score;
      }
      if (a.plannedEnd && b.plannedEnd) {
        return a.plannedEnd.getTime() - b.plannedEnd.getTime();
      }
      return String(a.refEap || "").localeCompare(String(b.refEap || ""));
    });

    return items.slice(0, 5);
  }

  function handleStartFocus_() {
    if (!lastFocusItems.length) {
      showFeedback("warn", "Sem foco recomendado no momento.");
      return;
    }

    var focus = lastFocusItems[0];
    var draft = buildCurrentDraft_();
    var selectedCodes = (Array.isArray(draft.selectedCodes) ? draft.selectedCodes.slice() : []);
    if (selectedCodes.indexOf(focus.projectCode) < 0) {
      selectedCodes.push(focus.projectCode);
    }

    populateProjects();
    selectProjectCodes_(selectedCodes);
    renderProjectSections(selectedCodes);
    applyDraftValuesToSections_(draft);

    var key = sanitizeKey(focus.projectCode);
    var refEl = document.getElementById("ref-" + key);
    var taskRefEl = document.getElementById("task-ref-" + key);
    var statusInput = document.querySelector('input[name="status-' + key + '"][value="EM_ANDAMENTO"]');
    if (refEl && cleanText(focus.refEap)) {
      refEl.value = cleanText(focus.refEap);
    }
    if (taskRefEl && cleanText(focus.refEap)) {
      taskRefEl.value = cleanText(focus.refEap);
    }
    if (statusInput) {
      statusInput.checked = true;
    }

    updateFormChecklist_();
    scheduleDraftSave_();
    showFeedback("ok", "Foco recomendado aplicado em " + focus.projectLabel + ". Revise e envie.");

    var section = document.querySelector('[data-project-code="' + focus.projectCode + '"]');
    if (section && typeof section.scrollIntoView === "function") {
      section.scrollIntoView({ behavior: "smooth", block: "start" });
    }
  }

  function loadHistoryFromServer_() {
    if (!historyPanelEl || !historyListEl) {
      return;
    }

    var professional = cleanText((professionalEl || {}).value);
    if (!professional) {
      historyItems = [];
      renderHistoryPanel_();
      return;
    }

    var appsScriptUrl = cleanText(cfg.appsScriptUrl);
    if (!appsScriptUrl || appsScriptUrl.indexOf("COLE_AQUI") >= 0) {
      historyItems = [];
      renderHistoryPanel_();
      return;
    }

    var callbackName = "vmHistoryCb_" + String(Date.now());
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
        if (data && data.status === "ok" && Array.isArray(data.history)) {
          historyItems = data.history;
        } else {
          historyItems = [];
        }
        renderHistoryPanel_();
      } finally {
        finish();
      }
    };

    var params = {
      action: "history",
      limit: 10,
      professional: professional,
      callback: callbackName,
      _: Date.now()
    };

    script.async = true;
    script.src = buildUrlWithParams(appsScriptUrl, params);
    script.onerror = function () {
      historyItems = [];
      renderHistoryPanel_();
      finish();
    };
    document.body.appendChild(script);
    setTimeout(finish, JSONP_TIMEOUT_MS);
  }

  function renderHistoryPanel_() {
    if (!historyPanelEl || !historyListEl) {
      return;
    }

    if (!historyItems.length) {
      historyPanelEl.classList.add("hidden");
      historyListEl.innerHTML = "";
      return;
    }

    historyPanelEl.classList.remove("hidden");
    historyListEl.innerHTML = historyItems
      .map(function (item) {
        var projects = Array.isArray(item.projects) ? item.projects : [];
        var preview = projects
          .slice(0, 3)
          .map(function (project) {
            var text = cleanText(project.projectCode) || cleanText(project.projectName) || "Projeto";
            if (cleanText(project.refEap)) {
              text += " - " + cleanText(project.refEap);
            }
            return "<li>" + escapeHtml(text) + "</li>";
          })
          .join("");
        var more = projects.length > 3 ? "<li>+" + String(projects.length - 3) + " item(ns)</li>" : "";

        return [
          '<article class="history-item">',
          '<div class="history-item-head">',
          '<strong>' + escapeHtml(cleanText(item.formId) || "Sem ID") + "</strong>",
          '<span>' + escapeHtml(cleanText(item.dateBr) || "-") + " | " + escapeHtml(cleanText(item.professional) || "-") + "</span>",
          "</div>",
          (cleanText(item.submittedAtBr)
            ? '<p class="history-submitted">Enviado em: ' + escapeHtml(cleanText(item.submittedAtBr)) + "</p>"
            : ""),
          '<p>' + escapeHtml(String(projects.length)) + " projeto(s) neste envio.</p>",
          "<ul>" + preview + more + "</ul>",
          "</article>"
        ].join("");
      })
      .join("");
  }

  function restoreEditContext_() {
    var data = getStorageJson_(STORAGE_KEYS.EDIT_CONTEXT);
    if (!data || typeof data !== "object") {
      pendingEditFormId = "";
      pendingEditPayload = null;
      return;
    }

    pendingEditFormId = "";
    pendingEditPayload = {
      formId: cleanText(data.formId),
      payload: data.payload || null
    };
  }

  function saveEditContext_(formId, payload) {
    if (!cleanText(formId) || !payload) {
      return;
    }
    pendingEditPayload = {
      formId: cleanText(formId),
      payload: payload
    };
    setStorageJson_(STORAGE_KEYS.EDIT_CONTEXT, pendingEditPayload);
    ensureEditButtonState_();
  }

  function clearEditMode_() {
    pendingEditFormId = "";
    if (submitButtonEl) {
      submitButtonEl.textContent = "Enviar registro diario";
    }
  }

  function ensureEditButtonState_() {
    if (!editLastEl) {
      return;
    }
    var professional = cleanText((professionalEl || {}).value);
    var context = pendingEditPayload || getStorageJson_(STORAGE_KEYS.EDIT_CONTEXT);
    if (!context || !context.payload) {
      editLastEl.classList.add("hidden");
      return;
    }

    var ctxProfessional = cleanText(context.payload.professional);
    if (professional && ctxProfessional && professional !== ctxProfessional) {
      editLastEl.classList.add("hidden");
      return;
    }
    editLastEl.classList.remove("hidden");
  }

  function handleEditLastSubmission_() {
    var context = pendingEditPayload || getStorageJson_(STORAGE_KEYS.EDIT_CONTEXT);
    if (!context || !context.payload || !cleanText(context.formId)) {
      showFeedback("warn", "Nenhum envio anterior disponivel para edicao neste navegador.");
      return;
    }

    var payload = context.payload;
    pendingEditFormId = cleanText(context.formId);
    applyPayloadToForm_(payload);
    if (submitButtonEl) {
      submitButtonEl.textContent = "Salvar edicao";
    }
    showFeedback("warn", "Modo edicao ativo para " + pendingEditFormId + ". Revise e clique em Salvar edicao.");
  }

  function applyPayloadToForm_(payload) {
    if (!payload || !Array.isArray(payload.projects) || !payload.projects.length) {
      return;
    }

    var professional = cleanText(payload.professional);
    var workDate = cleanText(payload.date);

    if (professionalEl && professional) {
      professionalEl.value = professional;
      rememberLastProfessional_(professional);
    }
    if (dateEl && workDate) {
      dateEl.value = workDate;
    }

    var selectedCodes = payload.projects
      .map(function (item) {
        return cleanText(item && item.projectCode);
      })
      .filter(Boolean);

    populateProjects();
    selectProjectCodes_(selectedCodes);
    renderProjectSections(selectedCodes);

    payload.projects.forEach(function (project) {
      var code = cleanText(project && project.projectCode);
      if (!code) {
        return;
      }
      applySectionValues_(code, {
        refEap: cleanText(project && project.refEap),
        taskRefEap: cleanText(project && project.taskRefEap),
        taskFreeText: cleanText(project && project.taskFreeText),
        statusCode: cleanText(project && project.statusCode),
        blockCode: cleanText(project && project.blockCode),
        blockDescription: cleanText(project && project.blockDescription),
        observation: cleanText(project && project.observation)
      });
    });

    updateFormChecklist_();
    renderDailyFocus_();
    scheduleDraftSave_();
  }

  function getOfflineQueue_() {
    var queue = getStorageJson_(STORAGE_KEYS.OFFLINE_QUEUE);
    if (!Array.isArray(queue)) {
      return [];
    }
    return queue.filter(function (item) {
      return item && item.payload;
    });
  }

  function setOfflineQueue_(queue) {
    setStorageJson_(STORAGE_KEYS.OFFLINE_QUEUE, Array.isArray(queue) ? queue : []);
    updateOfflineQueueUi_();
  }

  function enqueueOfflineSubmission_(payload) {
    var queue = getOfflineQueue_();
    queue.push({
      queuedAt: new Date().toISOString(),
      payload: payload
    });
    setOfflineQueue_(queue);
  }

  function updateOfflineQueueUi_() {
    if (!offlineQueueEl || !offlineQueueTextEl) {
      return;
    }
    var queue = getOfflineQueue_();
    if (!queue.length) {
      offlineQueueEl.classList.add("hidden");
      offlineQueueTextEl.textContent = "Sem envios pendentes.";
      return;
    }

    offlineQueueEl.classList.remove("hidden");
    offlineQueueTextEl.textContent = queue.length + " envio(s) pendente(s) aguardando internet.";
  }

  function flushOfflineQueue_() {
    if (queueRetryRunning) {
      return;
    }
    if (!navigator.onLine) {
      updateOfflineQueueUi_();
      return;
    }

    var appsScriptUrl = cleanText(cfg.appsScriptUrl);
    if (!appsScriptUrl || appsScriptUrl.indexOf("COLE_AQUI") >= 0) {
      return;
    }

    var queue = getOfflineQueue_();
    if (!queue.length) {
      updateOfflineQueueUi_();
      return;
    }

    queueRetryRunning = true;
    processOfflineQueueItem_(appsScriptUrl, queue, 0)
      .then(function (result) {
        setOfflineQueue_(result.remaining);
        if (result.sent > 0) {
          loadRefOptionsFromServer();
          loadPpmSnapshotFromServer();
          loadHistoryFromServer_();
          showFeedback("ok", "Envios pendentes reenviados: " + result.sent + ".");
        }
      })
      .finally(function () {
        queueRetryRunning = false;
      });
  }

  function processOfflineQueueItem_(url, queue, sentCount) {
    if (!queue.length) {
      return Promise.resolve({ remaining: [], sent: sentCount });
    }

    var current = queue[0];
    return sendPayload(url, current.payload)
      .then(function (result) {
        if (result.mode === "cors" && result.data && result.data.status === "error") {
          return { remaining: queue, sent: sentCount };
        }
        return processOfflineQueueItem_(url, queue.slice(1), sentCount + 1);
      })
      .catch(function () {
        return { remaining: queue, sent: sentCount };
      });
  }

  function shouldQueueSubmission_(error) {
    var message = cleanText(error && error.message).toLowerCase();
    if (!message) {
      return true;
    }
    if (message.indexOf("falha de comunicacao") >= 0) {
      return true;
    }
    if (message.indexOf("network") >= 0 || message.indexOf("fetch") >= 0 || message.indexOf("offline") >= 0) {
      return true;
    }
    return false;
  }

  function applyPortfolioFilter_(dataset) {
    var mode = cleanText(portfolioFilterEl && portfolioFilterEl.value);
    if (mode === "coord-critical") {
      return dataset
        .filter(function (item) {
          return item.semaforo === "red" || Number(item.atrasadasAteCorte || 0) > 0 || Number(item.bloqueada || 0) > 0;
        })
        .sort(compareMeetingPriority_);
    }

    if (mode === "coord-risk") {
      return dataset
        .filter(function (item) {
          return item.semaforo === "yellow" || item.semaforo === "red";
        })
        .sort(compareMeetingPriority_);
    }

    if (mode === "coord-blocked") {
      return dataset
        .filter(function (item) {
          return Number(item.bloqueada || 0) > 0;
        })
        .sort(compareMeetingPriority_);
    }

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

  function applyPortfolioSort_(dataset) {
    var items = Array.isArray(dataset) ? dataset.slice() : [];
    var mode = cleanText(portfolioSortEl && portfolioSortEl.value);

    if (mode !== "signature-asc" && mode !== "signature-desc") {
      return items;
    }

    var factor = mode === "signature-desc" ? -1 : 1;
    items.sort(function (a, b) {
      var aHasDate = a && a.signatureDate instanceof Date;
      var bHasDate = b && b.signatureDate instanceof Date;

      if (!aHasDate && !bHasDate) {
        return String(a && a.code ? a.code : "").localeCompare(String(b && b.code ? b.code : ""));
      }
      if (!aHasDate) {
        return 1;
      }
      if (!bHasDate) {
        return -1;
      }

      var diff = (a.signatureDate.getTime() - b.signatureDate.getTime()) * factor;
      if (diff !== 0) {
        return diff;
      }
      return String(a && a.code ? a.code : "").localeCompare(String(b && b.code ? b.code : ""));
    });

    return items;
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
      var signatureDate = getProjectSignatureDate_(code, snapshotProject, project);
      var daysSinceSignature = calculateDaysSince_(signatureDate, cutoffDate);

      data.push({
        code: code,
        label: cleanText(project.label) || code,
        signatureDate: signatureDate,
        daysSinceSignature: daysSinceSignature,
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
      return buildProjectScheduleFromSnapshot_(snapshotTasks);
    }

    return buildProjectScheduleFromOptions_(optionList);
  }

  function buildProjectScheduleFromOptions_(optionList) {
    var tasks = [];
    (optionList || []).forEach(function (item) {
      var refEap = cleanText(item && item.value);
      if (!refEap) {
        return;
      }

      var parsed = parseRefOptionLabel_(cleanText(item && item.label));
      tasks.push({
        refEap: refEap,
        phase: parsed.phase,
        task: parsed.task,
        status: normalizePortfolioStatus_(item && item.status),
        plannedStart: null,
        plannedEnd: null,
        realStart: null,
        realEnd: null,
        lastRecordDate: null,
        updatedAt: null,
        blocked: false,
        blockReason: ""
      });
    });

    tasks.sort(function (a, b) {
      return String(a.refEap).localeCompare(String(b.refEap));
    });
    return tasks;
  }

  function buildProjectScheduleFromSnapshot_(snapshotTasks) {
    var tasks = [];
    (snapshotTasks || []).forEach(function (snapshotTask) {
      var refEap = cleanText(snapshotTask && snapshotTask.refEap);
      if (!refEap) {
        return;
      }
      tasks.push({
        refEap: refEap,
        phase: cleanText(snapshotTask && snapshotTask.phase),
        task: cleanText(snapshotTask && snapshotTask.task),
        status: normalizePortfolioStatus_(snapshotTask && snapshotTask.status),
        plannedStart: parseDateFlexible_(snapshotTask && snapshotTask.plannedStart),
        plannedEnd: parseDateFlexible_(snapshotTask && snapshotTask.plannedEnd),
        realStart: parseDateFlexible_(snapshotTask && snapshotTask.realStart),
        realEnd: parseDateFlexible_(snapshotTask && snapshotTask.realEnd),
        lastRecordDate: parseDateFlexible_(snapshotTask && snapshotTask.lastRecordDate),
        updatedAt: parseDateFlexible_(snapshotTask && snapshotTask.updatedAt),
        blocked: Boolean(snapshotTask && snapshotTask.blocked),
        blockReason: cleanText(snapshotTask && snapshotTask.blockReason)
      });
    });

    tasks.sort(function (a, b) {
      return String(a.refEap).localeCompare(String(b.refEap));
    });
    return tasks;
  }

  function parseRefOptionLabel_(label) {
    var text = cleanText(label);
    if (!text) {
      return { phase: "", task: "" };
    }
    var parts = text.split("|").map(function (item) {
      return cleanText(item);
    });
    return {
      phase: parts.length > 1 ? parts[1] : "",
      task: parts.length > 2 ? parts.slice(2).join(" | ") : ""
    };
  }

  function getProjectSignatureDate_(projectCode, snapshotProject, project) {
    var fromProject = parseIsoDate_(cleanText(project && project.signatureDate));
    if (fromProject) {
      return fromProject;
    }

    var fromSnapshot = parseIsoDate_(cleanText(snapshotProject && snapshotProject.signatureDate));
    if (fromSnapshot) {
      return fromSnapshot;
    }

    var fromGlobalMap = parseIsoDate_(cleanText(PROJECT_SIGNATURE_DATE[projectCode]));
    if (fromGlobalMap) {
      return fromGlobalMap;
    }

    var fallback = parseIsoDate_(DEFAULT_SIGNATURE_DATE);
    if (fallback) {
      return fallback;
    }

    return getCutoffDate_();
  }

  function calculateDaysSince_(startDate, endDate) {
    if (!(startDate instanceof Date) || !(endDate instanceof Date)) {
      return null;
    }
    var start = cloneDate_(startDate);
    var end = cloneDate_(endDate);
    var diffMs = end.getTime() - start.getTime();
    if (diffMs < 0) {
      return 0;
    }
    return Math.floor(diffMs / (24 * 60 * 60 * 1000));
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
        var signatureMeta = buildSignatureMetaText_(item);

        return [
          '<article class="portfolio-card" data-contract-id="' + escapeAttr(item.code) + '" tabindex="0" role="button" aria-label="Abrir coordenacao de ' + escapeAttr(item.label) + '">',
          '<div class="portfolio-card-head">',
          "<h3>" + escapeHtml(item.label) + "</h3>",
          '<span class="portfolio-badge ' + escapeAttr(item.semaforo) + '">' + escapeHtml(labelSemaforo_(item.semaforo)) + "</span>",
          "</div>",
          '<p class="portfolio-contract-meta">' + escapeHtml(signatureMeta) + "</p>",
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
          '<button class="portfolio-config-btn" type="button" data-open-coord="' + escapeAttr(item.code) + '">Configurar projeto</button>',
          "</article>"
        ].join("");
      })
      .join("");
  }

  function handlePortfolioCardClick_(event) {
    var target = event && event.target;
    if (!target) {
      return;
    }

    var openButton = target.closest("[data-open-coord]");
    var card = target.closest(".portfolio-card");
    if (!openButton && !card) {
      return;
    }

    var contractId = "";
    if (openButton) {
      contractId = cleanText(openButton.getAttribute("data-open-coord"));
    } else if (card) {
      contractId = cleanText(card.getAttribute("data-contract-id"));
    }

    if (!contractId) {
      return;
    }

    if (openButton) {
      event.preventDefault();
      event.stopPropagation();
    }

    openCoordPanel_(contractId);
  }

  function handlePortfolioCardKeydown_(event) {
    if (!event) {
      return;
    }
    var key = event.key || "";
    if (key !== "Enter" && key !== " ") {
      return;
    }

    var card = event.target && event.target.closest ? event.target.closest(".portfolio-card") : null;
    if (!card) {
      return;
    }

    event.preventDefault();
    var contractId = cleanText(card.getAttribute("data-contract-id"));
    if (!contractId) {
      return;
    }

    openCoordPanel_(contractId);
  }

  function handleCoordModalKeydown_(event) {
    if (!coordState.open || !event) {
      return;
    }
    if ((event.key || "") === "Escape") {
      closeCoordPanel_();
    }
  }

  function openCoordPanel_(contractId) {
    var targetContractId = cleanText(contractId);
    var project = projectMap[targetContractId] || {};
    coordState.open = true;
    coordState.contractId = targetContractId;
    coordState.projectName = cleanText(project.label || targetContractId);
    coordState.signatureDate = cleanText(PROJECT_SIGNATURE_DATE[targetContractId] || "");
    coordState.daysSinceSignature = calculateDaysSinceFromIso_(coordState.signatureDate);
    coordState.deadlineSummary = cleanText(project.deadlineSummary);
    coordState.tasks = [];
    coordState.events = [];
    coordState.warnings = [];
    coordState.originalTasksByRef = {};
    coordState.originalEventsKey = "";

    if (coordModalEl) {
      coordModalEl.classList.remove("hidden");
      coordModalEl.setAttribute("aria-hidden", "false");
    }
    if (coordTitleEl) {
      coordTitleEl.textContent = "Configuracao do projeto " + targetContractId;
    }
    if (coordSubtitleEl) {
      coordSubtitleEl.textContent = "Carregando dados de coordenacao...";
    }
    if (coordContractMetaEl) {
      coordContractMetaEl.textContent = "Assinatura contratual: carregando...";
    }
    if (coordDeadlineMetaEl) {
      coordDeadlineMetaEl.textContent = "Prazos contratuais (resumo): carregando...";
    }
    if (coordTaskListEl) {
      coordTaskListEl.innerHTML = '<div class="portfolio-empty">Carregando tarefas do projeto...</div>';
    }
    if (coordEventsListEl) {
      coordEventsListEl.innerHTML = "";
    }
    clearCoordFeedback_();
    document.body.style.overflow = "hidden";

    loadCoordProjectFromServer_(targetContractId)
      .then(function (payload) {
        coordState.tasks = payload.tasks;
        coordState.events = payload.events;
        coordState.warnings = payload.warnings;
        coordState.projectName = resolvePreferredProjectName_(
          coordState.contractId,
          payload.projectName,
          coordState.projectName
        );
        coordState.signatureDate = payload.signatureDate || "";
        coordState.daysSinceSignature = payload.daysSinceSignature;
        coordState.deadlineSummary = payload.deadlineSummary || coordState.deadlineSummary;
        coordState.originalTasksByRef = {};
        payload.tasks.forEach(function (task) {
          coordState.originalTasksByRef[cleanText(task.refEap)] = {
            status: cleanText(task.status),
            responsible: cleanText(task.responsible),
            plannedStart: cleanText(task.plannedStart),
            plannedEnd: cleanText(task.plannedEnd),
            prazoDu: String(Number(task.prazoDu || 1)),
            predecessorRef: cleanText(task.predecessorRef),
            relationType: cleanText(task.relationType) || "FS",
            lagDu: String(Number(task.lagDu || 0)),
            percentReal: String(Number(task.percentReal || 0)),
            blocked: task.blocked ? "Sim" : "Nao",
            blockReason: cleanText(task.blockReason)
          };
        });
        coordState.originalEventsKey = coordEventsKey_(payload.events);
        renderCoordPanel_();
      })
      .catch(function (error) {
        if (coordSubtitleEl) {
          coordSubtitleEl.textContent = "Nao foi possivel carregar as configuracoes deste projeto.";
        }
        if (coordContractMetaEl) {
          coordContractMetaEl.textContent = "Assinatura contratual: nao foi possivel carregar.";
        }
        if (coordDeadlineMetaEl) {
          coordDeadlineMetaEl.textContent = "Prazos contratuais (resumo): nao foi possivel carregar.";
        }
        if (coordTaskListEl) {
          coordTaskListEl.innerHTML = '<div class="portfolio-empty">Erro ao carregar tarefas. Tente novamente.</div>';
        }
        showCoordFeedback_("error", cleanText(error && error.message) || "Erro ao carregar dados de coordenacao.");
      });
  }

  function closeCoordPanel_() {
    coordState.open = false;
    coordState.deadlineSummary = "";
    if (coordModalEl) {
      coordModalEl.classList.add("hidden");
      coordModalEl.setAttribute("aria-hidden", "true");
    }
    if (coordTaskListEl) {
      coordTaskListEl.innerHTML = "";
    }
    if (coordEventsListEl) {
      coordEventsListEl.innerHTML = "";
    }
    if (coordContractMetaEl) {
      coordContractMetaEl.textContent = "";
    }
    if (coordDeadlineMetaEl) {
      coordDeadlineMetaEl.textContent = "";
    }
    clearCoordFeedback_();
    document.body.style.overflow = "";
  }

  function loadCoordProjectFromServer_(contractId) {
    return new Promise(function (resolve, reject) {
      var appsScriptUrl = cleanText(cfg.appsScriptUrl);
      if (!appsScriptUrl || appsScriptUrl.indexOf("COLE_AQUI") >= 0) {
        reject(new Error("Backend nao configurado."));
        return;
      }

      var callbackName = "vmCoordCb_" + String(Date.now());
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
          if (!data || data.status !== "ok" || !data.project) {
            reject(new Error((data && data.message) || "Falha ao carregar configuracoes."));
            return;
          }
          resolve(normalizeCoordProjectPayload_(data.project));
        } finally {
          finish();
        }
      };

      script.async = true;
      script.src = buildUrlWithParams(appsScriptUrl, {
        action: "coord_project",
        contractId: contractId,
        callback: callbackName,
        _: Date.now()
      });
      script.onerror = function () {
        reject(new Error("Erro de conexao ao carregar configuracoes."));
        finish();
      };
      document.body.appendChild(script);
      setTimeout(function () {
        if (!settled) {
          reject(new Error("Tempo de resposta excedido ao carregar configuracoes."));
          finish();
        }
      }, JSONP_TIMEOUT_MS);
    });
  }

  function normalizeCoordProjectPayload_(project) {
    var tasks = Array.isArray(project.tasks) ? project.tasks : [];
    var events = Array.isArray(project.events) ? project.events : [];
    var warnings = Array.isArray(project.warnings) ? project.warnings : [];
    var contractId = cleanText(project.contractId);
    var signatureDate = cleanText(project.signatureDate);
    var deadlineSummary = cleanText(project.contractDeadlineSummary || project.deadlineSummary);
    var daysSinceSignature = null;

    if (!signatureDate && contractId) {
      signatureDate = cleanText(PROJECT_SIGNATURE_DATE[contractId] || "");
    }

    if (project && project.daysSinceSignature !== undefined && project.daysSinceSignature !== null && cleanText(project.daysSinceSignature) !== "") {
      daysSinceSignature = Number(project.daysSinceSignature);
      if (isNaN(daysSinceSignature) || daysSinceSignature < 0) {
        daysSinceSignature = null;
      }
    }
    if (daysSinceSignature === null) {
      daysSinceSignature = calculateDaysSinceFromIso_(signatureDate);
    }

    if (!deadlineSummary && contractId && projectMap[contractId]) {
      deadlineSummary = cleanText(projectMap[contractId].deadlineSummary);
    }

    return {
      contractId: contractId,
      projectName: cleanText(project.projectName),
      signatureDate: signatureDate,
      daysSinceSignature: daysSinceSignature,
      deadlineSummary: deadlineSummary,
      summary: project.summary || {},
      tasks: tasks.map(function (task) {
        return {
          refEap: cleanText(task.refEap),
          phase: cleanText(task.phase),
          task: cleanText(task.task),
          status: normalizePortfolioStatus_(task.status),
          responsible: cleanText(task.responsible),
          plannedStart: cleanText(task.plannedStart),
          plannedEnd: cleanText(task.plannedEnd),
          prazoDu: Number(task.prazoDu || 0),
          predecessorRef: cleanText(task.predecessorRef),
          relationType: cleanText(task.relationType) || "FS",
          lagDu: Number(task.lagDu || 0),
          percentReal: Number(task.percentReal || 0),
          blocked: !!task.blocked,
          blockReason: cleanText(task.blockReason),
          updatedAt: cleanText(task.updatedAt),
          lastRecordDate: cleanText(task.lastRecordDate)
        };
      }),
      events: events.map(function (event) {
        return {
          eventDate: cleanText(event.eventDate),
          eventType: cleanText(event.eventType),
          title: cleanText(event.title),
          responsible: cleanText(event.responsible),
          status: cleanText(event.status),
          observation: cleanText(event.observation)
        };
      }),
      warnings: warnings.map(function (warning) {
        return cleanText(warning);
      }).filter(Boolean)
    };
  }

  function calculateDaysSinceFromIso_(isoDate) {
    var signature = parseIsoDate_(isoDate);
    if (!signature) {
      return null;
    }
    return calculateDaysSince_(signature, getCutoffDate_());
  }

  function resolvePreferredProjectName_(contractId, incomingName, currentName) {
    var incoming = cleanText(incomingName);
    var current = cleanText(currentName);
    var localFromMap = cleanText((projectMap[contractId] || {}).label);
    var local = localFromMap || current;

    if (!incoming) {
      return local || contractId;
    }
    if (isContractDisplayName_(incoming)) {
      return incoming;
    }
    if (isContractDisplayName_(local)) {
      return local;
    }
    return incoming || local || contractId;
  }

  function renderCoordPanel_() {
    if (!coordTaskListEl || !coordEventsListEl) {
      return;
    }

    if (coordTitleEl) {
      coordTitleEl.textContent = "Configuracao do projeto " + coordState.contractId;
    }
    if (coordSubtitleEl) {
      coordSubtitleEl.textContent =
        (cleanText(coordState.projectName) || coordState.contractId) +
        " | Edite sequenciamento, datas, responsaveis, status, bloqueios e eventos.";
    }
    if (coordContractMetaEl) {
      coordContractMetaEl.textContent = buildCoordSignatureMeta_(
        cleanText(coordState.signatureDate),
        coordState.daysSinceSignature
      );
    }
    if (coordDeadlineMetaEl) {
      coordDeadlineMetaEl.innerHTML = buildCoordDeadlineMetaHtml_(cleanText(coordState.deadlineSummary));
    }

    renderCoordWarnings_();

    var statusOptions = ["NAO INICIADA", "EM ANDAMENTO", "BLOQUEADA", "CONCLUIDA"];
    var relationOptions = ["FS", "SS", "FF", "SF"];
    var professionals = Array.isArray(cfg.profissionais) ? cfg.profissionais : [];
    var tasksHtml = coordState.tasks
      .map(function (task, index) {
        var ref = cleanText(task.refEap);
        var plannedStart = formatDateInput_(task.plannedStart);
        var plannedEnd = formatDateInput_(task.plannedEnd);
        var duration = Math.max(1, Number(task.prazoDu || 1));
        var predecessorRef = cleanText(task.predecessorRef);
        var relationType = cleanText(task.relationType) || "FS";
        var lagDu = Number(task.lagDu || 0);
        var status = normalizePortfolioStatus_(task.status);
        var blocked = task.blocked ? "Sim" : "Nao";
        var tagClass = getCoordStatusTagClass_(status);
        var statusOptionsHtml = statusOptions
          .map(function (option) {
            return (
              '<option value="' +
              escapeAttr(option) +
              '"' +
              (option === status ? " selected" : "") +
              ">" +
              escapeHtml(option) +
              "</option>"
            );
          })
          .join("");
        var relationOptionsHtml = relationOptions
          .map(function (option) {
            return (
              '<option value="' +
              escapeAttr(option) +
              '"' +
              (option === relationType ? " selected" : "") +
              ">" +
              escapeHtml(option) +
              "</option>"
            );
          })
          .join("");
        var responsibleOptionsHtml = ['<option value="">Nao definido</option>']
          .concat(
            professionals.map(function (name) {
              var value = cleanText(name);
              return (
                '<option value="' +
                escapeAttr(value) +
                '"' +
                (value === cleanText(task.responsible) ? " selected" : "") +
                ">" +
                escapeHtml(value) +
                "</option>"
              );
            })
          )
          .join("");

        return [
          '<article class="coord-task-item" data-ref-eap="' + escapeAttr(ref) + '" data-status="' + escapeAttr(status) + '" data-index="' + String(index) + '">',
          '<div class="coord-task-main">',
          "<div>",
          "<strong>" + escapeHtml(ref + " - " + (task.task || "Sem descricao")) + "</strong>",
          "<p>" + escapeHtml(task.phase || "Sem fase") + "</p>",
          "</div>",
          '<span class="coord-task-tag ' + escapeAttr(tagClass) + '">' + escapeHtml(status) + "</span>",
          "</div>",
          '<div class="coord-task-grid">',
          '<div class="coord-field"><label>Status</label><select class="coord-status">' + statusOptionsHtml + "</select></div>",
          '<div class="coord-field"><label>Responsavel</label><select class="coord-responsible">' + responsibleOptionsHtml + "</select></div>",
          '<div class="coord-field"><label>Prazo DU</label><input class="coord-prazo-du" type="number" min="1" max="120" step="1" value="' + escapeAttr(String(duration)) + '" /></div>',
          '<div class="coord-field long"><label>Predecessora (REF EAP; use ; para multiplas)</label><input class="coord-predecessor" type="text" maxlength="200" value="' + escapeAttr(predecessorRef) + '" placeholder="Ex: CT117-WBS-1_1_1; CT117-WBS-1_1_2" /></div>',
          '<div class="coord-field"><label>Relacao</label><select class="coord-relation">' + relationOptionsHtml + "</select></div>",
          '<div class="coord-field"><label>Lag DU</label><input class="coord-lag-du" type="number" min="-30" max="30" step="1" value="' + escapeAttr(String(lagDu)) + '" /></div>',
          '<div class="coord-field"><label>Inicio planejado</label><input class="coord-planned-start" type="date" value="' + escapeAttr(plannedStart) + '" /></div>',
          '<div class="coord-field"><label>Fim planejado</label><input class="coord-planned-end" type="date" value="' + escapeAttr(plannedEnd) + '" /></div>',
          '<div class="coord-field"><label>% Real</label><input class="coord-percent-real" type="number" min="0" max="100" step="1" value="' + escapeAttr(String(Math.max(0, Math.min(100, Number(task.percentReal || 0))))) + '" /></div>',
          '<div class="coord-field"><label>Bloqueio</label><select class="coord-blocked"><option value="Nao"' + (blocked === "Nao" ? " selected" : "") + '>Nao</option><option value="Sim"' + (blocked === "Sim" ? " selected" : "") + ">Sim</option></select></div>",
          '<div class="coord-field long"><label>Desc. bloqueio</label><input class="coord-block-reason" type="text" maxlength="180" value="' + escapeAttr(cleanText(task.blockReason)) + '" /></div>',
          "</div>",
          "</article>"
        ].join("");
      })
      .join("");

    coordTaskListEl.innerHTML = tasksHtml || '<div class="portfolio-empty">Nenhuma tarefa encontrada para este projeto.</div>';
    coordTaskListEl.removeEventListener("change", handleCoordTaskFieldChange_);
    coordTaskListEl.addEventListener("change", handleCoordTaskFieldChange_);
    coordTaskListEl.addEventListener("input", handleCoordTaskFieldChange_);
    renderCoordEvents_(coordState.events);
    applyCoordFilter_();
  }

  function renderCoordWarnings_() {
    if (!coordAlertsEl) {
      return;
    }
    var warnings = Array.isArray(coordState.warnings) ? coordState.warnings : [];
    if (!warnings.length) {
      coordAlertsEl.classList.add("hidden");
      coordAlertsEl.innerHTML = "";
      return;
    }

    coordAlertsEl.classList.remove("hidden");
    coordAlertsEl.innerHTML =
      "<h5>Alertas de sequenciamento</h5><ul>" +
      warnings.map(function (warning) {
        return "<li>" + escapeHtml(cleanText(warning)) + "</li>";
      }).join("") +
      "</ul>";
  }

  function buildCoordSignatureMeta_(signatureDate, daysSinceSignature) {
    if (!cleanText(signatureDate)) {
      return "Assinatura contratual: nao informada | Dias de contrato: --";
    }

    var days = Number(daysSinceSignature);
    if (isNaN(days) || days < 0) {
      days = 0;
    }

    return (
      "Assinatura contratual: " +
      formatIsoDateToBr(signatureDate) +
      " | Dias de contrato: " +
      days
    );
  }

  function buildCoordDeadlineMetaHtml_(deadlineSummary) {
    var items = parseDeadlineSummaryItems_(deadlineSummary);
    if (!items.length) {
      return (
        '<span class="coord-deadline-title">Prazos contratuais (resumo)</span>' +
        '<span class="coord-deadline-empty">Nao cadastrado neste projeto.</span>'
      );
    }

    return (
      '<span class="coord-deadline-title">Prazos contratuais (resumo)</span>' +
      '<span class="coord-deadline-chips">' +
      items
        .map(function (item) {
          return (
            '<span class="coord-deadline-chip" tabindex="0" role="note" data-tip="' +
            escapeAttr(item.description) +
            '" aria-label="' +
            escapeAttr(item.description) +
            '">' +
            escapeHtml(item.label) +
            "</span>"
          );
        })
        .join("") +
      "</span>"
    );
  }

  function parseDeadlineSummaryItems_(deadlineSummary) {
    var summary = cleanText(deadlineSummary) || cleanText(DEFAULT_DEADLINE_SUMMARY);
    if (!summary) {
      return [];
    }

    if (summary.indexOf("|") >= 0) {
      return summary
        .split("|")
        .map(function (part) {
          return cleanText(part);
        })
        .filter(Boolean)
        .map(function (label) {
          return {
            label: label,
            description: buildDeadlineTooltipDescription_(label)
          };
        });
    }

    var detailed = parseDetailedDeadlineText_(summary);
    if (detailed.length) {
      return detailed;
    }

    return [
      {
        label: "PRAZOS (TEXTO)",
        description: summary
      }
    ];
  }

  function parseDetailedDeadlineText_(text) {
    var normalized = cleanText(text).replace(/\r/g, " ").replace(/\n+/g, " ").replace(/\s+/g, " ");
    if (!normalized) {
      return [];
    }

    var section41 = extractDeadlineSection_(normalized, /4\.1\b/i, /4\.2\b/i);
    var section42 = extractDeadlineSection_(normalized, /4\.2\b/i, /4\.3\b/i);
    var section43 = extractDeadlineSection_(normalized, /4\.3\b/i, /4\.4\b/i);
    var section44 = extractDeadlineSection_(normalized, /4\.4\b/i, /4\.5\b/i);
    var section45 = extractDeadlineSection_(normalized, /4\.5\b/i, /4\.6\b/i);
    var section46 = extractDeadlineSection_(normalized, /4\.6\b/i, /4\.7\b/i);
    var section47 = extractDeadlineSection_(normalized, /4\.7\b/i, /(?:5\.|$)/i);
    var items = [];
    var rafDu = [];

    function pushItem(sigla, description) {
      var cleanDesc = cleanText(description);
      if (!cleanDesc) {
        return "";
      }
      var du = extractDueDaysFromText_(cleanDesc);
      var label = sigla;
      if (du) {
        label += " " + du + "DU";
      }
      items.push({
        label: label,
        description: cleanDesc
      });
      return du;
    }

    var f1 = extractPhaseText_(section41, "I", "II");
    var f2 = extractPhaseText_(section41, "II", "III");
    var f3 = extractPhaseText_(section41, "III", "");
    pushItem("F1", f1);
    pushItem("F2", f2);
    pushItem("F3", f3);

    rafDu[0] = extractDueDaysFromText_(extractPhaseText_(section42, "I", "II"));
    rafDu[1] = extractDueDaysFromText_(extractPhaseText_(section42, "II", "III"));
    rafDu[2] = extractDueDaysFromText_(extractPhaseText_(section42, "III", ""));
    var rafDescription = cleanText(section42).replace(/^4\.2\.?\s*/i, "").trim();
    if (rafDescription) {
      var rafLabel = "RAF1/2/3";
      if (rafDu[0] || rafDu[1] || rafDu[2]) {
        rafLabel += " " + (rafDu[0] || "?") + "/" + (rafDu[1] || "?") + "/" + (rafDu[2] || "?") + "DU";
      }
      items.push({
        label: rafLabel,
        description: rafDescription
      });
    }

    pushItem("PRE", cleanText(section43).replace(/^4\.3\.?\s*/i, "").trim());
    pushItem("RFA", cleanText(section44).replace(/^4\.4\.?\s*/i, "").trim());
    pushItem("EXE", cleanText(section45).replace(/^4\.5\.?\s*/i, "").trim());
    pushItem("DET", cleanText(section46).replace(/^4\.6\.?\s*/i, "").trim());
    pushItem("RDE", cleanText(section47).replace(/^4\.7\.?\s*/i, "").trim());

    if (/recesso|natal|ano novo/i.test(normalized)) {
      items.push({
        label: "RECESSO",
        description: extractRecessoText_(normalized)
      });
    }

    return items;
  }

  function extractDeadlineSection_(text, startRegex, endRegex) {
    var full = cleanText(text);
    if (!full) {
      return "";
    }

    var startMatch = full.match(startRegex);
    if (!startMatch) {
      return "";
    }

    var startIndex = startMatch.index || 0;
    var tail = full.slice(startIndex);
    var endMatch = tail.match(endRegex);
    if (endMatch && endMatch.index > 0) {
      tail = tail.slice(0, endMatch.index);
    }
    return cleanText(tail);
  }

  function extractPhaseText_(sectionText, currentRoman, nextRoman) {
    var section = cleanText(sectionText);
    if (!section) {
      return "";
    }
    var pattern =
      "Fase\\s*" +
      currentRoman +
      "\\s*[:\\-]\\s*([\\s\\S]*?)" +
      (nextRoman ? "(?=Fase\\s*" + nextRoman + "\\s*[:\\-]|$)" : "$");
    var match = section.match(new RegExp(pattern, "i"));
    if (!match) {
      return "";
    }
    return cleanText(match[1]).replace(/^[\s;,-]+/, "").replace(/[\s;,-]+$/, "");
  }

  function extractDueDaysFromText_(text) {
    var source = cleanText(text);
    if (!source) {
      return "";
    }
    var match = source.match(/(\d{1,3})\s*(?:\([^)]+\))?\s*dias?\s*uteis?/i) || source.match(/(\d{1,3})\s*DU\b/i);
    return match ? cleanText(match[1]) : "";
  }

  function extractRecessoText_(text) {
    var normalized = cleanText(text);
    if (!normalized) {
      return "Nao sao contabilizados os prazos no periodo de recesso entre Natal e Ano Novo.";
    }
    var match = normalized.match(/[^.]*recesso[^.]*\./i);
    if (match) {
      return cleanText(match[0]);
    }
    return "Nao sao contabilizados os prazos no periodo de recesso entre Natal e Ano Novo.";
  }

  function buildDeadlineTooltipDescription_(tokenLabel) {
    var key = getDeadlineTokenKey_(tokenLabel);
    var description = "";

    if (key) {
      description = cleanText(deadlineLegendMap[key]);
    }

    if (!description && (key === "F1" || key === "F2" || key === "F3")) {
      description = cleanText(deadlineLegendMap["F1/F2/F3"]);
    }

    if (!description && key && key.indexOf("RAF") === 0) {
      description = cleanText(deadlineLegendMap["RAF1/2/3"] || deadlineLegendMap.RAF);
    }

    if (!description && /DU\b/i.test(tokenLabel)) {
      description = cleanText(deadlineLegendMap.DU);
    }

    if (!description && key === "RECESSO") {
      description = cleanText(deadlineLegendMap.RECESSO);
    }

    if (!description) {
      return cleanText(tokenLabel);
    }

    return cleanText(tokenLabel) + " - " + description;
  }

  function getDeadlineTokenKey_(tokenLabel) {
    var token = cleanText(tokenLabel).toUpperCase();
    if (!token) {
      return "";
    }

    if (token.indexOf("RECESSO") === 0) {
      return "RECESSO";
    }
    if (token.indexOf("RAF1/2/3") === 0) {
      return "RAF1/2/3";
    }
    if (token.indexOf("F1") === 0) {
      return "F1";
    }
    if (token.indexOf("F2") === 0) {
      return "F2";
    }
    if (token.indexOf("F3") === 0) {
      return "F3";
    }
    if (token.indexOf("PRE") === 0) {
      return "PRE";
    }
    if (token.indexOf("RFA") === 0) {
      return "RFA";
    }
    if (token.indexOf("EXE") === 0) {
      return "EXE";
    }
    if (token.indexOf("DET") === 0) {
      return "DET";
    }
    if (token.indexOf("RDE") === 0) {
      return "RDE";
    }
    return "";
  }

  function buildDeadlineLegendMap_(legendItems) {
    var map = Object.assign({}, DEFAULT_DEADLINE_LEGEND_MAP);
    if (!Array.isArray(legendItems)) {
      return map;
    }

    legendItems.forEach(function (item) {
      var sigla = cleanText(item && item.sigla).toUpperCase();
      var descricao = cleanText(item && item.descricao);
      if (!sigla || !descricao) {
        return;
      }

      map[sigla] = descricao;

      if (sigla === "F1/F2/F3") {
        map.F1 = descricao;
        map.F2 = descricao;
        map.F3 = descricao;
      }
      if (sigla === "RAF1/2/3") {
        map.RAF = descricao;
      }
      if (sigla.indexOf("RECESSO") >= 0) {
        map.RECESSO = descricao;
      }
      if (sigla.indexOf("DU") >= 0) {
        map.DU = descricao;
      }
    });

    return map;
  }

  function renderCoordEvents_(events) {
    if (!coordEventsListEl) {
      return;
    }

    var list = Array.isArray(events) ? events : [];
    if (!list.length) {
      coordEventsListEl.innerHTML = '<div class="portfolio-empty">Sem eventos programados para este projeto.</div>';
      return;
    }

    coordEventsListEl.innerHTML = list
      .map(function (event, index) {
        return buildCoordEventItemHtml_(event, index);
      })
      .join("");
  }

  function buildCoordEventItemHtml_(eventData, index) {
    var professionals = Array.isArray(cfg.profissionais) ? cfg.profissionais : [];
    var event = eventData || {};
    var responsibleOptions = ['<option value="">Nao definido</option>']
      .concat(
        professionals.map(function (name) {
          var value = cleanText(name);
          return (
            '<option value="' +
            escapeAttr(value) +
            '"' +
            (value === cleanText(event.responsible) ? " selected" : "") +
            ">" +
            escapeHtml(value) +
            "</option>"
          );
        })
      )
      .join("");

    return [
      '<article class="coord-event-item" data-event-index="' + String(index) + '">',
      '<div class="coord-event-grid">',
      '<div class="coord-field"><label>Data</label><input class="coord-event-date" type="date" value="' + escapeAttr(formatDateInput_(event.eventDate)) + '" /></div>',
      '<div class="coord-field"><label>Tipo</label><input class="coord-event-type" type="text" maxlength="60" value="' + escapeAttr(cleanText(event.eventType)) + '" placeholder="Aprovacao, reuniao..." /></div>',
      '<div class="coord-field"><label>Titulo</label><input class="coord-event-title" type="text" maxlength="120" value="' + escapeAttr(cleanText(event.title)) + '" placeholder="Ex: Aprovacao layout fase 1" /></div>',
      '<div class="coord-field"><label>Responsavel</label><select class="coord-event-responsible">' + responsibleOptions + "</select></div>",
      '<div class="coord-field"><label>Status</label><select class="coord-event-status"><option value="Planejado"' + (cleanText(event.status) === "Planejado" ? " selected" : "") + '>Planejado</option><option value="Concluido"' + (cleanText(event.status) === "Concluido" ? " selected" : "") + '>Concluido</option><option value="Cancelado"' + (cleanText(event.status) === "Cancelado" ? " selected" : "") + ">Cancelado</option></select></div>",
      "</div>",
      '<div class="coord-field coord-event-obs"><label>Observacao</label><textarea class="coord-event-observation" maxlength="220" placeholder="Detalhe rapido do evento">' + escapeHtml(cleanText(event.observation)) + "</textarea></div>",
      '<button class="coord-remove-event" type="button" data-remove-event="' + String(index) + '">Remover evento</button>',
      "</article>"
    ].join("");
  }

  function handleCoordTaskFieldChange_(event) {
    var target = event && event.target;
    if (!target) {
      return;
    }

    var item = target.closest(".coord-task-item");
    if (!item) {
      return;
    }

    var status = cleanText((item.querySelector(".coord-status") || {}).value) || "NAO INICIADA";
    item.setAttribute("data-status", status);
    var tagEl = item.querySelector(".coord-task-tag");
    if (tagEl) {
      tagEl.className = "coord-task-tag " + getCoordStatusTagClass_(status);
      tagEl.textContent = status;
    }
    applyCoordFilter_();
  }

  function applyCoordFilter_() {
    if (!coordTaskListEl) {
      return;
    }

    var mode = cleanText((coordFilterEl || {}).value) || "all";
    var today = getCutoffDate_();
    var items = coordTaskListEl.querySelectorAll(".coord-task-item");
    var i;

    for (i = 0; i < items.length; i += 1) {
      var item = items[i];
      var status = cleanText(item.getAttribute("data-status"));
      var blocked = cleanText((item.querySelector(".coord-blocked") || {}).value) === "Sim";
      var plannedEnd = parseDateFlexible_((item.querySelector(".coord-planned-end") || {}).value);
      var show = true;

      if (mode === "open") {
        show = status !== "CONCLUIDA";
      } else if (mode === "late") {
        show = status !== "CONCLUIDA" && plannedEnd && plannedEnd.getTime() < today.getTime();
      } else if (mode === "blocked") {
        show = blocked || status === "BLOQUEADA";
      }

      if (show) {
        item.classList.remove("hidden");
      } else {
        item.classList.add("hidden");
      }
    }
  }

  function addCoordEventRow_(eventData) {
    var current = collectCoordEventsFromForm_();
    current.push(eventData || {
      eventDate: "",
      eventType: "",
      title: "",
      responsible: cleanText((professionalEl || {}).value),
      status: "Planejado",
      observation: ""
    });
    renderCoordEvents_(current);
  }

  function handleCoordEventRemove_(event) {
    var target = event && event.target;
    if (!target) {
      return;
    }
    var removeIndex = cleanText(target.getAttribute("data-remove-event"));
    if (!removeIndex && removeIndex !== "0") {
      return;
    }

    var index = Number(removeIndex);
    if (isNaN(index)) {
      return;
    }

    var current = collectCoordEventsFromForm_();
    if (index < 0 || index >= current.length) {
      return;
    }
    current.splice(index, 1);
    renderCoordEvents_(current);
  }

  function handleCoordSave_() {
    var contractId = cleanText(coordState.contractId);
    if (!contractId) {
      showCoordFeedback_("error", "Projeto nao identificado para salvar.");
      return;
    }

    var appsScriptUrl = cleanText(cfg.appsScriptUrl);
    if (!appsScriptUrl || appsScriptUrl.indexOf("COLE_AQUI") >= 0) {
      showCoordFeedback_("error", "Backend nao configurado.");
      return;
    }

    var tasksNow = collectCoordTasksFromForm_();
    var taskUpdates = [];
    try {
      taskUpdates = buildCoordTaskUpdates_(tasksNow);
    } catch (validationError) {
      showCoordFeedback_("error", cleanText(validationError && validationError.message) || "Erro de validacao.");
      return;
    }
    var eventsNow = collectCoordEventsFromForm_();
    var eventsKey = coordEventsKey_(eventsNow);
    var hasEventsChanged = eventsKey !== coordState.originalEventsKey;

    if (!taskUpdates.length && !hasEventsChanged) {
      showCoordFeedback_("warn", "Nenhuma alteracao detectada para salvar.");
      return;
    }

    var editor = cleanText((professionalEl || {}).value) || "COORDENACAO";
    var payload = {
      action: "coord_update_project",
      contractId: contractId,
      projectName: cleanText(coordState.projectName),
      editor: editor,
      recalculate: true,
      updates: taskUpdates,
      events: eventsNow
    };

    coordSaveEl.disabled = true;
    coordSaveEl.textContent = "Salvando...";
    clearCoordFeedback_();

    sendPayload(appsScriptUrl, payload)
      .then(function (result) {
        if (result.mode === "cors" && result.data && result.data.status === "ok") {
          loadRefOptionsFromServer();
          loadPpmSnapshotFromServer();
          renderPortfolioOverview();
          return loadCoordProjectFromServer_(contractId).then(function (fresh) {
            coordState.tasks = fresh.tasks;
            coordState.events = fresh.events;
            coordState.projectName = fresh.projectName || coordState.projectName;
            coordState.originalTasksByRef = {};
            fresh.tasks.forEach(function (task) {
              coordState.originalTasksByRef[cleanText(task.refEap)] = {
                status: cleanText(task.status),
                responsible: cleanText(task.responsible),
                plannedStart: cleanText(task.plannedStart),
                plannedEnd: cleanText(task.plannedEnd),
                prazoDu: String(Number(task.prazoDu || 1)),
                predecessorRef: cleanText(task.predecessorRef),
                relationType: cleanText(task.relationType) || "FS",
                lagDu: String(Number(task.lagDu || 0)),
                percentReal: String(Number(task.percentReal || 0)),
                blocked: task.blocked ? "Sim" : "Nao",
                blockReason: cleanText(task.blockReason)
              };
            });
            coordState.originalEventsKey = coordEventsKey_(fresh.events);
            coordState.warnings = fresh.warnings || [];
            renderCoordPanel_();
            showCoordFeedback_(
              "ok",
              "Configuracoes salvas com sucesso. Tarefas atualizadas: " +
                String(Number(result.data.updatedTasks || 0)) +
                ". Eventos salvos: " +
                String(Number(result.data.savedEvents || 0)) +
                ". Datas recalculadas: " +
                String(Number(result.data.recalculatedDates || 0)) +
                "."
            );
          });
        }

        if (result.mode === "cors" && result.data && result.data.status === "error") {
          showCoordFeedback_("error", result.data.message || "Falha ao salvar configuracoes.");
          return;
        }

        showCoordFeedback_(
          "warn",
          "Salvo em modo de compatibilidade. Confirme na aba CONTROLE_EAP_ATUAL."
        );
      })
      .catch(function (error) {
        showCoordFeedback_("error", cleanText(error && error.message) || "Nao foi possivel salvar.");
      })
      .then(function () {
        coordSaveEl.disabled = false;
        coordSaveEl.textContent = "Salvar configuracoes do projeto";
      });
  }

  function collectCoordTasksFromForm_() {
    if (!coordTaskListEl) {
      return [];
    }

    var rows = coordTaskListEl.querySelectorAll(".coord-task-item");
    var items = [];
    var i;

    for (i = 0; i < rows.length; i += 1) {
      var row = rows[i];
      items.push({
        refEap: cleanText(row.getAttribute("data-ref-eap")),
        status: cleanText((row.querySelector(".coord-status") || {}).value),
        responsible: cleanText((row.querySelector(".coord-responsible") || {}).value),
        prazoDu: cleanText((row.querySelector(".coord-prazo-du") || {}).value),
        predecessorRef: cleanText((row.querySelector(".coord-predecessor") || {}).value),
        relationType: cleanText((row.querySelector(".coord-relation") || {}).value),
        lagDu: cleanText((row.querySelector(".coord-lag-du") || {}).value),
        plannedStart: cleanText((row.querySelector(".coord-planned-start") || {}).value),
        plannedEnd: cleanText((row.querySelector(".coord-planned-end") || {}).value),
        percentReal: cleanText((row.querySelector(".coord-percent-real") || {}).value),
        blocked: cleanText((row.querySelector(".coord-blocked") || {}).value),
        blockReason: cleanText((row.querySelector(".coord-block-reason") || {}).value)
      });
    }

    return items;
  }

  function buildCoordTaskUpdates_(tasksNow) {
    var updates = [];
    var i;
    for (i = 0; i < tasksNow.length; i += 1) {
      var item = tasksNow[i];
      var ref = cleanText(item.refEap);
      if (!ref) {
        continue;
      }

      var original = coordState.originalTasksByRef[ref] || {
        status: "",
        responsible: "",
        prazoDu: "1",
        predecessorRef: "",
        relationType: "FS",
        lagDu: "0",
        plannedStart: "",
        plannedEnd: "",
        percentReal: "",
        blocked: "Nao",
        blockReason: ""
      };

      var duration = Math.max(1, Math.min(120, Number(item.prazoDu || 1)));
      var lag = Math.max(-30, Math.min(30, Number(item.lagDu || 0)));
      var current = {
        status: normalizePortfolioStatus_(item.status),
        responsible: cleanText(item.responsible),
        prazoDu: String(duration),
        predecessorRef: cleanText(item.predecessorRef),
        relationType: cleanText(item.relationType) || "FS",
        lagDu: String(lag),
        plannedStart: cleanText(item.plannedStart),
        plannedEnd: cleanText(item.plannedEnd),
        percentReal: String(Math.max(0, Math.min(100, Number(item.percentReal || 0)))),
        blocked: cleanText(item.blocked) === "Sim" ? "Sim" : "Nao",
        blockReason: cleanText(item.blockReason)
      };

      if (current.predecessorRef) {
        var predecessorRefs = current.predecessorRef
          .split(/[;,]+/)
          .map(function (text) {
            return cleanText(text);
          })
          .filter(Boolean);
        if (predecessorRefs.indexOf(ref) >= 0) {
          throw new Error("A tarefa " + ref + " nao pode depender dela mesma.");
        }
      }

      var startDate = parseDateFlexible_(current.plannedStart);
      var endDate = parseDateFlexible_(current.plannedEnd);
      if (startDate && endDate && endDate.getTime() < startDate.getTime()) {
        throw new Error("Em " + ref + ", o fim planejado nao pode ser menor que o inicio planejado.");
      }

      if (
        current.status !== cleanText(original.status) ||
        current.responsible !== cleanText(original.responsible) ||
        current.prazoDu !== cleanText(original.prazoDu) ||
        current.predecessorRef !== cleanText(original.predecessorRef) ||
        current.relationType !== cleanText(original.relationType) ||
        current.lagDu !== cleanText(original.lagDu) ||
        current.plannedStart !== cleanText(original.plannedStart) ||
        current.plannedEnd !== cleanText(original.plannedEnd) ||
        current.percentReal !== cleanText(original.percentReal) ||
        current.blocked !== cleanText(original.blocked) ||
        current.blockReason !== cleanText(original.blockReason)
      ) {
        updates.push({
          refEap: ref,
          status: current.status,
          responsible: current.responsible,
          prazoDu: Number(current.prazoDu || 1),
          predecessorRef: current.predecessorRef,
          relationType: current.relationType,
          lagDu: Number(current.lagDu || 0),
          plannedStart: current.plannedStart,
          plannedEnd: current.plannedEnd,
          percentReal: Number(current.percentReal || 0),
          blocked: current.blocked,
          blockReason: current.blockReason
        });
      }
    }
    return updates;
  }

  function collectCoordEventsFromForm_() {
    if (!coordEventsListEl) {
      return [];
    }

    var rows = coordEventsListEl.querySelectorAll(".coord-event-item");
    var events = [];
    var i;
    for (i = 0; i < rows.length; i += 1) {
      var row = rows[i];
      var event = {
        eventDate: cleanText((row.querySelector(".coord-event-date") || {}).value),
        eventType: cleanText((row.querySelector(".coord-event-type") || {}).value),
        title: cleanText((row.querySelector(".coord-event-title") || {}).value),
        responsible: cleanText((row.querySelector(".coord-event-responsible") || {}).value),
        status: cleanText((row.querySelector(".coord-event-status") || {}).value) || "Planejado",
        observation: cleanText((row.querySelector(".coord-event-observation") || {}).value)
      };

      if (!event.eventDate && !event.eventType && !event.title && !event.observation) {
        continue;
      }
      events.push(event);
    }
    return events;
  }

  function coordEventsKey_(events) {
    var normalized = (events || []).map(function (event) {
      return {
        eventDate: cleanText(event && event.eventDate),
        eventType: cleanText(event && event.eventType),
        title: cleanText(event && event.title),
        responsible: cleanText(event && event.responsible),
        status: cleanText(event && event.status),
        observation: cleanText(event && event.observation)
      };
    });
    return JSON.stringify(normalized);
  }

  function showCoordFeedback_(type, message) {
    if (!coordFeedbackEl) {
      return;
    }
    coordFeedbackEl.classList.remove("hidden", "ok", "warn", "error");
    coordFeedbackEl.classList.add(type);
    coordFeedbackEl.textContent = message;
  }

  function clearCoordFeedback_() {
    if (!coordFeedbackEl) {
      return;
    }
    coordFeedbackEl.classList.add("hidden");
    coordFeedbackEl.classList.remove("ok", "warn", "error");
    coordFeedbackEl.textContent = "";
  }

  function getCoordStatusTagClass_(status) {
    if (status === "EM ANDAMENTO") {
      return "doing";
    }
    if (status === "BLOQUEADA") {
      return "blocked";
    }
    if (status === "CONCLUIDA") {
      return "done";
    }
    return "todo";
  }

  function formatDateInput_(value) {
    var parsed = parseDateFlexible_(value);
    if (!parsed) {
      return "";
    }
    return (
      String(parsed.getFullYear()) +
      "-" +
      String(parsed.getMonth() + 1).padStart(2, "0") +
      "-" +
      String(parsed.getDate()).padStart(2, "0")
    );
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

  function buildSignatureMetaText_(item) {
    var signatureDate = item && item.signatureDate instanceof Date ? item.signatureDate : null;
    if (!signatureDate) {
      return "Assinatura contratual: nao informada";
    }

    var days = Number(item && item.daysSinceSignature);
    if (isNaN(days) || days < 0) {
      days = 0;
    }

    return "Assinatura contratual: " + formatDateBr_(signatureDate) + " | " + days + " dias de contrato";
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
    ensureEditButtonState_();
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
    if (editLastEl) {
      editLastEl.classList.add("hidden");
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
    var previousProfessional = cleanText((professionalEl || {}).value);
    formEl.reset();
    if (previousProfessional && professionalEl) {
      professionalEl.value = previousProfessional;
    }
    setDefaultDate();
    populateProjects();
    renderProjectSections([]);
    if (bulkStatusEl) {
      bulkStatusEl.value = "";
    }
    if (bulkBlockEl) {
      bulkBlockEl.value = "";
    }
    updateFormChecklist_();
    ensureEditButtonState_();
    renderDailyFocus_();
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

  function getStorageJson_(key) {
    try {
      var raw = window.localStorage.getItem(key);
      if (!raw) {
        return null;
      }
      try {
        return JSON.parse(raw);
      } catch (parseError) {
        return raw;
      }
    } catch (error) {
      return null;
    }
  }

  function setStorageJson_(key, value) {
    try {
      window.localStorage.setItem(key, JSON.stringify(value));
    } catch (error) {
      // Ignora erros de armazenamento local (quota, privacidade, etc).
    }
  }

  function removeStorageKey_(key) {
    try {
      window.localStorage.removeItem(key);
    } catch (error) {
      // Sem acao.
    }
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
