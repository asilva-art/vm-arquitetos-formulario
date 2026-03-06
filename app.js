(function () {
  "use strict";

  var cfg = window.VM_FORM_CONFIG || {};
  var formEl = document.getElementById("daily-form");
  var professionalEl = document.getElementById("professional");
  var dateEl = document.getElementById("work-date");
  var projectListEl = document.getElementById("project-list");
  var projectSectionsEl = document.getElementById("project-sections");
  var feedbackEl = document.getElementById("feedback");
  var submitButtonEl = document.getElementById("submit-button");

  var projectMap = {};
  var projectList = cfg.projetos || [];
  var i;
  for (i = 0; i < projectList.length; i += 1) {
    if (projectList[i] && projectList[i].code) {
      projectMap[projectList[i].code] = projectList[i];
    }
  }

  init();

  function init() {
    if (!formEl) {
      return;
    }

    populateProfessionals();
    populateProjects();
    setDefaultDate();
    renderProjectSections([]);

    projectListEl.addEventListener("change", handleProjectSelectionChange);
    formEl.addEventListener("submit", handleSubmit);
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
    var html = (cfg.projetos || [])
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

  function handleProjectSelectionChange() {
    renderProjectSections(getSelectedProjectCodes());
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
        '<p class="field-title">Campo 4 - O que voce fez neste projeto hoje?</p>' +
        '<div class="option-list">' + buildTaskOptionsHtml(key) + "</div>" +
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
        '<textarea id="obs-' + key + '" maxlength="240" placeholder="Maximo de 2 linhas"></textarea>' +
        '<p class="hint">Opcional. Use apenas se houver algo importante.</p>' +
        "</div>";

      projectSectionsEl.appendChild(sectionEl);
      wireBlockToggle(key);
    });
  }

  function buildTaskOptionsHtml(key) {
    return (cfg.categoriasTarefa || [])
      .map(function (item, index) {
        var inputId = "task-" + key + "-" + index;
        return [
          '<label class="check-item" for="' + inputId + '">',
          '<input id="' + inputId + '" type="checkbox" name="tasks-' + key + '" value="' + escapeAttr(item) + '" />',
          "<span>" + escapeHtml(item) + "</span>",
          "</label>"
        ].join("");
      })
      .join("");
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
      var checkedTaskEls = document.querySelectorAll('input[name="tasks-' + key + '"]:checked');
      var checkedTasks = [];
      var taskIndex;
      for (taskIndex = 0; taskIndex < checkedTaskEls.length; taskIndex += 1) {
        checkedTasks.push(checkedTaskEls[taskIndex].value);
      }
      var statusEl = document.querySelector('input[name="status-' + key + '"]:checked');
      var blockEl = document.querySelector('input[name="block-' + key + '"]:checked');
      var blockDesc = ((document.getElementById("block-desc-" + key) || {}).value || "").trim();
      var obs = ((document.getElementById("obs-" + key) || {}).value || "").trim();

      if (!checkedTasks.length) {
        errors.push("Campo 4 obrigatorio na secao " + project.label + ".");
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
        tasks: checkedTasks,
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
          resetForm();
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

  function setDefaultDate() {
    var now = new Date();
    now.setMinutes(now.getMinutes() - now.getTimezoneOffset());
    dateEl.value = now.toISOString().slice(0, 10);
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
})();
