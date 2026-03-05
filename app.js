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

  var projectMap = new Map((cfg.projetos || []).map(function (p) { return [p.code, p]; }));

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
    (cfg.profissionais || []).forEach(function (name) {
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
      var project = projectMap.get(projectCode) || { code: projectCode, label: projectCode };
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

    blockInputs.forEach(function (input) {
      input.addEventListener("change", function () {
        if (input.value === "NAO") {
          blockWrap.classList.add("hidden");
          blockDesc.value = "";
        } else {
          blockWrap.classList.remove("hidden");
        }
      });
    });
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
      var project = projectMap.get(code) || { code: code, label: code };
      var key = sanitizeKey(code);
      var checkedTasks = Array.from(document.querySelectorAll('input[name="tasks-' + key + '"]:checked')).map(function (el) { return el.value; });
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
      .finally(function () {
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
    return Array.from(document.querySelectorAll('input[name="projects"]:checked')).map(function (el) { return el.value; });
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

window.VM_FORM_CONFIG = {
  appsScriptUrl:
    "https://script.google.com/macros/s/AKfycbzX7ZlICKM8wsdxc5vAlQqnPh7KVOwy-wejZOQuIkP6M5skKHEZRKcc3ppX046Zxlpiwg/exec",

  profissionais: [
    "Viviane",
    "Profissional 2",
    "Profissional 3",
    "Profissional 4",
    "Profissional 5",
    "Profissional 6",
    "Profissional 7",
    "Profissional 8",
    "Profissional 9",
    "Profissional 10"
  ],

  projetos: [
    { code: "CT-117", label: "CT-117 | PH Top Green" },
    { code: "CT-119", label: "CT-119 | KT Costa Laguna" },
    { code: "CT-120", label: "CT-120 | Alliance Sede NL" },
    { code: "CT-121", label: "CT-121 | VC Hermes Sabara" },
    { code: "CT-122", label: "CT-122 | IT Bernardo Guimaraes" },
    { code: "ADMIN-INTERNO", label: "Administrativo/Interno" }
  ],

  categoriasTarefa: [
    "Estudos Preliminares",
    "Anteprojeto",
    "Pre-Executivo",
    "Executivo",
    "Detalhamentos",
    "Acompanhamentos",
    "Aprovacoes e Marcos",
    "Comunicacao",
    "Outros"
  ],

  opcoesStatus: [
    { code: "CONCLUIDO", label: "Concluido hoje" },
    { code: "EM_ANDAMENTO", label: "Em andamento - continua amanha" },
    { code: "PAUSADO", label: "Pausado - aguardando" }
  ],

  opcoesBloqueio: [
    { code: "NAO", label: "Nao - tudo ok" },
    { code: "SIM_APROVACAO", label: "Sim - aguardando aprovacao do cliente" },
    { code: "SIM_ACESSO", label: "Sim - aguardando acesso ao local" },
    { code: "SIM_FORNECEDOR", label: "Sim - aguardando fornecedor" },
    { code: "SIM_OUTRO", label: "Sim - outro motivo" }
  ]
};

<!doctype html>
<html lang="pt-BR">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>VM + Arquitetos | Formulario Diario</title>
    <link rel="stylesheet" href="./styles.css" />
  </head>
  <body>
    <main class="container">
      <header class="page-header">
        <p class="brand">VM + ARQUITETOS</p>
        <h1>Formulario Diario de Projetos</h1>
        <p class="subtitle">
          Preencha 1 vez por dia, no fim do expediente. Tempo medio: 3 minutos.
        </p>
      </header>

      <section class="card">
        <form id="daily-form" novalidate>
          <div class="field">
            <label for="professional">Campo 1 - Nome do profissional</label>
            <select id="professional" name="professional" required>
              <option value="">Selecione...</option>
            </select>
          </div>

          <div class="field">
            <label for="work-date">Campo 2 - Data</label>
            <input id="work-date" name="workDate" type="date" required />
          </div>

          <div class="field">
            <p class="field-title">
              Campo 3 - Em qual(is) projeto(s) voce trabalhou hoje?
            </p>
            <div id="project-list" class="check-list"></div>
          </div>

          <div id="project-sections"></div>

          <div class="actions">
            <button id="submit-button" type="submit">Enviar registro diario</button>
          </div>
        </form>
      </section>

      <section id="feedback" class="feedback hidden" aria-live="polite"></section>
    </main>

    <script src="./config.js"></script>
    <script src="./app.js"></script>
  </body>
</html>

:root {
  --bg: #f4f3ef;
  --panel: #fffaf2;
  --panel-2: #ffffff;
  --text: #1f1f1f;
  --muted: #5d5d5d;
  --line: #d9d1c2;
  --accent: #8a4b08;
  --accent-strong: #6f3b06;
  --ok: #125c2d;
  --ok-bg: #dff5e8;
  --warn: #8a4b08;
  --warn-bg: #fff0da;
  --error: #912828;
  --error-bg: #ffe6e6;
}

* {
  box-sizing: border-box;
}

body {
  margin: 0;
  min-height: 100vh;
  font-family: "Source Sans 3", "Segoe UI", Tahoma, Arial, sans-serif;
  color: var(--text);
  background:
    radial-gradient(circle at 15% 10%, #f0e3cf 0, transparent 35%),
    radial-gradient(circle at 85% 0, #f6ead7 0, transparent 25%),
    var(--bg);
}

.container {
  max-width: 920px;
  margin: 0 auto;
  padding: 24px 16px 48px;
}

.page-header {
  margin-bottom: 16px;
}

.brand {
  margin: 0;
  font-size: 12px;
  letter-spacing: 0.14em;
  color: var(--muted);
}

h1 {
  margin: 6px 0 8px;
  font-size: 28px;
  line-height: 1.1;
}

.subtitle {
  margin: 0;
  color: var(--muted);
}

.card {
  background: linear-gradient(180deg, var(--panel) 0%, var(--panel-2) 100%);
  border: 1px solid var(--line);
  border-radius: 14px;
  padding: 18px;
  box-shadow: 0 12px 28px rgba(0, 0, 0, 0.06);
}

.field {
  margin-bottom: 16px;
}

label,
.field-title {
  display: block;
  margin-bottom: 8px;
  font-weight: 700;
  color: #292929;
}

select,
input[type="date"],
input[type="text"],
textarea {
  width: 100%;
  border: 1px solid #c8bea9;
  border-radius: 10px;
  background: #fffefb;
  color: #1f1f1f;
  padding: 11px 12px;
  font: inherit;
}

textarea {
  min-height: 68px;
  resize: vertical;
}

.check-list,
.option-list {
  display: grid;
  gap: 8px;
}

.check-item,
.radio-item {
  display: flex;
  align-items: flex-start;
  gap: 8px;
  border: 1px solid #ddd3bf;
  border-radius: 10px;
  background: #fffdf8;
  padding: 9px 10px;
}

.check-item input,
.radio-item input {
  margin-top: 2px;
}

.project-card {
  border: 1px solid #d7ceb9;
  border-radius: 12px;
  padding: 12px;
  background: #fffefb;
  margin: 14px 0;
}

.project-card h3 {
  margin: 0 0 10px;
  font-size: 17px;
}

.hint {
  margin: 6px 0 0;
  font-size: 13px;
  color: var(--muted);
}

.hidden {
  display: none !important;
}

.actions {
  margin-top: 22px;
}

button {
  appearance: none;
  border: 0;
  border-radius: 10px;
  background: var(--accent);
  color: #fff;
  font: inherit;
  font-weight: 700;
  padding: 12px 16px;
  cursor: pointer;
  width: 100%;
}

button:hover {
  background: var(--accent-strong);
}

button:disabled {
  opacity: 0.65;
  cursor: progress;
}

.feedback {
  margin-top: 14px;
  border-radius: 10px;
  border: 1px solid;
  padding: 11px 12px;
  font-weight: 600;
}

.feedback.ok {
  color: var(--ok);
  background: var(--ok-bg);
  border-color: #86d0a3;
}

.feedback.warn {
  color: var(--warn);
  background: var(--warn-bg);
  border-color: #f3c980;
}

.feedback.error {
  color: var(--error);
  background: var(--error-bg);
  border-color: #f2a8a8;
}

@media (max-width: 640px) {
  h1 {
    font-size: 23px;
  }

  .card {
    padding: 14px;
  }

  .project-card {
    padding: 10px;
  }
}
