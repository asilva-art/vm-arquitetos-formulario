# Backend (Google Apps Script)

## O que esta versao faz

- Continua recebendo o formulario diario e gravando em `⚙️ EXECUÇÃO`.
- Cria e mantem a aba `📌 CONTROLE_EAP_ATUAL` (uma linha por tarefa da EAP).
- Sincroniza automaticamente o controle a cada envio diario.
- Exige `REF EAP` por secao do formulario (rastreabilidade fina).
- Expõe `action=config` com opcoes de `REF EAP` por projeto para o frontend.
- Le contratos automaticamente da aba `CONTRATOS` (IDs `CT-...`) para montar os projetos ativos.
- Retorna `dailySummary` no `doPost` para exibir resumo objetivo ao usuario (feito no dia + pendencias + proximo foco).
- Mantem bloqueio de duplicata: maximo 1 envio por profissional por dia.
- Gera `ID ENVIO FORM.` no formato `FORM-AAAAMMDD-NNN`.
- Inclui reset total para reconstruir a base com `EAP_PADRAO_V1` (10 fases / 48 itens).

## Arquivos

- `Code.gs`: backend completo (envio + controle operacional).
- `appsscript.json`: manifesto do projeto Apps Script.

## Primeira ativacao (recomendada para base limpa)

1. No Apps Script, substituir `Code.gs` pelo arquivo deste projeto.
2. Salvar.
3. Executar a funcao `resetBancoDoZeroEapPadrao`.
4. Conferir:
   - `⚙️ EXECUÇÃO` limpa (dados antigos removidos)
   - `📅 EAP - LINHA DE BASE` recriada com a EAP padrao
   - `📌 CONTROLE_EAP_ATUAL` recriada com a EAP padrao

## Operacao diaria

- A equipe preenche o formulario normalmente.
- Cada envio atualiza:
  - `⚙️ EXECUÇÃO`
  - `📌 CONTROLE_EAP_ATUAL`

## Funcoes manuais disponiveis

- `setupControleEapAtual`: cria/prepara a aba de controle e carrega tarefas da EAP.
- `resetBancoDoZeroEapPadrao`: apaga dados operacionais e recria EAP/controle com o padrao.
- `sincronizarControleComExecucaoHistorica`: reprocesa historico da EXECUCAO para atualizar controle.
- `resumoControleEapAtual`: retorna contagem por status (nao iniciada, andamento, concluida, bloqueada).

## Ajustes comuns

- Projetos:
  - prioridade: aba `CONTRATOS` (id contrato + nome/apelido).
  - fallback: constante `PROJECT_MAP` (caso a aba nao exista ou esteja incompleta).
- Nome das abas:
  - `EXEC_SHEET_NAME`
  - `EAP_SHEET_NAME`
  - `CONTROL_SHEET_NAME`
  - `CONTRACTS_SHEET_NAME`
