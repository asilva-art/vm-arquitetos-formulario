# Backend (Google Apps Script)

## O que esta versao faz

- Continua recebendo o formulario diario e gravando em `⚙️ EXECUÇÃO`.
- Cria e mantem a aba `📌 CONTROLE_EAP_ATUAL` (uma linha por tarefa da EAP).
- Sincroniza automaticamente o controle a cada envio diario.
- Exige `REF EAP` por secao do formulario (rastreabilidade fina).
- Expõe `action=config` com opcoes de `REF EAP` por projeto para o frontend.
- Mantem bloqueio de duplicata: maximo 1 envio por profissional por dia.
- Gera `ID ENVIO FORM.` no formato `FORM-AAAAMMDD-NNN`.

## Arquivos

- `Code.gs`: backend completo (envio + controle operacional).
- `appsscript.json`: manifesto do projeto Apps Script.

## Primeira ativacao (1 vez)

1. No Apps Script, substituir `Code.gs` pelo arquivo deste projeto.
2. Salvar.
3. Executar a funcao `setupControleEapAtual`.
4. Executar a funcao `sincronizarControleComExecucaoHistorica`.
5. Conferir a nova aba `📌 CONTROLE_EAP_ATUAL` na planilha.

## Operacao diaria

- A equipe preenche o formulario normalmente.
- Cada envio atualiza:
  - `⚙️ EXECUÇÃO`
  - `📌 CONTROLE_EAP_ATUAL`

## Funcoes manuais disponiveis

- `setupControleEapAtual`: cria/prepara a aba de controle e carrega tarefas da EAP.
- `sincronizarControleComExecucaoHistorica`: reprocesa historico da EXECUCAO para atualizar controle.
- `resumoControleEapAtual`: retorna contagem por status (nao iniciada, andamento, concluida, bloqueada).

## Ajustes comuns

- Projetos mapeados:
  - constante `PROJECT_MAP`.
- Nome das abas:
  - `EXEC_SHEET_NAME`
  - `EAP_SHEET_NAME`
  - `CONTROL_SHEET_NAME`
