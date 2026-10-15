# Dashboard Matéria Prima — pasta única oficial

**Fonte de dados:** `Query - Dados - Lari.xlsx`, aba `query - Lari`.

**Dashboard oficial:** `Dash_Board_Materia_Prima.html`.

**JSON publicado:** `data/materia_prima.json`.

**Automação:** `ATUALIZAR_MATERIA_PRIMA.ps1`, iniciada com dois cliques em `INICIAR_ATUALIZACAO.cmd`.

**Repositório:** https://github.com/scremhellomy-netizen/DashBoardMateriaPrima

**Página publicada:** https://scremhellomy-netizen.github.io/DashBoardMateriaPrima/Dash_Board_Materia_Prima.html

## Primeiro uso

1. **Faça backup das duas pastas antigas e não as exclua antes de validar a primeira publicação.** Extraia o ZIP em uma única pasta, por exemplo `C:\Users\rlxjunior\Documents\DashBoardMateriaPrima_UNIFICADO`.
2. Confirme que o Windows tem **Excel desktop**, **Git for Windows**, internet e credenciais/permissão de escrita no repositório. A planilha precisa ter suas conexões/credenciais de dados funcionando nesse computador.
3. Dê dois cliques em `INICIAR_ATUALIZACAO.cmd` e acompanhe o console. Na primeira execução, o script inicializa o Git nessa mesma pasta e associa o repositório existente, sem depender dos caminhos das pastas antigas.
4. Confirme a mensagem `SUCESSO!`, abra o link publicado e clique em **Atualizar dados**. Aguarde alguns minutos caso o GitHub Pages ainda não tenha concluído a publicação.
5. Apenas após validar tudo, arquive ou exclua as pastas antigas. **Não mantenha duas pastas oficiais.**

## Funcionamento

`INICIAR_ATUALIZACAO.cmd` → abre Excel → executa `RefreshAll` e aguarda consultas → salva planilha → lê a aba `query - Lari` → valida colunas/IDs → gera `data/materia_prima.json` → versiona arquivos do dashboard e JSON → sincroniza com `main` → faz `git push`. O navegador usa o JSON relativo `data/materia_prima.json`; o botão **Atualizar dados** busca o arquivo publicado, sem abrir Excel no computador do visitante. A planilha não é enviada ao GitHub (regra `.gitignore`).

O JSON incluído no ZIP foi extraído do repositório fornecido, com **306 registros/IDs**; a outra cópia fornecida continha 303. O JSON será substituído pelo resultado atualizado da planilha quando a automação for executada.

## Novas métricas

- **Age:** dias corridos entre `Criado` e a data de hoje, em dias inteiros. Verde D+0–1; amarelo D+2–3; vermelho D+4 ou mais. O Age continua aumentando mesmo para registros finalizados, conforme solicitado.
- **SLA:** dias corridos entre `Criado` e `DATA DE FINALIZAÇÃO`, ou hoje se não finalizado. `OK` D+0–3; `Atenção` D+4–7; `Crítico` D+8–10; `Urgente` D+11 em diante. **A solicitação citava D+10 em duas faixas; foi reservado a Crítico para evitar duplicidade.** O filtro SLA aplica-se aos gráficos, indicadores, cards e tabela.
- **Eficiência operacional:** somente registros com data de finalização. Dentro do prazo: SLA D+0–3; fora do prazo: D+4 ou mais. **Pedidos** = IDs únicos; **quantidades** = soma de `qtd atendida` nos registros do grupo. Registros ainda abertos ficam fora dos dois cards.

## Cuidados e limites

- O pacote é uma consolidação estática revisada. **A atualização via Excel COM, autenticação Git e publicação no GitHub não foram executadas neste ambiente Linux**; é obrigatório testar no Windows antes de apagar as pastas antigas.
- Se o Git exigir autenticação, entre com a conta autorizada. Não compartilhe credenciais ou tokens em arquivos do projeto.
- O JSON é público no GitHub Pages. Confira se os campos da planilha podem ser publicados, inclusive comentários e nomes, antes de usar o fluxo com dados sensíveis.
- Se houver conflito de Git ou erro de consulta, a automação para com mensagem de erro; não interprete como publicação concluída.
- Para visualizar o HTML localmente com carregamento de JSON, use um servidor HTTP local; ao abrir diretamente como `file://`, navegadores podem bloquear `fetch`. O fluxo oficial é via GitHub Pages.
