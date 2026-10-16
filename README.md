# Análise IPA – Manaus Luz

Site: https://analise-ipa-manausluz.netlify.app/ (publicado pela Netlify a cada push na branch `main`).

## Como a página funciona

- `index.html` é fixo (visual, cards e gráfico). **Não precisa ser alterado no dia a dia.**
- `dados/dados.json` traz os números do cabeçalho e do gráfico:
  - `atualizado_em`, `periodo`, `os_fechadas`, `no_prazo`, `fora_do_prazo`, `ipa` (número, ex. 92.4), `nota` (texto, ex. "0,8")
  - `ipa_anterior` e `atualizado_anterior` (da atualização anterior, para a variação)
  - `dias`: lista `{dia, os, atrasos}` com os fechamentos e atrasos de cada dia do mês (a soma deve bater com `os_fechadas` e `fora_do_prazo`)
- `dados/analise.md` é o texto da análise do dia (exportado da aba nova do doc do IPA em markdown), **sem** o título, a linha de data e a tabela de KPIs do topo.
- `historico/AAAA-MM-DD_HHMM.md` guarda uma cópia de cada análise.

## Atualização diária (feita pela tarefa agendada)

1. Exportar a aba nova do doc em markdown e salvar em `dados/analise.md` (e cópia em `historico/`).
2. Atualizar `dados/dados.json` com os números do dia e os dados do gráfico.
3. `git commit` e `git push` na `main`. A Netlify republica sozinha em ~1 minuto.
