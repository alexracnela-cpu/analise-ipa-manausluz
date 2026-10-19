# Análise IPA – Manaus Luz

Site: https://analise-ipa-manausluz.netlify.app/ (publicado pela Netlify a cada push na branch `main`).

## Como a página funciona

- `index.html` é fixo (logo, tema, abas, cards e gráfico). **Não alterar no dia a dia.** `img/` também é fixo.
- Cada atualização é uma **aba** do site e tem dois arquivos em `historico/`, com o mesmo `<id>` (`AAAA-MM-DD_HHMM`):
  - `historico/<id>.md` – texto da análise (aba do doc do IPA exportada em markdown), **sem** o título `# ...`, a linha da data/menção e a primeira tabela de KPIs; o placeholder do gráfico vira `<!-- GRAFICO -->`.
  - `historico/<id>.json` – números do cabeçalho e do gráfico:
    `titulo`, `atualizado_em` ("dd/mm/aaaa HHhMM"), `periodo` ("01/mm a dd/mm/aaaa"), `mes`, `os_fechadas`, `no_prazo`, `fora_do_prazo`,
    `ipa` (2 casas, = 100*no_prazo/os_fechadas), `nota` (texto, ex. "0,8"), `ipa_anterior`/`atualizado_anterior` (da aba anterior; null no dia 01),
    `doc` (link do doc) e `dias` (lista `{dia, os, atrasos}` do mês; somas batem com `os_fechadas` e `fora_do_prazo`).
- `historico/indice.json` – lista das abas, **mais recente primeiro**: `[{"id": "<id>", "rotulo": "Até dd/mm HH:MM"}, ...]`. A primeira é a aberta por padrão (marcada como ATUAL).
- O IPA é sempre exibido com 2 casas decimais.

## Atualização diária (feita pela tarefa agendada)

1. Criar `historico/<id>.md` e `historico/<id>.json` da aba nova (copiar `titulo` e `doc` do json da aba anterior; `ipa_anterior` = `ipa` dela).
2. Inserir `{"id": "<id>", "rotulo": "Até dd/mm HH:MM"}` no **início** de `historico/indice.json`, mantendo as anteriores.
3. `git commit` e `git push` na `main`. A Netlify republica sozinha em ~1 minuto.

## Robô automático (GitHub Actions)

O workflow `.github/workflows/atualizar-ipa.yml` roda às **6h, 8h, 12h e 16h (Manaus)** e também pelo botão *Actions → Atualizar IPA → Run workflow*:

1. `robo/baixar_relatorio.py` entra no GisWorks (Manaus) e baixa o relatório "Ordens Executadas e Inspeções/Materiais por Município" do dia 01 do mês até agora.
2. `robo/atualizar_site.py` calcula o IPA (OS exceto origem/reclamante Obra, sem inspeções, campo "Fechada no Prazo", 2 casas) e cria a aba do dia com números, gráfico, recortes e equipes.
3. Faz commit em `historico/` na `main`; a Netlify republica sozinha.

- O robô mantém **uma aba automática por dia** (`"auto": true` no json): cada execução substitui a automática do mesmo dia. Abas feitas à mão nunca são alteradas nem apagadas.
- Login: secrets `GISWORKS_USUARIO` e `GISWORKS_SENHA` (Settings → Secrets and variables → Actions).
- Se falhar, o GitHub manda e-mail e o print da tela fica em *Artifacts → erro-gisworks* da execução.
- Faixas de nota: `FAIXAS_NOTA` no início de `robo/atualizar_site.py`.
