# Relatórios de Consumo de Água — Zangari Condomínios

Relatórios web (offline) de consumo de água do condomínio **AREPC**, gerados a
partir das planilhas mensais de leitura "Para Cobrança".

São duas visões, com o mesmo padrão visual e a marca Zangari:

| Página | O que mostra |
|---|---|
| [`relatorio/index.html`](relatorio/index.html) | **Relatório por unidade** — seletor de bloco/unidade, KPIs, evolução mês a mês (consumo m³ + valor R$) e detalhamento com as diferenças de cada mês. |
| [`relatorio/areas_comuns.html`](relatorio/areas_comuns.html) | **Relatório de áreas comuns** — visão geral (consumo total, custo, maior consumidor, ranking, evolução de ~3 anos) e um seletor para o histórico de cada ponto de medição (guaritas, quadras, ADM, etc.). |

As duas páginas se conectam pelo seletor **Unidades ↔ Áreas Comuns** no topo e têm
botão **Exportar PDF** (impressão A4 nítida em SVG).

## Como abrir

A pasta `relatorio/` é **self-contained** e funciona offline:

- **Duplo-clique** em `relatorio/index.html` — os dados carregam via `<script src>`
  (sem depender de servidor ou internet), então abre direto no navegador.
- Ou sirva localmente:
  ```bash
  python -m http.server 8000 --directory relatorio
  # abra http://localhost:8000
  ```

## Rodar com Docker

O site estático é empacotado num **nginx** leve (com gzip). Nada de backend.

```bash
docker compose up -d --build      # abre em http://localhost:8090
docker compose down               # para o container
```

Ou sem compose:

```bash
docker build -t relatorios-agua-zangari .
docker run -d -p 8090:80 --name relatorios-agua relatorios-agua-zangari
```

**Modo desenvolvimento** (monta `relatorio/` ao vivo, sem rebuild a cada edição):

```bash
docker compose --profile dev up relatorios-dev   # http://localhost:8081
```

> Ao atualizar os dados (novas leituras), rode os scripts Python para regenerar
> `relatorio/dados_*.js` e então `docker compose up -d --build` para reempacotar.
> Arquivos do container: `Dockerfile`, `nginx.conf`, `docker-compose.yml`, `.dockerignore`.

## Estrutura

```
relatorio/
  index.html              # relatório por unidade
  areas_comuns.html       # relatório de áreas comuns
  dados_unidades.js       # dados das unidades  (window.DADOS)
  dados_areas_comuns.js   # dados das áreas comuns (window.DADOS_AC)
  vendor/echarts.min.js   # biblioteca de gráficos (vendorizada, offline)
  assets/                 # logos Zangari

converter_para_json.py    # .xlsx "Para Cobrança"  -> json/ (por unidade)
gerar_dados_relatorio.py  # json/ -> relatorio/dados_unidades.js
gerar_areas_comuns.py     # aba "Áreas Comuns" do .xlsx -> relatorio/dados_areas_comuns.js
```

## Pipeline de dados

```
planilhas .xlsx ──converter_para_json.py──▶ json/ ──gerar_dados_relatorio.py──▶ relatorio/dados_unidades.js
                └─────────────────────── gerar_areas_comuns.py ─────────────────▶ relatorio/dados_areas_comuns.js
```

> **Observação:** as planilhas `.xlsx` de origem **não são versionadas** (contêm os
> dados brutos de cobrança) — veja o `.gitignore`. Os arquivos `dados_*.js` já
> committados são o snapshot de dados que faz os relatórios funcionarem. Para
> atualizar com novas leituras, coloque as planilhas na raiz e rode os scripts acima.

## Tratamento de dados

Os scripts tratam as inconsistências das planilhas de origem:

- Layouts diferentes entre meses (com/sem coluna "Releitura", colunas mapeadas por nome).
- Células de erro do Excel (`#VALUE!`/`#REF!`) e textos livres ("trancado", "12 L").
- Consumos negativos (correções de releitura) e medidores parados.
- Troca de hidrômetro (unidade renomeada mantida como série contínua).
- Nas áreas comuns: exclusão de medidor removido (anomalia) e **correção de um erro
  da planilha** (consumo mensal desatualizado em um dos pontos).

## Tecnologias

Python 3 + [openpyxl](https://openpyxl.readthedocs.io/) (extração) · HTML/CSS/JS
puro + [Apache ECharts](https://echarts.apache.org/) (renderizador SVG) · sem build,
sem dependências de runtime.
