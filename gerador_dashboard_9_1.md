# gerador_dashboard_9_1.py — Documentação Técnica Completa

**Versão:** 9.1  
**Arquivo:** `gerador_dashboard_9_1.py`  
**Linhas:** ~11.100  
**Linguagem:** Python 3.12+ / JavaScript (embutido)  
**Última atualização:** Abril 2026

---

## Sumário

1. [Visão Geral](#1-visão-geral)
2. [Dependências](#2-dependências)
3. [Arquitetura do Arquivo](#3-arquitetura-do-arquivo)
4. [Constantes Globais](#4-constantes-globais)
5. [Classe LaunchDataManager](#5-classe-launchdatamanager)
6. [Classe DashboardGenerator](#6-classe-dashboardgenerator)
   - [Inicialização e Permissões](#61-inicialização-e-permissões)
   - [Carregamento de Dados](#62-carregamento-de-dados)
   - [Processamento de Empreendimentos](#63-processamento-de-empreendimentos)
   - [Preparação de Dados para JSON](#64-preparação-de-dados-para-json)
   - [Crosstabs](#65-crosstabs)
   - [Geração de HTML](#66-geração-de-html)
7. [Sistema de Permissões e Perfis](#7-sistema-de-permissões-e-perfis)
8. [Estrutura do HTML Gerado](#8-estrutura-do-html-gerado)
9. [Menus e Submenus do Dashboard](#9-menus-e-submenus-do-dashboard)
10. [Métricas e Fórmulas de Cálculo](#10-métricas-e-fórmulas-de-cálculo)
11. [Funções JavaScript Principais](#11-funções-javascript-principais)
12. [Tabelas Geradas](#12-tabelas-geradas)
13. [Seção Insights](#13-seção-insights)
14. [Exportação PDF e XLSX](#14-exportação-pdf-e-xlsx)
15. [Arquitetura CSS/JS em Python](#15-arquitetura-cssjs-em-python)
16. [Saídas do Script](#16-saídas-do-script)
17. [Interface de Linha de Comando](#17-interface-de-linha-de-comando)
18. [Regras de Dados e Negócio](#18-regras-de-dados-e-negócio)
19. [Bugs Corrigidos na v9.1](#19-bugs-corrigidos-na-v91)
20. [Pontos de Atenção e Armadilhas](#20-pontos-de-atenção-e-armadilhas)

---

## 1. Visão Geral

O `gerador_dashboard_9_1.py` é um script Python standalone que lê dados do mercado imobiliário de Brasília a partir de um arquivo Excel e gera um **dashboard HTML completamente autossuficiente** — sem dependências de servidor para renderização, com todo o JavaScript, CSS e dados embutidos no arquivo HTML resultante.

O dashboard é publicado via Flask no Render e consumido por usuários com acesso controlado por perfil (OAuth Google). Cada perfil recebe um HTML diferente, com menus e seções habilitadas conforme suas permissões.

**Fluxo de alto nível:**

```
Excel (.xlsx)
    └─► load_data()
            └─► derive_fields()
            └─► LaunchDataManager.get_public_launch_counts()
            └─► compute_crosstabs_empreendimentos()
            └─► prepare_data_for_json()
                    └─► generate_html_template()
                            └─► create_html_structure()
                                    └─► Dashboard HTML (standalone)
                    └─► LaunchDataManager.generate_private_txt_report()
                            └─► lancamentos_residenciais.txt
                            └─► lancamentos_comerciais.txt
```

---

## 2. Dependências

### Python

| Biblioteca | Uso |
|---|---|
| `pandas` | Leitura do Excel, agregações, DataFrames |
| `numpy` | Operações numéricas, tratamento de NaN |
| `json` | Serialização de dados para embutir no HTML |
| `os`, `sys` | Caminhos de arquivo e saída |
| `unicodedata` | Normalização de strings (acentos) |
| `datetime` | Timestamps e formatação de datas |
| `tkinter` | Seletor de arquivo via GUI (opcional; fallback se ausente) |
| `argparse` | Interface de linha de comando |
| `user_permission_manager` | Sistema de permissões por perfil (módulo externo opcional) |

### JavaScript (CDN — carregadas no HTML gerado)

| Biblioteca | Versão | Uso |
|---|---|---|
| Chart.js | latest (jsdelivr) | Gráficos dos Insights |
| jsPDF | 2.5.1 | Exportação para PDF |
| jspdf-autotable | 3.5.28 | Tabelas no PDF |
| SheetJS (xlsx) | 0.17.0 | Exportação para Excel |

> **Atenção:** O dashboard depende de conexão com CDN para Chart.js, jsPDF e SheetJS. Em ambientes offline, os gráficos e exportações não funcionarão.

---

## 3. Arquitetura do Arquivo

```
gerador_dashboard_9_1.py
│
├── Constantes globais (L36–L106)
│   ├── RESIDENTIAL_REQUIRED_COLS
│   ├── COMMERCIAL_REQUIRED_COLS
│   ├── FAIXAS_VALOR
│   ├── FAIXAS_AREA
│   ├── MESES_ABREV / TRIMESTRES
│   ├── OFERTA_LANCAMENTOS / OFERTA_DISPONIVEIS / VENDIDOS / DISTRATO
│   └── EMPREENDIMENTO_SUFFIXES
│
├── class LaunchDataManager (L108–L255)
│   ├── get_public_launch_counts()
│   ├── get_private_launch_details()
│   ├── generate_private_txt_report()
│   └── _aggregate_* (helpers internos)
│
├── class DashboardGenerator (L257–L10982)
│   ├── __init__()
│   ├── load_permissions_config()
│   ├── get_user_by_profile()
│   ├── normalize_string() / format_ano_mes()
│   ├── categorize_value() / categorize_area() / get_trimestre()
│   ├── extract_empreendimento_name()
│   ├── get_projects_details()
│   ├── count_unique_projects()
│   ├── analyze_launches_by_company_and_neighborhood_with_empreendimentos()
│   ├── aggregate_projects_to_quarters() / _to_years() / _to_years_with_list()
│   ├── select_input_file()
│   ├── derive_fields()
│   ├── load_data()
│   ├── prepare_data_for_json()
│   ├── compute_crosstabs_empreendimentos()   ← corrigido na v9.1
│   ├── prepare_crosstabs_data_for_json()
│   ├── get_data_periods()
│   ├── load_menu_permissions()
│   ├── generate_html_template()
│   ├── create_html_structure()      ← gera todo o CSS + HTML + JS
│   │   ├── css_styles (string regular, L1546)
│   │   ├── html_body (f-string com dados, L~2900)
│   │   └── create_javascript_content()
│   ├── create_javascript_content()  ← JS puro em r"""..."""
│   ├── debug_january_2021()
│   ├── validate_html_txt_consistency()
│   ├── comprehensive_launch_debug()
│   ├── validate_launch_data_separation()
│   └── run()
│
├── load_permissions_config() (L10983)
└── main() (L10999)
```

---

## 4. Constantes Globais

### Colunas Obrigatórias

**Residencial (`RESIDENTIAL_REQUIRED_COLS`):**
`ANO_MES`, `ORIGEM_RECURSOS`, `ESTAGIO_OBRA`, `OFERTA_VENDA`, `BAIRRO`, `AREA`, `QUANTIDADE`, `QTD_QUARTOS`, `QTD_ELEVADORES`, `QTD_GARAGEM`, `TEMPO_FINANCIAMENTO`, `VALOR_MEDIO_M2`, `AREA_QUANTIDADE`, `AREA_VALOR`, `AREA_QUANTIDADE_VALOR`, `EMPREENDIMENTO`

**Comercial (`COMMERCIAL_REQUIRED_COLS`):**
Igual ao residencial, **exceto `QTD_QUARTOS`** (dados comerciais não têm esta coluna — importante para o fix da v9.1).

### Faixas de Valor (R$)

| Faixa | Label |
|---|---|
| 0 – 349.999 | `< 350.000` |
| 350.000 – 499.999 | `350.000 – 499.999` |
| 500.000 – 699.999 | `500.000 – 699.999` |
| 700.000 – 999.999 | `700.000 – 999.999` |
| 1.000.000 – 1.999.999 | `1.000.000 – 1.999.999` |
| ≥ 2.000.000 | `≥ 2.000.000` |

### Faixas de Área (m²)

Até 40m², 41–60m², 61–80m², 81–100m², 101–120m², 121–150m², 151–175m², 176–200m², mais de 200m².

### Tipos de OFERTA_VENDA

| Constante | Valores |
|---|---|
| `OFERTA_LANCAMENTOS` | `['OFERTADOS LANCAMENTOS']` |
| `OFERTA_DISPONIVEIS` | `['OFERTADOS DISPONIVEIS', 'OFERTADOS LANCAMENTOS']` |
| `VENDIDOS` | `['VENDIDOS', 'VENDIDOS - LANCADOS E VENDIDOS']` |
| `DISTRATO` | `['DISTRATO']` |

---

## 5. Classe LaunchDataManager

Responsável por separar **dados públicos** (contagens anonimizadas para o HTML) dos **dados privados** (detalhes completos com empresa/bairro para o TXT interno).

### get_public_launch_counts(df)

Retorna contagens de unidades e empreendimentos lançados, segmentadas por período:

```python
{
    "monthly_units": {AAAAMM: int},
    "quarterly_units": {AAAA_XT: int},
    "yearly_units": {AAAA: int},
    "monthly_projects": {AAAAMM: int},
    "quarterly_projects": {AAAA_XT: int},
    "yearly_projects": {AAAA: int}
}
```

A contagem de **empreendimentos** usa deduplicação anual: um empreendimento faseado que aparece em múltiplos meses do mesmo ano é contado apenas uma vez naquele ano. A contagem de **unidades** é soma direta das quantidades.

### get_private_launch_details(df)

Delega para `get_projects_details()` do `DashboardGenerator`. Retorna a estrutura completa com tríades `(empreendimento, empresa, bairro)` por período.

### generate_private_txt_report(details_data, output_file)

Delega para `aggregate_projects_to_years_with_list()`. Gera o arquivo TXT com empreendimentos lançados agrupados por ano.

---

## 6. Classe DashboardGenerator

### 6.1 Inicialização e Permissões

```python
DashboardGenerator(user_email=None, config_file="user_profiles.json")
```

- Se `user_email` for fornecido e o módulo `user_permission_manager` estiver disponível, autentica o usuário e ativa o controle de acesso.
- Sem email, executa em modo básico (sem controle de acesso — comportamento de admin).
- Instancia `LaunchDataManager` referenciando a si próprio.

**Atributos de dados carregados:**

| Atributo | Fonte no Excel |
|---|---|
| `residential_data` | Sheet 0 (obrigatório) |
| `commercial_data` | Sheet 1 (opcional) |
| `incc_data` | Sheet `'INCC'` (opcional) |
| `ipca_data` | Sheet `'IPCA'` (opcional) |
| `selic_data` | Sheet `'SELIC'` (opcional) |
| `juros_reais_data` | Sheet `'JUROS_REAIS'` (opcional) |

### 6.2 Carregamento de Dados

**`load_data(file_path)`**

Lê todas as sheets do Excel usando `pd.read_excel`. A sheet residencial é obrigatória; todas as demais usam `try/except` com fallback para DataFrame vazio. Após leitura, chama `derive_fields()` nos dados residenciais e comerciais.

**`derive_fields(df)`**

Adiciona ao DataFrame as colunas derivadas:

| Coluna | Derivação |
|---|---|
| `Faixa_Valor` | `categorize_value(AREA_VALOR)` |
| `Faixa_Area` | `categorize_area(AREA)` |
| `ANO_MES_NUM` | conversão numérica segura de `ANO_MES` |
| `ANO` | `ANO_MES_NUM // 100` |
| `MES` | `ANO_MES_NUM % 100` |
| `TRIMESTRE` | `get_trimestre(MES)` → `'1T'` a `'4T'` |

### 6.3 Processamento de Empreendimentos

**`extract_empreendimento_name(df)`**

Cria a coluna `EMPREENDIMENTO_AGRUPADO` a partir de `EMPREENDIMENTO`, aplicando em sequência:

1. Correções ortográficas comuns (`EMPPREENDIMENTO`, `EMPRENDIMENTO`, etc.)
2. Remoção de prefixos genéricos (`EMPREENDIMENTO `, `RESIDENCIAL `, `RES `)
3. Remoção de sufixos padronizados em ordem de especificidade crescente: sufixos compostos complexos (BL + especificação + DUPLEX), sufixos compostos médios, quartos/suítes, blocos simples, torres, especificações independentes (COBERTURA, GARDEN, DUPLEX...), numerações
4. Geração do `TERMO_PRINCIPAL` normalizado em maiúsculas

A lógica **não usa agrupamento por similaridade** — confia apenas na normalização determinística. A unicidade é estabelecida pela tríade `(EMPREENDIMENTO_AGRUPADO, EMPRESA, BAIRRO)`.

**`get_projects_details(df, period_col="ANO_MES")`**

Regra de lançamentos: considera apenas linhas com `OFERTA_VENDA == 'OFERTADOS LANCAMENTOS'` e `QUANTIDADE > 0`. A deduplicação é **por ano** — uma tríade que aparecer em vários meses do mesmo ano é contada apenas na primeira aparição naquele ano (não histórica global).

Retorna:
```python
{AAAAMM: [(empreendimento, empresa, bairro), ...]}
```

### 6.4 Preparação de Dados para JSON

**`prepare_data_for_json(df)`**

Aplica controle de acesso baseado em perfil — colunas sensíveis (nome da empresa, etc.) podem ser removidas dependendo do perfil do usuário autenticado. Converte tipos numéricos, substitui NaN por None e retorna `List[Dict]` compatível com `json.dumps`.

**`prepare_crosstabs_data_for_json(df)`**

Versão específica para a seção de Crosstabs, mantendo todas as colunas necessárias para as tabelas cruzadas por região.

### 6.5 Crosstabs

**`compute_crosstabs_empreendimentos(df_res, df_com)`**

Pré-computa contagem de empreendimentos únicos lançados estruturada como:
```python
{
    "residencial": {AAAAMM: {bairro: {quartos_str: count}}},
    "comercial":   {AAAAMM: {bairro: {'': count}}}
}
```

**Fix v9.1:** dados comerciais não têm a coluna `QTD_QUARTOS`. A função agora verifica a presença da coluna antes de iterar:
- Se `QTD_QUARTOS` existe (residencial): agrupa por valor de quartos, normalizando `≥4 → '4+'` e nulos → `''`.
- Se não existe (comercial): agrupa todos os empreendimentos do bairro sob a chave `''`.

### 6.6 Geração de HTML

**`generate_html_template()`**

Orquestra toda a serialização de dados para JSON e monta os parâmetros para `create_html_structure()`. Dados passados ao HTML:

| Variável JSON | Conteúdo |
|---|---|
| `residential_json` | Dados residenciais para gráficos/tabelas |
| `commercial_json` | Dados comerciais para gráficos/tabelas |
| `residential_crosstabs_json` | Dados residenciais com todas colunas para crosstabs |
| `commercial_crosstabs_json` | Dados comerciais com todas colunas para crosstabs |
| `crosstabs_empreendimentos_json` | Contagem empreendimentos por período×bairro×quartos |
| `projects_count_json` | Unidades lançadas por período (mensal/trimestral/anual) |
| `projects_count_empreendimentos_json` | Empreendimentos lançados por período |
| `launches_preprocessed_json` | Contagens pré-processadas para JS (substitui cálculo client-side) |
| `incc_json` / `ipca_json` / `selic_json` / `juros_reais_json` | Indicadores econômicos |
| `menu_config_json` | Perfil + menus e submenus permitidos |
| `max_period` | Último período disponível nos dados (AAAAMM como inteiro) |

**`create_html_structure(...)`**

Método com ~9.000 linhas que gera o HTML completo usando três tipos de strings Python:

| Tipo de conteúdo | Tipo de string Python | Motivo |
|---|---|---|
| CSS | String regular `"""..."""` | CSS usa `{}` literais para valores de propriedades |
| HTML com dados injetados | f-string `f"""..."""` | Variáveis Python interpoladas; `{{}}` para literais JS |
| JavaScript puro | Raw string `r"""..."""` | Evita que `\n`, `{}` do JS sejam interpretados pelo Python |

---

## 7. Sistema de Permissões e Perfis

### Arquivo de Configuração

`user_profiles.json` — fonte única de verdade para usuários e perfis.

```json
{
  "users": {
    "email@dominio.com": {
      "profile": "admin",
      "active": true,
      "name": "Nome do Usuário"
    }
  }
}
```

### Perfis Disponíveis

| Perfil | Menus Liberados | Submenus |
|---|---|---|
| `admin` | residencial, comercial, crosstabs, insights | Todos |
| `manager` | residencial, comercial | IVV, Oferta, Venda, VGO, VGV, VGL |
| `analyst` | residencial | IVV, Oferta, Venda |
| `viewer` | residencial | IVV, Oferta, Venda |

### Controle de Menu por Perfil

As permissões são carregadas de `dashboard_permissions.json` (se existir) ou de um fallback interno. São convertidas para JSON e injetadas no HTML como `const menuConfig = {...}`. A função JavaScript `applyMenuPermissions()` oculta no cliente os menus e submenus não permitidos ao perfil.

### Normalização de Submenus de Crosstabs

O arquivo de permissões pode usar nomes externos (ex: `ofertas_por_regiao`) que são mapeados para os identificadores internos JS (ex: `oferta_quantidade`) via `mapping_crosstabs` antes de serem serializados.

---

## 8. Estrutura do HTML Gerado

O arquivo HTML resultante é standalone e contém:

```
<html>
  <head>
    <style>  ← CSS completo (~1.400 linhas)
    <script> ← Dados embutidos como JSON em variáveis JS
    <script src="cdnjs...jspdf">
    <script src="cdnjs...autotable">
    <script src="cdnjs...xlsx">
    <script src="cdn.jsdelivr.net/chart.js">
  </head>
  <body>
    <!-- Topbar mobile -->
    <!-- Sidebar desktop com menus e submenus -->
    <!-- Filtros: Período, Bairro, Estagio Obra, Quartos, Faixa Valor, Área -->
    <!-- Área principal: tablesContainer / crossTablesContainer / insightsContainer -->
    <!-- Bottom nav mobile -->
    <script> ← JavaScript completo (~7.000 linhas)
  </body>
</html>
```

### Responsividade

O layout tem duas versões ativas:
- **Desktop:** sidebar lateral recolhível (280px / 60px collapsed), área de conteúdo à direita.
- **Mobile:** topbar superior, navegação inferior por abas, sidebar como overlay.

---

## 9. Menus e Submenus do Dashboard

### Menu Residencial e Comercial

Ordem de aparição dos submenus (v9.1):

1. IVV
2. Oferta
3. Venda
4. Lançamentos
5. Oferta em m²
6. Venda em m²
7. Preço de Oferta
8. Preço de Venda
9. **VGO** ← antes chamado "VGV sobre Ofertas"
10. **VGV** ← antes chamado "VGV sobre Vendas"
11. **VGL**
12. Distratos

### Menu Crosstabs

Submenus: IVV, Ofertas por Região, Vendas por Região, Unidades Lançadas, Empreendimentos Lançados, Preço de Oferta, Preço de Venda, Oferta em m², Venda em m², Gastos Pós-entrega por Região, Gastos Pós-entrega por Categoria.

### Menu Insights

Submenus: Indicadores Econômicos, Correlações.

---

## 10. Métricas e Fórmulas de Cálculo

### IVV — Índice de Velocidade de Vendas

```
IVV (%) = (Vendas / Ofertas) × 100
```

- **Vendas:** soma de `QUANTIDADE` onde `OFERTA_VENDA` ∈ `['VENDIDOS', 'VENDIDOS - LANCADOS E VENDIDOS']`
- **Ofertas:** soma de `QUANTIDADE` onde `OFERTA_VENDA` ∈ `['OFERTADOS DISPONIVEIS', 'OFERTADOS LANCAMENTOS']`

### VGL — Valor Geral de Lançamentos

Soma de `AREA_QUANTIDADE_VALOR` onde `OFERTA_VENDA == 'OFERTADOS LANCAMENTOS'`. Agregação por soma (fluxo mensal).

### VGO — Valor Geral de Ofertas (ex-"VGV sobre Ofertas")

Soma de `AREA_QUANTIDADE_VALOR` onde `OFERTA_VENDA` ∈ `['OFERTADOS DISPONIVEIS', 'OFERTADOS LANCAMENTOS']`. Usa **média** nos períodos trimestrais/anuais (não soma), para evitar dupla contagem de estoque que se mantém entre meses.

### VGV — Valor Geral de Vendas (ex-"VGV sobre Vendas")

Soma de `AREA_QUANTIDADE_VALOR` onde `OFERTA_VENDA` ∈ `['VENDIDOS', 'VENDIDOS - LANCADOS E VENDIDOS']`. Agregação por soma (fluxo de transações).

### Preço Médio Ponderado (R$/m²)

```
Preço Ponderado = Σ(AREA_QUANTIDADE_VALOR) / Σ(AREA_QUANTIDADE)
```

Calculado separadamente para ofertas e vendas.

### Gastos Pós-Entrega (Crosstabs — metodologia CBIC)

Estimativa de impacto econômico regional baseada em percentual do VGV:

| Padrão Construtivo | % do VGV | Bairros |
|---|---|---|
| Alto | 25% | Noroeste, Sudoeste, Jd. Botânico, Asa Sul, Asa Norte, Lago Norte |
| Médio | 15% | Águas Claras, Guará, Sobradinho, Park Sul, Taguatinga |
| Popular | 10% | Ceilândia, Planaltina, Santa Maria, Gama, Recanto das Emas, Samambaia |

Fonte: [Estudo CBIC 2021](https://cbic.org.br/wp-content/uploads/2021/02/pos-obraestudo-cbic.pdf)

### Normalização Base 100 (Insights)

Para comparação de séries de escalas diferentes no gráfico de Indicadores Econômicos:

```
Base100[t] = (Valor[t] / Média_últimos_12_meses) × 100
```

Após normalização, aplica-se Média Móvel de 3 meses para suavizar a série.

---

## 11. Funções JavaScript Principais

### Dados e Cálculo

| Função | Descrição |
|---|---|
| `calculateIVV(data)` | Calcula IVV por período a partir dos dados brutos |
| `calculateIndicator(data, ofertaTypes)` | Soma de QUANTIDADE por período para tipos de oferta dados |
| `calculatePeriodAggregations(data, ofertaTypes, isOferta)` | Agrega mensal → trimestral → anual por soma |
| `calculateIVVPeriodAggregations(data)` | Agrega IVV mensal → trimestral → anual (média ponderada) |
| `calculateAreaPeriodAggregations(data, ofertaTypes, isOferta)` | Agrega AREA_QUANTIDADE por período |
| `calculateValorPonderadoPeriodAggregations(data, ofertaTypes)` | Calcula preço médio ponderado por período |
| `calculateVGLVGVPeriodAggregations(data, ofertaTypes)` | Agrega AREA_QUANTIDADE_VALOR por soma (VGL e VGV) |
| `calculateVGLVGVPeriodAverages(data, ofertaTypes)` | Agrega AREA_QUANTIDADE_VALOR por **média** (VGO — evita dupla contagem de estoque) |
| `calculateUniqueProjects(data, ofertaTypes)` | Conta empreendimentos únicos (normaliza nome) |
| `calculateSecondaryVariables(residentialData)` | Calcula IVV, Oferta, Venda, VGL, VGO, VGV para Insights |
| `normalizeToBase100(values)` | Normaliza array para Base 100 usando média dos últimos 12 meses |
| `calculateRollingAverage(values, window)` | Média Móvel de N períodos |

### Interface e Navegação

| Função | Descrição |
|---|---|
| `applyMenuPermissions()` | Oculta menus/submenus conforme perfil do usuário |
| `initializeMenu()` | Monta o menu lateral e inicializa estado ativo |
| `buildCategoryNav()` | Constrói navegação de subcategorias dentro de cada menu |
| `toggleMainMenu(view, event)` | Abre/fecha seção do menu principal |
| `showCategory(cat)` | Filtra e exibe tabelas da categoria ativa |
| `assignTableCategories()` | Atribui `data-category` a cada card de tabela baseado no título |
| `handleViewClick(view, event)` | Gerencia mudança de view principal (residencial/comercial/crosstabs/insights) |
| `toggleSidebar()` | Colapsa/expande sidebar desktop |
| `toggleMobileSidebar()` / `toggleMobileFilters()` | Controles mobile |

### Filtros

| Função | Descrição |
|---|---|
| `getSelectedFilters()` | Coleta valores de todos os filtros ativos |
| `applyFilters()` | Aplica filtros e re-renderiza todas as tabelas |
| `generateCrossData(data)` | Processa dados para a seção Crosstabs (tabelas por região) |
| `populateCrossTabsFilters()` | Preenche dropdowns de filtro da seção Crosstabs |
| `applyCrossTabsFilters()` | Aplica filtros de período/tipologia nas Crosstabs |

### Tabelas

| Função | Descrição |
|---|---|
| `updateTables(data)` | Regenera todas as tabelas para os dados e filtros ativos |
| `createTable(title, data, isPercentage, projectsData, enterpriseData)` | Cria card de tabela mensal |
| `createQuarterlyTable(...)` | Cria card de tabela trimestral |
| `createYearlyTable(...)` | Cria card de tabela anual |
| `createTableMoney(...)` | Versão monetária (R$ Milhões) das tabelas mensais |
| `createQuarterlyTableMoney(...)` | Versão monetária trimestral |
| `createYearlyTableMoney(...)` | Versão monetária anual |
| `applyColoringToTableCells()` | Aplica heatmap de cores às células por intensidade de valor |
| `applyTrendColorsQuarterly()` / `applyTrendColorsYearly()` | Colorização de tendência (alta/baixa) |
| `sortCrossTable(element, category)` | Ordena colunas das crosstabs por clique no cabeçalho |

### Insights

| Função | Descrição |
|---|---|
| `buildEconomicIndicatorsMonthly(...)` | Monta série temporal dos indicadores econômicos |
| `renderEconomicIndicatorsChart(canvasId, primaryData, secondaryData)` | Renderiza gráfico Chart.js com indicadores + variáveis de mercado |
| `toggleSecondaryVariable(variable)` | Mostra/oculta série de mercado no gráfico |
| `renderCorrelationNarrative(containerId, ...)` | Gera análise textual de correlação ao clicar em série |
| `alignSecondaryData(primaryPeriods, secondaryPeriods, secondaryValues)` | Alinha séries de diferentes períodos |
| `normalizeToBase100(values)` | Normaliza série para Base 100 |
| `calculateRollingAverage(values, window)` | Média Móvel de 3 meses |

---

## 12. Tabelas Geradas

Para cada segmento (residencial/comercial), são geradas 33 tabelas em 3 períodos (mensal, trimestral, anual):

| # | Grupo | Título | Tipo |
|---|---|---|---|
| 1–3 | IVV | IVV Mensal/Trimestral/Anual | % (1 decimal) |
| 4–6 | Oferta | Ofertas Mensais/Trimestrais/Anuais (Unidades) | Inteiro |
| 7–9 | Venda | Vendas Mensais/Trimestrais/Anuais (Unidades) | Inteiro |
| 10–12 | Lançamentos | Lançamentos Mensais/Trimestrais/Anuais (Unidades [Empreendimentos]) | Inteiro |
| 13–15 | Oferta m² | Oferta Mensal/Trimestral/Anual (m²) | Inteiro |
| 16–18 | Venda m² | Venda Mensal/Trimestral/Anual (m²) | Inteiro |
| 19–21 | Preço Oferta | Preço de Oferta Mensal/Trimestral/Anual (R$/m²) | R$ (2 decimais) |
| 22–24 | Preço Venda | Preço de Venda Mensal/Trimestral/Anual (R$/m²) | R$ (2 decimais) |
| **25–27** | **VGO** | **VGO Mensal/Trimestral/Anual (R$ Milhões)** | R$ Mi (2 decimais) |
| **28–30** | **VGV** | **VGV Mensal/Trimestral/Anual (R$ Milhões)** | R$ Mi (2 decimais) |
| **31–33** | **VGL** | **VGL Mensal/Trimestral/Anual (R$ Milhões)** | R$ Mi (2 decimais) |
| 34–36 | Distratos | Distratos Mensais/Trimestrais/Anuais (Unidades) | Inteiro |

> A ordem VGO → VGV → VGL foi estabelecida na v9.1.

### Categorização Automática de Tabelas

A função `assignTableCategories()` atribui `data-category` a cada card de tabela analisando o texto do título em minúsculas:

| Padrão no título | Categoria atribuída |
|---|---|
| `'vgo'` | `vgv_ofertas` |
| `'vgv'` (sem 'vgo'/'vgl') | `vgv_vendas` |
| `'vgl'` | `vgl` |
| `'lanç'` ou `'lanc'` | `lancamentos` |
| `'oferta'` + `'m²'` (sem valor) | `oferta_m2` |
| `'venda'` + `'m²'` (sem valor) | `venda_m2` |
| `'oferta'` + `'preço'` ou `'valor'` | `valor_ponderado_oferta` |
| `'oferta'` (puro) | `oferta` |
| `'venda'` (puro) | `venda` |
| `'distrato'` | `distratos` |
| `'ivv'` | `ivv` |

---

## 13. Seção Insights

A seção Insights exibe:

1. **Card INCC-M:** Gráfico de linha comparando evolução do INCC-M com preço médio de oferta e preço médio de venda (ambos indexados). Fonte: IBGE.

2. **Card Indicadores Econômicos vs. Variáveis de Mercado:** Gráfico Chart.js com dois eixos Y:
   - Eixo primário (esquerda): SELIC, IPCA 12m, Juros Reais (linhas)
   - Eixo secundário (direita): IVV, Oferta, Venda, VGL, VGO, VGV, Lançamentos — todos normalizados em Base 100 com MM3m (barras, ocultas por padrão)
   
   O usuário clica em uma série na legenda para exibir/ocultar. Ao selecionar uma variável de mercado, o painel de correlação exibe a análise textual gerada pela função `renderCorrelationNarrative()`.

**Variáveis de mercado no gráfico (labels atuais):**

| Chave interna | Label no gráfico |
|---|---|
| `ivv` | IVV - Base 100 (MM 3m) |
| `oferta` | OFERTA - Base 100 (MM 3m) |
| `venda` | VENDA - Base 100 (MM 3m) |
| `vgl` | VGL - Base 100 (MM 3m) |
| `vgv_vendas` | VGV - Base 100 (MM 3m) |
| `vgv_ofertas` | VGO - Base 100 (MM 3m) |
| `lancamentos` | Lançamentos - Base 100 (MM 3m) |

---

## 14. Exportação PDF e XLSX

### PDF (`exportAllTablesToPDF(tipoImovel)`)

- Usa **jsPDF 2.5.1** + **jspdf-autotable 3.5.28** (CDN).
- Gera PDF A4 portrait com todas as tabelas visíveis da view/categoria ativa.
- Renderiza variações de tendência (tokens coloridos `+X,X%` / `-X,X%`) usando `drawRichLine()`.
- Detecta nota "Trimestre incompleto / Ano incompleto" via `findIncompleteNote()`.
- Nome do arquivo: `Relatorio_Completo_{TipoImovel}_{AAAA_MM}.pdf`.

### XLSX (`exportAllTablesToXLSX(tipoImovel)`)

- Usa **SheetJS 0.17.0** (CDN).
- Exporta todas as tabelas da view ativa como planilha Excel.
- Cada tabela vira uma aba separada no arquivo.

> **Nota:** Existe uma discrepância conhecida entre Chrome e Firefox na renderização de gráficos Chart.js para PDF. A abordagem canvas-to-image foi identificada como solução futura para padronizar o comportamento.

---

## 15. Arquitetura CSS/JS em Python

Esta é a regra de escaping mais crítica do arquivo — violá-la causa erros silenciosos ou falhas de renderização difíceis de depurar.

| Bloco | Tipo de string Python | Regra de escaping |
|---|---|---|
| `css_styles` | `"""..."""` regular | Chaves CSS `{color: red}` funcionam literalmente |
| HTML body | `f"""..."""` f-string | Variáveis Python: `{variavel}`. Literais JS: `{{variavel}}` |
| `js_code` | `r"""..."""` raw string | Nada é interpretado. `\n`, `{}`, `${}` funcionam literalmente |

**Exemplo prático:**

```python
# CSS — string regular
css_styles = """
.card { color: var(--primary-blue); }  ← {} literais, OK
"""

# HTML — f-string
html = f"""
<div data-max="{max_period}">           ← {max_period} = variável Python
<script>var x = {{}};</script>          ← {{}} = {} literal no JS
"""

# JavaScript — raw string
js_code = r"""
const obj = {key: 'value'};            ← {} literal, OK (raw string)
const tmpl = `Olá ${nome}`;            ← template literal JS, OK
"""
```

---

## 16. Saídas do Script

Para cada execução com um perfil, o script gera:

| Arquivo | Descrição |
|---|---|
| `templates/dashboard_{perfil}.html` | Dashboard HTML standalone do perfil |
| `lancamentos_residenciais.txt` | Lista privada de empreendimentos residenciais lançados por ano |
| `lancamentos_comerciais.txt` | Lista privada de empreendimentos comerciais lançados por ano |

No modo `--todos-perfis`, são gerados 4 dashboards:
- `templates/dashboard_admin.html`
- `templates/dashboard_manager.html`
- `templates/dashboard_analyst.html`
- `templates/dashboard_viewer.html`

---

## 17. Interface de Linha de Comando

```bash
python3 gerador_dashboard_9_1.py [input_file] [output_file] [--profile PERFIL] [--todos-perfis]
```

| Argumento | Descrição |
|---|---|
| `input_file` | Caminho para o Excel de entrada (opcional; abre GUI se omitido) |
| `output_file` | Caminho para o HTML de saída (opcional; gera timestamp se omitido) |
| `--profile {admin,manager,analyst,viewer}` | Gera dashboard para um perfil específico |
| `--todos-perfis` | Gera dashboards para todos os 4 perfis em sequência |

**Exemplos:**

```bash
# Modo interativo (abre seletor de arquivo)
python3 gerador_dashboard_9_1.py

# Arquivo específico, saída padrão com timestamp
python3 gerador_dashboard_9_1.py dados.xlsx

# Perfil específico
python3 gerador_dashboard_9_1.py dados.xlsx --profile manager

# Todos os perfis (fluxo de deploy)
python3 gerador_dashboard_9_1.py dados.xlsx --todos-perfis
```

---

## 18. Regras de Dados e Negócio

### Deduplicação de Lançamentos

- Um empreendimento é identificado pela tríade `(EMPREENDIMENTO_AGRUPADO, EMPRESA, BAIRRO)`.
- A deduplicação é **anual**: a mesma tríade que aparece em Janeiro e Março de 2025 é contada uma única vez em 2025, no mês de Janeiro (primeira aparição).
- Em meses diferentes de anos diferentes, é contada novamente (fases novas em ano novo).

### Oferta vs. Estoque

O VGO usa **média** nos agregados trimestrais/anuais, não soma, porque o estoque disponível em cada mês é um **estado**, não um **fluxo**. Somar Janeiro + Fevereiro + Março superestimaria o estoque do trimestre. O VGV e VGL usam **soma** porque representam transações/eventos.

### Campo EMPREENDIMENTO Vazio

Quando `EMPREENDIMENTO` está vazio ou nulo, `extract_empreendimento_name()` usa o valor `'N/A'`. Linhas com `EMPREENDIMENTO_AGRUPADO == 'N/A'` são excluídas da contagem de empreendimentos lançados.

---

## 19. Bugs Corrigidos na v9.1

### 1. Renomeação VGV Ofertas → VGO / VGV Vendas → VGV

**Problema:** Os nomes "VGV sobre Ofertas" e "VGV sobre Vendas" eram ambíguos — ambos tinham o prefixo VGV.

**Solução:** "VGV sobre Ofertas" passou a se chamar **VGO** (Valor Geral de Ofertas) e "VGV sobre Vendas" passou a se chamar simplesmente **VGV**. Todos os pontos afetados foram atualizados: labels de menu, labels de gráfico, títulos de tabelas, textos de interface, lógica de categorização de tabelas e filtros de exibição.

### 2. KeyError 'QTD_QUARTOS' em dados comerciais

**Problema:** `compute_crosstabs_empreendimentos()` tentava iterar por `QTD_QUARTOS` tanto para dados residenciais quanto comerciais. Como `QTD_QUARTOS` não existe nos dados comerciais, o script lançava `KeyError` ao executar `--todos-perfis`.

**Causa raiz:** A função `process()` interna era única e não distinguia o tipo de dado que estava processando.

**Solução:** Adicionada verificação `if 'QTD_QUARTOS' in b_data.columns` antes da iteração. Para dados sem a coluna (comercial), todos os empreendimentos do bairro são agrupados sob a chave `''`, mantendo a estrutura de dados compatível com o JavaScript do dashboard.

### 3. Reordenação de métricas

**Problema:** A ordem de aparição dos submenus VGL/VGV era inconsistente entre os arrays de permissão Python e os arrays de definição JavaScript.

**Solução:** Ordem padronizada para **VGO → VGV → VGL** em todos os pontos: arrays de permissão Python (admin e manager), `viewCategories` JS e geração sequencial das tabelas.

---

## 20. Pontos de Atenção e Armadilhas

### String Escaping

Nunca misture tipos de string. CSS em f-string ou JS em string regular causarão falhas silenciosas (chaves interpretadas como interpolação Python ou não).

### Dados Comerciais sem QTD_QUARTOS

Qualquer nova lógica que acesse colunas específicas de residencial deve verificar a presença da coluna antes de usar, pois dados comerciais têm um schema diferente.

### Deduplicação Anual de Empreendimentos

A regra de "primeiro aparecimento por ano" está em dois lugares: `get_projects_details()` (Python, para o TXT) e `calculateUniqueProjectsPeriodAggregations()` (JavaScript, legado). Os dados pré-processados em `launches_preprocessed_json` substituem o cálculo JS para a exibição nas tabelas — alterações nessa regra devem ser feitas no Python e refletidas no JSON.

### max_period

O valor `max_period` é determinado pelo maior `ANO_MES` nos dados residenciais + comerciais combinados e injetado no HTML. Ele controla qual período é considerado "incompleto" nas tabelas trimestrais/anuais. Se o dado vier com período futuro por engano, as tabelas podem marcar trimestres corretos como incompletos.

### CDN Dependência

Em redes corporativas restritivas, Chart.js, jsPDF e SheetJS podem não carregar. O dashboard funciona para navegação e filtros, mas gráficos e exportações falharão sem CDN.

### Compatibilidade PDF entre Navegadores

A exportação PDF funciona via `jsPDF` + DOM scraping. Há diferença de comportamento entre Chrome e Firefox, especialmente para gráficos Chart.js. A solução definitiva (canvas-to-image) ainda não foi implementada.
