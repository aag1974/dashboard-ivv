# Dashboard IVV

Dashboard de mercado imobiliário de Brasília — Índice de Velocidade de Vendas (IVV), ofertas, vendas, lançamentos, VGO/VGV/VGL, crosstabs por região e correlações com indicadores econômicos.

Produto da **Opinião Informação Estratégica**. Acesso restrito por OAuth Google + perfil de usuário.

URL de produção: https://dashboard-ivv.onrender.com

---

## Arquitetura

O sistema tem duas partes independentes:

### 1. Gerador (offline)

`gerador_dashboard_9_1.py` é um script Python standalone que lê um Excel com dados de mercado e gera **HTMLs autossuficientes** (CSS + JS + dados embutidos) — um por perfil. Os HTMLs gerados pesam ~110 MB cada e ficam em `templates/`.

Documentação técnica detalhada do gerador: [`gerador_dashboard_9_1.md`](./gerador_dashboard_9_1.md).

### 2. Servidor (online)

`server.py` é uma app Flask + Authlib que:

1. Autentica via OAuth Google.
2. Verifica se o e-mail está em `user_profiles.json` e ativo.
3. Lê o perfil do usuário e serve o `templates/dashboard_{perfil}.html` correspondente.
4. Garante **sessão única** por usuário (persistida em `sessions.json`).

Roda no Render via gunicorn (`Procfile`: `web: gunicorn server:app`).

```
Excel → gerador_dashboard_9_1.py → templates/dashboard_{admin,manager,analyst,viewer}.html
                                        │
                                        ▼
                                   server.py (Flask + OAuth)
                                        │
                                        ▼
                                  https://dashboard-ivv.onrender.com
```

---

## Estrutura do projeto

```
dashboard-ivv/
├── server.py                          # Flask + OAuth Google + sessão única
├── wsgi.py                            # Entry point WSGI
├── Procfile                           # Render: gunicorn server:app
├── requirements.txt                   # Flask, gunicorn, authlib, requests
│
├── gerador_dashboard_9_1.py           # Gerador dos HTMLs (~11.100 linhas)
├── gerador_dashboard_9_1.md           # Documentação técnica do gerador
│
├── user_profiles.json                 # Fonte única de usuários e perfis
├── user_permission_manager.py         # Auth + sanitização de dados
├── manage_users.py                    # CLI para gerenciar usuários
├── configurador_visual_permissoes.py  # GUI tkinter para editar permissões
├── dashboard_menu_permissions.json    # Permissões de menu por perfil
│
├── templates/
│   ├── index.html                     # Tela de login
│   ├── acesso_negado.html             # E-mail não autorizado
│   ├── dashboard_admin.html           # ~110 MB, gerado
│   ├── dashboard_manager.html         # ~110 MB, gerado
│   ├── dashboard_analyst.html         # ~110 MB, gerado
│   ├── dashboard_viewer.html          # ~110 MB, gerado
│   ├── lancamentos_residenciais.txt   # Listagem privada (admin)
│   └── lancamentos_comerciais.txt     # Listagem privada (admin)
│
└── .github/workflows/ping-render.yml  # Keep-alive (ping a cada 5min)
```

---

## Perfis e permissões

| Perfil | Visão | Submenus principais |
|---|---|---|
| `admin` | Acesso completo, dados de empresa e empreendimento | Tudo (residencial, comercial, crosstabs, insights) |
| `manager` | Métricas e dados financeiros agregados, sem nomes | Residencial + comercial completos, crosstabs, insights |
| `analyst` | Métricas de mercado básicas | IVV, oferta, venda, lançamentos, m², crosstabs limitadas |
| `viewer` | Dados públicos agregados | Apenas IVV, oferta, venda |

Definições em [`user_profiles.json`](./user_profiles.json) (perfis + lista de usuários) e [`dashboard_menu_permissions.json`](./dashboard_menu_permissions.json) (mapa de submenus por perfil).

---

## Métricas principais

| Métrica | Fórmula | Agregação trimestral/anual |
|---|---|---|
| **IVV** | (Vendas / Ofertas) × 100 | Média ponderada |
| **VGL** | Σ AREA_QUANTIDADE_VALOR onde OFERTA_VENDA = `OFERTADOS LANCAMENTOS` | Soma (fluxo) |
| **VGO** | Σ AREA_QUANTIDADE_VALOR de ofertas | **Média** (estoque é estado, não fluxo) |
| **VGV** | Σ AREA_QUANTIDADE_VALOR de vendas | Soma (fluxo) |
| **Preço Ponderado** | Σ AREA_QUANTIDADE_VALOR / Σ AREA_QUANTIDADE | — |

**Empreendimento** = tríade `(EMPREENDIMENTO_AGRUPADO, EMPRESA, BAIRRO)`. Deduplicação anual.

Detalhes completos das métricas, faixas e regras em [`gerador_dashboard_9_1.md`](./gerador_dashboard_9_1.md).

---

## Como regenerar os dashboards

Após receber um novo Excel mensal:

```bash
# Todos os perfis (fluxo padrão de deploy)
python3 gerador_dashboard_9_1.py dados.xlsx --todos-perfis

# Ou um perfil específico
python3 gerador_dashboard_9_1.py dados.xlsx --profile manager

# Ou modo interativo (abre seletor de arquivo)
python3 gerador_dashboard_9_1.py
```

Saídas geradas:

- `templates/dashboard_admin.html`
- `templates/dashboard_manager.html`
- `templates/dashboard_analyst.html`
- `templates/dashboard_viewer.html`
- `templates/lancamentos_residenciais.txt`
- `templates/lancamentos_comerciais.txt`

Em seguida: commit + push para `main` → Render faz deploy automático.

---

## Como rodar o servidor localmente

```bash
pip install -r requirements.txt

export GOOGLE_CLIENT_ID="..."
export GOOGLE_CLIENT_SECRET="..."
export SECRET_KEY="alguma-chave-aleatoria"

python3 server.py            # http://localhost:5000
# ou
gunicorn server:app          # produção
```

Variáveis de ambiente esperadas:

| Variável | Descrição |
|---|---|
| `GOOGLE_CLIENT_ID` | Client ID do OAuth Google |
| `GOOGLE_CLIENT_SECRET` | Client Secret do OAuth Google |
| `SECRET_KEY` | Chave para assinar cookies de sessão Flask |

---

## Gerenciamento de usuários

CLI interativa:

```bash
python3 manage_users.py
```

Permite listar, adicionar, editar perfil, ativar/desativar e testar autenticação. Edita diretamente `user_profiles.json`.

GUI para mapa de permissões de menu:

```bash
python3 configurador_visual_permissoes.py
```

---

## Deploy

- **Hospedagem:** Render (plano gratuito)
- **Entry point:** `gunicorn server:app` (via `Procfile`)
- **Deploy:** push em `main` → Render rebuilda automaticamente
- **Keep-alive:** GitHub Actions (`ping-render.yml`) faz `curl` a cada 5 minutos para evitar suspensão por inatividade

---

## Pontos de atenção

- **HTMLs grandes:** ~110 MB cada. Estão versionados no Git para simplificar o deploy no Render. Use Git LFS se for crescer mais.
- **String escaping no gerador:** `create_html_structure()` mistura CSS (string regular), HTML (f-string) e JS (raw string). Trocar o tipo causa falhas silenciosas — ver [`gerador_dashboard_9_1.md` §15](./gerador_dashboard_9_1.md).
- **Dados comerciais não têm `QTD_QUARTOS`:** qualquer iteração nessa coluna deve checar `if 'QTD_QUARTOS' in df.columns` (foi a causa do KeyError corrigido na v9.1).
- **CDN:** Chart.js, jsPDF e SheetJS são carregados via CDN nos HTMLs. Em redes corporativas restritivas, gráficos e exportações falham.
- **Sessão única:** o usuário só pode estar logado em uma aba/navegador por vez (`sessions.json`).

---

## Licença

Reprodução proibida. Material proprietário da Opinião Informação Estratégica.
