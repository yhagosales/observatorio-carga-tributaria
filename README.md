# 📊 Observatório Macrofiscal: Carga Tributária vs. Retorno Social e Fuga de Capitais no Brasil (1995–2024)

[![Licence](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![Status](https://img.shields.io/badge/status-active-success.svg)]()
[![GitHub Pages](https://img.shields.io/badge/demo-GitHub%20Pages-brightgreen.svg)](index.html)

Repositório analítico e painel interativo que cruza a trajetória da **Carga Tributária Bruta (% do PIB)** com indicadores de **qualidade de vida**, **eficiência dos gastos públicos (Saúde e Educação)** e **movimentação de capitais (IED e Conta Financeira)** ao longo de **30 anos de série histórica oficial (1995–2024)**.

---

## 🚀 Demonstração Visual (Live Demo)
O projeto inclui um **Dashboard Interativo em HTML5/Canvas**:
- Para visualizar online via **GitHub Pages**, acesse: `https://<seu-usuario>.github.io/<nome-do-repositorio>/`
- Ou abra localmente o arquivo [`index.html`](index.html) em qualquer navegador web (funciona 100% offline).

---

## 📂 Estrutura do Repositório

```text
├── index.html                           # Dashboard interativo executivo (HTML5 + Tailwind + Canvas)
├── analise_impacto_tributario.py        # Pipeline em Python: coleta (APIs), modelagem e exportação Excel
├── analise_impacto_tributario.xlsx      # Planilha executiva final com 4 abas e formatação C-Level
├── requirements.txt                     # Dependências do projeto Python
├── .gitignore                           # Arquivos e diretórios ignorados pelo Git
├── LICENSE                              # Licença MIT
└── README.md                            # Documentação técnica e guia do projeto
```

---

## 📈 Principais Achados Macroeconômicos

1. **Expansão da Carga sem Contrapartida Proporcional:**
   - A Carga Tributária Bruta avançou de **26,69% do PIB (1995)** para **33,65% do PIB (2024)** (+6,96 p.p.).
   - Em esforço de trabalho, o cidadão passou de **102 dias/ano** para **149 dias/ano** apenas para pagar impostos (+47 dias a mais).
2. **Severa Ineficiência Alocativa em Educação:**
   - O investimento público total saltou de **3,50% para 5,80% do PIB** (superando a média da OCDE de ~4,9%), enquanto a proficiência no **PISA Matemática** estagnou em **379 pontos** (64º de 81 países).
3. **Avanço em Saúde Primária vs. Dupla Tributação:**
   - A expectativa de vida média subiu de **68,5 para 76,6 anos** e a mortalidade infantil caiu 70%, impulsionadas pelo SUS e atenção primária, embora 51 milhões de brasileiros recorram a planos privados para evitar filas.
4. **Fuga de Capitais e Reação a Choques Tributários:**
   - A modelagem econométrica defasada (OLS) apontou que aumentos bruscos na carga tributária afetam negativamente a atratividade do Investimento Estrangeiro Direto líquido com 1 ano de defasagem ($\Delta \text{IED}_{t-1} = -0,345$).
5. **O Paradoxo da Infraestrutura:**
   - Enquanto a arrecadação cresceu 20 vezes nominalmente, o investimento público federal em infraestrutura despencou de **1,80% para menos de 0,60% do PIB**.
6. **Ranking IRBES (IBPT):**
   - O Brasil ocupa consecutivamente o **30º e último lugar** entre os 30 países com maior arrecadação do mundo no retorno em bem-estar social.

---

## 🛠️ Como Executar o Script em Python

### Pré-requisitos
- Python 3.9 ou superior instalado.

### 1. Clonar o Repositório
```bash
git clone https://github.com/<seu-usuario>/<nome-do-repositorio>.git
cd <nome-do-repositorio>
```

### 2. Instalar as Dependências
```bash
pip install -r requirements.txt
```

### 3. Rodar a Coleta e Gerar a Planilha
```bash
python analise_impacto_tributario.py
```
O script se conectará às APIs oficiais (Banco Central do Brasil e Banco Mundial), executará as regressões e salvará automaticamente o arquivo `analise_impacto_tributario.xlsx`.

---

## 🌐 Como Publicar no GitHub Pages (Site Online Gratuito)

Para que qualquer pessoa possa acessar o seu painel interativo pela internet:

1. Acesse o seu repositório no GitHub.
2. Clique na aba **Settings** (Configurações) no menu superior.
3. No menu lateral esquerdo, clique em **Pages**.
4. Na seção **Branch**:
   - Selecione a branch `main` (ou `master`).
   - Mantenha a pasta selecionada como `/(root)`.
5. Clique em **Save**.
6. Em cerca de 1 a 2 minutos, o GitHub exibirá no topo o link público do seu site:  
   `https://<seu-usuario>.github.io/<nome-do-repositorio>/`

---

## 🏛️ Fontes de Dados Oficiais
- **Banco Central do Brasil (SGS):** Séries 4380 (Carga Tributária), 433 (IPCA), 22885 (IED) e 23460 (Conta Financeira).
- **Banco Mundial (World Bank Open Data):** Indicadores `GC.TAX.TOTL.GD.ZS`, `SE.XPD.TOTL.GD.ZS`, `SP.DYN.LE00.IN`.
- **OCDE:** PISA (Programme for International Student Assessment).
- **INEP / MEC:** IDEB (Índice de Desenvolvimento da Educação Básica).
- **PNUD:** Índice de Desenvolvimento Humano (IDH).
- **IBPT:** Dias Trabalhados para Pagar Tributos e Ranking IRBES.
- **DIEESE:** Cesta Básica Nacional e horas de trabalho necessárias.

---

## 📜 Licença
Distribuído sob a licença MIT. Consulte [`LICENSE`](LICENSE) para mais informações.
