

import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import folium
import plotly.express as px

from io import BytesIO

from streamlit_folium import st_folium

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image
)
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4

# =========================
# CONFIG STREAMLIT
# =========================

st.set_page_config(
    page_title="EA Comercial Intelligence",
    page_icon="📊",
    layout="wide"
)

st.markdown("""
<style>

/* =========================
FUNDO GERAL
========================= */

.stApp {
    background: linear-gradient(
        135deg,
        #111827,
        #1f2937
    );
    color: white;
}

/* =========================
TEXTOS
========================= */

html,
body,
[class*="css"]  {
    color: white !important;
}

/* =========================
TÍTULOS
========================= */

h1, h2, h3, h4, h5, h6 {
    color: white !important;
}

/* =========================
SIDEBAR
========================= */

section[data-testid="stSidebar"] {
    background-color: #111827;
}

section[data-testid="stSidebar"] * {
    color: white !important;
}

/* =========================
CARDS
========================= */

.ea-card {
    background: rgba(255,255,255,0.04);
    border-radius: 14px;
    padding: 15px;
    border: 1px solid rgba(255,255,255,0.08);
    margin-bottom: 15px;
}

/* =========================
SECTION TITLE
========================= */

.ea-section-title {
    font-size: 24px;
    font-weight: 700;
    color: white;
    margin-bottom: 20px;
}

/* =========================
DATAFRAME
========================= */

[data-testid="stDataFrame"] {
    background-color: rgba(255,255,255,0.03);
    border-radius: 12px;
}

/* =========================
TABS
========================= */

button[data-baseweb="tab"] {
    color: white !important;
}

/* =========================
MÉTRICAS
========================= */

[data-testid="metric-container"] {
    background: rgba(255,255,255,0.04);
    border-radius: 12px;
    padding: 10px;
    border: 1px solid rgba(255,255,255,0.08);
}

/* =========================
LABELS
========================= */

label {
    color: white !important;
}

/* =========================
SELECTBOX
========================= */

div[data-baseweb="select"] * {
    color: white !important;
    background-color: #1f2937 !important;
}

/* =========================
INPUTS
========================= */

input {
    color: white !important;
    background-color: #1f2937 !important;
}

/* =========================
PLOTLY
========================= */

.js-plotly-plot .plotly .main-svg {
    background: transparent !important;
}

</style>
""", unsafe_allow_html=True)
# =========================
# VISUAL EA COMERCIAL INTELLIGENCE
# =========================

st.markdown(
    """
    <style>

    #MainMenu {visibility:hidden;}
    footer {visibility:hidden;}
    header {visibility:hidden;}

    .stApp{
        background:#f4f6f9;
    }

    .block-container{
        padding-top:1rem;
        padding-bottom:2rem;
        max-width:1450px;
    }

    html, body, [class*="css"]{
        font-family: 'Segoe UI', sans-serif;
    }

    /* SIDEBAR */

    section[data-testid="stSidebar"]{
        background: linear-gradient(180deg,#020617 0%, #071226 100%);
        border-right:1px solid rgba(255,255,255,0.06);
    }

    section[data-testid="stSidebar"] *{
        color:white;
    }

    /* HERO */

    .ea-hero{
        background:
        linear-gradient(
            135deg,
            #c40000 0%,
            #930000 45%,
            #1a0000 100%
        );

        border-radius:22px;

        padding:34px;

        display:flex;
        justify-content:space-between;
        align-items:center;

        box-shadow:
        0px 10px 35px rgba(0,0,0,0.18);

        margin-bottom:28px;

        color:white;
    }

    .ea-hero-left h1{
        color:white;
        margin:0;
        font-size:52px;
        font-weight:900;
        letter-spacing:-2px;
    }

    .ea-hero-left p{
        margin-top:6px;
        font-size:17px;
        color:#f3f4f6;
    }

    .ea-pill{
        display:inline-block;

        margin-top:18px;

        background:rgba(255,255,255,0.12);

        padding:10px 16px;

        border-radius:999px;

        border:1px solid rgba(255,255,255,0.12);

        font-size:14px;
    }

    .ea-hero-right{
        display:flex;
        gap:18px;
        flex-wrap:wrap;
    }

    .ea-status-card{

        background:rgba(255,255,255,0.08);

        border:1px solid rgba(255,255,255,0.14);

        border-radius:18px;

        padding:18px;

        min-width:220px;

        backdrop-filter: blur(10px);
    }

    .ea-status-card h4{
        margin:0;
        color:white;
        font-size:16px;
    }

    .ea-status-card p{
        margin-top:8px;
        color:#f3f4f6;
        font-size:14px;
    }

    .ea-dot{
        width:12px;
        height:12px;
        border-radius:999px;
        background:#22c55e;
        margin-bottom:10px;
    }

    /* KPI */

    div[data-testid="stMetric"]{

        background:white;

        border-radius:20px;

        padding:24px;

        border:none;

        box-shadow:
        0px 4px 18px rgba(15,23,42,0.06);

        transition:0.2s;
    }

    div[data-testid="stMetric"]:hover{
        transform:translateY(-3px);
    }

    div[data-testid="stMetricLabel"]{
        font-size:15px;
        font-weight:600;
        color:#6b7280;
    }

    div[data-testid="stMetricValue"]{
        font-size:34px;
        font-weight:800;
        color:#111827;
    }

    /* CARDS */

    .ea-card{

        background:white;

        border-radius:22px;

        padding:26px;

        box-shadow:
        0px 4px 18px rgba(15,23,42,0.05);

        margin-bottom:24px;
    }

    /* TÍTULOS */

    .ea-section-title{

        font-size:34px;

        font-weight:900;

        color:#111827;

        margin-top:20px;

        margin-bottom:20px;

        letter-spacing:-1px;
    }

    /* TABELAS */

    div[data-testid="stDataFrame"]{

        border-radius:18px;

        overflow:hidden;

        border:none;

        box-shadow:
        0px 4px 16px rgba(15,23,42,0.05);
    }

    /* ALERTAS */

    .stAlert{

        border-radius:18px;

        border:none;

        box-shadow:
        0px 4px 14px rgba(15,23,42,0.05);
    }

    /* ABAS */

    .stTabs [data-baseweb="tab-list"]{
        gap:10px;
    }

    .stTabs [data-baseweb="tab"]{

        background:white;

        border-radius:14px 14px 0px 0px;

        padding-left:22px;
        padding-right:22px;

        height:54px;

        font-weight:700;

        color:#6b7280;
    }

    .stTabs [aria-selected="true"]{

        background:#b30000 !important;

        color:white !important;
    }

    /* MOBILE */

    @media(max-width:900px){

        .ea-hero{
            flex-direction:column;
            align-items:flex-start;
        }

        .ea-hero-left h1{
            font-size:38px;
        }

        .ea-hero-right{
            width:100%;
            flex-direction:column;
        }

        .ea-status-card{
            width:100%;
        }

        .ea-section-title{
            font-size:28px;
        }
    }

    </style>
    """,
    unsafe_allow_html=True
)

# =========================
# DADOS GOOGLE SHEETS
# =========================

url = "https://docs.google.com/spreadsheets/d/1JQe8XfQfaPEOkHWHuQb7XvQ2aBuuC9UlwwBmZ-Lg2jI/export?format=csv"

df = pd.read_csv(url)

# limpar nomes das colunas
df.columns = df.columns.str.strip()

# =========================
# PADRONIZAR NOMES DA NOVA PLANILHA
# =========================

df.columns = df.columns.str.strip()

df = df.rename(columns={

    "Carimbo de data/hora": "Data",

    # CLIENTE
    "1. Nome do cliente": "Cliente",
    "1. Nome do cliente ": "Cliente",

    # RESPONSÁVEL
    "2. Nome do responsável pela compra": "Responsável Compra",
    "2. Nome do responsável pela compra ": "Responsável Compra",

    # WHATSAPP
    "3. WhatsApp para contato": "WhatsApp",
    "3. WhatsApp para contato ": "WhatsApp",

    # SEGMENTO
    "4. Segmento do cliente": "Segmento",
    "4. Segmento do cliente ": "Segmento",

    # CIDADE
    "5. Cidade": "Cidade",
    "5. Cidade ": "Cidade",

    # BAIRRO
    "6. Bairro": "Bairro",
    "6. Bairro ": "Bairro",

    # ENDEREÇO
    "7. Endereço": "Endereço",
    "7. Endereço ": "Endereço",

    # CEP
    "8. CEP": "CEP",
    "8. CEP ": "CEP",

    # FORNECEDOR
    "9. Quem é o principal fornecedor hoje?": "Concorrente",
    "9. Quem é o principal fornecedor hoje? ": "Concorrente",

    # TEMPO
    "10. Há quanto tempo trabalha com esse fornecedor?": "Tempo Fornecedor",
    "10. Há quanto tempo trabalha com esse fornecedor? ": "Tempo Fornecedor",

    # FATOR DECISÃO
    "11. O que mais pesa na decisão de compra?": "Fator Decisão",
    "11. O que mais pesa na decisão de compra? ": "Fator Decisão",

    # FORNECEDOR ATENDE
    "12. O fornecedor atual atende totalmente sua necessidade?": "Fornecedor Atende",
    "12. O fornecedor atual atende totalmente sua necessidade? ": "Fornecedor Atende",

    # DIFICULDADE FORNECEDOR
    "13. Qual principal dificuldade com o fornecedor atual?": "Dificuldade Fornecedor Atual",
    "13. Qual principal dificuldade com o fornecedor atual? ": "Dificuldade Fornecedor Atual",

    # PROBLEMAS
    "14. Quais problemas mais acontecem hoje?": "Problemas Atuais",
    "14. Quais problemas mais acontecem hoje? ": "Problemas Atuais",

    # QUALIDADE
    "15. Como você avalia a qualidade do fornecedor atual?": "Qualidade",
    "15. Como você avalia a qualidade do fornecedor atual? ": "Qualidade",

    # FREQUÊNCIA
    "16. Como você avalia a frequência das entregas?": "Frequência Entrega",
    "16. Como você avalia a frequência das entregas? ": "Frequência Entrega",

    # ATENDIMENTO
    "17. Como você avalia o atendimento comercial atual?": "Atendimento Atual",
    "17. Como você avalia o atendimento comercial atual? ": "Atendimento Atual",

    # JÁ COMPROU PIFPAF
    "18. Você já comprou da PifPaf?": "Ja Comprou PifPaf",
    "18. Você já comprou da PifPaf? ": "Ja Comprou PifPaf",

    # PIFPAF PREÇO
    "19. Como você avalia a PifPaf atualmente?  [Preço]": "PifPaf Preço",
    "19. Como você avalia a PifPaf atualmente? [Preço]": "PifPaf Preço",

    # PIFPAF QUALIDADE
    "19. Como você avalia a PifPaf atualmente?  [Qualidade]": "PifPaf Qualidade",
    "19. Como você avalia a PifPaf atualmente? [Qualidade]": "PifPaf Qualidade",

    # PIFPAF ENTREGA
    "19. Como você avalia a PifPaf atualmente?  [Entrega]": "PifPaf Entrega",
    "19. Como você avalia a PifPaf atualmente? [Entrega]": "PifPaf Entrega",

    # PIFPAF ATENDIMENTO
    "19. Como você avalia a PifPaf atualmente?  [Atendimento]": "PifPaf Atendimento",
    "19. Como você avalia a PifPaf atualmente? [Atendimento]": "PifPaf Atendimento",

    # PIFPAF VARIEDADE
    "19. Como você avalia a PifPaf atualmente?  [Variedade de produtos]": "PifPaf Variedade",
    "19. Como você avalia a PifPaf atualmente? [Variedade de produtos]": "PifPaf Variedade",

    # PIFPAF NEGOCIAÇÃO
    "19. Como você avalia a PifPaf atualmente?  [Negociação]": "PifPaf Negociação",
    "19. Como você avalia a PifPaf atualmente? [Negociação]": "PifPaf Negociação",

    # OPORTUNIDADE PIFPAF
    "20. O que faria você comprar mais da PifPaf?": "Oportunidade PifPaf",
    "20. O que faria você comprar mais da PifPaf? ": "Oportunidade PifPaf",

    # DIFICULDADE PIFPAF
    "21. Qual principal dificuldade com a PifPaf hoje?": "Dificuldade PifPaf",
    "21. Qual principal dificuldade com a PifPaf hoje? ": "Dificuldade PifPaf",

    # CATEGORIA
    "22. Qual categoria possui maior volume de vendas?": "Categoria Mais Vendida",
    "22. Qual categoria possui maior volume de vendas? ": "Categoria Mais Vendida",

    # VOLUME
    "23. Qual o volume médio mensal aproximado?": "Volume Mensal",
    "23. Qual o volume médio mensal aproximado? ": "Volume Mensal",

    # INTERESSE
    "Existe interesse em trocar fornecedor?": "Interesse em Trocar",

    # POTENCIAL
    "Potencial do cliente": "Potencial",

    # ESTADO
    "Estado": "Estado",

    # COORDENADAS
    "Latitude": "Latitude",
    "Longitude": "Longitude",
    "Endereco Completo": "Endereco Completo"

})

df.columns = df.columns.str.strip()

# =========================
# COLUNAS PADRÃO
# =========================

colunas_padrao = {
    "Data": "",
    "Cliente": "Não informado",
    "Responsável Compra": "",
    "WhatsApp": "",
    "Segmento": "Não informado",
    "Cidade": "Não informado",
    "Estado": "MG",
    "Bairro": "",
    "Endereço": "",
    "CEP": "",

    "Concorrente": "Não informado",
    "Tempo Fornecedor": "",
    "Fator Decisão": "",
    "Fornecedor Atende": "",
    "Dificuldade Fornecedor Atual": "",
    "Problemas Atuais": "",

    "Qualidade": 0,
    "Frequência Entrega": 0,
    "Atendimento Atual": 0,

    "Ja Comprou PifPaf": "",
    "PifPaf Preço": "",
    "PifPaf Qualidade": "",
    "PifPaf Entrega": "",
    "PifPaf Atendimento": "",
    "PifPaf Variedade": "",
    "PifPaf Negociação": "",

    "Oportunidade PifPaf": "",
    "Dificuldade PifPaf": "",

    "Categoria Mais Vendida": "Não informado",
    "Volume Mensal": 0,
    "Potencial": "Baixo",
    "Interesse em Trocar": "Não",

    "Status Comercial": "Prospectar",

    # compatibilidade com o código antigo
    "Índice de Atraso": 0,
    "Itens com Ruptura": 0
}

for coluna, valor in colunas_padrao.items():
    if coluna not in df.columns:
        df[coluna] = valor

for coluna, valor in colunas_padrao.items():
    df[coluna] = df[coluna].replace(["", "nan", "None"], valor)

# =========================
# CONVERTER TEXTO EM NÚMERO
# =========================

def converter_volume(valor):
    valor = str(valor).strip()

    if "Até" in valor:
        return 2000
    elif "2.000" in valor and "5.000" in valor:
        return 5000
    elif "5.000" in valor and "10.000" in valor:
        return 10000
    elif "Acima" in valor:
        return 15000
    else:
        valor = (
            valor
            .replace("R$", "")
            .replace(".", "")
            .replace(",", ".")
        )
        return pd.to_numeric(valor, errors="coerce")

df["Volume Mensal"] = df["Volume Mensal"].apply(converter_volume).fillna(0)

colunas_numericas = [
    "Qualidade",
    "Frequência Entrega",
    "Atendimento Atual",
    "Índice de Atraso",
    "Itens com Ruptura"
]

for coluna in colunas_numericas:
    df[coluna] = pd.to_numeric(df[coluna], errors="coerce").fillna(0)

# usar frequência como base de atraso operacional
df["Índice de Atraso"] = 5 - df["Frequência Entrega"]

# usar problemas atuais para estimar ruptura
df["Problemas Atuais"] = df["Problemas Atuais"].fillna("").astype("string").astype(str)

df["Itens com Ruptura"] = df["Problemas Atuais"].apply(
    lambda x: 5 if ("Ruptura" in str(x)) or ("Falta de produtos" in str(x)) else 0
)

# =========================
# COMPARAÇÃO LINHA A LINHA
# =========================

def nota_texto(valor):

    if pd.isna(valor):
        return 0

    valor = str(valor).strip().lower()

    mapa = {
        "péssimo": 1,
        "pessimo": 1,
        "ruim": 2,
        "regular": 3,
        "bom": 4,
        "excelente": 5,
        "ótimo": 5,
        "otimo": 5
    }

    if valor in mapa:
        return mapa[valor]

    try:
        return float(valor)

    except:
        return 0


colunas_avaliacao = [
    "Qualidade",
    "Frequência Entrega",
    "Atendimento Atual",
    "PifPaf Preço",
    "PifPaf Qualidade",
    "PifPaf Entrega",
    "PifPaf Atendimento",
    "PifPaf Variedade",
    "PifPaf Negociação"
]

for coluna in colunas_avaliacao:

    if coluna not in df.columns:
        df[coluna] = 0

    df[coluna] = df[coluna].apply(nota_texto)

    df[coluna] = pd.to_numeric(
        df[coluna],
        errors="coerce"
    ).fillna(0)


df["Score Fornecedor Atual"] = (
    df["Qualidade"] +
    df["Frequência Entrega"] +
    df["Atendimento Atual"]
) / 3


df["Score PifPaf"] = (
    df["PifPaf Preço"] +
    df["PifPaf Qualidade"] +
    df["PifPaf Entrega"] +
    df["PifPaf Atendimento"] +
    df["PifPaf Variedade"] +
    df["PifPaf Negociação"]
) / 6


df["Vantagem PifPaf"] = (
    df["Score PifPaf"] -
    df["Score Fornecedor Atual"]
)


df["Diagnóstico Comparativo"] = df["Vantagem PifPaf"].apply(
    lambda x: "PifPaf melhor posicionada" if x > 0.5
    else "Fornecedor atual melhor posicionado" if x < -0.5
    else "Disputa equilibrada"
)

# =========================
# PADRONIZAR PIFPAF
# =========================

df["Concorrente"] = df["Concorrente"].replace({
    "pifpaf": "PifPaf",
    "Pif paf": "PifPaf",
    "Pif Paf": "PifPaf",
    "pif paf": "PifPaf"
})

# st.success("✅ Dados carregados automaticamente do Google Sheets")

# =========================
# TOPO PREMIUM SEM HTML
# =========================

st.container(border=True)

col_titulo, col_pifpaf, col_status = st.columns([4,1,1])


with col_titulo:

    st.markdown("""
    # EA Comercial Intelligence

    CRM Analítico Comercial e Operacional

    Região: Vespasiano, Confins e São José da Lapa
    """)

st.success("🟢 Status: Online")


   # st.info("Base conectada: Google Forms + Google Sheets")
# =========================
# SIDEBAR / FILTROS
# =========================

st.sidebar.image(
    "assets/logo_ea.png",
    width=180
)

st.sidebar.caption("CRM Analítico Comercial")

st.sidebar.markdown("---")
st.sidebar.subheader("Filtros Estratégicos")

cidades_lista = sorted(df["Cidade"].dropna().unique())
clientes_lista = sorted(df["Cliente"].dropna().unique())
empresas_lista = sorted(df["Concorrente"].dropna().unique())
potenciais_lista = sorted(df["Potencial"].dropna().unique())

cidades = st.sidebar.multiselect(
    "Cidade",
    cidades_lista,
    default=cidades_lista
)

clientes = st.sidebar.multiselect(
    "Cliente",
    clientes_lista,
    default=clientes_lista
)

empresas = st.sidebar.multiselect(
    "Empresa / Concorrente",
    empresas_lista,
    default=empresas_lista
)

potenciais = st.sidebar.multiselect(
    "Potencial",
    potenciais_lista,
    default=potenciais_lista
)

st.sidebar.markdown("---")
st.sidebar.caption("Euler Antonio de Souza")
st.sidebar.caption("Status: Online")

df_filtrado = df[
    (df["Cidade"].isin(cidades)) &
    (df["Cliente"].isin(clientes)) &
    (df["Concorrente"].isin(empresas)) &
    (df["Potencial"].isin(potenciais))
]
# =========================
# ABAS DO SISTEMA
# =========================

aba_geral, aba_operacao, aba_clientes, aba_inteligencia, aba_relatorio, aba_mapa = st.tabs([
    "Visão Geral",
    "Operação",
    "Clientes",
    "Inteligência Comercial",
    "Relatório",
    "Mapa"
])

# =========================
# ABA 1 — VISÃO GERAL
# =========================

with aba_geral:

    st.markdown("## Indicadores Estratégicos")

    total_clientes = len(df_filtrado)

    score_fornecedor = round(
        df_filtrado["Score Fornecedor Atual"].mean(),
        1
    )

    score_pifpaf = round(
        df_filtrado["Score PifPaf"].mean(),
        1
    )

    vantagem_media = round(
        df_filtrado["Vantagem PifPaf"].mean(),
        1
    )

    clientes_alto = len(
        df_filtrado[
            df_filtrado["Potencial"] == "Alto"
        ]
    )

    col1, col2, col3, col4, col5 = st.columns(5)

    with col1:
        st.info(f"""
### Clientes
# {total_clientes}
Base total cadastrada
""")

    with col2:
        st.warning(f"""
### Score Fornecedor
# {score_fornecedor}
Fornecedor atual
""")

    with col3:
        st.success(f"""
### Score PifPaf
# {score_pifpaf}
Percepção da marca
""")

    with col4:
        st.info(f"""
### Vantagem PifPaf
# {vantagem_media}
Diferença média
""")

    with col5:
        st.info(f"""
### Alto Potencial
# {clientes_alto}
Clientes prioritários
""")

    # =========================
    # TABELA EXECUTIVA
    # =========================

    st.markdown("---")

    st.markdown(
        '<div class="ea-section-title">Dados da Pesquisa</div>',
        unsafe_allow_html=True
    )

    colunas_tabela = [
        "Data",
        "Cliente",
        "Segmento",
        "Concorrente",
        "Cidade",
        "Volume Mensal",
        "Potencial",
        "Oportunidade PifPaf",
        "Dificuldade PifPaf",

        "Score Fornecedor Atual",
        "Score PifPaf",
        "Vantagem PifPaf",
        "Diagnóstico Comparativo",

        "Status Comercial"
    ]

    colunas_existentes = [
        coluna for coluna in colunas_tabela
        if coluna in df_filtrado.columns
    ]

    st.dataframe(
        df_filtrado[colunas_existentes],
        use_container_width=True,
        hide_index=True
    )

    st.markdown('</div>', unsafe_allow_html=True)

# =========================
# ABA 2 — OPERAÇÃO
# =========================

with aba_operacao:

    st.markdown(
        '<div class="ea-section-title">Operação Comparativa</div>',
        unsafe_allow_html=True
    )

    st.markdown("### PifPaf x Fornecedor Atual")

    comparativo_score = pd.DataFrame({
        "Grupo": ["Fornecedor Atual", "PifPaf"],
        "Score Médio": [
            df_filtrado["Score Fornecedor Atual"].mean(),
            df_filtrado["Score PifPaf"].mean()
        ]
    })

    fig_score = px.bar(
        comparativo_score,
        x="Grupo",
        y="Score Médio",
        color="Grupo",
        text="Score Médio",
        title="Score Médio Comparativo"
    )

    fig_score.update_traces(
        texttemplate="%{text:.2f}",
        textposition="outside"
    )

    fig_score.update_layout(
        height=420,
        yaxis=dict(range=[0, 5]),
        template="plotly_dark",
        showlegend=False
    )

    st.plotly_chart(fig_score, use_container_width=True)

    st.markdown("---")
    st.markdown("### Diagnóstico por Cliente")

    tabela_comparativa = df_filtrado[
        [
            "Cliente",
            "Concorrente",
            "Cidade",
            "Score Fornecedor Atual",
            "Score PifPaf",
            "Vantagem PifPaf",
            "Diagnóstico Comparativo"
        ]
    ].sort_values(
        by="Vantagem PifPaf",
        ascending=False
    )

    st.dataframe(
        tabela_comparativa,
        use_container_width=True,
        hide_index=True
    )

    st.markdown("---")
    st.markdown("### Maiores Oportunidades para PifPaf")

    oportunidades = tabela_comparativa[
        tabela_comparativa["Diagnóstico Comparativo"] == "PifPaf melhor posicionada"
    ]

    if oportunidades.empty:
        st.warning("Nenhuma oportunidade clara identificada no momento.")
    else:
        st.success("Clientes onde a PifPaf está melhor posicionada:")
        st.dataframe(
            oportunidades,
            use_container_width=True,
            hide_index=True
        )

# =========================
# ABA — MAPA
# =========================

with aba_mapa:

    st.markdown("## Mapa Comercial")

    df_map = df_filtrado.copy()

    if "Latitude" not in df_map.columns:
        df_map["Latitude"] = None

    if "Longitude" not in df_map.columns:
        df_map["Longitude"] = None

    df_map["Latitude"] = (
        df_map["Latitude"]
        .astype(str)
        .str.replace(",", ".", regex=False)
    )

    df_map["Longitude"] = (
        df_map["Longitude"]
        .astype(str)
        .str.replace(",", ".", regex=False)
    )

    df_map["Latitude"] = pd.to_numeric(
        df_map["Latitude"],
        errors="coerce"
    )

    df_map["Longitude"] = pd.to_numeric(
        df_map["Longitude"],
        errors="coerce"
    )

    df_map = df_map.dropna(
        subset=["Latitude", "Longitude"]
    )

    if df_map.empty:

        st.warning(
            "Nenhuma coordenada encontrada."
        )

    else:

        mapa = folium.Map(
            location=[
                df_map["Latitude"].mean(),
                df_map["Longitude"].mean()
            ],
            zoom_start=10,
            tiles="CartoDB positron"
        )

        for _, row in df_map.iterrows():

            atraso = row["Índice de Atraso"]
            ruptura = row["Itens com Ruptura"]
            qualidade = row["Qualidade"]

            risco = "Baixo"
            cor = "green"

            if atraso >= 4 or ruptura >= 5 or qualidade <= 2:
                risco = "Crítico"
                cor = "red"

            elif atraso >= 2 or ruptura >= 3:
                risco = "Atenção"
                cor = "orange"

            folium.CircleMarker(
                location=[
                    row["Latitude"],
                    row["Longitude"]
                ],
                radius=12,
                popup=f"""
                <b>Cliente:</b> {row['Cliente']}<br>
                <b>Cidade:</b> {row['Cidade']}<br>
                <b>Fornecedor:</b> {row['Concorrente']}<br>
                <b>Potencial:</b> {row['Potencial']}<br>
                <b>Risco:</b> {risco}
                """,
                color=cor,
                fill=True,
                fill_color=cor,
                fill_opacity=0.7
            ).add_to(mapa)

        st_folium(
            mapa,
            width=1200,
            height=540
        )
# =========================
# ABA 4 — CLIENTES
# =========================

with aba_clientes:

    st.markdown(
        '<div class="ea-section-title">Gargalos Estratégicos por Cliente</div>',
        unsafe_allow_html=True
    )

    gargalos = df_filtrado[
        (
            df_filtrado["Diagnóstico Comparativo"]
            == "Fornecedor atual melhor posicionado"
        )
    ]

    equilibrados = df_filtrado[
        (
            df_filtrado["Diagnóstico Comparativo"]
            == "Disputa equilibrada"
        )
    ]

    oportunidades = df_filtrado[
        (
            df_filtrado["Diagnóstico Comparativo"]
            == "PifPaf melhor posicionada"
        )
    ]

    if gargalos.empty and equilibrados.empty:

        st.success(
            "A PifPaf apresenta posicionamento competitivo positivo na base analisada."
        )

    else:

        if not gargalos.empty:

            st.error(
                f"{len(gargalos)} cliente(s) apresentam vantagem para o fornecedor atual."
            )

            for _, row in gargalos.iterrows():

                st.error(f"""
📍 {row['Cidade']}

Cliente: {row['Cliente']}

Fornecedor Atual: {row['Concorrente']}

Score Fornecedor: {round(row['Score Fornecedor Atual'], 2)}

Score PifPaf: {round(row['Score PifPaf'], 2)}

Vantagem PifPaf: {round(row['Vantagem PifPaf'], 2)}
""")

        if not equilibrados.empty:

            st.warning(
                f"{len(equilibrados)} cliente(s) estão em disputa equilibrada."
            )

            for _, row in equilibrados.iterrows():

                st.warning(f"""
📍 {row['Cidade']}

Cliente: {row['Cliente']}

Fornecedor Atual: {row['Concorrente']}

Disputa equilibrada entre fornecedor atual e PifPaf.
""")

        if not oportunidades.empty:

            st.success(
                f"{len(oportunidades)} cliente(s) apresentam alta oportunidade para avanço comercial da PifPaf."
            )

            for _, row in oportunidades.iterrows():

                st.success(f"""
📍 {row['Cidade']}

Cliente: {row['Cliente']}

Fornecedor Atual: {row['Concorrente']}

Score Fornecedor: {round(row['Score Fornecedor Atual'], 2)}

Score PifPaf: {round(row['Score PifPaf'], 2)}

Vantagem PifPaf: {round(row['Vantagem PifPaf'], 2)}
""")
# =========================
# RANKING DE QUALIDADE
# =========================

st.markdown("---")
st.markdown(
    '<div class="ea-section-title">Ranking de Qualidade</div>',
    unsafe_allow_html=True
)

ranking_qualidade = df_filtrado.sort_values(
    by="Qualidade",
    ascending=False
)

st.dataframe(
    ranking_qualidade[
        [
            "Cliente",
            "Concorrente",
            "Cidade",
            "Qualidade",
            "Potencial"
        ]
    ],
    use_container_width=True,
    hide_index=True
)

# =========================
# INSIGHTS
# =========================

st.markdown("---")
st.markdown(
    '<div class="ea-section-title">Insights Estratégicos</div>',
    unsafe_allow_html=True
)

melhor_qualidade = df_filtrado.loc[df_filtrado["Qualidade"].idxmax()]
pior_atraso = df_filtrado.loc[df_filtrado["Índice de Atraso"].idxmax()]
maior_ruptura = df_filtrado.loc[df_filtrado["Itens com Ruptura"].idxmax()]

st.success(f"✅ Melhor qualidade: {melhor_qualidade['Concorrente']}")
st.warning(f"⚠️ Maior atraso: {pior_atraso['Concorrente']}")
st.error(f"🚨 Maior ruptura: {maior_ruptura['Concorrente']}")

# =========================
# SCORE COMERCIAL PREMIUM
# =========================

st.markdown("---")
st.markdown("## Ranking de Oportunidades")

df_score = df_filtrado[
    df_filtrado["Concorrente"] != "PifPaf"
].copy()

if df_score.empty:

    st.warning("Não existem concorrentes suficientes para gerar ranking de oportunidades.")

else:

    scores = []

    for _, row in df_score.iterrows():

        pontos = 0
        pontos += row["Índice de Atraso"] * 10
        pontos += row["Itens com Ruptura"] * 6

        if row["Qualidade"] <= 2:
            pontos += 20

        if row["Potencial"] == "Alto":
            pontos += 30

        elif row["Potencial"] == "Médio":
            pontos += 15

        if row["Interesse em Trocar"] == "Sim":
            pontos += 35

        scores.append(pontos)

    df_score["Score Comercial"] = scores

    def classificar_prioridade(score):

        if score >= 90:
            return "Alta"

        elif score >= 60:
            return "Média"

        return "Baixa"

    df_score["Prioridade"] = df_score["Score Comercial"].apply(classificar_prioridade)

    ranking_score = df_score.sort_values(
        by="Score Comercial",
        ascending=False
    )

    top_oportunidade = ranking_score.iloc[0]

    st.markdown("### Principal Oportunidade Comercial")

    st.success(f"Cliente: {top_oportunidade['Cliente']}")
    st.write(f"Score Comercial: {top_oportunidade['Score Comercial']}")
    st.write(f"Prioridade: {top_oportunidade['Prioridade']}")
    st.write(f"Concorrente Atual: {top_oportunidade['Concorrente']}")

    st.dataframe(
        ranking_score[
            [
                "Cliente",
                "Concorrente",
                "Cidade",
                "Potencial",
                "Score Comercial",
                "Prioridade"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )

    # =========================
    # COMPARATIVO PIFPAF
    # =========================

    st.markdown("---")

    st.markdown(
        '<div class="ea-section-title">Comparativo PifPaf x Fornecedor Atual</div>',
        unsafe_allow_html=True
    )

    score_fornecedor = df_filtrado["Score Fornecedor Atual"].mean()
    score_pifpaf = df_filtrado["Score PifPaf"].mean()
    vantagem_media = df_filtrado["Vantagem PifPaf"].mean()

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown("### Score Fornecedor")
        st.write(f"Fornecedor Atual: {round(score_fornecedor, 2)}")

    with col2:
        st.markdown("### Score PifPaf")
        st.write(f"PifPaf: {round(score_pifpaf, 2)}")

    with col3:
        st.markdown("### Vantagem PifPaf")
        st.write(f"Diferença média: {round(vantagem_media, 2)}")

    if score_pifpaf > score_fornecedor:
        st.success("A PifPaf apresenta melhor posicionamento médio na base analisada.")
    elif score_pifpaf < score_fornecedor:
        st.warning("O fornecedor atual apresenta melhor posicionamento médio que a PifPaf.")
    else:
        st.info("A disputa entre PifPaf e fornecedor atual está equilibrada.")
    # =========================
    # OPORTUNIDADES E AÇÕES RECOMENDADAS
    # =========================

    st.markdown("---")

    st.markdown(
        '<div class="ea-section-title">Oportunidades e Ações Recomendadas</div>',
        unsafe_allow_html=True
    )

    for _, row in df_filtrado.iterrows():

        cliente = row["Cliente"]
        empresa = row["Concorrente"]
        atraso = row["Índice de Atraso"]
        ruptura = row["Itens com Ruptura"]
        qualidade = row["Qualidade"]
        potencial = row["Potencial"]

        categoria = (
            row["Categoria Mais Vendida"]
            if "Categoria Mais Vendida" in df_filtrado.columns
            else "categoria Indicadores Operacionais"
        )

        if empresa == "PifPaf":

            if atraso >= 4 or ruptura >= 5 or qualidade <= 2:

                st.error(f"""
Cliente: {cliente}

Gargalo interno identificado na operação da PifPaf.

Atraso: {atraso}
Rupturas: {ruptura}
Qualidade: {qualidade}

Ações recomendadas:
- Revisar rota de entrega
- Avaliar aumento da frequência de abastecimento
- Verificar cadeia fria e conservação
- Reduzir ruptura no cliente
- Criar plano de acompanhamento comercial prioritário
""")

            else:

                st.success(f"""
Cliente: {cliente}

A operação da PifPaf apresenta estabilidade operacional.

Ações recomendadas:
- Manter nível de serviço
- Preservar relacionamento comercial
- Monitorar indicadores preventivamente
""")

        else:

            if atraso >= 4 or ruptura >= 5 or qualidade <= 2:

                st.warning(f"""
Cliente: {cliente}

Concorrente atual: {empresa}

Falhas operacionais identificadas no fornecedor atual.

Atraso: {atraso}
Rupturas: {ruptura}
Qualidade: {qualidade}

Oportunidade para PifPaf:
- Realizar visita comercial
- Apresentar proposta PifPaf
- Destacar estabilidade logística
- Oferecer mix estratégico de {categoria}
- Usar falhas do concorrente como argumento comercial
""")

            elif potencial == "Alto":

                st.info(f"""
Cliente: {cliente}

Concorrente atual: {empresa}

Cliente de alto potencial, porém com baixa vulnerabilidade operacional no momento.

Ações recomendadas:
- Manter relacionamento ativo
- Monitorar fornecedor atual
- Mapear oportunidades futuras
- Acompanhar volume e categoria comprada
""")

            else:

                st.success(f"""
Cliente: {cliente}

Concorrente atual: {empresa}

Fornecedor atual apresenta estabilidade operacional.

Ação recomendada:
- Monitorar periodicamente
- Manter cliente no radar comercial
""")

# =========================
# ANÁLISE ESTRATÉGICA AVANÇADA
# =========================

st.markdown("---")

st.markdown(
    '<div class="ea-section-title">Análise Estratégica Avançada</div>',
    unsafe_allow_html=True
)

analises = []

score_fornecedor = df_filtrado["Score Fornecedor Atual"].mean()
score_pifpaf = df_filtrado["Score PifPaf"].mean()
vantagem_media = df_filtrado["Vantagem PifPaf"].mean()
volume_total = df_filtrado["Volume Mensal"].sum()

clientes_prioritarios = len(
    df_filtrado[df_filtrado["Potencial"] == "Alto"]
)

interesse_troca = len(
    df_filtrado[
        df_filtrado["Interesse em Trocar"]
        .astype(str)
        .str.contains("Sim|Talvez", case=False, na=False)
    ]
)

pifpaf_melhor = len(
    df_filtrado[
        df_filtrado["Diagnóstico Comparativo"] == "PifPaf melhor posicionada"
    ]
)

fornecedor_melhor = len(
    df_filtrado[
        df_filtrado["Diagnóstico Comparativo"] == "Fornecedor atual melhor posicionado"
    ]
)

equilibrado = len(
    df_filtrado[
        df_filtrado["Diagnóstico Comparativo"] == "Disputa equilibrada"
    ]
)

if score_pifpaf > score_fornecedor:

    analises.append(
        f"A PifPaf apresenta score médio superior ao fornecedor atual "
        f"({round(score_pifpaf, 2)} contra {round(score_fornecedor, 2)}). "
        "Isso indica vantagem competitiva percebida na base analisada."
    )

elif score_pifpaf < score_fornecedor:

    analises.append(
        f"O fornecedor atual apresenta score médio superior à PifPaf "
        f"({round(score_fornecedor, 2)} contra {round(score_pifpaf, 2)}). "
        "Isso indica necessidade de reforçar proposta comercial, negociação ou percepção de valor."
    )

else:

    analises.append(
        "A PifPaf e o fornecedor atual apresentam score médio equivalente. "
        "O cenário indica disputa equilibrada."
    )

analises.append(
    f"Na análise cliente por cliente, a PifPaf está melhor posicionada em "
    f"{pifpaf_melhor} cliente(s), o fornecedor atual lidera em "
    f"{fornecedor_melhor} cliente(s), e há disputa equilibrada em "
    f"{equilibrado} cliente(s)."
)

if clientes_prioritarios > 0:

    analises.append(
        f"Existem {clientes_prioritarios} cliente(s) de alto potencial que devem receber atenção comercial prioritária."
    )

if interesse_troca > 0:

    analises.append(
        f"Foram identificados {interesse_troca} cliente(s) com abertura para troca ou ampliação de fornecimento."
    )

if vantagem_media > 0.5:

    analises.append(
        "A vantagem média da PifPaf é relevante. A recomendação é acelerar visitas comerciais, apresentar proposta estruturada e explorar os pontos fortes percebidos."
    )

elif vantagem_media < -0.5:

    analises.append(
        "A desvantagem média da PifPaf indica necessidade de ação corretiva em preço, entrega, atendimento ou negociação antes de buscar conversão direta."
    )

else:

    analises.append(
        "A vantagem média está próxima do equilíbrio. A recomendação é trabalhar relacionamento, negociação consultiva e acompanhamento contínuo."
    )

if volume_total >= 10000:

    analises.append(
        "O volume mensal analisado demonstra potencial relevante para crescimento regional."
    )

for texto in analises:

    st.info(texto)

# =========================
# ABA — INTELIGÊNCIA COMERCIAL
# =========================

with aba_inteligencia:

    st.markdown("## Inteligência Comercial")

    # =========================
    # SCORE DE CONVERSÃO INTELIGENTE
    # =========================

    df_score = df_filtrado.copy()

    df_score["Score Conversão"] = (
        (
            df_score["Vantagem PifPaf"].fillna(0) * 25
        )
        +
        (
            df_score["Score PifPaf"].fillna(0) * 10
        )
    )

    # bônus por interesse

    df_score.loc[
        df_score["Interesse em Trocar"]
        .astype(str)
        .str.contains("Sim", case=False, na=False),

        "Score Conversão"
    ] += 30

    df_score.loc[
        df_score["Interesse em Trocar"]
        .astype(str)
        .str.contains("Talvez", case=False, na=False),

        "Score Conversão"
    ] += 15

    # bônus por potencial

    df_score.loc[
        df_score["Potencial"]
        .astype(str)
        .str.contains("Alto", case=False, na=False),

        "Score Conversão"
    ] += 20

    ranking = df_score.sort_values(
        by="Score Conversão",
        ascending=False
    )

    # =========================
    # FATORES DE DECISÃO
    # =========================

    st.markdown("---")
    st.subheader("Fatores de Decisão de Compra")

    fatores = (
        df_filtrado["Fator Decisão"]
        .astype(str)
        .str.split(",")
        .explode()
        .str.strip()
        .value_counts()
        .reset_index()
    )

    fatores.columns = ["Fator", "Quantidade"]

    fig_fatores = px.bar(
        fatores,
        x="Fator",
        y="Quantidade",
        color="Quantidade",
        template="plotly_dark",
        title="O que mais pesa na decisão de compra"
    )

    st.plotly_chart(fig_fatores, use_container_width=True)

    # =========================
    # PROBLEMAS DOS CONCORRENTES
    # =========================

    st.markdown("---")
    st.subheader("Principais Problemas dos Concorrentes")

    problemas = (
        df_filtrado["Problemas Atuais"]
        .astype(str)
        .str.split(",")
        .explode()
        .str.strip()
        .value_counts()
        .reset_index()
    )

    problemas.columns = ["Problema", "Quantidade"]

    fig_problemas = px.bar(
        problemas,
        x="Problema",
        y="Quantidade",
        color="Quantidade",
        template="plotly_dark",
        title="Problemas identificados nos fornecedores"
    )

    st.plotly_chart(fig_problemas, use_container_width=True)

    # =========================
    # OPORTUNIDADES PIFPAF
    # =========================

    st.markdown("---")
    st.subheader("O que faria comprar mais da PifPaf")

    oportunidades = (
        df_filtrado["Oportunidade PifPaf"]
        .astype(str)
        .str.split(",")
        .explode()
        .str.strip()
        .value_counts()
        .reset_index()
    )

    oportunidades.columns = ["Oportunidade", "Quantidade"]

    fig_oportunidades = px.bar(
        oportunidades,
        x="Oportunidade",
        y="Quantidade",
        color="Quantidade",
        template="plotly_dark",
        title="Oportunidades de crescimento PifPaf"
    )

    st.plotly_chart(fig_oportunidades, use_container_width=True)

    # =========================
    # DIFICULDADES PIFPAF
    # =========================

    st.markdown("---")
    st.subheader("Principais Dificuldades com a PifPaf")

    dificuldades = (
        df_filtrado["Dificuldade PifPaf"]
        .astype(str)
        .str.split(",")
        .explode()
        .str.strip()
        .value_counts()
        .reset_index()
    )

    dificuldades.columns = ["Dificuldade", "Quantidade"]

    fig_dificuldades = px.bar(
        dificuldades,
        x="Dificuldade",
        y="Quantidade",
        color="Quantidade",
        template="plotly_dark",
        title="Principais dificuldades com a PifPaf"
    )

    st.plotly_chart(fig_dificuldades, use_container_width=True)

    # =========================
    # PARTICIPAÇÃO DOS CONCORRENTES
    # =========================

    st.markdown("---")
    st.subheader("Participação dos Concorrentes")

    ranking_concorrentes = (
        df_filtrado["Concorrente"]
        .value_counts()
        .reset_index()
    )

    ranking_concorrentes.columns = ["Concorrente", "Clientes"]

    fig_concorrentes = px.pie(
        ranking_concorrentes,
        names="Concorrente",
        values="Clientes",
        template="plotly_dark",
        title="Participação dos fornecedores"
    )

    st.plotly_chart(fig_concorrentes, use_container_width=True)

    # =========================
    # POTENCIAL POR CIDADE
    # =========================

    st.markdown("---")
    st.subheader("Potencial Comercial por Cidade")

    potencial_cidade = (
        df_filtrado.groupby("Cidade")["Cliente"]
        .count()
        .reset_index()
    )

    potencial_cidade.columns = ["Cidade", "Clientes"]

    fig_cidade = px.bar(
        potencial_cidade,
        x="Cidade",
        y="Clientes",
        color="Clientes",
        template="plotly_dark",
        title="Clientes mapeados por cidade"
    )

    st.plotly_chart(fig_cidade, use_container_width=True)

    # =========================
    # RADAR — PERCEPÇÃO PIFPAF
    # =========================

    st.markdown("---")
    st.subheader("Radar de Percepção da PifPaf")

    colunas_radar = [
        "PifPaf Preço",
        "PifPaf Qualidade",
        "PifPaf Entrega",
        "PifPaf Atendimento",
        "PifPaf Variedade",
        "PifPaf Negociação"
    ]

    colunas_existentes_radar = [
        coluna for coluna in colunas_radar
        if coluna in df_filtrado.columns
    ]

    if len(colunas_existentes_radar) > 0:

        df_radar = df_filtrado[colunas_existentes_radar].copy()

        for coluna in colunas_existentes_radar:
            df_radar[coluna] = pd.to_numeric(
                df_radar[coluna],
                errors="coerce"
            ).fillna(0)

        radar_media = df_radar.mean().reset_index()
        radar_media.columns = ["Indicador", "Nota"]

        fig_radar = px.line_polar(
            radar_media,
            r="Nota",
            theta="Indicador",
            line_close=True,
            template="plotly_dark",
            title="Percepção média da PifPaf"
        )

        fig_radar.update_traces(fill="toself")

        st.plotly_chart(fig_radar, use_container_width=True)

    # =========================
    # HEATMAP
    # =========================

    st.markdown("---")
    st.subheader("Mapa de Força Comercial por Cidade e Concorrente")

    heatmap = (
        df_filtrado
        .groupby(["Cidade", "Concorrente"])["Cliente"]
        .count()
        .reset_index()
    )

    heatmap.columns = ["Cidade", "Concorrente", "Clientes"]

    fig_heatmap = px.density_heatmap(
        heatmap,
        x="Concorrente",
        y="Cidade",
        z="Clientes",
        template="plotly_dark",
        title="Concentração de clientes por cidade e fornecedor"
    )

    st.plotly_chart(fig_heatmap, use_container_width=True)

    # =========================
    # FUNIL COMERCIAL
    # =========================

    st.markdown("---")
    st.subheader("Funil Comercial")

    total_mapeados = len(df_filtrado)

    interessados = len(
        df_filtrado[
            df_filtrado["Interesse em Trocar"]
            .astype(str)
            .str.contains("Sim|Talvez", case=False, na=False)
        ]
    )

    alto_potencial = len(
        df_filtrado[
            df_filtrado["Potencial"]
            .astype(str)
            .str.contains("Alto", case=False, na=False)
        ]
    )

    oportunidades_qtd = len(
        df_filtrado[
            df_filtrado["Oportunidade PifPaf"].astype(str).str.len() > 3
        ]
    )

    funil = pd.DataFrame({
        "Etapa": [
            "Clientes mapeados",
            "Com abertura",
            "Alto potencial",
            "Oportunidade identificada"
        ],
        "Quantidade": [
            total_mapeados,
            interessados,
            alto_potencial,
            oportunidades_qtd
        ]
    })

    fig_funil = px.funnel(
        funil,
        x="Quantidade",
        y="Etapa",
        template="plotly_dark",
        title="Funil de oportunidades comerciais"
    )

    st.plotly_chart(fig_funil, use_container_width=True)

    # =========================
    # MATRIZ DE OPORTUNIDADE
    # =========================

    st.markdown("---")
    st.subheader("Matriz de Oportunidade Comercial")

    st.dataframe(
        ranking[
            [
                "Cliente",
                "Cidade",
                "Concorrente",
                "Volume Mensal",
                "Score Conversão",
                "Diagnóstico Comparativo"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )

    # =========================
    # RANKING DE CONVERSÃO
    # =========================

    st.markdown("---")
    st.subheader("Ranking de Conversão Comercial")

    st.dataframe(
        ranking[
            [
                "Cliente",
                "Cidade",
                "Concorrente",
                "Potencial",
                "Interesse em Trocar",
                "Score Conversão"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )
# =========================
# ABA 5 — RELATÓRIO / APRESENTAÇÃO
# =========================

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
from io import BytesIO

with aba_relatorio:

    st.markdown(
        '<div class="ea-section-title">Apresentação Executiva</div>',
        unsafe_allow_html=True
    )

    if st.button("Gerar Apresentação PowerPoint"):

        prs = Presentation()
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)

        vermelho = RGBColor(153, 27, 27)
        preto = RGBColor(17, 24, 39)
        cinza = RGBColor(107, 114, 128)

        def add_title(slide, title, subtitle=""):
            box = slide.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12), Inches(0.8))
            tf = box.text_frame
            tf.text = title
            p = tf.paragraphs[0]
            p.font.size = Pt(30)
            p.font.bold = True
            p.font.color.rgb = preto

            if subtitle:
                sub = slide.shapes.add_textbox(Inches(0.65), Inches(1.15), Inches(11.5), Inches(0.4))
                sub_tf = sub.text_frame
                sub_tf.text = subtitle
                sub_tf.paragraphs[0].font.size = Pt(14)
                sub_tf.paragraphs[0].font.color.rgb = cinza

        def add_metric(slide, x, y, title, value):
            shape = slide.shapes.add_shape(
                1, Inches(x), Inches(y), Inches(2.4), Inches(1.1)
            )
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(245, 245, 245)
            shape.line.color.rgb = RGBColor(220, 220, 220)

            tf = shape.text_frame
            tf.text = title
            tf.paragraphs[0].font.size = Pt(11)
            tf.paragraphs[0].font.color.rgb = cinza

            p = tf.add_paragraph()
            p.text = str(value)
            p.font.size = Pt(24)
            p.font.bold = True
            p.font.color.rgb = vermelho

        total_clientes = len(df_filtrado)
        volume_total = df_filtrado["Volume Mensal"].sum()
        score_fornecedor = round(df_filtrado["Score Fornecedor Atual"].mean(), 2)
        score_pifpaf = round(df_filtrado["Score PifPaf"].mean(), 2)
        vantagem_media = round(df_filtrado["Vantagem PifPaf"].mean(), 2)

        oportunidades = df_filtrado[
            df_filtrado["Diagnóstico Comparativo"] == "PifPaf melhor posicionada"
        ]

        gargalos = df_filtrado[
            df_filtrado["Diagnóstico Comparativo"] == "Fornecedor atual melhor posicionado"
        ]

        equilibrados = df_filtrado[
            df_filtrado["Diagnóstico Comparativo"] == "Disputa equilibrada"
        ]

        # SLIDE 1
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_title(slide, "EA Comercial Intelligence", "Apresentação Executiva — Análise PifPaf x Fornecedor Atual")

        tx = slide.shapes.add_textbox(Inches(0.8), Inches(2.0), Inches(11.5), Inches(2.0))
        tf = tx.text_frame
        tf.text = "Diagnóstico comercial, competitivo e operacional com base nas respostas do Forms."
        tf.paragraphs[0].font.size = Pt(24)
        tf.paragraphs[0].font.bold = True
        tf.paragraphs[0].font.color.rgb = preto

        add_metric(slide, 0.8, 4.2, "Clientes", total_clientes)
        add_metric(slide, 3.4, 4.2, "Score Fornecedor", score_fornecedor)
        add_metric(slide, 6.0, 4.2, "Score PifPaf", score_pifpaf)
        add_metric(slide, 8.6, 4.2, "Vantagem PifPaf", vantagem_media)

        # SLIDE 2
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_title(slide, "Resumo Estratégico", "Leitura geral da base analisada")

        resumo = f"""
Foram analisados {total_clientes} clientes.

Score médio do fornecedor atual: {score_fornecedor}
Score médio da PifPaf: {score_pifpaf}
Vantagem média da PifPaf: {vantagem_media}

Clientes onde a PifPaf está melhor posicionada: {len(oportunidades)}
Clientes onde o fornecedor atual lidera: {len(gargalos)}
Clientes em disputa equilibrada: {len(equilibrados)}

Volume mensal estimado: R$ {volume_total:,.0f}
"""

        box = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.8), Inches(4.8))
        tf = box.text_frame
        tf.text = resumo
        for p in tf.paragraphs:
            p.font.size = Pt(18)
            p.font.color.rgb = preto

        # SLIDE 3
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_title(slide, "Comparativo Geral", "Fornecedor Atual x PifPaf")

        rows = 3
        cols = 3
        table = slide.shapes.add_table(rows, cols, Inches(1), Inches(1.8), Inches(11), Inches(1.8)).table

        headers = ["Indicador", "Fornecedor Atual", "PifPaf"]
        values = [
            ["Score Médio", score_fornecedor, score_pifpaf],
            ["Diferença", "-", vantagem_media]
        ]

        for i, h in enumerate(headers):
            table.cell(0, i).text = h

        for r, row in enumerate(values, start=1):
            for c, val in enumerate(row):
                table.cell(r, c).text = str(val)

        for row in table.rows:
            for cell in row.cells:
                cell.text_frame.paragraphs[0].font.size = Pt(14)

        # SLIDE 4
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_title(slide, "Ranking Comparativo por Cliente", "Clientes ordenados pela vantagem da PifPaf")

        ranking = df_filtrado[
            [
                "Cliente",
                "Concorrente",
                "Cidade",
                "Score Fornecedor Atual",
                "Score PifPaf",
                "Vantagem PifPaf",
                "Diagnóstico Comparativo"
            ]
        ].sort_values(by="Vantagem PifPaf", ascending=False).head(8)

        table = slide.shapes.add_table(
            len(ranking) + 1,
            7,
            Inches(0.3),
            Inches(1.5),
            Inches(12.7),
            Inches(4.8)
        ).table

        headers = list(ranking.columns)

        for c, h in enumerate(headers):
            table.cell(0, c).text = h

        for r, (_, row) in enumerate(ranking.iterrows(), start=1):
            for c, h in enumerate(headers):
                table.cell(r, c).text = str(round(row[h], 2)) if isinstance(row[h], float) else str(row[h])

        for row in table.rows:
            for cell in row.cells:
                cell.text_frame.paragraphs[0].font.size = Pt(8)

        # SLIDE 5
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_title(slide, "Maiores Oportunidades", "Clientes onde a PifPaf está melhor posicionada")

        texto = ""
        if oportunidades.empty:
            texto = "Nenhuma oportunidade clara identificada no momento."
        else:
            for _, row in oportunidades.head(6).iterrows():
                texto += f"• {row['Cliente']} — {row['Cidade']} | Vantagem: {round(row['Vantagem PifPaf'], 2)}\n"

        box = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.8), Inches(4.8))
        tf = box.text_frame
        tf.text = texto
        for p in tf.paragraphs:
            p.font.size = Pt(18)

        # SLIDE 6
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_title(slide, "Gargalos Estratégicos", "Clientes onde o fornecedor atual ainda lidera")

        texto = ""
        if gargalos.empty:
            texto = "Nenhum gargalo estratégico crítico identificado."
        else:
            for _, row in gargalos.head(6).iterrows():
                texto += f"• {row['Cliente']} — {row['Concorrente']} | Diferença: {round(row['Vantagem PifPaf'], 2)}\n"

        box = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.8), Inches(4.8))
        tf = box.text_frame
        tf.text = texto
        for p in tf.paragraphs:
            p.font.size = Pt(18)

        # SLIDE 7
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_title(slide, "Plano de Ação Comercial", "Recomendações práticas")

        plano = """
1. Priorizar clientes onde a PifPaf já aparece melhor posicionada.
2. Criar abordagem consultiva para clientes em disputa equilibrada.
3. Para clientes onde o fornecedor atual lidera, reforçar preço, entrega, variedade e negociação.
4. Trabalhar propostas por categoria de maior volume.
5. Acompanhar evolução dos scores a cada nova pesquisa.
"""

        box = slide.shapes.add_textbox(Inches(0.8), Inches(1.5), Inches(11.8), Inches(4.8))
        tf = box.text_frame
        tf.text = plano
        for p in tf.paragraphs:
            p.font.size = Pt(19)

        # SLIDE 8
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        add_title(slide, "Conclusão Estratégica", "EA Comercial Intelligence")

        conclusao = """
A análise permite identificar com clareza onde a PifPaf possui vantagem competitiva,
onde há disputa equilibrada e onde o fornecedor atual ainda está mais forte.

O modelo transforma respostas do Forms em inteligência comercial prática,
orientando visitas, negociações e priorização de clientes.
"""

        box = slide.shapes.add_textbox(Inches(0.8), Inches(1.7), Inches(11.8), Inches(4.8))
        tf = box.text_frame
        tf.text = conclusao
        for p in tf.paragraphs:
            p.font.size = Pt(22)
            p.font.color.rgb = preto

        arquivo_pptx = BytesIO()
        prs.save(arquivo_pptx)
        arquivo_pptx.seek(0)

        st.download_button(
            label="Baixar Apresentação PowerPoint",
            data=arquivo_pptx,
            file_name="apresentacao_ea_intelligence.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )