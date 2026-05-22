import streamlit as st
import pandas as pd
import sqlite3
import os
import json
from datetime import datetime
from zoneinfo import ZoneInfo

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(BASE_DIR, "crm_pedidos.db")


def conectar():
    return sqlite3.connect(DB_NAME)


def criar_tabela_estoque():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS estoque_diario (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            data_importacao TEXT,
            arquivo_origem TEXT,
            aba_origem TEXT,
            centro TEXT,
            material TEXT,
            descricao TEXT,
            categoria TEXT,
            estoque_total REAL,
            dados_json TEXT
        )
    """)

    conn.commit()
    conn.close()


def identificar_coluna(df_base, opcoes):
    for coluna in opcoes:
        if coluna in df_base.columns:
            return coluna
    return None


def preparar_estoque(df_base):
    df_base = df_base.copy()
    df_base.columns = df_base.columns.astype(str).str.strip()
    df_base = df_base.dropna(how="all")

    coluna_material = identificar_coluna(
        df_base,
        ["Material", "Codigo", "Código", "Cod Material", "Código Material"]
    )

    coluna_descricao = identificar_coluna(
        df_base,
        ["Texto breve de material", "Produto", "Descrição", "Descricao", "Material Descrição"]
    )

    coluna_categoria = identificar_coluna(
        df_base,
        ["Categorias", "Categoria", "Grupo", "Linha"]
    )

    coluna_centro = identificar_coluna(
        df_base,
        ["Centro", "CD", "Unidade"]
    )

    hoje = datetime.now(ZoneInfo("America/Sao_Paulo"))

    meses_pt = {
        "01": "jan", "02": "fev", "03": "mar", "04": "abr",
        "05": "mai", "06": "jun", "07": "jul", "08": "ago",
        "09": "set", "10": "out", "11": "nov", "12": "dez"
    }

    data_pt = f"{hoje.strftime('%d')}.{meses_pt[hoje.strftime('%m')]}"

    possiveis_datas_hoje = [
        data_pt,
        hoje.strftime("%d/%m"),
        hoje.strftime("%d-%m"),
        hoje.strftime("%d.%m"),
    ]

    coluna_total = identificar_coluna(df_base, possiveis_datas_hoje)

    if coluna_total is None:
        coluna_total = identificar_coluna(
            df_base,
            ["Total", "Estoque", "Estoque Total", "Saldo", "Qtd", "Quantidade"]
        )

    if coluna_material is None:
        df_base["Material"] = ""
        coluna_material = "Material"

    if coluna_descricao is None:
        df_base["Texto breve de material"] = ""
        coluna_descricao = "Texto breve de material"

    if coluna_categoria is None:
        df_base["Categorias"] = "Não informado"
        coluna_categoria = "Categorias"

    if coluna_centro is None:
        df_base["Centro"] = "Não informado"
        coluna_centro = "Centro"

    if coluna_total is None:
        df_base["Total"] = 0
        coluna_total = "Total"

    df_base["_material"] = df_base[coluna_material].astype(str)
    df_base["_descricao"] = df_base[coluna_descricao].astype(str)
    df_base["_categoria"] = df_base[coluna_categoria].astype(str)
    df_base["_centro"] = df_base[coluna_centro].astype(str)
    df_base["_estoque_total"] = pd.to_numeric(
        df_base[coluna_total],
        errors="coerce"
    ).fillna(0)

    return df_base, {
        "material": coluna_material,
        "descricao": coluna_descricao,
        "categoria": coluna_categoria,
        "centro": coluna_centro,
        "total": coluna_total
    }


def salvar_estoque_no_banco(df_base, nome_arquivo, aba_origem):
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("DELETE FROM estoque_diario")
    conn.commit()

    data_importacao = datetime.now(
        ZoneInfo("America/Sao_Paulo")
    ).strftime("%d/%m/%Y %H:%M:%S")

    for _, row in df_base.iterrows():
        cursor.execute("""
            INSERT INTO estoque_diario (
                data_importacao,
                arquivo_origem,
                aba_origem,
                centro,
                material,
                descricao,
                categoria,
                estoque_total,
                dados_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            data_importacao,
            nome_arquivo,
            aba_origem,
            row.get("_centro", ""),
            row.get("_material", ""),
            row.get("_descricao", ""),
            row.get("_categoria", ""),
            float(row.get("_estoque_total", 0)),
            json.dumps(row.to_dict(), ensure_ascii=False, default=str)
        ))

    conn.commit()
    conn.close()

    return data_importacao


st.set_page_config(
    page_title="Estoque EA CRM",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

criar_tabela_estoque()

st.markdown("""
<style>
.stApp {
    background: linear-gradient(180deg, #07111f 0%, #0b1220 100%);
    color: #f8fafc;
}

.block-container {
    padding-top: 1rem;
    padding-left: 1rem;
    padding-right: 1rem;
    max-width: 100% !important;
}

h1, h2, h3, p, label, span {
    color: #f8fafc !important;
}

[data-testid="stSidebar"] {
    background: #050b16;
}

[data-testid="stSidebar"] * {
    color: #f8fafc !important;
}

.stTabs [data-baseweb="tab-list"] {
    gap: 12px;
    border-bottom: 1px solid rgba(255,255,255,0.12);
}

.stTabs [data-baseweb="tab"] {
    background: #111827;
    color: #cbd5e1;
    border-radius: 12px 12px 0 0;
    padding: 12px 20px;
    font-weight: 700;
}

.stTabs [aria-selected="true"] {
    background: #b30000 !important;
    color: white !important;
}

[data-testid="stMetric"] {
    background: linear-gradient(145deg, #111827, #1e293b);
    border: 1px solid rgba(255,255,255,0.10);
    border-radius: 20px;
    padding: 20px;
    box-shadow: 0 8px 28px rgba(0,0,0,0.35);
}

[data-testid="stMetricLabel"] {
    color: #94a3b8 !important;
    font-weight: 700;
}

[data-testid="stMetricValue"] {
    color: #ffffff !important;
    font-size: 34px !important;
    font-weight: 900;
}

input, textarea, div[data-baseweb="select"] > div {
    background: #111827 !important;
    color: #f8fafc !important;
    border: 1px solid rgba(255,255,255,0.18) !important;
    border-radius: 14px !important;
}

.stButton > button {
    background: #b30000 !important;
    color: #ffffff !important;
    border-radius: 12px !important;
    border: none !important;
    font-weight: 800 !important;
    padding: 0.7rem 1.2rem !important;
}

.stButton > button:hover {
    background: #d00000 !important;
}

[data-testid="stDataFrame"] {
    border-radius: 16px;
    overflow: hidden;
    border: 1px solid rgba(255,255,255,0.12);
}

.premium-header {
    background: linear-gradient(145deg, rgba(15,23,42,0.95), rgba(2,6,23,0.95));
    padding: 24px;
    border-radius: 24px;
    border: 1px solid rgba(255,255,255,0.10);
    margin-bottom: 22px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.35);
}

.premium-title {
    font-size: 38px;
    font-weight: 900;
    color: white;
    margin-bottom: 4px;
}

.premium-subtitle {
    color: #94a3b8;
    font-size: 16px;
}

.status-pill {
    background: rgba(22,163,74,0.18);
    color: #4ade80;
    padding: 12px 18px;
    border-radius: 16px;
    border: 1px solid rgba(74,222,128,0.25);
    font-weight: 800;
    display: inline-block;
}

.search-box {
    background: rgba(15,23,42,0.95);
    border: 1px solid rgba(255,255,255,0.10);
    border-radius: 20px;
    padding: 22px;
    margin-top: 18px;
    margin-bottom: 18px;
}

@media(max-width: 900px) {
    .block-container {
        padding-left: 1rem;
        padding-right: 1rem;
    }

    .premium-title {
        font-size: 34px;
        line-height: 1.1;
    }

    [data-testid="stMetricValue"] {
        font-size: 28px !important;
    }

    .stTabs [data-baseweb="tab"] {
        font-size: 15px;
        padding: 10px 12px;
    }

    input {
        font-size: 17px !important;
        min-height: 48px !important;
    }
}

@media (max-width: 768px) {
    .block-container {
        padding-left: 0.7rem !important;
        padding-right: 0.7rem !important;
        max-width: 100% !important;
    }

    [data-testid="stDataFrame"] {
        overflow-x: auto;
        width: 100% !important;
    }

    table {
        min-width: 900px;
    }

    [data-testid="stMetricValue"] {
        font-size: 24px !important;
    }

    .premium-title {
        font-size: 28px !important;
        line-height: 1.1;
    }

    .search-box {
        padding: 14px !important;
    }
}
</style>
""", unsafe_allow_html=True)

USUARIOS = {
    "euler": "123456",
    "estoque": "123456",
    "vendas": "123456"
}

if "logado_estoque" not in st.session_state:
    st.session_state.logado_estoque = False

if not st.session_state.logado_estoque:
    st.markdown('<div class="premium-header">', unsafe_allow_html=True)
    st.markdown('<div class="premium-title">🔐 Login Estoque EA CRM</div>', unsafe_allow_html=True)
    st.markdown('<div class="premium-subtitle">Acesso restrito para consulta e importação de estoque</div>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    usuario = st.text_input("Usuário")
    senha = st.text_input("Senha", type="password")

    if st.button("Entrar"):
        if usuario in USUARIOS and senha == USUARIOS[usuario]:
            st.session_state.logado_estoque = True
            st.session_state.usuario_estoque = usuario
            st.rerun()
        else:
            st.error("Usuário ou senha incorretos.")

    st.stop()

st.sidebar.markdown("## EA Comercial")
st.sidebar.markdown("### Intelligence")
st.sidebar.success(f"Usuário: {st.session_state.usuario_estoque}")

if st.sidebar.button("Sair"):
    st.session_state.logado_estoque = False
    st.rerun()

st.markdown("""
<div class="premium-header">
    <div class="premium-title">📦 Estoque Diário</div>
    <div class="premium-subtitle">Consulta e importação rápida de estoque para celular</div>
    <br>
    <span class="status-pill">● Status: Online</span>
</div>
""", unsafe_allow_html=True)

aba_consultar, aba_importar = st.tabs([
    "Consultar Estoque",
    "Importar Estoque"
])

with aba_importar:

    st.header("📤 Importar Planilha de Estoque")

    arquivo_estoque = st.file_uploader(
        "Envie a planilha de estoque",
        type=["xlsx", "xls"]
    )

    if arquivo_estoque:
        try:
            excel_estoque = pd.ExcelFile(arquivo_estoque)

            abas_validas = [
                aba for aba in excel_estoque.sheet_names
                if not str(aba).startswith("_")
            ]

            aba_padrao = "Geral" if "Geral" in abas_validas else abas_validas[0]

            aba_selecionada = st.selectbox(
                "Selecione a aba da planilha",
                abas_validas,
                index=abas_validas.index(aba_padrao)
            )

            df_original = pd.read_excel(
                arquivo_estoque,
                sheet_name=aba_selecionada
            )

            df_estoque, colunas_estoque = preparar_estoque(df_original)

            st.success("Estoque carregado com sucesso!")

            chave_upload = f"{arquivo_estoque.name}_{aba_selecionada}"

            if st.session_state.get("ultimo_estoque_salvo") != chave_upload:
                data_importacao = salvar_estoque_no_banco(
                    df_estoque,
                    arquivo_estoque.name,
                    aba_selecionada
                )

                st.session_state["ultimo_estoque_salvo"] = chave_upload

                st.success(
                    f"Estoque substituído e salvo automaticamente em {data_importacao}."
                )
            else:
                st.info("Este estoque já foi salvo nesta sessão.")

            col1, col2, col3, col4 = st.columns(4)

            col1.metric("Produtos", len(df_estoque))
            col2.metric("Estoque Total", f'{df_estoque["_estoque_total"].sum():,.2f}')
            col3.metric("Zerados", len(df_estoque[df_estoque["_estoque_total"] <= 0]))
            col4.metric("Aba", aba_selecionada)

            busca = st.text_input(
                "🔎 Pesquisar na planilha carregada",
                placeholder="Digite código, descrição ou categoria"
            )

            categorias = sorted(
                df_estoque["_categoria"].dropna().astype(str).unique()
            )

            categoria = st.selectbox("Categoria", ["Todas"] + categorias)

            df_resultado = df_estoque.copy()

            if busca:
                df_resultado = df_resultado[
                    df_resultado["_material"].str.contains(busca, case=False, na=False)
                    |
                    df_resultado["_descricao"].str.contains(busca, case=False, na=False)
                    |
                    df_resultado["_categoria"].str.contains(busca, case=False, na=False)
                ]

            if categoria != "Todas":
                df_resultado = df_resultado[
                    df_resultado["_categoria"].astype(str) == categoria
                ]

            colunas_exibicao = [
                colunas_estoque["centro"],
                colunas_estoque["material"],
                colunas_estoque["descricao"],
                colunas_estoque["categoria"],
                colunas_estoque["total"]
            ]

            colunas_exibicao = list(dict.fromkeys(colunas_exibicao))

            for _, row in df_exibir.iterrows():

    estoque_cor = "#22c55e"

    if row["Estoque"] <= 0:
        estoque_cor = "#ef4444"

    st.markdown(f"""
    <div style="
        background: linear-gradient(145deg,#111827,#0f172a);
        border:1px solid rgba(255,255,255,0.08);
        border-radius:18px;
        padding:18px;
        margin-bottom:14px;
        box-shadow:0 6px 18px rgba(0,0,0,0.25);
    ">

        <div style="
            display:flex;
            justify-content:space-between;
            align-items:center;
            margin-bottom:10px;
        ">

            <div style="
                font-size:15px;
                color:#94a3b8;
                font-weight:700;
            ">
                {row["Centro"]}
            </div>

            <div style="
                background:{estoque_cor};
                color:white;
                padding:6px 12px;
                border-radius:999px;
                font-size:14px;
                font-weight:800;
            ">
                {row["Estoque"]}
            </div>

        </div>

        <div style="
            font-size:17px;
            font-weight:800;
            color:white;
            line-height:1.3;
            margin-bottom:8px;
        ">
            {row["Produto"]}
        </div>

        <div style="
            color:#38bdf8;
            font-size:15px;
            font-weight:700;
        ">
            Código: {row["Código"]}
        </div>

        <div style="
            margin-top:10px;
            color:#cbd5e1;
            font-size:14px;
        ">
            Categoria: {row["Categoria"]}
        </div>

    </div>
    """, unsafe_allow_html=True)
        except Exception as erro:
            st.error("Erro ao ler a planilha de estoque.")
            st.exception(erro)

with aba_consultar:

    st.header("🔎 Consultar Último Estoque Salvo")

    conn = conectar()

    df_salvo = pd.read_sql_query("""
        SELECT
            data_importacao,
            arquivo_origem,
            aba_origem,
            centro,
            material,
            descricao,
            categoria,
            estoque_total
        FROM estoque_diario
    """, conn)

    conn.close()

    if df_salvo.empty:
        st.warning("Nenhum estoque importado ainda.")
        st.stop()

    ultima_data = df_salvo["data_importacao"].iloc[0]

    st.success(f"Última atualização: {ultima_data}")

    df_ultimo = df_salvo.copy()

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Produtos", len(df_ultimo))
    col2.metric("Estoque Total", f"{df_ultimo['estoque_total'].sum():,.2f}")
    col3.metric("Zerados", len(df_ultimo[df_ultimo["estoque_total"] <= 0]))
    col4.metric("Categorias", df_ultimo["categoria"].nunique())

    st.markdown('<div class="search-box">', unsafe_allow_html=True)

    busca_salva = st.text_input(
        "🔎 Buscar produto no estoque salvo",
        placeholder="Digite código, descrição ou categoria"
    )

    categorias_salvas = sorted(
        df_ultimo["categoria"]
        .dropna()
        .astype(str)
        .unique()
    )

    categoria_salva = st.selectbox(
        "Filtrar categoria",
        ["Todas"] + categorias_salvas
    )

    st.markdown('</div>', unsafe_allow_html=True)

    df_resultado_salvo = df_ultimo.copy()

    if busca_salva:
        df_resultado_salvo = df_resultado_salvo[
            df_resultado_salvo["material"].astype(str).str.contains(busca_salva, case=False, na=False)
            |
            df_resultado_salvo["descricao"].astype(str).str.contains(busca_salva, case=False, na=False)
            |
            df_resultado_salvo["categoria"].astype(str).str.contains(busca_salva, case=False, na=False)
        ]

    if categoria_salva != "Todas":
        df_resultado_salvo = df_resultado_salvo[
            df_resultado_salvo["categoria"].astype(str) == categoria_salva
        ]

    df_exibir = df_resultado_salvo[
        [
            "centro",
            "material",
            "descricao",
            "categoria",
            "estoque_total"
        ]
    ].rename(columns={
        "centro": "Centro",
        "material": "Código",
        "descricao": "Produto",
        "categoria": "Categoria",
        "estoque_total": "Estoque"
    })

    st.dataframe(
        df_exibir,
        use_container_width=True,
        hide_index=True
    )