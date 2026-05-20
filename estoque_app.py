import streamlit as st
import pandas as pd
import sqlite3
import os
import json
from datetime import datetime

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

    data_importacao = datetime.now().strftime("%d/%m/%Y %H:%M")

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

# =========================
# LOGIN SIMPLES
# =========================

USUARIOS = {
    "euler": "123456",
    "estoque": "123456",
    "vendas": "123456"
}

if "logado_estoque" not in st.session_state:
    st.session_state.logado_estoque = False

if not st.session_state.logado_estoque:
    st.title("🔐 Login Estoque EA CRM")

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

st.sidebar.success(f"Usuário: {st.session_state.usuario_estoque}")

if st.sidebar.button("Sair"):
    st.session_state.logado_estoque = False
    st.rerun()

# =========================
# APP ESTOQUE
# =========================

st.title("📦 Estoque Diário")
st.caption("Consulta e importação rápida de estoque para celular")

aba_importar, aba_consultar = st.tabs([
    "Importar Estoque",
    "Consultar Estoque"
])

# =========================
# IMPORTAR ESTOQUE
# =========================

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

            col1, col2, col3, col4 = st.columns(4)

            with col1:
                st.metric("Produtos na base", len(df_estoque))

            with col2:
                st.metric("Estoque Total", f'{df_estoque["_estoque_total"].sum():,.2f}')

            with col3:
                st.metric("Itens zerados", len(df_estoque[df_estoque["_estoque_total"] <= 0]))

            with col4:
                st.metric("Aba lida", aba_selecionada)

            st.markdown("---")

            busca = st.text_input(
                "🔎 Pesquisar na planilha carregada",
                placeholder="Digite código, descrição ou categoria"
            )

            categorias = sorted(
                df_estoque["_categoria"]
                .dropna()
                .astype(str)
                .unique()
            )

            categoria = st.selectbox(
                "Categoria",
                ["Todas"] + categorias
            )

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

            st.dataframe(
                df_resultado[colunas_exibicao],
                use_container_width=True,
                hide_index=True
            )

            if st.button("💾 Salvar estoque diário no banco"):
                data_importacao = salvar_estoque_no_banco(
                    df_estoque,
                    arquivo_estoque.name,
                    aba_selecionada
                )

                st.success(f"Estoque salvo com sucesso em {data_importacao}.")

        except Exception as erro:
            st.error("Erro ao ler a planilha de estoque.")
            st.exception(erro)

# =========================
# CONSULTAR ESTOQUE SALVO
# =========================

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
        ORDER BY id DESC
    """, conn)

    conn.close()

    if df_salvo.empty:
        st.warning("Nenhum estoque importado ainda.")
        st.stop()

    ultima_data = df_salvo["data_importacao"].iloc[0]

    df_ultimo = df_salvo[
        df_salvo["data_importacao"] == ultima_data
    ].copy()

    st.success(f"Última atualização: {ultima_data}")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.metric("Produtos", len(df_ultimo))

    with col2:
        st.metric("Estoque Total", f"{df_ultimo['estoque_total'].sum():,.2f}")

    with col3:
        st.metric("Zerados", len(df_ultimo[df_ultimo["estoque_total"] <= 0]))

    st.markdown("---")

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

    st.dataframe(
        df_resultado_salvo[
            [
                "centro",
                "material",
                "descricao",
                "categoria",
                "estoque_total"
            ]
        ],
        use_container_width=True,
        hide_index=True
    )