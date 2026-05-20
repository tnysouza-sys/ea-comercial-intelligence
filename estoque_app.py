import streamlit as st
import pandas as pd
import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_NAME = os.path.join(BASE_DIR, "crm_pedidos.db")

def conectar():
    return sqlite3.connect(DB_NAME)

st.set_page_config(
    page_title="Estoque EA CRM",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.title("📦 Estoque Diário")
st.caption("Consulta rápida para celular")

conn = conectar()

df_estoque = pd.read_sql_query("""
    SELECT
        data_importacao,
        centro,
        material,
        descricao,
        categoria,
        estoque_total
    FROM estoque_diario
    ORDER BY descricao
""", conn)

conn.close()

if df_estoque.empty:
    st.warning("Nenhum estoque carregado ainda.")
    st.stop()

ultima_carga = df_estoque["data_importacao"].iloc[0]

st.success(f"Última atualização: {ultima_carga}")

busca = st.text_input("🔎 Buscar produto", placeholder="Digite código, descrição ou categoria")

df_filtrado = df_estoque.copy()

if busca:
    df_filtrado = df_filtrado[
        df_filtrado["material"].astype(str).str.contains(busca, case=False, na=False)
        |
        df_filtrado["descricao"].astype(str).str.contains(busca, case=False, na=False)
        |
        df_filtrado["categoria"].astype(str).str.contains(busca, case=False, na=False)
    ]

col1, col2, col3 = st.columns(3)

with col1:
    st.metric("Produtos", len(df_filtrado))

with col2:
    st.metric("Estoque Total", f"{df_filtrado['estoque_total'].sum():,.0f}")

with col3:
    st.metric("Zerados", len(df_filtrado[df_filtrado["estoque_total"] <= 0]))

st.markdown("---")

st.dataframe(
    df_filtrado[
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