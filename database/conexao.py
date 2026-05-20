import sqlite3
import os

from dotenv import load_dotenv

load_dotenv()

DB_NAME = "data/crm_pedidos.db"

def conectar():
    return sqlite3.connect(DB_NAME)

def criar_tabelas():
    conn = conectar()
    cursor = conn.cursor()

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS clientes (
            codigo_cliente TEXT PRIMARY KEY,
            nome TEXT,
            email TEXT,
            telefone TEXT,
            cnpj TEXT,
            endereco TEXT,
            cidade TEXT,
            estado TEXT,
            metodo_pagamento TEXT
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS pedidos (
            numero_pedido TEXT PRIMARY KEY,
            codigo_cliente TEXT,
            valor_total REAL,
            origem TEXT,
            FOREIGN KEY (codigo_cliente)
            REFERENCES clientes(codigo_cliente)
        )
    """)

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS itens_pedido (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_pedido TEXT,
            codigo_produto TEXT,
            produto TEXT,
            quantidade REAL,
            preco REAL,
            total_sem_st REAL,
            st REAL,
            total_com_st REAL,
            FOREIGN KEY (numero_pedido)
            REFERENCES pedidos(numero_pedido)
        )
    """)

    conn.commit()
    conn.close()


def testar_conexao():
    try:
        conn = conectar()
        print("Banco conectado com sucesso.")
        conn.close()
    except Exception as e:
        print("Erro:", e)