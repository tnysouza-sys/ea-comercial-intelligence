import pdfplumber
import re

def limpar_valor(valor):
    if not valor:
        return 0.0
    valor = valor.replace("R$", "").replace(".", "").replace(",", ".").strip()
    try:
        return float(valor)
    except:
        return 0.0


def extrair_texto_pdf(arquivo_pdf):
    texto = ""
    with pdfplumber.open(arquivo_pdf) as pdf:
        for pagina in pdf.pages:
            texto += pagina.extract_text() + "\n"
    return texto


def extrair_dados_pedido(arquivo_pdf):
    texto = extrair_texto_pdf(arquivo_pdf)

    pedido = re.search(r"Pedido SAP\s+(\d+)", texto)
    codigo_cliente = re.search(r"Codigo do Cliente\s+(\d+)", texto)
    nome = re.search(r"Nome\s+(.+)", texto)
    email = re.search(r"E-mail\s+(.+)", texto)
    cnpj = re.search(r"CNPJ\s+(\d+)", texto)
    endereco = re.search(r"Endereço de Entrega\s+(.+)", texto)
    metodo_pagamento = re.search(r"Metodo de Pagamento\s+(.+)", texto)

    total = re.search(r"Total Valor c/ ST\s+R\$ [\d.,]+\s+R\$ [\d.,]+\s+R\$ ([\d.,]+)", texto)

    dados = {
        "numero_pedido": pedido.group(1) if pedido else "",
        "codigo_cliente": codigo_cliente.group(1) if codigo_cliente else "",
        "nome": nome.group(1).strip() if nome else "",
        "email": email.group(1).strip() if email else "",
        "cnpj": cnpj.group(1).strip() if cnpj else "",
        "endereco": endereco.group(1).strip() if endereco else "",
        "metodo_pagamento": metodo_pagamento.group(1).strip() if metodo_pagamento else "",
        "valor_total": limpar_valor(total.group(1)) if total else 0.0,
        "texto_completo": texto
    }

    return dados