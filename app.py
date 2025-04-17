from flask import Flask, render_template, request, send_file
from openpyxl import load_workbook
from docx import Document
import os
import io
from datetime import datetime

app = Flask(__name__)

# Carregar os dados da planilha apenas uma vez
DADOS_COOPERADOS = []

def carregar_dados_uma_vez():
    global DADOS_COOPERADOS
    path = os.path.join(os.path.dirname(__file__), "TESTAR.xlsx")
    wb = load_workbook(path, data_only=True)
    ws = wb.active

    colunas = [cell.value for cell in ws[1]]

    for row in ws.iter_rows(min_row=2, values_only=True):
        dados = dict(zip(colunas, row))
        nome = row[5]  # Coluna F (índice 5)
        if nome:  # Só adiciona se houver nome na coluna F
            dados["Nome"] = nome
            DADOS_COOPERADOS.append(dados)

def buscar_cooperado_por_nome(nome_busca):
    for dados in DADOS_COOPERADOS:
        if dados.get("Nome") and dados["Nome"].strip().lower() == nome_busca.strip().lower():
            return dados
    return None

def substituir_campos(doc, dados):
    for p in doc.paragraphs:
        for chave, valor in dados.items():
            if chave and isinstance(valor, str):
                p.text = p.text.replace(f"<<{chave}>>", valor)

    for tabela in doc.tables:
        for row in tabela.rows:
            for cell in row.cells:
                for chave, valor in dados.items():
                    if chave and isinstance(valor, str):
                        cell.text = cell.text.replace(f"<<{chave}>>", valor)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        nome = request.form["nome"]
        tipo = request.form["tipo"]
        rg_manual = request.form["rg"]
        hora_atual = datetime.now().strftime("%H:%M")

        dados = buscar_cooperado_por_nome(nome)

        if not dados:
            return "Cooperado não encontrado."

        # Adiciona informações extras manuais
        dados["RG"] = rg_manual
        dados["HORA"] = hora_atual

        # Tratamento para documento PJ
        if tipo == "PJ":
            modelo_path = os.path.join("modelo_pj.docx")
            doc = Document(modelo_path)
            dados["EMPRESA"] = dados.get("EMPRESA", "")
            dados["PESSOAJURIDICA"] = dados.get("PESSOAJURIDICA", "")  # CNPJ
            dados["CHAVE"] = dados.get("CHAVE", "")
        else:
            modelo_path = os.path.join("modelo.docx")
            doc = Document(modelo_path)
            # Se não for PJ, removemos os campos que não existem
            dados.pop("EMPRESA", None)
            dados.pop("PESSOAJURIDICA", None)
            dados.pop("CHAVE", None)

        substituir_campos(doc, dados)

        output = io.BytesIO()
        doc.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name=f"{nome}_{tipo}.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    return render_template("index.html")

if __name__ == "__main__":
    carregar_dados_uma_vez()
    app.run(debug=True)
