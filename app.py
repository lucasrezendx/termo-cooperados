from flask import Flask, render_template, request, send_file
import pandas as pd
from docx import Document
from datetime import datetime
import pytz
import io
import os

app = Flask(__name__)

# Carregar a planilha uma vez ao iniciar
df = pd.read_excel("cooperados.xlsx")

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        nome = request.form["nome"]
        modelo = request.form["modelo"]
        rg = request.form["rg"]
        hora_manual = request.form.get("hora")
        empresa = request.form.get("empresa")
        chave = request.form.get("chave")

        # Verificar se o nome existe na base
        linha = df[df['Nome'].astype(str).str.strip().str.lower() == nome.lower()]
        if linha.empty:
            return "Nome não encontrado na base de dados.", 400

        dados = linha.iloc[0].to_dict()

        # Obter data e hora no fuso de Brasília
        fuso_brasilia = pytz.timezone("America/Sao_Paulo")
        hora_atual = datetime.now(fuso_brasilia).strftime("%H:%M")
        data_atual = datetime.now(fuso_brasilia).strftime("%d/%m/%Y")

        # Se o usuário preencher a hora manualmente, usamos ela
        hora_final = hora_manual if hora_manual else hora_atual

        if modelo == "PJ":
            doc = Document("modelo_pj.docx")
            substituicoes = {
                "NOME": dados.get("Nome", ""),
                "CPF": str(dados.get("CPF", "")),
                "RG": rg,
                "EMPRESA": empresa,
                "PESSOAJURIDICA": chave,
                "ENDERECO": dados.get("Endereço", ""),
                "DATA": data_atual,
                "HORA": hora_final
            }
        else:  # PF ou AGRO
            doc = Document("modelo.docx")
            substituicoes = {
                "NOME": dados.get("Nome", ""),
                "CPF": str(dados.get("CPF", "")),
                "RG": rg,
                "ENDERECO": dados.get("Endereço", ""),
                "DATA": data_atual,
                "HORA": hora_final
            }

        # Substituir texto nos parágrafos
        for paragrafo in doc.paragraphs:
            for chave, valor in substituicoes.items():
                if chave in paragrafo.text:
                    paragrafo.text = paragrafo.text.replace(chave, valor)

        # Substituir em tabelas
        for tabela in doc.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    for chave, valor in substituicoes.items():
                        if chave in celula.text:
                            celula.text = celula.text.replace(chave, valor)

        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        return send_file(
            buffer,
            as_attachment=True,
            download_name=f"termo_{modelo.lower()}_{nome.replace(' ', '_')}.docx"
        )

    return render_template("index.html")
