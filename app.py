
from flask import Flask, render_template, request, send_file, redirect, url_for
import pandas as pd
from docx import Document
from datetime import datetime
import os
from io import BytesIO

app = Flask(__name__)

def formatar_cpf(cpf):
    cpf = ''.join(filter(str.isdigit, str(cpf)))
    return f'{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}' if len(cpf) == 11 else cpf

def substituir_texto_formatado(paragrafos, substituicoes):
    for paragrafo in paragrafos:
        for chave, valor in substituicoes.items():
            if chave in paragrafo.text:
                for run in paragrafo.runs:
                    if chave in run.text:
                        run.text = run.text.replace(chave, valor)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        tipo = request.form["tipo"]
        df = pd.read_excel("cooperados.xlsx", dtype=str).fillna("")

        now = datetime.now()
        data_atual = now.strftime("%d/%m/%Y")
        hora_atual = now.strftime("%H:%M")

        if tipo in ["PF", "AGRO"]:
            nome = request.form["nome"]
            linha = df[df["Nome"].str.strip().str.lower() == nome.strip().lower()]

            if linha.empty:
                return "Cooperado não encontrado."

            dados = linha.iloc[0]

            substituicoes = {
                "NOMECOOPERADO": dados["Nome"],
                "ESTADOCIVIL": dados["Estado Civil"],
                "OCUPACAO": dados["Ocupação"],
                "CPFCOOPERADO": formatar_cpf(dados["CPF/CNPJ"]),
                "ENDERECO": dados["Endereço"],
                "CEP": dados["CEP"],
                "CIDADE": dados["Cidade"],
                "DATA": data_atual,
                "HORA": hora_atual,
                "RGCOOPERADO": request.form["rg"],
                "APELIDODISPOSITIVO": request.form["apelido"],
                "MODELODISPOSITIVO": request.form["modelo"],
                "CHAVEMULTICANAL": request.form["chave"],
                "NOMECOLABORADOR": request.form["colaborador"],
                "CPFCOLABORADOR": formatar_cpf(request.form["cpf_colaborador"]),
            }

            doc = Document("modelo.docx")

        else:  # PJ
            nome_empresa = request.form["nome_empresa"]
            linha = df[df["Nome"].str.strip().str.lower() == nome_empresa.strip().lower()]

            if linha.empty:
                return "Empresa não encontrada."

            dados = linha.iloc[0]

            substituicoes = {
                "NOMEDAEMPRESA": nome_empresa,
                "PESSOAJURIDICA": formatar_cpf(dados["CPF/CNPJ"]),
                "LUGAR": dados["Endereço"],
                "CITY": dados["Cidade"],
                "NOMECOOPERADO": request.form["nome_cooperado"],
                "ESTADOCIVIL": request.form["estado_civil"],
                "OCUPACAO": request.form["ocupacao"],
                "CPFCOOPERADO": formatar_cpf(request.form["cpf"]),
                "RGCOOPERADO": request.form["rg"],
                "ENDERECO": request.form["endereco"],
                "CEP": request.form["cep"],
                "CIDADE": request.form["cidade"],
                "APELIDODISPOSITIVO": request.form["apelido"],
                "MODELODISPOSITIVO": request.form["modelo"],
                "CHAVEMULTICANAL": request.form["chave"],
                "DATA": data_atual,
                "HORA": hora_atual,
                "LOCAL": request.form["local"],
                "NOMECOLABORADOR": request.form["colaborador"],
                "CPFCOLABORADOR": formatar_cpf(request.form["cpf_colaborador"]),
            }

            doc = Document("modelo_pj.docx")

        substituir_texto_formatado(doc.paragraphs, substituicoes)
        for tabela in doc.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    substituir_texto_formatado(celula.paragraphs, substituicoes)

        output = BytesIO()
        doc.save(output)
        output.seek(0)

        return send_file(output, as_attachment=True, download_name="documento_preenchido.docx")

    return render_template("index.html")
