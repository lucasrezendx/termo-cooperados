from flask import Flask, render_template, request, send_file
from docx import Document
from datetime import datetime
from openpyxl import load_workbook
from io import BytesIO
import os

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

def carregar_dados(nome_busca):
    excel_path = os.path.join(os.path.dirname(__file__), "cooperados.xlsx")
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active

    colunas = [cell.value for cell in ws[1]]
    for row in ws.iter_rows(min_row=2, values_only=True):
        dados = dict(zip(colunas, row))
        if dados["Nome"] and dados["Nome"].strip().lower() == nome_busca.strip().lower():
            return dados
    return None

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        tipo = request.form["tipo"]

        now = datetime.now()
        data_atual = now.strftime("%d/%m/%Y")
        hora_atual = now.strftime("%H:%M")

        if tipo in ["PF", "AGRO"]:
            nome = request.form["nome"]
            dados = carregar_dados(nome)

            if not dados:
                return "Cooperado não encontrado."

            substituicoes = {
                "NOMECOOPERADO": dados.get("Nome", ""),
                "ESTADOCIVIL": dados.get("Estado Civil", ""),
                "OCUPACAO": dados.get("Ocupação", ""),
                "CPFCOOPERADO": formatar_cpf(dados.get("CPF/CNPJ", "")),
                "ENDERECO": dados.get("Endereço", ""),
                "CEP": dados.get("CEP", ""),
                "CIDADE": dados.get("Cidade", ""),
                "DATA": data_atual,
                "HORA": hora_atual,
                "RGCOOPERADO": request.form["rg"],
                "APELIDODISPOSITIVO": request.form["apelido"],
                "MODELODISPOSITIVO": request.form["modelo"],
                "CHAVEMULTICANAL": request.form["chave"],
                "NOMECOLABORADOR": request.form["colaborador"],
                "CPFCOLABORADOR": formatar_cpf(request.form["cpf_colaborador"]),
            }

            doc_path = os.path.join(os.path.dirname(__file__), "modelo.docx")
            doc = Document(doc_path)

        else:  # PJ
            nome_empresa = request.form["nome_empresa"]
            dados = carregar_dados(nome_empresa)

            if not dados:
                return "Empresa não encontrada."

            substituicoes = {
                "NOMEDAEMPRESA": nome_empresa,
                "PESSOAJURIDICA": formatar_cpf(dados.get("CPF/CNPJ", "")),
                "LUGAR": dados.get("Endereço", ""),
                "CITY": dados.get("Cidade", ""),
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

            doc_path = os.path.join(os.path.dirname(__file__), "modelo_pj.docx")
            doc = Document(doc_path)

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
