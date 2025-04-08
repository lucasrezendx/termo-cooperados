from flask import Flask, render_template, request, send_file
from docx import Document
from datetime import datetime
from openpyxl import load_workbook
from io import BytesIO

app = Flask(__name__)

def formatar_cpf(cpf):
    cpf = ''.join(filter(str.isdigit, str(cpf)))
    return f'{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}' if len(cpf) == 11 else cpf

def formatar_rg(rg):
    rg = ''.join(filter(str.isdigit, str(rg)))
    return f'{rg[:2]}.{rg[2:5]}.{rg[5:8]}-{rg[8:]}' if len(rg) == 9 else rg

def formatar_cep(cep):
    cep = ''.join(filter(str.isdigit, str(cep)))
    return f'{cep[:5]}-{cep[5:]}' if len(cep) == 8 else cep

def substituir_texto_formatado(paragrafos, substituicoes):
    for paragrafo in paragrafos:
        for chave, valor in substituicoes.items():
            if chave in paragrafo.text:
                for run in paragrafo.runs:
                    if chave in run.text:
                        run.text = run.text.replace(chave, valor)

def carregar_dados(nome_busca):
    wb = load_workbook("cooperados.xlsx", data_only=True)
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
        tipo = request.form.get("tipo", "").upper()
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
                "CEP": formatar_cep(dados.get("CEP", "")),
                "CIDADE": dados.get("Cidade", ""),
                "DATA": data_atual,
                "HORA": hora_atual,
                "RGCOOPERADO": formatar_rg(request.form["rg"]),
                "APELIDODISPOSITIVO": request.form["apelido"],
                "MODELODISPOSITIVO": request.form["modelo"],
                "CHAVEMULTICANAL": request.form["chave"],
                "LOCAL": request.form["local"],
                "NOMECOLABORADOR": request.form["colaborador"],
                "CPFCOLABORADOR": formatar_cpf(request.form["cpf_colaborador"]),
            }

            doc = Document("modelo.docx")

        elif tipo == "PJ":
            nome_empresa = request.form["nome_empresa"]
            dados = carregar_dados(nome_empresa)

            if not dados:
                return "Empresa não encontrada."

            substituicoes = {
                "EMPRESA": nome_empresa.strip(),
                "PESSOAJURIDICA": formatar_cpf(dados.get("CPF/CNPJ", "")),
                "LUGAR": dados.get("Endereço", ""),
                "CITY": dados.get("Cidade", ""),
                "NOMECOOPERADO": request.form["nome_cooperado"],
                "ESTADOCIVIL": request.form["estado_civil"],
                "OCUPACAO": request.form["ocupacao"],
                "CPFCOOPERADO": formatar_cpf(request.form["cpf"]),
                "RGCOOPERADO": formatar_rg(request.form["rg"]),
                "ENDERECO": request.form["endereco"],
                "CEP": formatar_cep(request.form["cep"]),
                "CIDADE": request.form["cidade"],
                "APELIDODISPOSITIVO": request.form["apelido"],
                "MODELODISPOSITIVO": request.form["modelo"],
                "CHAVEMULTICANAL": request.form["chave"],
                "LOCAL": request.form["local"],
                "DATA": data_atual,
                "HORA": hora_atual,
                "NOMECOLABORADOR": request.form["colaborador"],
                "CPFCOLABORADOR": formatar_cpf(request.form["cpf_colaborador"]),
            }

            doc = Document("modelo_pj.docx")
        else:
            return "Tipo inválido."

        substituir_texto_formatado(doc.paragraphs, substituicoes)
        for tabela in doc.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    substituir_texto_formatado(celula.paragraphs, substituicoes)

        output = BytesIO()
        doc.save(output)
        output.seek(0)

        nome_arquivo = f"TERMO_{substituicoes['NOMECOOPERADO'].replace(' ', '_')}_{tipo}.docx"
        return send_file(output, as_attachment=True, download_name=nome_arquivo)

    return render_template("index.html")
