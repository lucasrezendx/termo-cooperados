from flask import Flask, render_template, request, send_file
from docx import Document
from datetime import datetime
from openpyxl import load_workbook
from io import BytesIO
import os
import pytz

app = Flask(__name__)

# Formata CPF ou CNPJ automaticamente
def formatar_cpf_cnpj(valor):
    valor = ''.join(filter(str.isdigit, str(valor)))
    if len(valor) == 11:
        return f"{valor[:3]}.{valor[3:6]}.{valor[6:9]}-{valor[9:]}"
    elif len(valor) == 14:
        return f"{valor[:2]}.{valor[2:5]}.{valor[5:8]}/{valor[8:12]}-{valor[12:]}"
    return valor

def formatar_rg(rg):
    rg = ''.join(filter(str.isdigit, str(rg)))
    return f'{rg[:2]}.{rg[2:5]}.{rg[5:8]}-{rg[8:]}' if len(rg) == 9 else rg

def formatar_cep(cep):
    cep = ''.join(filter(str.isdigit, str(cep)))
    return f'{cep[:5]}-{cep[5:]}' if len(cep) == 8 else cep

# NOVA FUNÇÃO ROBUSTA para substituir textos mesmo divididos em runs
def substituir_texto_formatado(paragrafos, substituicoes):
    for paragrafo in paragrafos:
        texto_completo = "".join(run.text for run in paragrafo.runs)
        novo_texto = texto_completo
        for chave, valor in substituicoes.items():
            if chave in novo_texto:
                novo_texto = novo_texto.replace(chave, valor)
        if novo_texto != texto_completo:
            for run in paragrafo.runs:
                run.text = ""
            if paragrafo.runs:
                paragrafo.runs[0].text = novo_texto

def carregar_dados(nome_busca):
    path = os.path.join(os.path.dirname(__file__), "TESTAR.xlsx")
    wb = load_workbook(path, data_only=True)
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
        now = datetime.now(pytz.timezone("America/Sao_Paulo"))
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
                "CPFCOOPERADO": formatar_cpf_cnpj(dados.get("CPF/CNPJ", "")),
                "ENDERECO": dados.get("Endereço", ""),
                "CEP": formatar_cep(dados.get("CEP", "")),
                "CIDADE": dados.get("Cidade", ""),
                "DATA": data_atual,
                "HORA": hora_atual,
                "RGCOOPERADO": formatar_rg(request.form["rg"]),
                "APELIDODISPOSITIVO": request.form["apelido"],
                "MODELODISPOSITIVO": request.form["modelo"],
                "LOCAL": request.form["local"],
                "NOMECOLABORADOR": request.form["colaborador"],
                "CPFCOLABORADOR": formatar_cpf_cnpj(request.form["cpf_colaborador"]),
            }
            doc = Document("modelo.docx")

        elif tipo == "PJ":
            empresa = request.form["nome_empresa"]
            dados_empresa = carregar_dados(empresa)
            if not dados_empresa:
                return "Empresa não encontrada."

            substituicoes = {
                "EMPRESA": empresa,
                "NOMECOOPERADO": request.form["nome_cooperado"],
                "ESTADOCIVIL": request.form["estado_civil"],
                "OCUPACAO": request.form["ocupacao"],
                "CPFCOOPERADO": formatar_cpf_cnpj(request.form["cpf"]),
                "RGCOOPERADO": formatar_rg(request.form["rg_cooperado"]),
                "APELIDODISPOSITIVO": request.form["apelido"],
                "MODELODISPOSITIVO": request.form["modelo"],
                "CHAVEMULTICANAL": request.form["chave"],
                "LOCAL": request.form["local"],
                "NOMECOLABORADOR": request.form["colaborador"],
                "CPFCOLABORADOR": formatar_cpf_cnpj(request.form["cpf_colaborador"]),
                "LUGAR": dados_empresa.get("Endereço", ""),
                "CITY": dados_empresa.get("Cidade", ""),
                "ENDERECO": dados_empresa.get("Endereço", ""),
                "CEP": formatar_cep(dados_empresa.get("CEP", "")),
                "CIDADE": dados_empresa.get("Cidade", ""),
                "PESSOAJURIDICA": formatar_cpf_cnpj(dados_empresa.get("CPF/CNPJ", "")),
                "DATA": data_atual,
                "HORA": hora_atual,
            }
            doc = Document("modelo_pj.docx")

        else:
            return "Tipo inválido. Use PF, AGRO ou PJ."

        # Substituir em parágrafos e em células de tabelas
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
