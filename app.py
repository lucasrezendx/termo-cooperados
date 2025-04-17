from flask import Flask, render_template, request, send_file
import pandas as pd
from docx import Document
from datetime import datetime
import os

app = Flask(__name__)

# Carrega a planilha de cooperados
df = pd.read_excel("cooperados.xlsx")

def carregar_dados(nome):
    """Busca os dados do cooperado pelo nome."""
    linha = df[df["Nome"].str.lower() == nome.lower()]
    if not linha.empty:
        return linha.iloc[0].to_dict()
    return None

def formatar_cpf(cpf):
    """Formata CPF ou CNPJ."""
    cpf = ''.join(filter(str.isdigit, str(cpf)))
    if len(cpf) == 11:
        return f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
    elif len(cpf) == 14:
        return f"{cpf[:2]}.{cpf[2:5]}.{cpf[5:8]}/{cpf[8:12]}-{cpf[12:]}"
    return cpf

def formatar_rg(rg):
    """Formata RG."""
    rg = ''.join(filter(str.isdigit, str(rg)))
    if len(rg) >= 9:
        return f"{rg[:2]}.{rg[2:5]}.{rg[5:8]}-{rg[8:]}"
    return rg

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        tipo = request.form["tipo"]
        data_atual = datetime.now().strftime("%d/%m/%Y")
        hora_atual = datetime.now().strftime("%H:%M:%S")

        if tipo in ["PF", "AGRO"]:
            nome = request.form["nome"]
            dados = carregar_dados(nome)
            if not dados:
                return "Cooperado não encontrado."

            substituicoes = {
                "NOMECOOPERADO": nome,
                "RGCOOPERADO": formatar_rg(request.form["rg"]),
                "APELIDODISPOSITIVO": request.form["apelido"],
                "MODELODISPOSITIVO": request.form["modelo"],
                "LOCAL": request.form["local"],
                "NOMECOLABORADOR": request.form["colaborador"],
                "CPFCOLABORADOR": formatar_cpf(request.form["cpf_colaborador"]),
                "LUGAR": dados.get("Endereço", ""),
                "CITY": dados.get("Cidade", ""),
                "PESSOAFISICA": formatar_cpf(dados.get("CPF/CNPJ", "")),
                "DATA": data_atual,
                "HORA": hora_atual,
            }

            doc = Document("modelo.docx")

        elif tipo == "PJ":
            empresa = request.form["empresa"]
            dados_empresa = carregar_dados(empresa)
            if not dados_empresa:
                return "Empresa não encontrada."

            substituicoes = {
                "NOMECOOPERADO": request.form["nome"],
                "ESTADOCIVIL": request.form["estado_civil"],
                "OCUPACAO": request.form["ocupacao"],
                "CPFCOOPERADO": formatar_cpf(request.form["cpf"]),
                "RGCOOPERADO": formatar_rg(request.form["rg"]),
                "EMPRESA": empresa,
                "APELIDODISPOSITIVO": request.form["apelido"],
                "MODELODISPOSITIVO": request.form["modelo"],
                "CHAVEMULTICANAL": request.form["chave"],
                "LOCAL": request.form["local"],
                "NOMECOLABORADOR": request.form["colaborador"],
                "CPFCOLABORADOR": formatar_cpf(request.form["cpf_colaborador"]),
                "LUGAR": dados_empresa.get("Endereço", ""),
                "CITY": dados_empresa.get("Cidade", ""),
                "PESSOAJURIDICA": formatar_cpf(dados_empresa.get("CPF/CNPJ", "")),
                "DATA": data_atual,
                "HORA": hora_atual,
            }

            doc = Document("modelo_pj.docx")
        else:
            return "Tipo de cooperado inválido."

        # Substitui os campos no documento
        for par in doc.paragraphs:
            for chave, valor in substituicoes.items():
                if chave in par.text:
                    par.text = par.text.replace(chave, valor)

        # Substituir também nas tabelas, se houver
        for tabela in doc.tables:
            for linha in tabela.rows:
                for celula in linha.cells:
                    for chave, valor in substituicoes.items():
                        if chave in celula.text:
                            celula.text = celula.text.replace(chave, valor)

        nome_arquivo = f"termo_{tipo.lower()}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        caminho = os.path.join("downloads", nome_arquivo)
        os.makedirs("downloads", exist_ok=True)
        doc.save(caminho)

        return send_file(caminho, as_attachment=True)

    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)
