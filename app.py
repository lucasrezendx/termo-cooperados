from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from docx import Document
from datetime import datetime

app = Flask(__name__)

# Função para ler os dados do Excel
def buscar_dados(nome):
    df = pd.read_excel("cooperados.xlsx")
    dados = df[df["Nome"] == nome].iloc[0]
    return dados

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/gerar', methods=['POST'])
def gerar_documento():
    # Dados do formulário
    nome = request.form['nome']
    apelido_dispositivo = request.form['apelido']
    modelo_dispositivo = request.form['modelo']
    chave_multicanal = request.form['chave']
    nome_colaborador = request.form['colaborador']
    cpf_colaborador = request.form['cpf']
    tipo_pessoa = request.form['tipo_pessoa']

    # Busca os dados do cooperado
    dados_cooperado = buscar_dados(nome)

    # Escolhe o modelo do documento conforme o tipo de pessoa
    if tipo_pessoa == 'PF' or tipo_pessoa == 'AGRO':
        doc = Document("modelo.docx")
    elif tipo_pessoa == 'PJ':
        doc = Document("modelo_pj.docx")

    # Preenche os campos do documento
    for p in doc.paragraphs:
        if 'NOME_COOPERADO' in p.text:
            p.text = p.text.replace('NOME_COOPERADO', dados_cooperado["Nome"])
        if 'APELIDO_DISPOSITIVO' in p.text:
            p.text = p.text.replace('APELIDO_DISPOSITIVO', apelido_dispositivo)
        if 'MODELO_DISPOSITIVO' in p.text:
            p.text = p.text.replace('MODELO_DISPOSITIVO', modelo_dispositivo)
        if 'CHAVE_MULTICANAL' in p.text:
            p.text = p.text.replace('CHAVE_MULTICANAL', chave_multicanal)
        if 'NOME_COLABORADOR' in p.text:
            p.text = p.text.replace('NOME_COLABORADOR', nome_colaborador)
        if 'CPF_COLABORADOR' in p.text:
            p.text = p.text.replace('CPF_COLABORADOR', cpf_colaborador)

    # Preenche a data e hora
    data_atual = datetime.now().strftime("%d/%m/%Y %H:%M")
    for p in doc.paragraphs:
        if 'DATA_HORA' in p.text:
            p.text = p.text.replace('DATA_HORA', data_atual)

    # Salva o documento preenchido
    arquivo_gerado = "documento_preenchido.docx"
    doc.save(arquivo_gerado)

    # Verifica se o arquivo foi criado corretamente
    if not os.path.exists(arquivo_gerado):
        return "Erro: o documento não foi criado.", 500

    return send_file(
        arquivo_gerado,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name="documento_preenchido.docx"
    )

if __name__ == '__main__':
    app.run(debug=True)
