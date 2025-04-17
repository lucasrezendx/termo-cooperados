from flask import Flask, render_template, request, send_file
import os
from docx import Document
import pandas as pd
from datetime import datetime

app = Flask(__name__)

# Função para gerar o termo de cooperado
@app.route('/gerar', methods=['POST'])
def gerar():
    # Coleta dados do formulário
    nome = request.form['nome']
    tipo = request.form['tipo']
    apelido = request.form['apelido']
    modelo = request.form['modelo']
    chave = request.form['chave']
    colaborador = request.form['colaborador']
    cpf = request.form['cpf']
    local = request.form['local']
    nome_colaborador = request.form['nome_colaborador']
    cpf_colaborador = request.form['cpf_colaborador']

    # Carregar base de dados
    df = pd.read_excel('cooperados.xlsx')  # Certifique-se que o arquivo esteja na pasta correta
    cooperado = df[df['Nome'] == nome].iloc[0]

    # Abre o documento modelo dependendo do tipo de cooperado
    if tipo == 'PJ':
        doc = Document('modelo_pj.docx')  # Substitua com o caminho correto do modelo PJ
    else:
        doc = Document('modelo.docx')  # Substitua com o caminho correto do modelo PF/AGRO

    # Substitui campos no documento conforme os dados do formulário
    for para in doc.paragraphs:
        if 'NOME' in para.text:
            for run in para.runs:
                run.text = run.text.replace('NOME', nome)
        if 'APELIDO' in para.text:
            for run in para.runs:
                run.text = run.text.replace('APELIDO', apelido)
        if 'MODELO' in para.text:
            for run in para.runs:
                run.text = run.text.replace('MODELO', modelo)
        if 'CHAVE' in para.text:
            for run in para.runs:
                run.text = run.text.replace('CHAVE', chave)
        if 'COLABORADOR' in para.text:
            for run in para.runs:
                run.text = run.text.replace('COLABORADOR', colaborador)
        if 'CPF' in para.text:
            for run in para.runs:
                run.text = run.text.replace('CPF', cpf)

    # Salva o documento gerado
    termo_file = f"termo_{nome.replace(' ', '_')}.docx"
    doc.save(termo_file)

    # Envia o arquivo para download
    return send_file(termo_file, as_attachment=True)

# Rota principal
@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
