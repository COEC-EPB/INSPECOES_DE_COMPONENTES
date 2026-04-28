from flask import Flask, request, send_file
from flask_cors import CORS
import pandas as pd
import os

app = Flask(__name__)
CORS(app)

@app.route('/')
def home():
    return "API online 🚀"

@app.route('/processar', methods=['POST'])
def processar():
    if 'file' not in request.files:
        return {"erro": "Nenhum arquivo enviado"}, 400

    file = request.files['file']

    df = pd.read_csv(file)

    # EXEMPLO DE PROCESSAMENTO
    df['resultado'] = df.iloc[:, 0] * 2

    output = "resultado.xlsx"
    df.to_excel(output, index=False)

    return send_file(output, as_attachment=True)

if __name__ == '__main__':
    app.run()
