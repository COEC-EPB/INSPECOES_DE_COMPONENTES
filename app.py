from flask import Flask, request, send_file
from flask_cors import CORS
import pandas as pd

app = Flask(__name__)
CORS(app)

@app.route('/')
def home():
    return "API online 🚀"

@app.route('/processar', methods=['POST'])
def processar():

    arquivos = request.files.getlist("meses")
    ipeo_file = request.files.get("ipeo")

    if not arquivos or not ipeo_file:
        return {"erro": "Envie arquivos de meses e IPEO"}, 400

    # 🔹 CONCATENAR MESES
    lista_df = []
    for f in arquivos:
        df = pd.read_excel(f)
        lista_df.append(df)

    df_meses = pd.concat(lista_df, ignore_index=True)

    # 🔹 IPEO
    df_ipeo = pd.read_excel(ipeo_file)

    # 🔹 JOIN
    df = df_meses.merge(
        df_ipeo,
        on=["MATRICULA", "MES"],
        how="inner"
    )

    # 🔹 COLUNAS
    colunas = [
        "%DI", "%ROE", "%RNT", "%IOC", "%ISF", "%ROV",
        "EMPRESA", "MES", "REGIONAL", "POLO PRESTADOR",
        "MATRICULA", "NOME FUNCIONARIO", "%IPEO",
        "% Produtividade", "%Eficiencia", "% Utilização", "TMS"
    ]

    df = df[colunas]

    # 🔹 AGRUPAMENTO
    df_final = df.groupby([
        "EMPRESA", "MES", "REGIONAL", "POLO PRESTADOR",
        "MATRICULA", "NOME FUNCIONARIO"
    ], as_index=False).mean()

    # 🔹 EXPORTAR
    output = "resultado.xlsx"
    df_final.to_excel(output, index=False)

    return send_file(output, as_attachment=True)
