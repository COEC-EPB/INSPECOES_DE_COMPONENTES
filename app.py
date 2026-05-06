from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import unicodedata
import io
import os

app = Flask(__name__)
CORS(app)


@app.route('/')
def home():
    return "API online 🚀"


# 🔹 NORMALIZAR COLUNAS
def normalizar(df):
    df.columns = [
        unicodedata.normalize('NFKD', str(c))
        .encode('ASCII', 'ignore')
        .decode('ASCII')
        .strip()
        .upper()
        for c in df.columns
    ]
    return df


# 🔹 LER EXCEL
def ler_excel(file):
    df_raw = pd.read_excel(file, header=None)

    for i, row in df_raw.iterrows():
        if row.astype(str).str.contains("FUNCIONARIO", case=False).any():
            df = pd.read_excel(file, header=i)
            return normalizar(df)

    # fallback
    df = pd.read_excel(file)
    return normalizar(df)


# 🔹 EXTRAIR MES
def extrair_mes(nome):
    nome = nome.upper()
    meses = ["JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"]
    for m in meses:
        if m in nome:
            return m
    return None


# 🔹 SEPARAR FUNCIONARIO
def separar(df):
    col = next((c for c in df.columns if "FUNCIONARIO" in c), None)

    if not col:
        raise Exception("Coluna FUNCIONARIO não encontrada")

    split = df[col].astype(str).str.split(" - ", n=1, expand=True)

    df["MATRICULA"] = split[0].str.strip()
    df["NOME"] = split[1].str.strip()

    return df


# 🔹 LIMPAR NUMÉRICOS
def limpar(df):
    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = (
                df[c].astype(str)
                .str.replace("%","")
                .str.replace(",",".")
                .str.strip()
            )
            df[c] = pd.to_numeric(df[c], errors="ignore")
    return df


@app.route('/processar', methods=['POST'])
def processar():
    try:
        arquivos = request.files.getlist("meses")
        ipeo_file = request.files.get("ipeo")

        if not arquivos or not ipeo_file:
            return jsonify({"erro": "Envie todos os arquivos"}), 400

        lista = []

        # 🔥 PROCESSAR MESES
        for f in arquivos:
            df = ler_excel(f)

            mes = extrair_mes(f.filename)
            if not mes:
                return jsonify({"erro": f"Mês não identificado no arquivo {f.filename}"}), 400

            df = separar(df)
            df["MES"] = mes

            lista.append(df)

        df_meses = pd.concat(lista, ignore_index=True)

        # 🔥 VALIDAR MATRÍCULA
        if "MATRICULA" not in df_meses.columns:
            return jsonify({"erro": "Erro ao gerar MATRICULA"}), 400

        # 🔥 IPEO
        df_ipeo = ler_excel(ipeo_file)

        if "MATRICULA" not in df_ipeo.columns:
            return jsonify({"erro": "IPEO sem coluna MATRICULA"}), 400

        if "MES" not in df_ipeo.columns:
            df_ipeo["MES"] = "SEM MES"

        # 🔥 PADRONIZAR
        df_meses["MATRICULA"] = df_meses["MATRICULA"].astype(str).str.strip()
        df_ipeo["MATRICULA"] = df_ipeo["MATRICULA"].astype(str).str.strip()

        df_meses["MES"] = df_meses["MES"].astype(str)
        df_ipeo["MES"] = df_ipeo["MES"].astype(str)

        # 🔥 MERGE
        df = pd.merge(df_meses, df_ipeo, on=["MATRICULA","MES"], how="left")

        df = limpar(df)

        # 🔥 AGRUPAMENTO
        def agregar(g):
            res = {}

            res["EMPRESA"] = g.get("EMPRESA", ["SEM DADO"]).iloc[0]
            res["MES"] = g["MES"].iloc[0]
            res["MATRICULA"] = g["MATRICULA"].iloc[0]
            res["NOME"] = g["NOME"].iloc[0]

            # PRESTADOR POR TMS
            if "TMS" in g and "PRESTADOR" in g:
                resumo = g.groupby("PRESTADOR")["TMS"].sum()
                res["PRESTADOR"] = resumo.idxmax()
            else:
                res["PRESTADOR"] = "SEM DADO"

            if "REGIONAL" in g:
                res["REGIONAL"] = g["REGIONAL"].value_counts().idxmax()

            if "POLO" in g:
                res["POLO"] = g["POLO"].value_counts().idxmax()

            for c in g.columns:
                if pd.api.types.is_numeric_dtype(g[c]):
                    res[c] = g[c].mean()

            return pd.Series(res)

        df_final = df.groupby(["MES","MATRICULA","NOME"]).apply(agregar).reset_index(drop=True)

        # 🔥 EXPORTAR
        output = io.BytesIO()
        df_final.to_excel(output, index=False)
        output.seek(0)

        return send_file(output, as_attachment=True, download_name="resultado.xlsx")

    except Exception as e:
        print("ERRO:", str(e))
        return jsonify({"erro": str(e)}), 500
