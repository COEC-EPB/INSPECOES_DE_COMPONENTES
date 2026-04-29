from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import unicodedata

app = Flask(__name__)
CORS(app)


@app.route('/')
def home():
    return "API online 🚀"


# 🔹 NORMALIZAR COLUNAS
def normalizar_colunas(df):
    def limpar(col):
        col = str(col).strip().upper()
        col = unicodedata.normalize('NFKD', col).encode('ASCII', 'ignore').decode('ASCII')
        return col
    df.columns = [limpar(c) for c in df.columns]
    return df


# 🔹 ENCONTRAR COLUNA APROXIMADA
def col(df, nome):
    nome = nome.upper()
    for c in df.columns:
        if nome in c:
            return c
    return None


# 🔹 LER EXCEL (HEADER INTELIGENTE)
def ler_excel(file):
    df_raw = pd.read_excel(file, header=None)

    header_row = None

    for i, row in df_raw.iterrows():
        valores = row.astype(str).str.upper()

        if valores.str.contains("MATR").any():
            header_row = i
            break

    if header_row is None:
        raise ValueError("Não encontrou cabeçalho com MATR")

    df = pd.read_excel(file, header=header_row)
    df = normalizar_colunas(df)

    print("COLUNAS DETECTADAS:", df.columns.tolist())

    return df


# 🔹 EXTRAIR MES
def extrair_mes(nome):
    nome = nome.upper()

    mapa = {
        "JAN": "JAN", "FEV": "FEV", "MAR": "MAR", "ABR": "ABR",
        "MAI": "MAI", "JUN": "JUN", "JUL": "JUL", "AGO": "AGO",
        "SET": "SET", "OUT": "OUT", "NOV": "NOV", "DEZ": "DEZ"
    }

    for k in mapa:
        if k in nome:
            return mapa[k]

    return None


# 🔹 SEPARAR FUNCIONARIO (MATRICULA - NOME)
def separar_funcionario(df, col_func):

    serie = df[col_func].astype(str)

    split = serie.str.split("-", n=1, expand=True)

    df["MATRICULA"] = split[0].str.strip()

    if split.shape[1] > 1:
        df["NOME"] = split[1].str.strip()
    else:
        df["NOME"] = ""

    df["MATRICULA"] = df["MATRICULA"].str.replace(".0", "", regex=False)

    return df


# 🔹 PADRONIZAR CHAVES
def padronizar_chaves(df):
    df["MATRICULA"] = df["MATRICULA"].astype(str).str.replace(".0", "", regex=False).str.strip()
    df["MES"] = df["MES"].astype(str).str.upper().str.strip()
    return df


# 🔥 ROTA PRINCIPAL
@app.route('/processar', methods=['POST'])
def processar():
    try:

        arquivos = request.files.getlist("meses")
        ipeo_file = request.files.get("ipeo")

        if not arquivos or not ipeo_file:
            return jsonify({"erro": "Envie arquivos"}), 400

        # 🔹 MESES
        lista = []

        for f in arquivos:
            df = ler_excel(f)

            mes = extrair_mes(f.filename)
            if not mes:
                return jsonify({"erro": f"Mês inválido: {f.filename}"}), 400

            df["MES"] = mes

            col_func = col(df, "FUNCIONARIO")
            if not col_func:
                return jsonify({"erro": f"FUNCIONARIO não encontrado em {f.filename}"}), 400

            df = separar_funcionario(df, col_func)

            lista.append(df)

        df_meses = pd.concat(lista, ignore_index=True)

        # 🔹 IPEO
        df_ipeo = ler_excel(ipeo_file)

        # 🔹 PADRONIZAÇÃO
        df_meses = padronizar_chaves(df_meses)
        df_ipeo = padronizar_chaves(df_ipeo)

        print("MES MESES:", df_meses["MES"].unique())
        print("MES IPEO:", df_ipeo["MES"].unique())

        print("MATRICULAS MESES:", df_meses["MATRICULA"].head())
        print("MATRICULAS IPEO:", df_ipeo["MATRICULA"].head())

        # 🔹 MERGE
        df = df_meses.merge(df_ipeo, on=["MATRICULA", "MES"], how="inner")

        print("LINHAS APÓS MERGE:", len(df))

        if df.empty:
            return jsonify({"erro": "Merge não encontrou dados (verifique MATRÍCULA/MES)"}), 400

        # 🔹 COLUNAS DINÂMICAS
        def get(nome):
            return col(df, nome)

        colunas_grupo = [
            get("EMPRESA"),
            "MES",
            get("REGIONAL"),
            get("PRESTADOR"),
            "MATRICULA",
            get("N
