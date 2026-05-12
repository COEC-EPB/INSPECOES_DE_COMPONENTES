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

    header_row = None

    # 🔍 procura a linha que tem FUNCIONARIO ou MATRÍCULA
    for i, row in df_raw.iterrows():
        row_str = row.astype(str).str.upper()

        if row_str.str.contains("FUNCIONARIO").any() or \
           row_str.str.contains("MATRICULA").any():
            header_row = i
            break

    # fallback inteligente
    if header_row is None:
        for i, row in df_raw.iterrows():
            row_str = row.astype(str).str.upper()
            if row_str.str.contains("%").any():
                header_row = i
                break

    # último fallback
    if header_row is None:
        header_row = 2

    print(f"✅ HEADER DETECTADO NA LINHA: {header_row}")

    df = pd.read_excel(file, header=header_row)

    # normalizar colunas
    df.columns = [
        unicodedata.normalize('NFKD', str(c))
        .encode('ASCII', 'ignore')
        .decode('ASCII')
        .strip()
        .upper()
        for c in df.columns
    ]

    print("📊 COLUNAS CORRETAS:", df.columns.tolist())

    return df


# 🔹 EXTRAIR MES
def extrair_mes(nome):
    nome = nome.upper()
    meses = ["JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"]
    for m in meses:
        if m in nome:
            return m
    return None


# 🔹 SEPARAR FUNCIONARIO
def separar(df, nome_arquivo=""):

    possiveis = ["FUNCIONARIO", "FUNCIONÁRIO", "NOME", "COLABORADOR"]

    col = None

    for p in possiveis:
        for c in df.columns:
            if p in c:
                col = c
                break

        if col:
            break

    if not col:
        raise Exception(f"Coluna funcionário não encontrada: {nome_arquivo}")

    texto = df[col].astype(str)

    # 🔥 matrícula
    df["MATRICULA"] = (
        texto
        .str.extract(r"(\d+)")[0]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )

    # 🔥 nome
    df["NOME"] = (
        texto
        .str.replace(r"^\d+\s*-\s*", "", regex=True)
        .str.strip()
    )

    return df


# 🔹 LIMPAR NUMÉRICOS
def limpar(df):

    # 🔥 apenas colunas numéricas conhecidas
    colunas_numericas = [
        "% PRODUTIVIDADE",
        "% EFICIENCIA",
        "% UTILIZACAO",
        "TMS",
        "% DI",
        "% ROE",
        "% RNT",
        "% IOC",
        "% ISF",
        "% ROV",
        "% IPEO"
    ]

    for c in colunas_numericas:

        if c in df.columns:

            df[c] = (
                df[c]
                .astype(str)
                .str.replace("%", "", regex=False)
                .str.replace(",", ".", regex=False)
                .str.strip()
            )

            df[c] = pd.to_numeric(df[c], errors="coerce")

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

            df = separar(df, f.filename)
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


        # 🔥 MAPA DE MESES
        mapa_meses = {
            "1": "JAN", "01": "JAN", 1: "JAN",
            "2": "FEV", "02": "FEV", 2: "FEV",
            "3": "MAR", "03": "MAR", 3: "MAR",
            "4": "ABR", "04": "ABR", 4: "ABR",
            "5": "MAI", "05": "MAI", 5: "MAI",
            "6": "JUN", "06": "JUN", 6: "JUN",
            "7": "JUL", "07": "JUL", 7: "JUL",
            "8": "AGO", "08": "AGO", 8: "AGO",
            "9": "SET", "09": "SET", 9: "SET",
            "10": "OUT", 10: "OUT",
            "11": "NOV", 11: "NOV",
            "12": "DEZ", 12: "DEZ",
        }
        
        # 🔹 PADRONIZAR MES IPEO
        df_ipeo["MES"] = df_ipeo["MES"].astype(str).str.strip()
        df_ipeo["MES"] = df_ipeo["MES"].map(mapa_meses).fillna(df_ipeo["MES"])

        # 🔹 PADRONIZAR MES MIP
        df_meses["MES"] = df_meses["MES"].astype(str).str.strip()
        df_meses["MES"] = df_meses["MES"].map(mapa_meses).fillna(df_meses["MES"])
        # 🔥 PADRONIZAR
        df_meses["MATRICULA"] = df_meses["MATRICULA"].astype(str).str.strip()
        df_ipeo["MATRICULA"] = df_ipeo["MATRICULA"].astype(str).str.strip()

        df_meses["MES"] = df_meses["MES"].astype(str)
        df_ipeo["MES"] = df_ipeo["MES"].astype(str)

        print("MES MIP:", df_meses["MES"].unique())
        print("MES IPEO:", df_ipeo["MES"].unique())

        df_meses["MATRICULA"] = df_meses["MATRICULA"].astype(str).str.replace(".0","", regex=False).str.strip()
        df_ipeo["MATRICULA"] = df_ipeo["MATRICULA"].astype(str).str.replace(".0","", regex=False).str.strip()
        # 🔥 MERGE

        print(df_meses[["MES","MATRICULA","NOME"]].head())

        print(df_ipeo[["MES","MATRICULA"]].head())
        df = pd.merge(df_meses, df_ipeo, on=["MATRICULA","MES"], how="left")

        print(df.head())
        df = limpar(df)

        # 🔥 AGRUPAMENTO
       # 🔥 AGRUPAMENTO
        def agregar(g):
            res = {}
        
            res["EMPRESA"] = g["EMPRESA"].iloc[0] if "EMPRESA" in g else "SEM DADO"
            res["MES"] = g["MES"].iloc[0] if "MES" in g else "SEM MES"
            res["MATRICULA"] = g["MATRICULA"].iloc[0] if "MATRICULA" in g else "SEM MATRICULA"
            res["NOME"] = g["NOME"].iloc[0] if "NOME" in g else ""
        
            if "PRESTADOR" in g.columns and "TMS" in g.columns:
                g_valid = g.dropna(subset=["PRESTADOR", "TMS"])
        
                if not g_valid.empty:
                    resumo = g_valid.groupby("PRESTADOR")["TMS"].sum()
                    res["PRESTADOR"] = resumo.idxmax() if not resumo.empty else "SEM DADO"
                else:
                    res["PRESTADOR"] = "SEM DADO"
            else:
                res["PRESTADOR"] = "SEM DADO"
        
            if "REGIONAL" in g.columns:
                r = g["REGIONAL"].dropna()
                res["REGIONAL"] = r.value_counts().idxmax() if not r.empty else "SEM DADO"
        
            if "POLO" in g.columns:
                p = g["POLO"].dropna()
                res["POLO"] = p.value_counts().idxmax() if not p.empty else "SEM DADO"
        
            for c in g.columns:
                if pd.api.types.is_numeric_dtype(g[c]):
                    res[c] = g[c].mean()
        
            return pd.Series(res)
        df = df.dropna(subset=["MATRICULA"])

        df = df[df["MATRICULA"] != ""]    

        df_final = df.groupby(["MES","MATRICULA","NOME"]).apply(agregar).reset_index(drop=True)

        # 🔥 EXPORTAR
        output = io.BytesIO()
        df_final.to_excel(output, index=False)
        output.seek(0)

        return send_file(output, as_attachment=True, download_name="resultado.xlsx")

    except Exception as e:
        print("ERRO:", str(e))
        return jsonify({"erro": str(e)}), 500
