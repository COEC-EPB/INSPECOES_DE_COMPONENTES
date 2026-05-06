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
def normalizar_colunas(df):
    def limpar(col):
        col = str(col).strip().upper()
        col = unicodedata.normalize('NFKD', col).encode('ASCII', 'ignore').decode('ASCII')
        return col
    df.columns = [limpar(c) for c in df.columns]
    return df


# 🔹 BUSCA FLEXÍVEL
def col(df, nome):
    nome = nome.upper()
    for c in df.columns:
        if nome in c:
            return c
    return None


# 🔹 LER EXCEL
def ler_excel(file):
    df_raw = pd.read_excel(file, header=None)

    header_row = None
    for i, row in df_raw.iterrows():
        if row.astype(str).str.upper().str.contains("FUNCIONARIO").any():
            header_row = i
            break

    if header_row is None:
        header_row = 2

    df = pd.read_excel(file, header=header_row)
    df = normalizar_colunas(df)

    return df


# 🔹 EXTRAIR MÊS
def extrair_mes(nome):
    nome = nome.upper()
    meses = ["JAN","FEV","MAR","ABR","MAI","JUN",
             "JUL","AGO","SET","OUT","NOV","DEZ"]
    for m in meses:
        if m in nome:
            return m
    return None


# 🔹 SEPARAR FUNCIONARIO
def separar_funcionario(df, col_func):
    df[col_func] = df[col_func].astype(str)
    split = df[col_func].str.split(" - ", n=1, expand=True)

    df["MATRICULA"] = split[0].str.strip().str.replace(".0", "", regex=False)
    df["NOME"] = split[1].str.strip() if split.shape[1] > 1 else ""

    return df


# 🔹 PADRONIZAR CHAVES
def padronizar_chaves(df):
    df["MATRICULA"] = df["MATRICULA"].astype(str).str.strip()
    df["MES"] = df["MES"].astype(str).str.upper().str.strip()
    return df


# 🔹 LIMPAR NUMÉRICOS
def limpar_numericos(df):
    for c in df.columns:
        if df[c].dtype == "object":
            tem_percent = df[c].astype(str).str.contains("%").any()

            df[c] = (
                df[c].astype(str)
                .str.replace("%","", regex=False)
                .str.replace(",",".", regex=False)
                .str.strip()
            )

            df[c] = pd.to_numeric(df[c], errors="ignore")

            if tem_percent:
                df[c] = df[c] / 100

    return df


# 🔥 ESCOLHER PRESTADOR POR PESO (TMS)
def escolher_prestador(grupo, df):
    col_prestador = col(df, "PRESTADOR")
    col_tms = col(df, "TMS")

    if not col_prestador:
        return "SEM DADO"

    if col_tms and col_tms in grupo:
        resumo = grupo.groupby(col_prestador)[col_tms].sum()
        if not resumo.empty:
            return resumo.idxmax()

    return grupo[col_prestador].dropna().iloc[0]


@app.route('/processar', methods=['POST'])
def processar():
    try:
        arquivos = request.files.getlist("meses")
        ipeo_file = request.files.get("ipeo")

        if not arquivos or not ipeo_file:
            return jsonify({"erro": "Envie arquivos"}), 400

        lista = []

        for f in arquivos:
            df = ler_excel(f)

            mes = extrair_mes(f.filename)
            if not mes:
                return jsonify({"erro": f"Mês inválido"}), 400

            df["MES"] = mes

            col_func = col(df, "FUNCIONARIO")
            if not col_func:
                return jsonify({"erro": "FUNCIONARIO não encontrado"}), 400

            df = separar_funcionario(df, col_func)
            lista.append(df)

        df_meses = pd.concat(lista, ignore_index=True)
        df_ipeo = ler_excel(ipeo_file)

        df_meses = padronizar_chaves(df_meses)
        df_ipeo = padronizar_chaves(df_ipeo)

        # 🔹 MERGE
        df = pd.merge(df_meses, df_ipeo, on=["MATRICULA", "MES"], how="left")

        col_empresa = col(df, "EMPRESA")

        if not col_empresa:
            # tenta alternativas comuns
            for tentativa in ["SGLEMP", "EMP", "EMPRESA_X", "EMPRESA_Y"]:
                col_empresa = col(df, tentativa)
                if col_empresa:
                    break
        
        if not col_empresa:
            df["EMPRESA"] = "SEM DADO"
            col_empresa = "EMPRESA"

        # 🔹 PRESTADOR PROPRIO
        col_prestador = col(df, "PRESTADOR")
        if col_prestador:
            df[col_prestador] = df[col_prestador].fillna("PROPRIO")
            df.loc[df[col_prestador] == "", col_prestador] = "PROPRIO"

        # 🔹 RESOLVER _x _y
        for c in list(df.columns):
            if c.endswith("_x"):
                base = c[:-2]
                y_col = base + "_y"

                if y_col in df.columns:
                    df[base] = df[c].combine_first(df[y_col])
                else:
                    df[base] = df[c]

        df = df[[c for c in df.columns if not c.endswith("_x") and not c.endswith("_y")]]

        df = limpar_numericos(df)

        # 🔹 CORRIGIR %
        colunas_percentuais = ["DI","ROE","RNT","IOC","ISF","ROV","IPEO"]

        for nome in colunas_percentuais:
            col_real = col(df, nome)
            if col_real:
                df[col_real] = df[col_real].apply(
                    lambda x: x/100 if pd.notnull(x) and x > 1 else x
                ).fillna(0)

        # 🔥 AGRUPAMENTO POR MÊS COM REGRA INTELIGENTE

        def agregar(grupo):
            resultado = {}

            resultado["EMPRESA"] = grupo[col(df,"EMPRESA")].iloc[0]
            resultado["MES"] = grupo["MES"].iloc[0]
            resultado["MATRICULA"] = grupo["MATRICULA"].iloc[0]
            resultado["NOME"] = grupo[col(df,"NOME")].iloc[0]

            # 🔥 PRESTADOR CORRETO
            resultado["PRESTADOR"] = escolher_prestador(grupo, df)

            # 🔹 REGIONAL E POLO
            col_regional = col(df,"REGIONAL")
            col_polo = col(df,"POLO")

            if col_regional:
                resultado["REGIONAL"] = grupo[col_regional].dropna().value_counts().idxmax()

            if col_polo:
                resultado["POLO"] = grupo[col_polo].dropna().value_counts().idxmax()

            # 🔹 MÉDIAS
            for coluna in grupo.columns:
                if pd.api.types.is_numeric_dtype(grupo[coluna]):
                    resultado[coluna] = grupo[coluna].mean()

            return pd.Series(resultado)

        df_final = df.groupby([
            col(df,"EMPRESA"),
            "MES",
            "MATRICULA",
            col(df,"NOME")
        ]).apply(agregar).reset_index(drop=True)

        # 🔹 EXPORTAR
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False)

        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="resultado.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        print("ERRO:", str(e))
        return jsonify({"erro": str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8080))
    app.run(host='0.0.0.0', port=port)
