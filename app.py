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


# 🔹 BUSCAR COLUNA FLEXÍVEL
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


# 🔹 PADRONIZAR
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


# 🔥 ESCOLHER PRESTADOR POR TMS
def escolher_prestador(grupo, col_prestador, col_tms):
    if col_prestador and col_tms and col_tms in grupo:
        resumo = grupo.groupby(col_prestador)[col_tms].sum()
        if not resumo.empty:
            return resumo.idxmax()

    return grupo[col_prestador].dropna().iloc[0] if col_prestador else "SEM DADO"


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
                continue

            df["MES"] = mes

            col_func = col(df, "FUNCIONARIO")
            if not col_func:
                continue

            df = separar_funcionario(df, col_func)
            lista.append(df)

        if not lista:
            return jsonify({"erro": "Nenhum arquivo válido encontrado"}), 400
        df_meses = pd.concat(lista, ignore_index=True)

        # 🔥 GARANTIR MES NO DF_CORRETO
        if "MES" not in df_meses.columns:
            df_meses["MES"] = "SEM MES"
        
        df_meses["MES"] = df_meses["MES"].astype(str).str.upper().str.strip()


        
        
        df_ipeo = ler_excel(ipeo_file)

        if "MES" not in df_ipeo.columns:
            df_ipeo["MES"] = "SEM MES"
        
        df_ipeo["MES"] = df_ipeo["MES"].astype(str).str.upper().str.strip()

        df_meses = padronizar_chaves(df_meses)
        df_ipeo = padronizar_chaves(df_ipeo)

        # 🔹 MERGE
        df = pd.merge(df_meses, df_ipeo, on=["MATRICULA", "MES"], how="left")

        # 🔹 RESOLVER COLUNAS DUPLICADAS
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

        # 🔹 COLUNAS FIXAS
        col_empresa = col(df,"EMPRESA")

        if not col_empresa:
            # tenta alternativas
            for tentativa in ["SGLEMP", "EMP", "EMPRESA_X", "EMPRESA_Y"]:
                col_empresa = col(df, tentativa)
                if col_empresa:
                    break
        
        # se ainda não encontrou → cria
        if not col_empresa:
            df["EMPRESA"] = "SEM DADO"
            col_empresa = "EMPRESA"
        col_nome = col(df,"NOME")
        col_regional = col(df,"REGIONAL")
        col_polo = col(df,"POLO")
        col_prestador = col(df,"PRESTADOR")
        col_tms = col(df,"TMS")

        if "EMPRESA" not in df.columns:
            df["EMPRESA"] = "SEM DADO"

        # 🔥 AGRUPAMENTO
        def agregar(grupo):
            res = {}

            res["EMPRESA"] = grupo[col_empresa].iloc[0] if col_empresa in grupo else "SEM DADO"
            res["MES"] = grupo["MES"].iloc[0] if "MES" in grupo else "SEM MES"
            res["MATRICULA"] = grupo["MATRICULA"].iloc[0]
            res["NOME"] = grupo[col_nome].iloc[0] if col_nome else ""

            # 🔥 PRESTADOR POR PESO
            res["PRESTADOR"] = escolher_prestador(grupo, col_prestador, col_tms)

            if col_regional:
                res["REGIONAL"] = grupo[col_regional].dropna().value_counts().idxmax()

            if col_polo:
                res["POLO"] = grupo[col_polo].dropna().value_counts().idxmax()

            # 🔹 MÉDIAS
            for c in grupo.columns:
                if pd.api.types.is_numeric_dtype(grupo[c]):
                    res[c] = grupo[c].mean()

            return pd.Series(res)

        grupo_cols = ["MATRICULA"]

        if "MES" in df.columns:
            grupo_cols.insert(0, "MES")
        else:
            df["MES"] = "SEM MES"
            grupo_cols.insert(0, "MES")

        if col_empresa:
            grupo_cols.insert(0, col_empresa)
        
        if col_nome:
            grupo_cols.append(col_nome)
        
        df_final = df.groupby(grupo_cols).apply(agregar).reset_index(drop=True)

        # 🔹 EXPORTAR
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False)

        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="resultado.xlsx"
        )

    except Exception as e:
        print("ERRO:", str(e))
        return jsonify({"erro": str(e)}), 500


if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8080))
    app.run(host='0.0.0.0', port=port)
