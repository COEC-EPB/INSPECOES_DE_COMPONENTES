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

    print("HEADER:", header_row)
    print("COLUNAS:", df.columns.tolist())

    return df


# 🔹 EXTRAIR MES
def extrair_mes(nome):
    nome = nome.upper()
    meses = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN",
             "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]

    for m in meses:
        if m in nome:
            return m
    return None


# 🔹 SEPARAR FUNCIONARIO
def separar_funcionario(df, col_func):
    split = df[col_func].astype(str).str.split("-", n=1, expand=True)

    df["MATRICULA"] = split[0].str.strip().str.replace(".0", "", regex=False)
    df["NOME"] = split[1].str.strip() if split.shape[1] > 1 else ""

    return df


# 🔹 PADRONIZAR CHAVES
def padronizar_chaves(df):
    df["MATRICULA"] = df["MATRICULA"].astype(str).str.strip()
    df["MES"] = df["MES"].astype(str).str.upper().str.strip()
    return df


# 🔥 LIMPAR NUMÉRICOS (REMOVE % E VÍRGULA)
def limpar_numericos(df):
    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = (
                df[c]
                .astype(str)
                .str.replace("%", "", regex=False)
                .str.replace(",", ".", regex=False)
                .str.strip()
            )
            # tenta converter
            df[c] = pd.to_numeric(df[c], errors="ignore")

    return df


# 🔥 ROTA PRINCIPAL
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
                return jsonify({"erro": f"Mês inválido: {f.filename}"}), 400

            df["MES"] = mes

            col_func = col(df, "FUNCIONARIO")
            if not col_func:
                return jsonify({"erro": f"FUNCIONARIO não encontrado em {f.filename}"}), 400

            df = separar_funcionario(df, col_func)
            lista.append(df)

        df_meses = pd.concat(lista, ignore_index=True)

        df_ipeo = ler_excel(ipeo_file)

        df_meses = padronizar_chaves(df_meses)
        df_ipeo = padronizar_chaves(df_ipeo)

        df = pd.merge(df_meses, df_ipeo, on=["MATRICULA", "MES"], how="inner")

        if df.empty:
            return jsonify({"erro": "Merge vazio"}), 400

        # 🔥 LIMPAR NÚMEROS
        df = limpar_numericos(df)

        # 🔹 COLUNAS DE AGRUPAMENTO
        colunas_grupo = [
            col(df, "EMPRESA"),
            "MES",
            col(df, "REGIONAL"),
            col(df, "PRESTADOR"),
            col(df, "POLO"),
            "MATRICULA",
            col(df, "NOME")
        ]

        colunas_grupo = [c for c in colunas_grupo if c]

        # 🔥 AGG DINÂMICO (NUNCA MAIS QUEBRA)
        agg_dict = {}

        for c in df.columns:
            if c in colunas_grupo:
                continue
            elif pd.api.types.is_numeric_dtype(df[c]):
                agg_dict[c] = "mean"
            else:
                agg_dict[c] = "first"

        df_final = df.groupby(colunas_grupo, as_index=False).agg(agg_dict)

        # 🔹 RENOMEAR
        df_final = df_final.rename(columns={
            "MES": "MÊS",
            "MATRICULA": "MATRÍCULA",
            "% PRODUTIVIDADE": "% Produtividade",
            "% EFICIENCIA": "% Eficiência",
            "% UTILIZACAO": "% Utilização",
            "% DI": "% DI",
            "% ROE": "% ROE",
            "% RNT": "% RNT",
            "% IOC": "% IOC",
            "% ISF": "% ISF",
            "% ROV": "% ROV",
            "% IPEO": "% IPEO"
        })

        colunas_final = [
            "EMPRESA",
            "MÊS",
            "REGIONAL",
            "PRESTADOR",
            "MATRÍCULA",
            "NOME",
            "% Utilização",
            "% Produtividade",
            "% Eficiência",
            "TMS",
            "% DI",
            "% ROE",
            "% RNT",
            "% IOC",
            "% ISF",
            "% ROV",
            "% IPEO",
            "POLO"
        ]
        
        # pega só as que existem (evita erro)
        colunas_existentes = [c for c in colunas_final if c in df_final.columns]
        
        df_final = df_final[colunas_existentes]

        # 🔥 EXPORTAR EXCEL
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
