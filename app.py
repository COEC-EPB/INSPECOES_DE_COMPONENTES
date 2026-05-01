from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import unicodedata
import io
import os
import re

app = Flask(__name__)
CORS(app)


@app.route('/')
def home():
    return "API online 🚀"


# ==============================
# NORMALIZAÇÃO ROBUSTA
# ==============================
def norm(s):
    s = str(s).upper().strip()
    s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("ASCII")
    s = s.replace("%", "").strip()
    s = re.sub(r"\s+", " ", s)
    return s


def normalizar_colunas(df):
    df.columns = [norm(c) for c in df.columns]
    return df


def find_col(df, alvo):
    alvo = norm(alvo)
    cols = {norm(c): c for c in df.columns}

    if alvo in cols:
        return cols[alvo]

    for k, original in cols.items():
        if k.startswith(alvo):
            return original

    return None


# ==============================
# LEITURA INTELIGENTE
# ==============================
def ler_excel(file):
    df_raw = pd.read_excel(file, header=None)

    header_row = None
    for i, row in df_raw.iterrows():
        if row.astype(str).str.upper().str.contains("FUNCIONARIO").any():
            header_row = i
            break

    if header_row is None:
        header_row = 0

    df = pd.read_excel(file, header=header_row)
    df = normalizar_colunas(df)
    return df


def extrair_mes(nome):
    nome = nome.upper()
    meses = ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN",
             "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]
    for m in meses:
        if m in nome:
            return m
    return None


# ==============================
# FUNCIONÁRIO
# ==============================
def separar_funcionario(df):
    col_func = find_col(df, "FUNCIONARIO")
    split = df[col_func].astype(str).str.split("-", n=1, expand=True)
    df["MATRICULA"] = split[0].str.strip().str.replace(".0", "", regex=False)
    df["NOME FUNCIONARIO"] = split[1].str.strip()
    return df


def limpar_numericos(df):
    for c in df.columns:
        df[c] = (
            df[c].astype(str)
            .str.replace("%", "", regex=False)
            .str.replace(",", ".", regex=False)
            .str.strip()
        )
        df[c] = pd.to_numeric(df[c], errors="ignore")
    return df


# ==============================
# ROTA PRINCIPAL
# ==============================
@app.route('/processar', methods=['POST'])
def processar():
    try:
        meses_files = request.files.getlist("meses")
        ipeo_file = request.files.get("ipeo")

        if not meses_files or not ipeo_file:
            return jsonify({"erro": "Envie os arquivos"}), 400

        dfs = []

        for f in meses_files:
            df = ler_excel(f)
            df["MES"] = extrair_mes(f.filename)
            df = separar_funcionario(df)
            dfs.append(df)

        df_meses = pd.concat(dfs, ignore_index=True)

        df_ipeo = ler_excel(ipeo_file)

        # CHAVES
        for df in [df_meses, df_ipeo]:
            df["MATRICULA"] = df["MATRICULA"].astype(str)
            df["MES"] = df["MES"].astype(str)

        # MERGE
        df = pd.merge(df_meses, df_ipeo, on=["MATRICULA", "MES"], how="inner")

        if df.empty:
            return jsonify({"erro": "Merge vazio"}), 400

        df = limpar_numericos(df)

        # ==============================
        # AGRUPAMENTO
        # ==============================
        col_grupo = [
            find_col(df, "EMPRESA"),
            "MES",
            find_col(df, "REGIONAL"),
            find_col(df, "PRESTADOR"),
            find_col(df, "POLO"),
            "MATRICULA",
            "NOME FUNCIONARIO"
        ]
        col_grupo = [c for c in col_grupo if c]

        agg = {}
        for c in df.columns:
            if c in col_grupo:
                continue
            agg[c] = "mean" if pd.api.types.is_numeric_dtype(df[c]) else "first"

        df = df.groupby(col_grupo, as_index=False).agg(agg)

        # ==============================
        # CÁLCULO DOS %
        # ==============================
        hh_disp = find_col(df, "HH DISPONIVEL")

        mapa_hh = {
            "% DI":  "HH IMPROD DI",
            "% ROE": "HH IMPROD ROE",
            "% RNT": "HH IMPROD RNT",
            "% IOC": "HH IMPROD IOC",
            "% ISF": "HH IMPROD ISF",
            "% ROV": "HH IMPROD ROV"
        }

        for pct, hh in mapa_hh.items():
            c_hh = find_col(df, hh)
            if hh_disp and c_hh:
                df[pct] = (df[c_hh] / df[hh_disp]).round(4)

        # ==============================
        # RENOMEAR
        # ==============================
        df = df.rename(columns={
            "MES": "MÊS",
            "MATRICULA": "MATRÍCULA",
            "NOME FUNCIONARIO": "NOME FUNCIONÁRIO"
        })

        # ==============================
        # ORDEM FINAL
        # ==============================
        ordem = [
            "EMPRESA", "MÊS", "REGIONAL", "PRESTADOR",
            "MATRÍCULA", "NOME FUNCIONÁRIO",
            "UTILIZACAO", "PRODUTIVIDADE", "EFICIENCIA",
            "TMS", "% DI", "% ROE", "% RNT",
            "% IOC", "% ISF", "% ROV", "% IPEO", "POLO"
        ]

        lookup = {norm(c): c for c in df.columns}
        cols_final = [lookup[norm(c)] for c in ordem if norm(c) in lookup]
        df = df[cols_final]

        # ==============================
        # EXPORTAR
        # ==============================
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False)

        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="resultado.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        return jsonify({"erro": str(e)}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
