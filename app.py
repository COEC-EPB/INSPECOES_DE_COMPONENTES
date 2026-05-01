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


# 🔹 ENCONTRAR COLUNA APROXIMADA
def col(df, nome):
    nome = nome.upper()
    for c in df.columns:
        if nome in c:
            return c
    return None


# 🔹 LER EXCEL (HEADER AUTOMÁTICO)
def ler_excel(file):

    df_raw = pd.read_excel(file, header=None)

    header_row = None

    for i, row in df_raw.iterrows():
        valores = row.astype(str).str.upper()

        # 🔥 procurar linha REAL de header
        if (
            valores.str.contains("FUNC").any() and
            valores.str.contains("PROD").any()
        ):
            header_row = i
            break

    # fallback seguro
    if header_row is None:
        print("⚠️ HEADER NÃO ENCONTRADO → tentando linha 5")
        header_row = 5

    df = pd.read_excel(file, header=header_row)
    df = normalizar_colunas(df)

    print("✅ HEADER CORRETO:", header_row)
    print("📊 COLUNAS:", df.columns.tolist())

    return df

# 🔹 EXTRAIR MES DO NOME DO ARQUIVO
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


# 🔹 SEPARAR FUNCIONARIO
def separar_funcionario(df, col_func):
    split = df[col_func].astype(str).str.split("-", n=1, expand=True)

    df["MATRICULA"] = split[0].str.strip().str.replace(".0", "", regex=False)
    df["NOME"] = split[1].str.strip() if split.shape[1] > 1 else ""

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

        print("MES MESES:", df_meses["MES"].unique())
        print("MES IPEO:", df_ipeo["MES"].unique())

        df = df_meses.merge(df_ipeo, on=["MATRICULA", "MES"], how="inner")

        print("LINHAS APÓS MERGE:", len(df))

        if df.empty:
            return jsonify({"erro": "Merge vazio → MATRÍCULA ou MES não batem"}), 400

        def get(nome):
            return col(df, nome)

        colunas_grupo = [
            get("EMPRESA"),
            "MES",
            get("REGIONAL"),
            get("PRESTADOR"),
            get("POLO"),
            "MATRICULA",
            get("NOME")
        ]

        colunas_grupo = [c for c in colunas_grupo if c is not None]

        df = df.fillna(0)

        colunas_numericas = [
            "UTILIZACAO", "PRODUTIVIDADE", "EFICIENCIA", "TMS",
            "DI", "ROE", "RNT", "IOC", "ISF", "ROV", "IPEO"
        ]

        for col_num in colunas_numericas:
            col_real = col(df, col_num)
            if col_real:
                df[col_real] = pd.to_numeric(df[col_real], errors="coerce")

        df_final = df.groupby(colunas_grupo, as_index=False).mean(numeric_only=True)

        print("COLUNAS FINAIS:", df_final.columns.tolist())

        mapa_saida = {
            "MES": "MÊS",
            "MATRICULA": "MATRÍCULA",
            "PRODUTIVIDADE": "% Produtividade",
            "EFICIENCIA": "% Eficiência",
            "UTILIZACAO": "% Utilização",
            "DI": "% DI",
            "ROE": "% ROE",
            "RNT": "% RNT",
            "IOC": "% IOC",
            "ISF": "% ISF",
            "ROV": "% ROV",
            "IPEO": "% IPEO"
        }

        df_final = df_final.rename(columns=mapa_saida)

        ordem = [
            "EMPRESA", "MÊS", "REGIONAL", "PRESTADOR", "MATRÍCULA", "NOME",
            "% Utilização", "% Produtividade", "% Eficiência", "TMS",
            "% DI", "% ROE", "% RNT", "% IOC", "% ISF", "% ROV", "% IPEO", "POLO"
        ]
        
        for col_nome in ordem:
            if col_nome not in df_final.columns:
                df_final[col_nome] = 0
            
        ordem_existente = [c for c in ordem if c in df_final.columns]
        df_final = df_final[ordem_existente]

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


# 🔥 PORTA PARA RAILWAY
if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8080))
    app.run(host='0.0.0.0', port=port)
