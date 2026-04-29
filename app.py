from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import unicodedata

app = Flask(__name__)
CORS(app)

@app.route('/')
def home():
    return "API online 🚀"


# 🔹 EXTRAIR MES DO NOME DO ARQUIVO
def extrair_mes(nome):
    nome = nome.upper()

    if "JAN" in nome: return "JAN"
    elif "FEV" in nome: return "FEV"
    elif "MAR" in nome: return "MAR"
    elif "ABR" in nome: return "ABR"
    elif "MAI" in nome: return "MAI"
    elif "JUN" in nome: return "JUN"
    elif "JUL" in nome: return "JUL"
    elif "AGO" in nome: return "AGO"
    elif "SET" in nome: return "SET"
    elif "OUT" in nome: return "OUT"
    elif "NOV" in nome: return "NOV"
    elif "DEZ" in nome: return "DEZ"
    else: return None


# 🔹 LIMPAR TEXTO
def limpar(texto):
    texto = str(texto).strip().upper()
    texto = unicodedata.normalize('NFKD', texto).encode('ASCII', 'ignore').decode('ASCII')
    return texto


# 🔹 NORMALIZAR COLUNAS
def normalizar(df):
    df.columns = [limpar(c) for c in df.columns]
    return df


# 🔹 LER EXCEL COM HEADER AUTOMÁTICO
def ler_excel(file):
    df_raw = pd.read_excel(file, header=None)

    header_row = 0

    for i, row in df_raw.iterrows():
        texto = " ".join(row.astype(str).apply(limpar))
        if "FUNCIONARIO" in texto:
            header_row = i
            break

    df = pd.read_excel(file, header=header_row)
    return normalizar(df)


# 🔹 ENCONTRAR COLUNA
def col(df, nome):
    nome = limpar(nome).replace(" ", "")
    for c in df.columns:
        if nome in limpar(c).replace(" ", ""):
            return c
    return None


# 🔹 SEPARAR MATRICULA E NOME
def separar_funcionario(df, coluna):
    split = df[coluna].astype(str).str.split(" - ", n=1, expand=True)

    df["MATRICULA"] = split[0].str.strip()
    df["NOME"] = split[1].str.strip()

    # evita problema tipo 3080117.0
    df["MATRICULA"] = df["MATRICULA"].str.replace(r"\.0$", "", regex=True)

    return df


@app.route('/processar', methods=['POST'])
def processar():
    try:
        arquivos = request.files.getlist("meses")
        ipeo_file = request.files.get("ipeo")

        if not arquivos or not ipeo_file:
            return jsonify({"erro": "Envie arquivos"}), 400

        # 🔹 PROCESSAR MESES
        lista = []

        for f in arquivos:
            try:
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

            except Exception as e:
                return jsonify({"erro": f"Erro {f.filename}: {str(e)}"}), 400

        df_meses = pd.concat(lista, ignore_index=True)

        # 🔹 PROCESSAR IPEO
        try:
            df_ipeo = ler_excel(ipeo_file)
        except Exception as e:
            return jsonify({"erro": f"Erro no IPEO: {str(e)}"}), 400

        # garantir string
        df_meses["MATRICULA"] = df_meses["MATRICULA"].astype(str)
        df_ipeo["MATRICULA"] = df_ipeo["MATRICULA"].astype(str)

        # 🔹 AJUSTAR MES IPEO (se vier número)
        if df_ipeo["MES"].dtype != "object":
            mapa = {
                1: "JAN", 2: "FEV", 3: "MAR", 4: "ABR",
                5: "MAI", 6: "JUN", 7: "JUL", 8: "AGO",
                9: "SET", 10: "OUT", 11: "NOV", 12: "DEZ"
            }
            df_ipeo["MES"] = df_ipeo["MES"].map(mapa)

        # 🔹 MERGE FINAL
        df = df_meses.merge(df_ipeo, on=["MATRICULA", "MES"], how="inner")

        # 🔹 AGRUPAMENTO FINAL
        df_final = df.groupby(
            ["EMPRESA", "MES", "REGIONAL", "POLO", "MATRICULA", "NOME"],
            as_index=False
        ).mean(numeric_only=True)

        # 🔹 EXPORTAR
        output = "resultado.xlsx"
        df_final.to_excel(output, index=False)

        return send_file(output, as_attachment=True)

    except Exception as e:
        print("ERRO:", str(e))
        return jsonify({"erro": str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True)
