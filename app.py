from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import tempfile
import os
import xlwt

app = Flask(__name__)
CORS(app)


print("xlwt está instalado correctamente.")


def leer_excel(archivo):
    filename = archivo.filename.lower()
    if filename.endswith(".xls"):
        return pd.read_excel(archivo, engine="xlrd")
    else:
        return pd.read_excel(archivo, engine="openpyxl")

@app.route("/procesar", methods=["POST"])
def procesar():
    try:
        files = request.files
        archivo_real = files.get("stock_real")
        archivo_ecommerce = files.get("stock_ecommerce")

        if not archivo_real or not archivo_ecommerce:
            return jsonify({"error": "Faltan archivos"}), 400

        stock_real = leer_excel(archivo_real)
        stock_ecommerce = leer_excel(archivo_ecommerce)

        stock_real.columns = stock_real.columns.str.strip()
        stock_ecommerce.columns = stock_ecommerce.columns.str.strip()

        stock_real = stock_real.rename(columns={'Codigo Interno': 'Codigo'})
        stock_ecommerce = stock_ecommerce.rename(columns={'Art.': 'Codigo'})

        merged = pd.merge(
            stock_ecommerce,
            stock_real[['Codigo', 'Stock']],
            on='Codigo',
            how='left'
        )

        filtrado = merged[
            (merged['Stock Web'] != merged['Stock']) & (~merged['Stock'].isna())
        ].copy()

        filtrado['Stock Web'] = filtrado['Stock'].astype(int)

        resultado = filtrado[['ID', 'Codigo', 'Stock Web']].rename(columns={'Codigo': 'Art.'})

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp:
            resultado.to_excel(tmp.name, index=False, sheet_name='Actualización', engine='xlwt')
            tmp.seek(0)
            return send_file(
                tmp.name,
                mimetype="application/vnd.ms-excel",
                as_attachment=True,
                download_name="stock_actualizado.xls"
            )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
