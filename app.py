from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import tempfile
import os

app = Flask(__name__)
CORS(app)

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

        stock_real_reducido = stock_real[['Art.', 'Stock Web', '$ web']]

        merged = pd.merge(
            stock_ecommerce,
            stock_real_reducido,
            on='Art.',
            how='left',
            suffixes=('', '_nuevo')
        )

        merged['Stock Web'] = merged['Stock Web_nuevo'].combine_first(merged['Stock Web'])
        merged['$ web'] = merged['$ web_nuevo'].combine_first(merged['$ web'])

        merged = merged.drop(columns=['Stock Web_nuevo', '$ web_nuevo'])

        columnas_finales = [
            'ID', 'Art.', 'Nombre', 'Stock Local', 'Stock Web',
            '$ ML', '$ web', '$ FT', 'U$S', 'Categoria', 'Marca',
            'Dimensiones', 'Tags', 'GTIN', 'Proveedor', 'Condicion',
            'Estado Web', 'IVA', 'Imp. int.'
        ]

        resultado = merged[columnas_finales]

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
