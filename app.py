from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pandas as pd
import tempfile

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
        archivo_ecommerce = files.get("articulos")

        if not archivo_real or not archivo_ecommerce:
            return jsonify({"error": "Faltan archivos"}), 400

        stock_real = leer_excel(archivo_real)
        stock_ecommerce = leer_excel(archivo_ecommerce)

        stock_real.columns = stock_real.columns.str.strip()
        stock_ecommerce.columns = stock_ecommerce.columns.str.strip()

        stock_real_reducido = stock_real[['Codigo Interno', 'Stock', 'Precio Final', 'IVA']]
        stock_real_reducido.rename(columns={
            'Codigo Interno': 'Art.',
            'Stock': 'Stock_nuevo',
            'Precio Final': '$ web_nuevo',
            'IVA': 'IVA_nuevo'
        }, inplace=True)

        iva_mapping = {'1.21': 21, '1.105': 10.5}
        stock_real_reducido['IVA_nuevo'] = stock_real_reducido['IVA_nuevo'].astype(str).map(iva_mapping)


        merged = pd.merge(
            stock_ecommerce,
            stock_real_reducido,
            on='Art.',
            how='left'
        )

        merged['Stock Web_orig'] = merged['Stock Web']
        merged['$ web_orig'] = merged['$ web']
        merged['IVA_orig'] = merged['IVA']

        merged['Stock Web'] = merged['Stock_nuevo'].combine_first(merged['Stock Web'])
        merged['$ web'] = merged['$ web_nuevo'].combine_first(merged['$ web'])
        merged['IVA'] = merged['IVA_nuevo'].combine_first(merged['IVA'])

        modificados = merged[
            (merged['Stock Web'] != merged['Stock Web_orig']) |
            (merged['$ web'] != merged['$ web_orig']) |
            (merged['IVA'] != merged['IVA_orig'])
        ].copy()

        modificados = modificados.drop(columns=[
            'Stock_nuevo', '$ web_nuevo', 'IVA_nuevo',
            'Stock Web_orig', '$ web_orig', 'IVA_orig'
        ])

        columnas_finales = [
            'ID', 'Art.', 'Nombre', 'Stock Local', 'Stock Web',
            '$ ML', '$ web', '$ FT', 'U$S', 'Categoria', 'Marca',
            'Dimensiones', 'Tags', 'GTIN', 'Proveedor', 'Condicion',
            'Estado Web', 'IVA', 'Imp. int.'
        ]

        resultado = modificados[columnas_finales]

        if 'GTIN' in resultado.columns:
            resultado['GTIN'] = resultado['GTIN'].fillna('').astype(str)

        with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as tmp:
            with pd.ExcelWriter(tmp.name, engine='xlwt') as writer:
                resultado.to_excel(writer, index=False, sheet_name='Actualización')
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
