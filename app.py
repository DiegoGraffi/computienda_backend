from flask import Flask, request, send_file
from flask_cors import CORS
import pandas as pd
from io import BytesIO

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
    files = request.files
    archivo_real = files.get("stock_real")
    archivo_ecommerce = files.get("stock_ecommerce")

    if not archivo_real or not archivo_ecommerce:
        return {"error": "Faltan archivos"}, 400

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

    filtrado = merged[(merged['Stock Web'] != merged['Stock']) & (~merged['Stock'].isna())].copy()
    filtrado['Stock Web'] = filtrado['Stock'].astype(int)

    resultado = filtrado[['ID', 'Codigo', 'Stock Web']].rename(columns={'Codigo': 'Art.'})

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlwt') as writer:
        resultado.to_excel(writer, index=False, sheet_name='Actualización')
    output.seek(0)

    return send_file(
        output,
        mimetype="application/vnd.ms-excel",  
        as_attachment=True,
        download_name="stock_actualizado.xls" 
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)