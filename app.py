# app.py
from flask import Flask, request, send_file, make_response # Agregada make_response
import pandas as pd 
import numpy as np
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
from io import BytesIO
from flask_cors import CORS 

app = Flask(__name__)
CORS(app) 

# --- FUNCIONES AUXILIARES ---

def normalizar_porcentaje(valor):
    try:
        if pd.isnull(valor):
            return 0.0
        if isinstance(valor, str):
            valor = valor.strip().replace('%', '').replace(' ', '')
        valor = float(valor)
        if valor > 1:
            valor /= 100
        return round(valor, 4)
    except:
        return 0.0

def ultimos_6_digitos(valor):
    try:
        valor_str = str(valor)
        return valor_str[-6:]
    except:
        return np.nan

# --- RUTA DE PROCESAMIENTO ---

@app.route('/procesar_archivo', methods=['POST'])
def procesar_archivo():
    if 'file' not in request.files:
        return {"error": "No se encontró el archivo"}, 400
    file = request.files['file']

    try:
        # Usamos engine='openpyxl' para evitar dependencia de drivers de Windows/VB
        df = pd.read_excel(file, engine='openpyxl')
    except Exception as e:
        return {"error": f"Error al leer el archivo Excel: {e}"}, 500
    
    # Parámetros de tolerancia
    tolerancia_100 = 0.01 
    rango_min = 0.8769
    rango_max = 0.8832 

    # Normalizar %
    df['%'] = df['%'].apply(normalizar_porcentaje)

    # Eliminar facturas con pago específico
    facturas_con_pe = df.loc[df['RECEIPT_METHOD'] == 'PE_F391501_BANC_5414_PEN', 'XXX'].unique()
    df = df[~df['XXX'].isin(facturas_con_pe)]

    # Aplicar filtros
    filtro_transaccion = ~df['TRANSACTION_TYPE'].str.contains('BOL', na=False)
    df_filtrado = df[filtro_transaccion].copy()

    # Separar por moneda
    df_usd = df_filtrado[df_filtrado['CURRENCY'] == 'USD'].copy()
    df_pen = df_filtrado[df_filtrado['CURRENCY'] == 'PEN'].copy()

    # Filtrar montos mínimos
    df_pen = df_pen[df_pen['INVOICE_AMOUNT'] >= 700]
    df_usd = df_usd[df_usd['INVOICE_AMOUNT_FUNCTIONAL'] >= 700]

    # Identificar filas con porcentaje ~100%
    autodetracciones_pen = df_pen[df_pen['%'].between(1 - tolerancia_100, 1 + tolerancia_100)].copy()
    autodetracciones_usd = df_usd[df_usd['%'].between(1 - tolerancia_100, 1 + tolerancia_100)].copy()

    autodetracciones_pen['CURRENCY'] = 'PEN'
    autodetracciones_usd['CURRENCY'] = 'USD'

    df_autodetracciones = pd.concat([autodetracciones_pen, autodetracciones_usd], ignore_index=True)
    
    # Quitar duplicados en AUTODETRACCIONES
    df_autodetracciones = df_autodetracciones.drop_duplicates(subset=['XXX'])

    # Eliminar filas 100% de PEN y USD originales
    df_pen = df_pen[~df_pen['%'].between(1 - tolerancia_100, 1 + tolerancia_100)]
    df_usd = df_usd[~df_usd['%'].between(1 - tolerancia_100, 1 + tolerancia_100)]

    # Eliminar filas con porcentaje entre 87.69% y 88.32%
    df_pen = df_pen[~df_pen['%'].between(rango_min, rango_max)]
    df_usd = df_usd[~df_usd['%'].between(rango_min, rango_max)]

    # Crear nueva columna con últimos 6 dígitos de XXX
    df_pen['XXX_last6'] = df_pen['XXX'].apply(ultimos_6_digitos)
    df_usd['XXX_last6'] = df_usd['XXX'].apply(ultimos_6_digitos)

    # Columnas a mantener
    columnas_a_mantener = [
        'INVOICE_CUSTOMER_NAME',
        'INVOICE_CUSTOMER_TAXPAYER_ID',
        'XXX',
        'INVOICE_DATE',
        'CURRENCY',
        'INVOICE_AMOUNT'
    ]

    # Seleccionar y limpiar DataFrames
    df_autodetracciones = df_autodetracciones[columnas_a_mantener]
    
    df_pen = df_pen[columnas_a_mantener + ['XXX_last6']]
    df_pen = df_pen.drop_duplicates(subset=['XXX_last6'])
    df_pen = df_pen.sort_values(by='XXX_last6')
    
    df_usd = df_usd[columnas_a_mantener + ['XXX_last6']]
    df_usd = df_usd.drop_duplicates(subset=['XXX_last6'])
    df_usd = df_usd.sort_values(by='XXX_last6')
    
    # 4. Guardar los datos en un búfer (BytesIO)
    output = BytesIO()
    nombre_salida = 'SGS Autodetracciones_filtrado_final.xlsx'

    # Guardar a Excel en el búfer
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_filtrado.to_excel(writer, sheet_name='Todos_los_datos', index=False)
        df_usd.to_excel(writer, sheet_name='USD', index=False)
        df_pen.to_excel(writer, sheet_name='PEN', index=False)
        df_autodetracciones.to_excel(writer, sheet_name='AUTODETRACCIONES', index=False)
    
    # Regresar al inicio del búfer
    output.seek(0)
    
    # 5. Devolver el archivo al usuario usando make_response (Más robusto)
    response = make_response(output.getvalue())
    response.headers['Content-Type'] = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    response.headers['Content-Disposition'] = f'attachment; filename="{nombre_salida}"'
    return response

if __name__ == '__main__':
    # Ejecutar en el puerto 5001
    app.run(debug=True, port=5001)