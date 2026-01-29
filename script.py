from flask import Flask, request, send_file
import pandas as pd
import io

app = Flask(__name__)

@app.route('/upload', methods=['POST'])
def upload_file():
    uploaded_file = request.files['file']
    if uploaded_file and uploaded_file.filename.endswith('.txt'):
        # Leer archivo de texto con pandas (detectando separadores comunes)
        try:
            df = pd.read_csv(uploaded_file, sep=None, engine='python')
        except Exception:
            df = pd.read_csv(uploaded_file, delimiter='\t', engine='python')

        # Convertir a Excel en memoria
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        return send_file(
            output,
            as_attachment=True,
            download_name='resultado.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    return "Formato no soportado", 400

if __name__ == '__main__':
    app.run(debug=True)
