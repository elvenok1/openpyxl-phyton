import io
from flask import Flask, request, send_file, jsonify
from openpyxl import Workbook
# El script del usuario ahora podría necesitar estas importaciones,
# por lo que el servidor ya no las "inyecta".
# from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
# from openpyxl.chart import LineChart, Reference

# --- Inicialización de la Aplicación Flask ---
app = Flask(__name__)

# --- Endpoint Principal para Ejecutar Código Openpyxl ---
@app.route('/create-excel-from-script', methods=['POST'])
def create_excel_from_script():
    """
    Recibe un script de Python, lo ejecuta en el ámbito local de esta función
    y devuelve el archivo Excel resultante.
    
    ADVERTENCIA: La ejecución de código dinámico de esta manera es riesgosa.
    El script tendrá acceso a las variables locales 'wb' y 'ws'.
    """
    try:
        python_code = request.get_data(as_text=True)

        if not python_code:
            return jsonify({"error": "No se proporcionó ningún script de Python."}), 400

        # Se crean las variables que estarán disponibles para el script ejecutado
        wb = Workbook()
        ws = wb.active
        ws.title = "Hoja1 Excel Genius"

        # --- EJECUCIÓN DIRECTA ---
        # El código se ejecuta en el ámbito actual. Tiene acceso a 'wb', 'ws',
        # 'python_code', 'request', y cualquier otra variable local.
        # El script del usuario ahora debe gestionar sus propias importaciones.
        exec(python_code)

        # Guarda el libro de trabajo modificado en un buffer en memoria
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        # Envía el buffer como un archivo Excel para descargar
        return send_file(
            buffer,
            as_attachment=True,
            download_name='reporte_dinamico.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        # Captura cualquier error durante la ejecución y lo reporta
        print(f"Error durante la ejecución del script: {e}")
        return jsonify({"error": f"Error interno del servidor al ejecutar el script: {str(e)}"}), 500

if __name__ == '__main__':
    # Ejecuta la aplicación Flask
    app.run(debug=True, port=5001)
