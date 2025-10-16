import io
from flask import Flask, request, send_file, jsonify
from openpyxl import Workbook

# --- Inicialización de la Aplicación Flask ---
app = Flask(__name__)

# --- Endpoint Principal para Ejecutar Código Openpyxl ---
@app.route('/create-excel-from-script', methods=['POST'])
def create_excel_from_script():
    """
    Recibe un script de Python en el cuerpo de la solicitud, lo ejecuta en un
    contexto que tiene acceso a un libro y hoja de trabajo de openpyxl,
    y devuelve el archivo Excel resultante.
    """
    try:
        # 1. Recibe el código completo del cuerpo de la solicitud.
        # El método get_data() lee todo el stream antes de continuar.
        python_code = request.get_data(as_text=True)

        if not python_code:
            return jsonify({"error": "No se proporcionó ningún script de Python."}), 400

        # Crea un nuevo libro y hoja de trabajo para cada solicitud
        wb = Workbook()
        ws = wb.active
        ws.title = "Hoja1 Excel Genius"

        # --- CONTEXTO DE EJECUCIÓN CONTROLADO ---
        # Se define un diccionario que contendrá las variables y módulos
        # disponibles para el script del usuario.
        execution_context = {
            'wb': wb,
            'ws': ws,
            # Módulos y clases de uso común de openpyxl
            'Font': __import__('openpyxl.styles', fromlist=['Font']).Font,
            'PatternFill': __import__('openpyxl.styles', fromlist=['PatternFill']).PatternFill,
            'Alignment': __import__('openpyxl.styles', fromlist=['Alignment']).Alignment,
            'Border': __import__('openpyxl.styles', fromlist=['Border']).Border,
            'Side': __import__('openpyxl.styles', fromlist=['Side']).Side,
            'LineChart': __import__('openpyxl.chart', fromlist=['LineChart']).LineChart,
            'Reference': __import__('openpyxl.chart', fromlist=['Reference']).Reference,
        }

        # Ejecuta el script de Python proporcionado dentro del contexto definido
        exec(python_code, execution_context)

        # Guarda el libro de trabajo modificado en un buffer en memoria
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0) # Regresa al inicio del buffer para la lectura

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
