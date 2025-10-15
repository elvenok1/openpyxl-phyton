import io
from flask import Flask, request, send_file, jsonify
from openpyxl import Workbook

# --- Inicialización de la Aplicación Flask ---
app = Flask(__name__)

# --- Endpoint Principal para Ejecutar Código Openpyxl ---
@app.route('/create-excel-from-script', methods=['POST'])
def create_excel_from_script():
    """
    Este endpoint recibe un script de Python como texto plano.
    El script se ejecuta en el servidor, dándole acceso a un
    libro de trabajo (workbook) y una hoja (worksheet) de openpyxl
    previamente inicializados.
    """
    try:
        # Obtener el código Python del cuerpo de la solicitud
        python_code = request.get_data(as_text=True)
        if not python_code:
            return jsonify({"error": "No se proporcionó ningún script de Python en el cuerpo de la solicitud."}), 400

        # Crear un nuevo libro de trabajo (workbook) y seleccionar la hoja activa
        wb = Workbook()
        ws = wb.active
        ws.title = "Hoja1 Excel Genius"

        # Crear un diccionario para el contexto de ejecución del script
        # Le pasamos el workbook (wb) y la worksheet (ws) para que el script pueda usarlos
        execution_context = {
            'wb': wb,
            'ws': ws,
            # NOTA: Puedes pre-importar y pasar más módulos si lo deseas.
            # Por ejemplo, si tus scripts usan siempre los mismos componentes de openpyxl:
            'Font': __import__('openpyxl.styles', fromlist=['Font']).Font,
            'PatternFill': __import__('openpyxl.styles', fromlist=['PatternFill']).PatternFill,
            'Alignment': __import__('openpyxl.styles', fromlist=['Alignment']).Alignment,
            'Border': __import__('openpyxl.styles', fromlist=['Border']).Border,
            'Side': __import__('openpyxl.styles', fromlist=['Side']).Side,
            'LineChart': __import__('openpyxl.chart', fromlist=['LineChart']).LineChart,
            'Reference': __import__('openpyxl.chart', fromlist=['Reference']).Reference,
        }

        # Ejecutar el código recibido en el contexto que hemos creado
        # ¡ADVERTENCIA! exec() es muy potente pero puede ser inseguro.
        # Asegúrate de confiar plenamente en la fuente del código.
        exec(python_code, execution_context)

        # Guardar el libro de trabajo resultante en un buffer en memoria
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        # Enviar el archivo Excel como respuesta
        return send_file(
            buffer,
            as_attachment=True,
            download_name='reporte_dinamico.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        # Si algo sale mal en el script del usuario, devolvemos un error claro
        print(f"Error durante la ejecución del script: {e}")
        return jsonify({"error": f"Error interno del servidor al ejecutar el script: {str(e)}"}), 500

if __name__ == '__main__':
    # Usamos el puerto 5001 para evitar conflictos si tienes el otro servicio corriendo
    app.run(debug=True, port=5001)
