import io
from flask import Flask, request, send_file, jsonify
# Ya no necesitamos pre-importar nada de openpyxl aquí para el agente.

app = Flask(__name__)

@app.route('/process-excel', methods=['POST'])
def process_excel():
    """
    Recibe un script de Python que define una función 'generar_excel()'.
    Ejecuta esta función, que debe devolver un objeto Workbook de openpyxl,
    y envía el archivo Excel resultante.
    
    Este enfoque es seguro en un entorno controlado y da máxima flexibilidad al agente.
    """
    try:
        python_code = request.get_data(as_text=True)

        if not python_code:
            return jsonify({"error": "No se proporcionó ningún script de Python."}), 400

        # Preparamos un diccionario de locales para la ejecución.
        # Esto nos permite capturar la función definida por el agente.
        local_scope = {}
        
        # 1. Ejecutamos el código del agente para DEFINIR la función 'generar_excel'.
        exec(python_code, globals(), local_scope)

        # 2. Verificamos que la función fue definida.
        if 'generar_excel' not in local_scope:
            return jsonify({"error": "El script no definió la función 'generar_excel()'."}), 400

        # 3. LLAMAMOS a la función y obtenemos el objeto workbook.
        generar_excel_func = local_scope['generar_excel']
        wb = generar_excel_func() # El contrato es que la función devuelve el workbook.

        # 4. Guardamos el workbook en memoria y lo enviamos.
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)

        return send_file(
            buffer,
            as_attachment=True,
            download_name='resultado_excel.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        # Captura cualquier error, tanto en 'exec' como en la llamada a la función.
        error_message = f"Error durante la ejecución del script: {type(e).__name__}: {e}"
        print(error_message)
        # Devolvemos un error 500 con el mensaje para facilitar el debug.
        return jsonify({"error": error_message}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5001)
