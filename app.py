from flask import Flask, jsonify, render_template, send_file, request # Framework Flask y utilidades para manejo de JSON, renderizado de plantillas, envío de archivos y solicitudes HTTP
import subprocess # Ayuda en la ejecución de procesos externos (se utiliza para llamar a generator.py)
import os # Manejo de archivos y directorios
import threading # Permite ejecutar tareas en hilos paralelos
import time # Manejo de tiempos de espera (sleep)
from werkzeug.utils import secure_filename
import shutil # Operaciones de alto nivel con archivos (copiar, mover, etc...)

# Configuración del directorio del historial
# Se define la carpeta donde se guardará el historial de los reportes generados
HISTORIAL_DIR = os.path.join(os.getcwd(), "reports_historial")
if not os.path.exists(HISTORIAL_DIR):
    os.makedirs(HISTORIAL_DIR)

# Configuración del directorio de subida y extensiones permitidas
# Se define la carpeta donde se guardará el archivo excel subido por el usuario con la información de los cursos
# Se declaran la extensiones validas para el archivo
UPLOAD_FOLDER = os.path.join(os.getcwd(), "excel_upload")
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Se crea una instancia de la aplicación Flask
app = Flask(__name__, template_folder="./public")

# =============================================================
# Ruta principal (Home)
# Objetivo: Renderizar la interfaz web principal (index.html).
# =============================================================
@app.route('/')
def home():
    return render_template("index.html")

# =====================================================================================================
# Endpoint: /upload_excel
# Objetivo: Actualizar la base de datos con la información de los cursos y participantes
# sin tener que acceder directamente al servidor.
# - Se recibe el archivo Excel desde la web, lo valida y lo guarda o reemplaza en una
#   ubicación accesible para que el proceso de generación de reportes utilice los datos actualizados.
# - Devuelve una respuesta JSON indicando ya sea el éxito o el error.
# =====================================================================================================
@app.route('/upload_excel', methods=['POST'])
def upload_excel():
    if 'excel_file' not in request.files:
        return jsonify({"error": "No se encontró el archivo en la solicitud"}), 400
    
    file = request.files['excel_file']
    if file.filename == "":
        return jsonify({"error": "No se seleccionó ningún archivo"}), 400
    
    if file and allowed_file(file.filename):
        # Se eliminan los archivos previos en UPLOAD_FOLDER para que quede solo el último que se subio
        for f in os.listdir(UPLOAD_FOLDER):
            os.remove(os.path.join(UPLOAD_FOLDER, f))

        filename = secure_filename(file.filename)
        file_path = os.path.join(UPLOAD_FOLDER, filename)
        try:
            file.save(file_path)
            DATA_FILE = os.path.join(os.getcwd(), "db_excel.xlsx")
            shutil.copy(file_path, DATA_FILE)
            return jsonify({"message": f"Archivo '{filename}' subido y actualizado exitosamente."})
        except Exception as e:
            return jsonify({"error": f"Error al guardar el archivo: str{e}"}), 500
    else:
        return jsonify({"error": "Tipo de archivo no permitido. Solo se permiten archivos Excel."}), 400
    
# ======================================================================================================
# Endpoint: /current_excel
# Objetivo: Devolver el nombre del ultimo archivo Excel subido por el usuario con la intención de que
# el usuario sepa de que archivo se generaran los documentos en caso de querer generarlos al instante.
# ======================================================================================================   
@app.route('/current_excel', methods=['GET'])
def current_excel():
    try:
        files = os.listdir(UPLOAD_FOLDER)
        if len(files) == 1:
            current_filename = files[0]
            return jsonify({ "filename": current_filename })
        else:
            return jsonify({"filename": None})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# ===============================================================================================
# Función: allowed_file
# Objetivo: Verificar que el archivo subido por el usuario contenga las extensiones permitidas,
# en este caso solo Excel es permitido.
# ===============================================================================================
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# ==========================================================================================
# Endpoint: /download_zip
# Objetivo: Permitir la descarga del archivo ZIP generado.
# - Recibe el nombre del archivo a través de un parámetro GET.
# - Envía el archivo como descarga adjunta.
# - Posteriormente, lanza un hilo para eliminar el ZIP del servidor después de un retraso.
# ==========================================================================================
@app.route('/download_zip', methods=['GET'])
def download_zip():
    file_name = request.args.get("file")
    zip_path = os.path.join(os.getcwd(), file_name)
    if os.path.exists(zip_path):
        response = send_file(zip_path, as_attachment=True)
        # Se lanza la eliminación del archivo ZIP en un hilo separado para no bloquear la respuesta
        threading.Thread(target=delayed_delete, args=(zip_path,)).start()
        return response
    return jsonify({"error": "Archivo ZIP no encontrado, revise el formato del archivo de datos"}), 500

# =============================================================================
# Función: delayed_delete
# Objetivo: Esperar 5 segundos y luego eliminar el archivo ZIP del servidor
# =============================================================================
def delayed_delete(zip_path):
    time.sleep(5) # Se espera 5 segundos antes de eliminar el ZIP del servidor
    try:
        if os.path.exists(zip_path):
            os.remove(zip_path)
            print(f"Archivo ZIP eliminado del servidor: {zip_path}")
    except Exception as e:
        app.logger.error(f"Error al eliminar el archivo ZIP: {e}");
# ===========================================================================
# Endpoint: /historial
# Objetivo: Listar los archivos presentes en el directorio del historial.
# - Devuelve un JSON con la lista de archivos.
# ===========================================================================
@app.route('/historial', methods=['GET'])
def list_historial():
    try:
        # Se verifica que el directorio del historial exista en tiempo de ejecución
        if not os.path.exists(HISTORIAL_DIR):
            return jsonify({"error": "El directorio del historial no existe"}), 500
        # Se listan todos los archivos del directorio HISTORIAL_DIR
        files = os.listdir(HISTORIAL_DIR)
        return jsonify({"historial": files})
    except Exception as e:
        return jsonify({"error": "No se pudo obtener el historial de los documentos", "details": str(e)}), 500

# =================================================================================
# Endpoint: /clean_historial
# Objetivo: Eliminar todos los archivos presentes en el directorio del historial.
# - Se invoca mediante una solicitud POST.
# =================================================================================
@app.route('/clean_historial', methods=['POST'])
def clean_historial():
    try:
        for f in os.listdir(HISTORIAL_DIR):
            file_path = os.path.join(HISTORIAL_DIR, f)
            os.remove(file_path)
        return jsonify({"message": "Historial limpiado existosamente."})
    except Exception as e:
        return jsonify({"error": "Error al limpiar el historial", "details": str(e)}), 500

# =================================================================================================================
# Endpoint: /generate
# Objetivo: Ejecutar el script generator.py para generar reportes.
# - Recibe datos JSON (por ejemplo, el mes a procesar).
# - Llama generator.py como proceso externo y captura su salida.
# - Analiza la salida para extraer información (como el nombre del ZIP generado) y construye una respuesta JSON.
# =================================================================================================================
@app.route('/generate', methods=['POST'])
def report_generator():
    try:
        data = request.get_json(silent=True) or {} # Se obtiene el JSON enviado en la solicitud, sin lanzar error si es inválido
        # Se obtiene el párametro 'month' del JSON de la solicitud. Si no se envía, se deja vacío para que generator.py use el mes actual por defecto
        month = data.get("month")
        if month is None:
            month = "" # Se puede dejar vacio, de modo que generator.py use el mes actual por defecto
        # Se ejecuta generator.py pasando el argumento "month" como un proceso externo y se captura su salida
        result = subprocess.run(['python', 'generator.py', str(month)], capture_output=True, text=True)

        # Si la ejecución fue exitosa (código de salida 0), se procesa la salida para extraer los datos relevantes
        if result.returncode == 0:
            # Se divide la salida en líneas para buscar información específica
            output_lines = result.stdout.splitlines()
            zip_path = None
            summary = {}

            for line in output_lines:
                if line.startswith("Mes procesado:"):
                    summary["month"] = line.split(":", 1)[1].strip()
                elif line.startswith("Total de documentos generados:"):
                    summary["total_docs"] = line.split(":", 1)[1].strip()
                elif line.startswith("Cursos sin participantes:"):
                    summary["courses_without_participants"] = line.split(":", 1)[1].strip()
                elif line.startswith("ZIP generado:"):
                    zip_path = line.split(":", 1)[1].strip()
                    summary["zip_path"] = zip_path
            
            # Si se encontró el ZIP y el archivo existe, se construye la URL de descarga y se devuelve el resumen
            if zip_path and os.path.exists(zip_path):
                # Se construye el URL de descarga
                download_url = f"/download_zip?file={os.path.basename(zip_path)}"
                summary["download_url"] = download_url
                return jsonify({
                    "message": "Reportes generados correctamente",
                    "summary": summary
                })

            return jsonify({"error": "Archivo ZIP no encontrado. Revise el formato del archivo Excel."}), 500
        # Si hubo algún error en la ejecución, se devuelve un mensaje de error con los detalles de la salida del error
        else:
            return jsonify({"error": "Error al generar los reportes", "details": result.stderr}), 500 # 500 es un código HTTP que indica un error interno en el servidor
    # Captura errores inesperados y los devuelve como respuesta en formato JSON
    except Exception as e:
        return jsonify({"error": "Error inesperado", "details": str(e)}), 500    

# Punto de entrada de la app Flask, se inicia la aplicación en modo depuración en el puerto 5000
if __name__ == '__main__':
    app.run(debug=True, port=5000)