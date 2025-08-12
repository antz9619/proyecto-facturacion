from flask import Flask, request, render_template, send_file
import pandas as pd
import locale
import os  # Importar el módulo os para manejar variables de entorno

# Configurar el formato de moneda para Argentina
try:
    locale.setlocale(locale.LC_ALL, 'es_AR.UTF-8')
except locale.Error:
    locale.setlocale(locale.LC_ALL, '')  # Usa la configuración predeterminada si 'es_AR.UTF-8' no está disponible

app = Flask(__name__)

# Configuración de tarifas
TARIFARIO = {
    "0 - 0.5": [5152.81, 6661.83, 6906.47, 7292.16, 5878.09],
    "0.5 - 1": [5236.35, 6765.25, 7009.81, 7395.62, 5969.35],
    "1 - 2": [5352.46, 6972.59, 7292.90, 8578.14, 6152.29],
    "2 - 5": [5694.33, 7754.53, 8501.36, 10647.87, 6842.24],
    "5 - 10": [7329.07, 9708.43, 12399.24, 15113.05, 8566.25],
    "10 - 15": [8802.00, 12361.42, 16602.02, 21342.60, 10907.14],
    "15 - 20": [10824.53, 15833.26, 22172.12, 29397.97, 13970.52],
    "20 - 25": [12017.31, 18277.13, 26328.19, 35827.09, 16126.88],
    "Excedente": [244.43, 517.86, 599.75, 1039.16, 456.94]
}

DESCRIPCION_ZONA = {
    1: "Local",
    2: "Regional",
    3: "Nacional 1",
    4: "Nacional 2",
    5: "Principales"
}

@app.route("/", methods=["GET"])
def formulario():
    return render_template("index.html")

def obtener_precio(rango, zona):
    return TARIFARIO.get(rango, [0]*5)[zona - 1] if 1 <= zona <= 5 else 0

def determinar_rango(peso):
    if peso <= 0.5: return "0 - 0.5"
    elif peso <= 1: return "0.5 - 1"
    elif peso <= 2: return "1 - 2"
    elif peso <= 5: return "2 - 5"
    elif peso <= 10: return "5 - 10"
    elif peso <= 15: return "10 - 15"
    elif peso <= 20: return "15 - 20"
    elif peso <= 25: return "20 - 25"
    return "Excedente"

@app.route("/procesar", methods=["POST"])
def procesar_archivo():
    if "archivo" not in request.files:
        return "Error: No se subió ningún archivo."
    
    archivo = request.files["archivo"]
    if archivo.filename == "":
        return "Error: Ningún archivo fue seleccionado."

    try:
        # Leer y normalizar datos
        df = pd.read_excel(archivo)
        df.columns = df.columns.str.strip()  # Normalizar nombres de columnas
        
        # Validar columnas requeridas
        columnas_requeridas = {
        'GramRea': 'GramRea',
        'GramAfo': 'GramAfo',  # Mantener esta columna para el cálculo del aforado
        'Zo': 'Zo',  
        'Precio Unitario': 'Precio Unitario',
        'G u i a': 'G u i a'  # Incluir la columna "Guia" para validar duplicados
        }
        
        # Renombrar columnas según necesidad
        for original, nuevo in columnas_requeridas.items():
            if original in df.columns:
                df.rename(columns={original: nuevo}, inplace=True)
        
        # Verificar columnas esenciales
        columnas_esenciales = ['GramRea','GramAfo', 'Zo', 'Precio Unitario', 'G u i a']
        for col in columnas_esenciales:
            if col not in df.columns:
                return f"Error: Falta la columna requerida: {col}"

        # Cálculos principales
        df['kilos PESO REAL'] = df['GramRea'] / 1000
        df['PESO VOLUMETRICO'] = df['GramAfo'] / 1000
        df['PESO A LIQUIDAR'] = df[['kilos PESO REAL', 'PESO VOLUMETRICO']].max(axis=1)
        
        # Determinar rangos y precios
        df['Rango'] = df['PESO A LIQUIDAR'].apply(determinar_rango)
        df['Descripcion'] = df['Zo'].map(DESCRIPCION_ZONA)
        
        # Cálculo de precios
        def calcular_precio(row):
            if row['PESO A LIQUIDAR'] > 25:
                excedente = (row['PESO A LIQUIDAR'] - 25) * obtener_precio("Excedente", row['Zo'])
                return obtener_precio("20 - 25", row['Zo']) + excedente
            return obtener_precio(row['Rango'], row['Zo'])
        
        df['RANGO PRECIO A COBRAR'] = df.apply(calcular_precio, axis=1)
        
        # Diferencias y revisiones
        df['DIFERENCIA'] = df['Precio Unitario'] - df['RANGO PRECIO A COBRAR']
        df['REVISIÓN'] = df['DIFERENCIA'].apply(lambda x: "ok" if x == 0 else "Validar con el courier el cobro")

        # Validar datos repetidos en la columna "Guia"
        df['Repetido'] = df['G u i a'].duplicated(keep=False)  # Marca duplicados como True

        # Resumir las Diferencias por Zonas y agregar recuento de duplicados
        resumen = df.groupby('Descripcion').agg(
            Total_Negativas=('DIFERENCIA', lambda x: x[x < 0].sum()),
            Total_Positivas=('DIFERENCIA', lambda x: x[x > 0].sum()),
            Total_OK=('REVISIÓN', lambda x: (x == "ok").sum()),
            Total_Repetidos=('G u i a', lambda x: x.duplicated(keep=False).sum())  # Contar duplicados en "Guia"
        ).reset_index()

        # Calcular totales y agregar la fila de totales al resumen
        totales = {
            'Descripcion': 'Total',
            'Total_Negativas': resumen['Total_Negativas'].sum(),
            'Total_Positivas': resumen['Total_Positivas'].sum(),
            'Total_OK': resumen['Total_OK'].sum(),
            'Total_Repetidos': resumen['Total_Repetidos'].sum()
        }
        resumen = pd.concat([resumen, pd.DataFrame([totales])], ignore_index=True)

        # Generar nombre de archivo basado en el original
        nombre_original = os.path.splitext(archivo.filename)[0]
        nombre_procesado = f"{nombre_original}_procesado.xlsx"

        # Guardar los datos procesados y el resumen en un único archivo Excel
        with pd.ExcelWriter(nombre_procesado, engine="openpyxl") as writer:
            df.to_excel(writer, sheet_name="Datos Procesados", index=False)
            resumen.to_excel(writer, sheet_name="Resumen", index=False)

        return f'''
     <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Archivo Procesado</title>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 0;
                height: 100vh;
                background: linear-gradient(135deg, #2E7DFF, #65C7F7);
                display: flex;
                justify-content: center;
                align-items: center;
            }}
            .container {{
                background: white;
                border-radius: 15px;
                box-shadow: 0 4px 10px rgba(0, 0, 0, 0.1);
                padding: 30px;
                text-align: center;
                width: 80%;
                max-width: 600px;
            }}
            h1 {{
                font-size: 2.5em;
                color: #333;
                margin-bottom: 10px;
            }}
            p {{
                font-size: 1.2em;
                color: #666;
                margin-bottom: 20px;
            }}
            a {{
                display: block;
                color: #4CAF50;
                text-decoration: none;
                margin: 10px 0;
                font-size: 1.2em;
            }}
            a:hover {{
                color: #45a049;
                text-decoration: underline;
            }}
            button {{
                background-color: #4CAF50;
                color: white;
                padding: 10px 20px;
                border: none;
                border-radius: 5px;
                font-size: 1.2em;
                cursor: pointer;
                margin-top: 20px;
                transition: background-color 0.3s, transform 0.2s;
            }}
            button:hover {{
                background-color: #45a049;
                transform: scale(1.05);
            }}
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Archivo procesado con éxito</h1>
            <p>Tu archivo ha sido procesado. Puedes descargar los resultados a continuación.</p>
            <a href="/descargar/{nombre_procesado}">Descargar archivo procesado</a>
            <button onclick="window.location.href='/'">Volver al formulario</button>
        </div>
    </body>
    </html>
'''

    except Exception as e:
        return f"Error procesando el archivo: {str(e)}"

@app.route("/descargar/<filename>")
def descargar(filename):
    try:
        return send_file(filename, as_attachment=True)
    except Exception as e:
        return f"Error al descargar: {str(e)}"

if __name__ == "__main__":
     app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))

