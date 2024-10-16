from flask import Flask, request, send_file, render_template
from reportlab.lib.pagesizes import inch
from reportlab.pdfgen import canvas
import pandas as pd
import os
import re
import zipfile
from datetime import datetime

app = Flask(__name__)

# Ruta del logo en la misma carpeta del código
LOGO_PATH = "logo.png"

# Función para generar el PDF
def generar_pdf(datos, productos, output_filename):
    # Tamaño del PDF: 6 pulgadas de alto por 4 de ancho
    width, height = 4 * inch, 6 * inch

    # Crear el PDF
    c = canvas.Canvas(output_filename, pagesize=(width, height))

    # Datos del PDF
    nombre = datos['Cliente']
    fecha = datos['Fecha A Entregar']
    metodo_pago = datos['Metodo De Pago']
    calle = datos['Dirección de entrega/Calle']
    calle2 = datos['Dirección de entrega/Calle2']
    total = datos['Total']

    # Establecer márgenes reducidos
    margin = 0.3 * inch
    y_position = height - margin  # Posición inicial de Y

    # Blanco y negro
    c.setStrokeColorRGB(0, 0, 0)
    c.setFillColorRGB(0, 0, 0)

    # Logo en la esquina superior derecha
    if os.path.exists(LOGO_PATH):
        c.drawImage(LOGO_PATH, width - 0.85 * inch, height - 0.85 * inch, width=0.8 * inch, height=0.8 * inch, mask='auto')

    # Ajustar el tamaño del rectángulo negro al largo del nombre del cliente
    text_width = c.stringWidth(nombre, "Helvetica-Bold", 16) + 10
    c.setFillColorRGB(0, 0, 0)  # Fondo negro
    c.rect(margin, y_position - 25, text_width, 20, fill=1)  # Dibujar el rectángulo negro ajustado al nombre
    c.setFillColorRGB(1, 1, 1)  # Texto blanco
    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin + 5, y_position - 20, nombre)
    y_position -= 40

    # Fecha de entrega en negrita
    c.setFont("Helvetica-Bold", 9)
    c.setFillColorRGB(0, 0, 0)  # Volver al texto negro
    c.drawString(margin, y_position, "Fecha de entrega:")
    c.setFont("Helvetica", 9)
    c.drawString(margin + 80, y_position, f"{fecha}")
    y_position -= 15

    # Dirección en negrita
    c.setFont("Helvetica-Bold", 9)
    c.drawString(margin, y_position, "Dirección:")
    c.setFont("Helvetica", 9)
    c.drawString(margin + 60, y_position, f"{calle}")
    if pd.notna(calle2):
        y_position -= 12
        c.drawString(margin + 60, y_position, f"{calle2}")
    y_position -= 15

    # Método de pago en negrita
    c.setFont("Helvetica-Bold", 9)
    c.drawString(margin, y_position, "Método de pago:")
    c.setFont("Helvetica", 9)
    c.drawString(margin + 80, y_position, f"{metodo_pago}")
    y_position -= 20

    # Tabla de productos
    c.setFont("Helvetica-Bold", 9)
    c.drawString(margin, y_position, "Producto")
    c.drawString(margin + 115, y_position, "Unidad")
    c.drawString(margin + 170, y_position, "Cantidad")
    c.drawString(margin + 220, y_position, "Subtotal")
    y_position -= 15

    square_size = 5  # Tamaño del cuadrado
    c.setFont("Helvetica", 7)

    for producto, cantidad, subtotal in productos:
        # Validar que el producto no sea NaN
        if pd.isna(producto):
            continue

        # Dibujar cuadrado
        c.rect(margin - 10, y_position, square_size, square_size, stroke=1, fill=0)

        producto, unidad = extraer_producto_unidad(producto)
        c.drawString(margin, y_position, producto)  # Mostrar "(Trans)" dentro del nombre del producto
        if unidad:
            c.drawString(margin + 120, y_position, unidad)
        c.drawString(margin + 188, y_position, str(cantidad))
        # Si el subtotal es igual a NaN cambiarlo a 0
        if pd.isna(subtotal):
            subtotal = 0
        c.drawString(margin + 230, y_position, f"$ {subtotal * cantidad}")

        # Dibujar una línea fina debajo de cada producto
        c.setLineWidth(0.5)
        c.line(margin, y_position - 5, width - 0.1 * inch, y_position - 5)

        y_position -= 15

    # Total
    y_position -= 10
    c.setFont("Helvetica-Bold", 10)
    if pd.isna(total):
            total = 0
    c.drawString(margin + 185, y_position, f"Total: $ {total}")
    y_position -= 30

    # Guardar PDF
    c.showPage()
    c.save()

# Función para extraer producto y unidad y reemplazar "En Transición" por "(Trans)"
def extraer_producto_unidad(producto):
    # Asegurarse de que el producto no sea NaN (tipo float)
    if not isinstance(producto, str):
        return "Producto no especificado", "1 unidad"

    # Reemplazar "En Transición" por "(Trans)" dentro del nombre del producto
    producto = producto.replace("En Transición", "")

    # Caso especial para "Huevos Pastoreo Libre"
    if "Huevos Pastoreo Libre" in producto:
        match = re.search(r"\((\d+)/\d+/(\d+)\).*?Maple \((\d+)", producto)
        if match:
            unidades = match.group(3) + " Unidades"
            return "Huevos Pastoreo Libre", unidades

    # Caso general
    match = re.search(r"(.+?) \((.*?)\)$", producto)
    if match:
        producto = match.group(1)  # Nombre del producto
        unidad = match.group(2)  # Unidad (ej: "1 Kilo")
        return producto, unidad

    # Si no hay paréntesis, retornar "1 unidad" por defecto
    return producto, "1 unidad"

# Función para leer Excel y detectar clientes
def leer_excel(file_path):
    df = pd.read_excel(file_path)
    
    datos_list = []
    cliente_actual = None
    productos = []

    # Iterar sobre cada fila del archivo Excel
    for i, row in df.iterrows():
        if pd.notna(row['Cliente']):
            # Si es un nuevo cliente, procesar el anterior
            if cliente_actual is not None:
                datos_list.append((cliente_actual, productos))
                productos = []

            # Capturar los datos del nuevo cliente
            cliente_actual = {
                'Cliente': row['Cliente'],
                'Fecha A Entregar': row['Fecha A Entregar'],
                'Metodo De Pago': row['Metodo De Pago'],
                'Dirección de entrega/Calle': row['Dirección de entrega/Calle'],
                'Dirección de entrega/Calle2': row['Dirección de entrega/Calle2'],
                'Total': row['Total']
            }
        else:
            # Si no hay un nombre de cliente, asumir que es una línea de producto
            productos.append((row['Líneas del pedido/Producto'], row['Líneas del pedido/Cantidad'], row['Líneas del pedido/Subtotal']))

    # Añadir el último cliente procesado
    if cliente_actual is not None:
        datos_list.append((cliente_actual, productos))

    return datos_list

# Ruta principal para cargar archivo y generar ZIP con PDFs
@app.route("/", methods=["GET", "POST"])
def upload_file():
    if request.method == "POST":
        if 'file' not in request.files:
            return "No se subió un archivo"
        
        file = request.files['file']
        if file.filename == '':
            return "El archivo no tiene nombre"
        
        # Guardar archivo Excel temporalmente
        file_path = os.path.join(os.getcwd(), file.filename)
        file.save(file_path)

        # Leer archivo Excel
        datos_list = leer_excel(file_path)

        # Crear carpeta temporal para los PDFs
        pdf_folder = "temp_pdfs"
        os.makedirs(pdf_folder, exist_ok=True)

        # Generar un PDF por cada cliente
        for datos, productos in datos_list:
            output_filename = f"{datos['Cliente'].replace(' ', '_')}_Pedido.pdf"
            output_path = os.path.join(pdf_folder, output_filename)
            generar_pdf(datos, productos, output_path)

        # Crear un archivo ZIP con todos los PDFs
        fecha_entrega = datos_list[0][0]['Fecha A Entregar'].strftime('%Y-%m-%d')
        zip_filename = f"pedidos_{fecha_entrega}.zip"
        with zipfile.ZipFile(zip_filename, 'w') as zipf:
            for pdf_file in os.listdir(pdf_folder):
                zipf.write(os.path.join(pdf_folder, pdf_file), pdf_file)

        # Limpiar la carpeta temporal
        for pdf_file in os.listdir(pdf_folder):
            os.remove(os.path.join(pdf_folder, pdf_file))
        os.rmdir(pdf_folder)

        # Enviar el archivo ZIP como respuesta
        return send_file(zip_filename, as_attachment=True)

    # Formulario para cargar archivo
    return '''
    <!doctype html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Generar PDF desde Excel</title>
        <style>
            body {
                font-family: Arial, sans-serif;
                background-color: #f7f7f7;
                margin: 0;
                padding: 0;
                display: flex;
                justify-content: center;
                align-items: center;
                height: 100vh;
            }
            .container {
                background-color: white;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
            }
            h1 {
                margin-bottom: 20px;
            }
            form {
                display: flex;
                flex-direction: column;
            }
            input[type="file"] {
                margin-bottom: 20px;
            }
            input[type="submit"] {
                background-color: #4CAF50;
                color: white;
                padding: 10px;
                border: none;
                border-radius: 4px;
                cursor: pointer;
                font-size: 16px;
            }
            input[type="submit"]:hover {
                background-color: #45a049;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <h1>Subir archivo Excel</h1>
            <form method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept=".xlsx">
                <input type="submit" value="Generar PDF y Descargar ZIP">
            </form>
        </div>
    </body>
    </html>
    '''

if __name__ == "__main__":
    app.run(debug=True)
