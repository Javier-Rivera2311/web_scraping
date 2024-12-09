import requests
from bs4 import BeautifulSoup
import pandas as pd
from openpyxl import load_workbook

# URL de los productos
url = "https://listado.mercadolibre.cl/tinta-plotter-hp#D[A:tinta%20plotter%20hp]"

response = requests.get(url)

soup = BeautifulSoup(response.content, 'html.parser')

productos = soup.find_all('div', class_="poly-card__content")

nombres = []
precios = []
envios = []
urls = []

for producto in productos:
    titulo = producto.find('h2', class_="poly-box poly-component__title")
    nombre = titulo.text
    precio = producto.find('span', class_="andes-money-amount andes-money-amount--cents-superscript").text
    envio = producto.find('div', class_="poly-component__shipping")
    envio_texto = envio.text if envio else "Envío no especificado"
    enlace = titulo.find('a')['href']

    nombres.append(nombre)
    precios.append(precio)
    envios.append(envio_texto)
    urls.append(enlace)

    print(f"nombre: {nombre} | precio: {precio} ")

# Crear el DataFrame
df = pd.DataFrame({
    'Nombre': nombres,
    'Precio': precios,
    'Envio': envios,
    'URL': urls
})

# Convertir los precios a formato numérico
df['Precio'] = df['Precio'].replace({'\$': '', '\.': ''}, regex=True).astype(int)

# Ordenar por precios de menor a mayor
df = df.sort_values(by='Precio')

# Guardar el archivo Excel
archivo_excel = 'productos_mercadolibre.xlsx'
df.to_excel(archivo_excel, index=False)

#Ajustar el tamaño de las columnas en excel

# Ajustar el ancho de las columnas
workbook = load_workbook(archivo_excel)
worksheet = workbook.active

for column in worksheet.columns:
    max_length = 0
    column_letter = column[0].column_letter  # Obtener la letra de la columna
    for cell in column:
        try:
            if cell.value:  # Solo si la celda tiene valor
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass
    adjusted_width = max_length + 2  # Margen adicional
    worksheet.column_dimensions[column_letter].width = adjusted_width

# Guardar los ajustes
workbook.save(archivo_excel)

print("Su información fue guardada y ajustada.")