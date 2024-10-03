from os.path import abspath, dirname, join, splitext
from os import makedirs, walk
from playwright.sync_api import sync_playwright, ElementHandle
import pandas as pd
from lxml import html
import re
import xlwings as xw

python_ubicacion = abspath(dirname(__file__))
carpeta_html = join(python_ubicacion, "html")
carpeta_imagen = join(python_ubicacion, "imagen")

makedirs(carpeta_html, exist_ok=True)
makedirs(carpeta_imagen, exist_ok=True)

def exportar_imagen(producto: ElementHandle, nombre: str):
	elem = producto.query_selector("//descendant::img[@class='item_image']")
	ruta_imagen = join(carpeta_imagen, f"{nombre}.png")
	elem.screenshot(path=ruta_imagen)

def exportar_html(html_text, nombre):
	ruta_elem = join(carpeta_html, f"{nombre}.html")
	with open(ruta_elem, 'w', encoding='utf8') as f:
		f.write(str(html_text))

def extraer():
	with sync_playwright() as p:
		browser = p.chromium.launch(headless=False)
		context = browser.new_context()
		page = context.new_page()

		link = 'https://gpperu.com/'
		page.goto(link, wait_until='load')
		page.wait_for_timeout(1000)

		xpath_categoria = "//select[@id='select_category']"
		xpath_options = "//option[contains(text(),'Categorías')]/following-sibling::option"
		xpath_marcas = "//li[contains(text(), 'Marca')]/following-sibling::li/a"
		xpath_producto = "//img[contains(@alt, 'Oferta')]/ancestor::div[@class='item']"

		elem_options = page.query_selector_all(xpath_options)
		value_options = [i.get_attribute("value") for i in elem_options]
  
		for value in value_options[7:10]:
			page.select_option(selector=xpath_categoria, value=value)
			page.wait_for_timeout(1000)
   
			elem_marcas = page.query_selector_all(xpath_marcas)
			marcas = [
	   			{
			  		"link": i.get_attribute("href"),
					"marca": i.inner_text().split()[0]
		   		} for i in elem_marcas
		  	]

			for data_marca in marcas:
				marca = data_marca["marca"]
				link = data_marca["link"]
	
				new_page = context.new_page()
				new_page.goto(link, wait_until='load')
				new_page.wait_for_timeout(1000)
	
				productos = new_page.query_selector_all(xpath_producto)
				for index, producto in enumerate(productos):
					html_text = producto.inner_html()

					nombre = f"{value}_{marca}_{index + 1}"
					exportar_imagen(producto, nombre)
					exportar_html(html_text, nombre)


		page.close()
		browser.close()

def acumular():
	data_acumulada = []
	
	for root, _, files in walk(carpeta_html):
		for file in files:
			name, _ = splitext(file)
			ruta_archivo = join(root, file)
			with open(ruta_archivo, "r", encoding='utf8') as f:
				html_text = f.read()
			respuesta = procesar_archivo(html_text, name)
			data_acumulada.append(respuesta)
   
	df = pd.DataFrame(data_acumulada)
	ruta_respuesta = join(python_ubicacion, "respuesta.xlsx")

	presentar(df, ruta_respuesta)

def exportar_excel(df, ruta_respuesta):
    df["precio"] = pd.to_numeric(df["precio"], errors='coerce')
    df["precio_promo"] = pd.to_numeric(df["precio_promo"], errors='coerce')
    df["imagen"] = ''

    with pd.ExcelWriter(ruta_respuesta, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)

        worksheet = writer.sheets['Sheet1']
       
        worksheet.set_column('A:A', 10)
        worksheet.set_column('B:B', 20)
        worksheet.set_column('C:C', 40)
        worksheet.set_column('D:D', 40)
        worksheet.set_column('E:E', 15)
        worksheet.set_column('F:F', 15)
        worksheet.set_column('G:G', 30)
        
        worksheet.freeze_panes(1, 0)
        
        for i, _ in enumerate(df.index, start=1):
            worksheet.set_row(i, 80)

def insertar_imagen_excel(ruta_respuesta):
    wb = xw.Book(ruta_respuesta)
    sheet = wb.sheets.active
    image_values = sheet.range('G2').expand('down').value
    
    for image_index, image_path in enumerate(image_values):
        current_cell = sheet.range("H2").offset(row_offset=image_index, column_offset=0)
        
        cell_width = current_cell.width
        cell_height = current_cell.height
        left = current_cell.left + (cell_width - 50) / 2
        top = current_cell.top + (cell_height - 80) / 2
        
        picture = sheet.pictures.add(
            image_path,
            left=left,
            top=top,
            width=50,
            height=80,
        )
        picture.api.Placement = 1
    
    sheet.range('G:G').api.EntireColumn.Hidden = True
        
    wb.save()
    wb.close()

def presentar(df, ruta_respuesta):
	exportar_excel(df, ruta_respuesta)
	insertar_imagen_excel(ruta_respuesta)

def procesar_archivo(html_text, name):
	html_tree = html.fromstring(html_text)
	dato_marca = html_tree.xpath("//span[@class='item_brand']/text()")[0]
	dato_sub_categoria = html_tree.xpath("//p[@class='item_sub_category']/text()")[0]
	dato_nombre = html_tree.xpath("//h2[@class='item_name']/text()")[0]
	dato_descripcion = html_tree.xpath("//p[@class='item_description']/text()")[0]
	dato_precio = html_tree.xpath("//span[@class='item_price' and position()=2]/text()")[0]
	dato_precio_promo = html_tree.xpath("//span[contains(@class,'promo')]/text()")[0]

	nombre_tratado = dato_nombre.replace("\n", "").replace("\t", "")
	descripcion_tratado = dato_descripcion.replace("\n", "").replace("\t", "")
 
	patron = r'\d+[.,]?\d*'
	precio_tratado = re.findall(patron, dato_precio)[0].replace(",", "")
	precio_promo_tratado = re.findall(patron, dato_precio_promo)[0].replace(",", "")

	return {
		"marca": dato_marca,
		"sub_categoria": dato_sub_categoria,
		"nombre": nombre_tratado,
		"descripcion": descripcion_tratado,
		"precio": precio_tratado,
		"precio_promo": precio_promo_tratado,
		"path_imagen": join(carpeta_imagen, f"{name}.png")
	}

def main():
	extraer()
	acumular()


if __name__ == '__main__':
	main()












