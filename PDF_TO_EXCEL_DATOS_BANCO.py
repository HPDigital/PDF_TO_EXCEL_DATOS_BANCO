"""
PDF_TO_EXCEL_DATOS_BANCO
"""

#!/usr/bin/env python
# coding: utf-8

# In[2]:


import fitz  # PyMuPDF
import pandas as pd
import re

# Ruta al archivo PDF


# Función para extraer texto de cada página del PDF usando PyMuPDF
def extract_text_from_pdf(pdf_path):
    document = fitz.open(pdf_path)
    text = ""
    for page_num in range(len(document)):
        page = document.load_page(page_num)
        text += page.get_text()
    return text



def orden_valores(pdf_text, path_out):
# Dividir el texto en líneas
    lines = pdf_text.split('\n')

    # Filtrar las líneas relevantes
    keywords = [
        'INGRESOS OPERACIONALES', 'COSTOS', 'RESULTADO FINACIERO BRUTO', 
        'Gastos de Comercialización', 'Gastos Administrativos', 'EGRESOS OPERACIONALES', 
        'RESULTADO OPERATIVO', 'Otros Ingresos', 'Rendimiento por Inversiones', 
        'INGRESOS NO OPERACIONALES', 'Cargos por Diferencia de Cambio', 'Otros Egresos', 
        'Ajuste por inflación y tenencia de bienes', 'EGRESOS NO OPERACIONALES', 
        'RESULTADO NO OPERACIONAL', 'RESULTADO NETO DESPUES DE OPERACIONES', 
        'Ingresos de Gestiones Anteriores', 'Gastos de Gestiones Anteriores', 
        'RESULTADO DE GESTIONES ANTERIORES', 'Ingresos Extraordinarios', 
        'Gastos Extraordinarios', 'RESULTADO EXTRAORDINARIO', 
        'RESULTADO DE OPERACION NETO', 'Gastos Financieros', 
        'UTILIDAD ANTES DE IMPUESTOS', 'Impuesto a las Utilidades de las Empresas', 
        'UTILIDAD NETA DE LA GESTION'
    ]

    relevant_lines = []
    for line in lines:
        if any(keyword in line for keyword in keywords):
            relevant_lines.append(line)

    # Crear un DataFrame a partir de las líneas filtradas
    data = {'Descripción': [], 'Monto': []}
    i = 0
    while i < len(lines):
        line = lines[i]
        if any(keyword in line for keyword in keywords):
            description = line.strip()
            i += 1
            # Verificar si la siguiente línea contiene el monto
            if i < len(lines) and re.match(r'-?\d[\d,]*$', lines[i].strip()):
                amount = lines[i].strip().replace(',', '')
                i += 1
            else:
                amount = ''
            data['Descripción'].append(description)
            data['Monto'].append(amount)
        else:
            i += 1


    df = pd.DataFrame(data)
    # Guardar el DataFrame en un archivo Excel
    excel_path = path_out
    df.to_excel(excel_path, index=False)
    print(f"\nDatos extraídos y guardados en {excel_path}")
    return df




direcciones = ["C:\\Users\\HP\\Downloads\\202312_CMI_EEFF_ER.pdf"]
# path_out = r"C:\Users\HP\Downloads\202312_CMI_EEFF_ER.xlsx"


def transformar(pdf_path,path_out):
    pdf_text = extract_text_from_pdf(pdf_path)
    data= orden_valores(pdf_text, path_out)
    print(":)")


for direc_in in direcciones :
    direc_out = direc_in.replace("pdf","xlsx")
    transformar(direc_in, direc_out)



# In[ ]:






if __name__ == "__main__":
    pass
