"""
Lector de facturas en PDF.

Recorre una carpeta con PDFs, busca en cada archivo los textos clave
definidos en CAMPOS y extrae el número que aparece en esa línea.
Guarda el resultado en un Excel con una fila por archivo.

Uso:
    pip install -r requirements.txt
    python lector_facturas.py
"""
import os
import re

import pandas as pd
import pdfplumber

CARPETA_PDFS = "pdfs"
ARCHIVO_SALIDA = "resultado_facturas.xlsx"

# Nombre de la columna en Excel -> texto a buscar y qué número de la línea tomar
# (0 = primer número de la línea, 1 = segundo, etc.)
CAMPOS = {
    "Campo A": {"texto": "Texto clave A", "numero": 0},
    "Campo B": {"texto": "Texto clave B", "numero": 0},
    "Campo C": {"texto": "Texto clave C", "numero": 1},
}

PATRON_NUMERO = r"\d+(?:,\d{3})*(?:\.\d{2})?"

# Texto cuyas líneas se suman para la columna "Total"
TEXTO_SUMA = "Texto clave de suma"


def extraer_valor(texto_pdf, texto_buscado, posicion):
    """Devuelve el número en la posición indicada de la primera línea que contiene el texto."""
    for linea in texto_pdf.split("\n"):
        if texto_buscado.lower() in linea.lower():
            numeros = re.findall(PATRON_NUMERO, linea)
            if len(numeros) > posicion:
                return numeros[posicion]
    return ""


def extraer_total(texto_pdf, texto_buscado):
    """Suma todos los números de las líneas que contienen el texto."""
    total = 0.0
    for linea in texto_pdf.split("\n"):
        if texto_buscado.lower() in linea.lower():
            numeros = re.findall(PATRON_NUMERO, linea)
            if numeros:
                total += float(numeros[-1].replace(",", ""))
    return f"{total:.2f}"


def leer_texto_pdf(ruta):
    texto = ""
    with pdfplumber.open(ruta) as pdf:
        for pagina in pdf.pages:
            contenido = pagina.extract_text()
            if contenido:
                texto += contenido + "\n"
    return texto


def main():
    resultados = []
    for archivo in sorted(os.listdir(CARPETA_PDFS)):
        if not archivo.lower().endswith(".pdf"):
            continue
        texto = leer_texto_pdf(os.path.join(CARPETA_PDFS, archivo))
        fila = {"archivo": archivo}
        for columna, config in CAMPOS.items():
            fila[columna] = extraer_valor(texto, config["texto"], config["numero"])
        fila["Total"] = extraer_total(texto, TEXTO_SUMA)
        resultados.append(fila)

    df = pd.DataFrame(resultados)
    df.to_excel(ARCHIVO_SALIDA, index=False)
    print(f"Listo: {len(resultados)} archivos procesados -> {ARCHIVO_SALIDA}")


if __name__ == "__main__":
    main()
