# Lector de facturas en PDF

Script en Python que lee una carpeta de PDFs, busca textos clave en cada archivo y guarda los valores encontrados en un Excel, con una fila por archivo.

## Qué hace
- Recorre la carpeta `pdfs/` y procesa todos los PDF.
- Busca líneas que contengan un texto clave y extrae el número que corresponde.
- Suma las líneas de un texto clave para calcular un total.
- Genera `resultado_facturas.xlsx`.

## Cómo usarlo
1. Instalá las dependencias: `pip install -r requirements.txt`
2. Creá una carpeta `pdfs/` junto al script y copiá ahí los PDF.
3. Editá en `lector_facturas.py` el diccionario `CAMPOS` con los textos clave de tus documentos.
4. Ejecutá: `python lector_facturas.py`

## Limitaciones
- Los números se leen en formato `1,234.50` (coma para miles, punto para decimales). Si tus documentos usan `1.234,50`, hay que ajustar `PATRON_NUMERO`.
- Depende de que el texto del PDF sea seleccionable. Un PDF escaneado como imagen no se puede leer con este script.
- Los textos clave deben coincidir con el contenido real del documento.

## Tecnologías
Python, pdfplumber, pandas, openpyxl.
