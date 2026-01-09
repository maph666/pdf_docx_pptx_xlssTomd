# mdtox GUI

📁 **Archivos necesarios para su funcionamiento**

- `gui.py` — Interfaz gráfica (Tkinter): seleccionar archivo de entrada y generar `.md`.
- `converter.py` — Lógica principal de conversión (soporta PDF, DOCX, PPTX, XLSX → Markdown).
- `api.py` — Interfaz de línea de comandos (usa `markitdown` bajo el capó).
- `app.py` — Punto de entrada alternativo / lanzador de la aplicación.
- `requirements.txt` — Dependencias necesarias para ejecutar las conversiones.
- `tests/` — Scripts de prueba y ejemplos (`tests/run_input_to_md_tests.py`, `tests/sample.md`).
- `README.md` — Documentación del proyecto (este archivo).
- `ven/` — (Virtualenv local; no se guarda en VCS normalmente) contenedor del entorno virtual usado en desarrollo.

Aplicación sencilla para convertir documentos a Markdown usando `markitdown`.

Compatibilidades añadidas

- Ahora soporta como entrada: **PDF**, **Word (.docx)**, **PowerPoint (.pptx)** y **Excel (.xlsx**) y convierte a **Markdown (.md)**.

Requisitos

- Python 3.8+
- Crear y activar un virtualenv

Instalación

1. python -m venv .venv
2. .\.venv\Scripts\activate
3. pip install -r requirements.txt

Nota: Si prefieres instalar manualmente las dependencias necesarias para las conversiones, instala:

```
pip install python-docx html2docx python-pptx openpyxl markdown
```

Uso (GUI)

1. Ejecuta `python gui.py`.
2. En "Input file" selecciona un archivo **.pdf**, **.docx**, **.pptx** o **.xlsx**.
3. El campo "Output .md" se rellena por defecto con el mismo nombre y la extensión `.md` (puedes cambiar la ruta de salida).
4. Pulsa "Convert" para generar el archivo Markdown.

Uso (línea de comandos)

- Convertir un PDF a Markdown:

```
python api.py input.pdf -o output.md
```

- Convertir un Word a Markdown:

```
python api.py input.docx -o output.md
```

Pruebas

- Se incluyen pruebas básicas en `tests/run_input_to_md_tests.py`. Puedes ejecutarlas después de instalar dependencias:

```
ven\Scripts\python.exe tests\run_input_to_md_tests.py
```

Notas y limitaciones

- La conversión intenta extraer texto, títulos y tablas de forma simple. No conserva estilos complejos ni formateo avanzado.
- Para PowerPoint: cada diapositiva se convierte a un bloque; el título de la diapositiva se convierte en un encabezado H1 en el Markdown.
- Para Excel: cada hoja se exporta como una tabla Markdown (encabezado = primera fila).

Contribuciones

- Pull requests bienvenidas. Sugiero agregar tests adicionales y mejorar el manejo de listas y estilos en DOCX/PPTX.

Licencia

- MIT

