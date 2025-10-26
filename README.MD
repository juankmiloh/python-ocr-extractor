
# Proyecto OCR TRM

Este proyecto permite extraer automáticamente el número de factura y la tasa de cambio (TRM) desde archivos PDF utilizando OCR, sin necesidad de tener Python instalado en el sistema.

---

## ✅ Requisitos

- Windows (con acceso a PowerShell o CMD)
- Acceso a carpetas locales, pero **sin necesidad de permisos de administrador**
- Python embebido (`python-3.13.9-embed-amd64`)
- Tesseract OCR
- Poppler para Windows

---

## 📁 Estructura recomendada

```
TRM_PYTHON/
├── main.py                     ← Script principal OCR
├── requirements.txt            ← Lista de dependencias
├── get_trm.bat                 ← Script para ejecutar todo automáticamente
├── venv/                       ← Entorno virtual creado previamente
├── tesseract/
│   ├── tesseract.exe
│   └── tessdata/
│       └── spa.traineddata     ← Soporte para idioma español
├── poppler/
│   └── Library/bin/pdftoppm.exe
└── PWC1/                       ← Carpeta con PDFs a procesar (Ubica todos los PDF que contengan en su nombre de archivo "Declaracion Importacion")
```

---

## ⚙️ Pasos realizados

### 1. Descargar Python embebido

- Desde [python.org → embedded zip](https://www.python.org/downloads/windows/)
- Extraído en:  
  `C:\Users\JUANK\Desktop\TRM_PYTHON\python-3.13.9-embed-amd64`

### 2. Activar compatibilidad de librerías en Python embebido

- Editar `python313._pth`:
  - Descomentar la línea: `import site`
  - Asegurarse de que incluya `Lib` y `Scripts`

---

### 3. Instalar pip

```cmd
C:\Users\JUANK\Desktop\TRM_PYTHON\python-3.13.9-embed-amd64\python.exe get-pip.py
```

---

### 4. Crear entorno virtual

```cmd
C:\Users\JUANK\Desktop\TRM_PYTHON\python-3.13.9-embed-amd64\python.exe -m virtualenv venv
```

---

### 5. Instalar dependencias

```cmd
call venv\Scripts\activate.bat
pip install -r requirements.txt
```

---

### 6. Descargar e integrar Poppler

- Descargar desde: [https://github.com/oschwartz10612/poppler-windows/releases](https://github.com/oschwartz10612/poppler-windows/releases)
- Extraer `pdftoppm.exe` en:  
  `C:\Users\JUANK\Desktop\TRM_PYTHON\poppler\Library\bin`

---

### 7. Descargar e integrar Tesseract

- Descargar desde: [https://github.com/UB-Mannheim/tesseract/wiki](https://github.com/UB-Mannheim/tesseract/wiki)
- Instalar en:  
  `C:\Users\JUANK\Desktop\TRM_PYTHON\tesseract\`
- Crear carpeta `tessdata` dentro y colocar `spa.traineddata` descargado desde:  
  [https://github.com/tesseract-ocr/tessdata](https://github.com/tesseract-ocr/tessdata)

---

### 8. Modificaciones en `main.py`

Agregar al inicio del archivo:

```python
import os
import pytesseract

# Configurar ruta a tesseract
base_dir = os.path.dirname(os.path.abspath(__file__))
pytesseract.pytesseract.tesseract_cmd = os.path.join(base_dir, "Tesseract-OCR", "tesseract.exe")
os.environ["TESSDATA_PREFIX"] = os.path.join(base_dir, "Tesseract-OCR", "tessdata")
```

Y dentro del bucle de procesamiento:

```python
gray_image = ImageOps.grayscale(image)
text = pytesseract.image_to_string(gray_image, lang='spa')
```

También debes pasar `poppler_path`:

```python
pages = convert_from_path(pdf_path, poppler_path=os.path.join(base_dir, "poppler", "Library", "bin"))
```

---

### 9. Script de ejecución automática (`get_trm.bat`)
```cmd
.\get_trm.bat
```

---

## ✅ Resultado esperado

Al ejecutar `get_trm.bat`, el sistema:

- Solicita la ruta al usuario. NANA Por favor, ingresa la ruta de la carpeta a procesar: Ej. C:\Users\JUANK\Desktop\TRM_PYTHON\PWC1
- Convierte cada PDF en imágenes (con Poppler)
- Aplica OCR en español (con Tesseract)
- Extrae número de factura y TRM
- Guarda los resultados en un Excel con timestamp

---

## 📌 Observaciones

- No se necesita instalación de Python en el sistema
- No requiere variables de entorno (`PATH`) globales
- Ideal para entornos corporativos restringidos

---

## EJECUCION MANUAL SI YA SE TIENE EL ENTORNO VIRTUAL CREADO
.\venv\Scripts\activate
pip install -r requirements.txt
python main.py --ruta  "C:\Users\JUANK\Desktop\TRM_PYTHON\PWC1"