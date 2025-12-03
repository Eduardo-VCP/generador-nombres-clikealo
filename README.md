# Generador de Nombres de Productos

Sistema automatizado para generar nombres descriptivos y estructurados de productos consultando MongoDB y aplicando transformaciones según esquemas configurables. El sistema procesa múltiples categorías de productos y genera archivos Excel con nombres formateados y detección de datos faltantes.


**Características principales:**
- ✅ Generación automática de nombres de productos basados en especificaciones técnicas
- ✅ Soporte para múltiples categorías de productos 
- ✅ Exportación a archivos Excel con pestañas diferentes
- ✅ Detección y marcado visual de datos faltantes en Excel
- ✅ Procesamiento de múltiples schemas simultáneamente

## 🚀 Cómo Levantar el Proyecto

### Prerrequisitos

- **Python 3.7+** instalado
- **MongoDB** accesible (local o remoto)
- **Pymongo 4.0+** instalado - Cliente Python para conectarse a MongoDB y leer los productos
- **Openpyxl 3.0+** instalado - Librería para generar y formatear archivos Excel con los resultados

### Instalación de Dependencias

`requirements.txt`:

```txt
pymongo>=4.0.0
openpyxl>=3.0.0
```

E instalar con:

```bash
pip install -r requirements.txt
```
## 🚦 Quick Start

### 1. Clonar el Repositorio

```bash
git clone https://github.com/Eduardo-VCP/generador-nombres-clikealo.git
cd generador-nombres-clikealo
```

### 2. Configurar Conexión a MongoDB

Editar el archivo `generador-nombres.py` o `general_archive/generar-nombres.py` y configurar:

```python
MONGO_URI = 'MONGO_URI PARA LA BDD EN PRODUCCIO'  # URI de conexión a MongoDB
DB_NAME = 'BASE DE DATOS'  # o 'development' según el entorno
```

### 3. Ejecutar el Script

**Versión (múltiples schemas):**

```bash
python general_archive/generar-nombres.py
```

### 4. Verificar Salida

El archivo `productos_output.xlsx` se generará en el directorio actual.

## 📁 Estructura del Proyecto

```
generador-nombres-clikealo/
├── generador-nombres.py          # Script principal (versión simple, un schema)
├── general_archive/              # Versión avanzada con múltiples schemas
│   ├── generar-nombres.py        # Script que procesa múltiples schemas
│   └── schemas/                  # Esquemas JSON de configuración
│       ├── schemaAllInOne.json
│       ├── schemaImpresora.json
│       ├── schemaLaptop.json
│       ├── schemaMonitor.json
│       └── schemaMouseGamer.json
└── README.md
```

## ⚙️ Configuración

### Variables de Configuración

#### Versión Avanzada (`general_archive/generar-nombres.py`)

```python
MONGO_URI = 'MONGO_URI PARA LA BDD EN PRODUCCION'
DB_NAME = 'BASE DE DATOS'
# Los schemas se cargan automáticamente desde el directorio 'schemas/'
```
## 📤 Salida Excel

El archivo Excel generado incluye:

- **Encabezado verde**: Nombres de columnas con fondo verde y texto blanco
- **Columna SKU**: Identificador del producto
- **Columna Nombre Completo**: Nombre generado según el schema
- **Columnas individuales**: Una columna por cada campo definido en el schema
- **Colores indicativos**: Verde (completo) / Amarillo (faltantes)
- **Ancho automático**: Las columnas se ajustan automáticamente al contenido


## 🔗 Tecnologías

- **Python 3.7+**
- **PyMongo** - Cliente MongoDB para Python
- **OpenPyXL** - Generación y manipulación de archivos Excel
- **JSON** - Configuración de schemas
