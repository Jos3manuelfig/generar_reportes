# 📊 Extractor de Datos Excel - Formulario WhatsApp

Sistema simple para extraer y procesar datos de archivos Excel locales con filtros por fecha.

## 🚀 Características

- ✅ Lectura directa de archivos Excel (.xlsx, .xls)
- 📅 Filtrado por rango de fechas
- 📊 Exportación a Excel con formato personalizado
- 🎨 Interfaz web intuitiva con Streamlit
- 📋 Mapeo automático de columnas
- 🔧 Script de línea de comandos

## 📋 Columnas Incluidas

El sistema extrae y organiza los siguientes datos:

| Columna | Descripción |
|---------|-------------|
| ID | Identificador único |
| CODIGO DE POSTE CAMPO | Código del poste en campo |
| DISTRITO UBICACION | Distrito de ubicación |
| LATITUD Y | Coordenada de latitud |
| LONGITUD X | Coordenada de longitud |
| PROPIETARIO | Propietario del poste |
| FECHA | Fecha de inspección/trabajo |
| EMPRESA EJECUTORA | Empresa que ejecuta el trabajo |
| CUADRILLA | Cuadrilla asignada |
| ¿EXISTE APOYO? | Verificación de apoyo |
| NUMERO DE CABLES | Cantidad de cables |
| ¿SE REVISO EN CAMPO? | Verificación de revisión en campo |
| ¿TRABAJO EJECUTADO? | Estado del trabajo |
| ¿CUMPLE EL DMS? | Cumplimiento del DMS |
| OBSERVACIONES | Comentarios del trabajo |
| REQUIERE MEJORA | Si requiere mejoras |
| GRAVEDAD DEL CASO | Nivel de gravedad |
| MEJORA WIN EMPRESAS | Mejoras WIN empresas |

## 🛠️ Instalación

### 1. Instalar dependencias

```bash
pip install -r requirements.txt
```

## 🚀 Uso

### Opción 1: Interfaz Web (Recomendado)

```bash
streamlit run app_excel.py
```

La aplicación se abrirá en tu navegador en `http://localhost:8501`

**Pasos:**
1. Sube tu archivo Excel en la barra lateral
2. Configura filtros de fecha (opcional)
3. Haz clic en "Extraer Datos"
4. Descarga el archivo Excel generado

### Opción 2: Línea de Comandos

```bash
python extract_excel.py "NORMALIZACION DE RED (respuestas).xlsx" 2024-06-11 2024-06-11
```

**Sintaxis:**
```bash
python extract_excel.py "ruta/archivo.xlsx" [fecha_inicio] [fecha_fin]
```

**Ejemplos:**
```bash
# Extraer todos los datos
python extract_excel.py "datos.xlsx"

# Extraer datos de hoy
python extract_excel.py "datos.xlsx" 2024-06-11 2024-06-11

# Extraer datos de un rango
python extract_excel.py "datos.xlsx" 2024-06-01 2024-06-30
```

## 📁 Estructura del Proyecto

```
ExtratordeDatos/
├── app_excel.py              # Aplicación web para Excel
├── excel_extractor.py        # Motor de extracción Excel
├── extract_excel.py          # Script de línea de comandos
├── config.py                 # Configuración
├── requirements.txt          # Dependencias
├── README_EXCEL.md           # Este archivo
└── reporte_postes_*.xlsx     # Archivos de salida (generados)
```

## 🔧 Personalización

### Modificar el mapeo de columnas

Edita el archivo `config.py`:

```python
OUTPUT_COLUMNS = [
    'ID',
    'CODIGO DE POSTE CAMPO',
    # ... más columnas
]
```

### Agregar nuevas columnas

1. Agrega la nueva columna a `OUTPUT_COLUMNS` en `config.py`
2. El sistema automáticamente buscará columnas similares en tu Excel

## 📝 Notas Importantes

- ✅ **Sin credenciales**: No necesitas configurar Google API
- 📁 **Archivos locales**: Trabaja directamente con archivos Excel
- 🔄 **Mapeo automático**: Detecta columnas similares automáticamente
- 💾 **Almacenamiento**: Los archivos Excel se guardan en la carpeta del proyecto

## 🤝 Solución de Problemas

### Error: "No se encontró columna de fecha"

**Solución**: Verifica que tu Excel tenga una columna que contenga "fecha" o "date" en el nombre.

### Error: "El archivo no existe"

**Solución**: Verifica la ruta del archivo Excel y asegúrate de que esté en la carpeta correcta.

### Error: "Formato de fecha debe ser YYYY-MM-DD"

**Solución**: Usa el formato correcto: `2024-06-11` (no `11-06-2024`).

---

**Desarrollado para extracción de datos de formularios WhatsApp 📱** 