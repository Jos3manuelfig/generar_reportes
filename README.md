# Generador de Reportes Formateados

Sistema simplificado para generar reportes Excel con formato específico.

## 📋 Requisitos

```bash
pip install pandas openpyxl
```

## 🚀 Uso

### Comando básico:
```bash
python generar_reporte.py "archivo.xlsx"
```

### Con filtros:
```bash
python generar_reporte.py "archivo.xlsx" [fecha_inicio] [fecha_fin] [distrito] [cuadrilla]
```

## 📋 Ejemplos de uso:

### 1. Todos los datos (sin filtros):
```bash
python generar_reporte.py "datos.xlsx"
```

### 2. Un día específico:
```bash
python generar_reporte.py "datos.xlsx" 2025-07-02 2025-07-02
```

### 3. Rango de fechas:
```bash
python generar_reporte.py "datos.xlsx" 2025-06-30 2025-07-02
```

### 4. Con distrito específico:
```bash
python generar_reporte.py "datos.xlsx" 2025-07-02 2025-07-02 "SAN MIGUEL"
```

### 5. Con distrito y cuadrilla:
```bash
python generar_reporte.py "datos.xlsx" 2025-07-02 2025-07-02 "SAN MIGUEL" "CU2"
```

## 📅 Formato de fecha:
- **Formato requerido:** `YYYY-MM-DD`
- **Ejemplos:** `2025-06-30`, `2025-07-01`, `2025-07-02`

## 📊 Columnas del reporte generado:

1. CODIGO DE POSTE CAMPO
2. DISTRITO UBICACION
3. LATITUD Y
4. LONGITUD X
5. PROPIETARIO
6. FECHA (formato DD-MM-YYYY)
7. EMPRESA EJECUTORA
8. CUADRILLA
9. ¿EXISTE APOYO?
10. NUMERO DE CABLES
11. ¿SE REVISO EN CAMPO?
12. ¿TRABAJO EJECUTADO?
13. ¿CUMPLE EL DMS?
14. OBSERVACIONES Y/O COMENTARIOS DEL TRABAJO O INSPECCIÓN

## 💡 Características:

- ✅ **Automático:** Detecta si el archivo necesita encabezados
- ✅ **Flexible:** Filtros por fecha, distrito y cuadrilla
- ✅ **Formato exacto:** Genera el reporte con las columnas específicas
- ✅ **Excel optimizado:** Ajusta automáticamente el ancho de columnas
- ✅ **Timestamp:** Cada reporte tiene fecha y hora única

## 📁 Archivos generados:

- `reportes/reporte_formateado_YYYYMMDD_HHMMSS.xlsx` - Reporte final
- `archivo_encabezados.xlsx` - Archivo original con encabezados (si es necesario)

## 🎯 Casos de uso comunes:

### Para el 2 de julio:
```bash
python generar_reporte.py "prueba2_encabezados.xlsx" 2025-07-02 2025-07-02
```

### Para San Miguel:
```bash
python generar_reporte.py "prueba2_encabezados.xlsx" 2025-07-02 2025-07-02 "SAN MIGUEL"
```

### Para cuadrilla CU2:
```bash
python generar_reporte.py "prueba2_encabezados.xlsx" 2025-07-02 2025-07-02 "" "" "CU2"
```
