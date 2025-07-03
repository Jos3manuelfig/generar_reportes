#!/usr/bin/env python3
"""
Script simplificado para generar reportes formateados de Excel
Uso: python generar_reporte.py "archivo.xlsx" [fecha_inicio] [fecha_fin] [distrito] [cuadrilla]
"""

import pandas as pd
import sys
import os
from datetime import datetime

def agregar_encabezados(input_file, output_file=None):
    """
    Agrega encabezados a un archivo Excel que no los tiene
    """
    try:
        # Leer el archivo sin encabezados
        df = pd.read_excel(input_file, header=None)
        
        print(f"📊 Archivo leído: {input_file}")
        print(f"📋 Filas: {len(df)}")
        print(f"📋 Columnas: {len(df.columns)}")
        
        # Detectar automáticamente qué columnas podrían ser fechas
        date_columns = []
        for i, col in enumerate(df.columns):
            sample_values = df[col].dropna().head(10)
            if len(sample_values) > 0:
                if pd.api.types.is_datetime64_any_dtype(sample_values):
                    date_columns.append(i)
                    print(f"📅 Columna {i} parece contener fechas")
        
        # Crear encabezados básicos
        headers = []
        for i in range(len(df.columns)):
            if i in date_columns:
                headers.append("FECHA")
            else:
                headers.append(f"COLUMNA_{i+1}")
        
        # Asignar los encabezados
        df.columns = headers
        
        # Guardar con encabezados
        if output_file is None:
            output_file = input_file.replace('.xlsx', '_encabezados.xlsx')
        
        df.to_excel(output_file, index=False)
        print(f"✅ Archivo guardado con encabezados: {output_file}")
        
        return output_file
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

def generar_reporte_formateado(input_file, start_date=None, end_date=None, district=None, crew=None):
    """
    Genera un reporte con el formato específico solicitado
    """
    try:
        # Verificar si el archivo necesita encabezados
        try:
            df = pd.read_excel(input_file)
            # Si llega aquí, el archivo ya tiene encabezados
            print(f"📊 Archivo con encabezados: {input_file}")
        except:
            # Si hay error, agregar encabezados
            print("📋 Agregando encabezados al archivo...")
            input_file = agregar_encabezados(input_file)
            if input_file is None:
                return None
            df = pd.read_excel(input_file)
        
        print(f"📋 Total de registros: {len(df)}")
        
        # Crear DataFrame con el formato específico
        formatted_df = pd.DataFrame()
        
        # Mapear columnas según el formato solicitado
        formatted_df['CODIGO DE POSTE CAMPO'] = df['COLUMNA_4']  # Código del poste
        formatted_df['DISTRITO UBICACION'] = df['COLUMNA_6']     # Distrito
        formatted_df['LATITUD Y'] = df['COLUMNA_7']              # Latitud
        formatted_df['LONGITUD X'] = df['COLUMNA_9']             # Longitud
        formatted_df['PROPIETARIO'] = df['COLUMNA_11']           # Propietario
        formatted_df['FECHA'] = df['FECHA.1'].dt.strftime('%d-%m-%Y')  # Fecha en formato DD-MM-YYYY
        formatted_df['EMPRESA EJECUTORA'] = df['COLUMNA_13']     # Empresa ejecutora
        formatted_df['CUADRILLA'] = df['COLUMNA_14']             # Cuadrilla
        formatted_df['¿EXISTE APOYO?'] = df['COLUMNA_15']        # ¿Existe apoyo?
        formatted_df['NUMERO DE CABLES'] = df['COLUMNA_16']      # Número de cables
        formatted_df['¿SE REVISO EN CAMPO?'] = df['COLUMNA_17']  # ¿Se revisó en campo?
        formatted_df['¿TRABAJO EJECUTADO?'] = df['COLUMNA_18']   # ¿Trabajo ejecutado?
        formatted_df['¿CUMPLE EL DMS?'] = df['COLUMNA_19']       # ¿Cumple el DMS?
        formatted_df['OBSERVACIONES Y/O COMENTARIOS DEL TRABAJO O INSPECCIÓN'] = df['COLUMNA_20']  # Observaciones
        
        print(f"✅ DataFrame formateado creado con {len(formatted_df)} registros")
        
        # Aplicar filtros
        if start_date:
            start_datetime = pd.to_datetime(start_date)
            mask = pd.to_datetime(df['FECHA.1']).dt.date >= start_datetime.date()
            formatted_df = formatted_df[mask]
            print(f"📅 Filtrado por fecha de inicio: {start_date}")
        
        if end_date:
            end_datetime = pd.to_datetime(end_date)
            mask = pd.to_datetime(df['FECHA.1']).dt.date <= end_datetime.date()
            formatted_df = formatted_df[mask]
            print(f"📅 Filtrado por fecha de fin: {end_date}")
        
        if district:
            mask = df['COLUMNA_6'].str.contains(district, case=False, na=False)
            formatted_df = formatted_df[mask]
            print(f"🏘️ Filtrado por distrito: {district}")
        
        if crew:
            mask = df['COLUMNA_14'].str.contains(crew, case=False, na=False)
            formatted_df = formatted_df[mask]
            print(f"👥 Filtrado por cuadrilla: {crew}")
        
        print(f"📊 Registros finales: {len(formatted_df)}")
        
        if len(formatted_df) == 0:
            print("❌ No hay registros que coincidan con los filtros especificados")
            return None
        
        # Guardar reporte en carpeta reportes
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"reportes/reporte_formateado_{timestamp}.xlsx"
        
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            formatted_df.to_excel(writer, sheet_name='Reporte', index=False)
            
            # Ajustar ancho de columnas
            worksheet = writer.sheets['Reporte']
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        print(f"✅ Reporte formateado guardado: {output_file}")
        
        # Mostrar primeras filas como ejemplo
        print(f"\n📋 Primeras 5 filas del reporte:")
        print(formatted_df.head().to_string(index=False))
        
        return output_file
        
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

def main():
    print("🚀 Generador de Reportes Formateados")
    print("=" * 50)
    
    # Verificar argumentos
    if len(sys.argv) < 2:
        print("❌ Error: Debes especificar la ruta del archivo Excel")
        print("\n📋 Uso: python generar_reporte.py 'archivo.xlsx' [fecha_inicio] [fecha_fin] [distrito] [cuadrilla]")
        print("\n📋 Ejemplos:")
        print("  python generar_reporte.py 'datos.xlsx'")
        print("  python generar_reporte.py 'datos.xlsx' 2025-07-02 2025-07-02")
        print("  python generar_reporte.py 'datos.xlsx' 2025-07-02 2025-07-02 'SAN MIGUEL'")
        print("  python generar_reporte.py 'datos.xlsx' 2025-07-02 2025-07-02 'SAN MIGUEL' 'CU2'")
        print("\n📅 Formato de fecha: YYYY-MM-DD (ej: 2025-07-02)")
        return
    
    # Obtener ruta del archivo
    excel_file = sys.argv[1]
    
    # Verificar que el archivo existe
    if not os.path.exists(excel_file):
        print(f"❌ Error: El archivo '{excel_file}' no existe")
        print("💡 Asegúrate de que el archivo esté en la carpeta correcta")
        return
    
    print(f"📁 Archivo Excel: {excel_file}")
    
    # Procesar argumentos opcionales
    start_date = None
    end_date = None
    district = None
    crew = None
    
    # Fechas
    if len(sys.argv) > 2:
        try:
            start_date = datetime.strptime(sys.argv[2], '%Y-%m-%d').date()
            print(f"📅 Fecha de inicio: {start_date}")
        except ValueError:
            print("❌ Error: Formato de fecha debe ser YYYY-MM-DD")
            return
    
    if len(sys.argv) > 3:
        try:
            end_date = datetime.strptime(sys.argv[3], '%Y-%m-%d').date()
            print(f"📅 Fecha de fin: {end_date}")
        except ValueError:
            print("❌ Error: Formato de fecha debe ser YYYY-MM-DD")
            return
    
    # Distrito
    if len(sys.argv) > 4:
        district = sys.argv[4]
        print(f"🏘️ Distrito: {district}")
    
    # Cuadrilla
    if len(sys.argv) > 5:
        crew = sys.argv[5]
        print(f"👥 Cuadrilla: {crew}")
    
    # Generar reporte
    try:
        output_file = generar_reporte_formateado(excel_file, start_date, end_date, district, crew)
        
        if output_file:
            print(f"\n🎉 ¡Reporte generado exitosamente!")
            print(f"📁 Archivo: {output_file}")
            print("💡 Abre el archivo en Excel para ver el reporte completo")
        else:
            print("❌ No se pudo generar el reporte")
            
    except Exception as e:
        print(f"❌ Error: {str(e)}")

if __name__ == "__main__":
    main() 