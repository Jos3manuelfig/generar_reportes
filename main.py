import re
import pandas as pd

# Leer texto desde archivo
with open("mensajes.txt", "r", encoding="utf-8") as file:
    texto = file.read()

# Buscar la fecha
fecha_match = re.search(r"Fecha:\s*(\d{2}/\d{2}/\d{2})", texto)
fecha = fecha_match.group(1) if fecha_match else "Fecha no encontrada"

# Buscar bloques de materiales
bloques = re.findall(r"Materiales.*?:.*?(?=Materiales|\Z)", texto, re.DOTALL | re.IGNORECASE)

data = []

for bloque in bloques:
    # Extraer número de poste
    poste_match = re.search(r"Materiales\s+(\d+):", bloque, re.IGNORECASE)
    poste = poste_match.group(1) if poste_match else "Desconocido"

    # Buscar línea de POSTE
    poste_line_match = re.search(rf"POSTE\s+{poste}(?:\s+(\w+))?(?:\s+(\w+))?", bloque, re.IGNORECASE)
    if poste_line_match:
        cod1 = poste_line_match.group(1)
        cod2 = poste_line_match.group(2)

        if cod1 and cod1.isdigit():
            codigo = cod1
            propietario = cod2 if cod2 and cod2.upper() in ["LDS", "OPT", "ENEL"] else ""
        else:
            codigo = "S/C"
            propietario = cod1 if cod1 and cod1.upper() in ["LDS", "OPT", "ENEL"] else ""
    else:
        codigo = "S/C"
        propietario = ""

    # Contar etiquetas (cables)
    etiquetas = sum(int(e) for e in re.findall(r"(\d+)\s+etiqueta", bloque, re.IGNORECASE))

    data.append({
        "Fecha": fecha,
        "Número de Poste": poste,
        "Código": codigo,
        "Propietario": propietario,
        "Cantidad de Etiquetas": etiquetas
    })

# Crear DataFrame
df = pd.DataFrame(data)

# Exportar a Excel
df.to_excel("reporte_postes.xls", index=False)
print("✅ Archivo 'reporte_postes.xlsx' generado correctamente.")
