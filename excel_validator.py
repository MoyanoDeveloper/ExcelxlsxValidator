import pandas as pd

# Leer el archivo Excel
try:
    data = pd.read_excel("datosNumero.xlsx")
except FileNotFoundError:
    print("El archivo 'datosNumero.xlsx' no se encontró.")
    exit()
except Exception as e:
    print(f"Ocurrió un error al leer el archivo: {e}")
    exit()

# Imprimir los nombres de las columnas para depuración
print("Nombres de columnas:", data.columns.tolist())

# Limpiar los nombres de las columnas eliminando espacios
data.columns = data.columns.str.strip()

# Validaciones básicas
def validar_filas(data):
    errores = []
    # Verifica si las columnas necesarias existen
    required_columns = ['NOMBRE', 'MONTO']
    for col in required_columns:
        if col not in data.columns:
            errores.append(f"Columna faltante: '{col}'")

    # Verifica los datos en las filas si las columnas existen
    if not errores:  # Solo valida filas si no hay errores de columnas
        for index, row in data.iterrows():
            fila_info = row.to_dict()  # Convertir la fila a diccionario para mostrar su contenido

            # Verifica que NOMBRE no esté vacío y sea un string
            if pd.isnull(row['NOMBRE']) or not isinstance(row['NOMBRE'], str):
                errores.append(
                    f"Fila {index + 1}: 'NOMBRE' debe ser un string. Fila: {fila_info}"
                )

            # Verifica que MONTO no esté vacío y sea un número
            if pd.isnull(row['MONTO']) or not isinstance(row['MONTO'], (int, float)):
                errores.append(
                    f"Fila {index + 1}: 'MONTO' debe ser un número. La fila con error es la fila {index + 1}: {fila_info}"
                )

    return errores

# Verificar errores
errores = validar_filas(data)
if errores:
    print("Errores encontrados:")
    for error in errores:
        print(error)
else:
    print("Datos validados")