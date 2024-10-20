import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, Text
import os

def cargar_archivo():
    """Abrir el diálogo de selección de archivo y cargar Excel."""
    archivo = filedialog.askopenfilename(
        title="Selecciona un archivo Excel",
        filetypes=[("Archivos Excel", "*.xlsx *.xls")],
    )
    if archivo:
        try:
            global data
            data = pd.read_excel(archivo)
            text_area.insert(tk.END, "Archivo cargado con éxito.\n")
            text_area.insert(tk.END, f"Columnas: {data.columns.tolist()}\n\n")
            validar_datos()
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al leer el archivo: {e}")

def validar_datos():
    """Validar los datos cargados y mostrar los resultados."""
    errores = validar_filas(data)
    if errores:
        text_area.insert(tk.END, "Errores encontrados:\n")
        for error in errores:
            text_area.insert(tk.END, error + "\n")
        
        # Preguntar si se desea guardar los errores
        guardar = messagebox.askyesno("Errores Encontrados", "Errores Encontrados ¿Deseas guardar los errores en un archivo Excel?")
        if guardar:
            guardar_errores_excel(errores)
    else:
        text_area.insert(tk.END, "Datos validados correctamente.\n")

def validar_filas(data):
    """Función para validar los datos del Excel."""
    errores = []
    required_columns = ['NOMBRE', 'MONTO']

    for col in required_columns:
        if col not in data.columns:
            errores.append(f"Columna faltante: '{col}'")

    if not errores:
        for index, row in data.iterrows():
            fila_info = row.to_dict()
            if pd.isnull(row['NOMBRE']) or not isinstance(row['NOMBRE'], str):
                errores.append(f"Fila {index + 1}: 'NOMBRE' debe ser un string. {fila_info}")
            if pd.isnull(row['MONTO']) or not isinstance(row['MONTO'], (int, float)):
                errores.append(f"Fila {index + 1}: 'MONTO' debe ser un número. {fila_info}")

    return errores

def guardar_errores_excel(errores):
    """Guardar los errores en un archivo Excel."""
    nombre_archivo = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Archivos Excel", "*.xlsx")],
        title="Guardar errores como"
    )
    if nombre_archivo:
        try:
            df_errores = pd.DataFrame(errores, columns=["Errores"])
            df_errores.to_excel(nombre_archivo, index=False)
            messagebox.showinfo("Éxito", f"Errores guardados en: {os.path.basename(nombre_archivo)}")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar el archivo: {e}")

# Configuración de la ventana principal
ventana = tk.Tk()
ventana.title("Validador de Datos Excel")
ventana.geometry("600x400")

# Botón para cargar archivo
btn_cargar = tk.Button(ventana, text="Cargar Archivo", command=cargar_archivo)
btn_cargar.pack(pady=10)

# Área de texto para mostrar resultados
text_area = Text(ventana, wrap=tk.WORD, width=70, height=20)
text_area.pack(padx=10, pady=10)

# Iniciar la aplicación
ventana.mainloop()