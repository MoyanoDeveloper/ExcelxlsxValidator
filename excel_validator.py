import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, Text
import os
import matplotlib.pyplot as plt
import seaborn as sns

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
            text_area.delete(1.0, tk.END)  # Limpiar el área de texto
            text_area.insert(tk.END, "Archivo cargado con éxito.\n")
            text_area.insert(tk.END, f"Columnas: {data.columns.tolist()}\n\n")
            validar_datos()
        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error al leer el archivo: {e}")

def validar_datos():
    """Validar los datos cargados y mostrar los resultados."""
    errores = validar_filas(data)
    text_area.delete(1.0, tk.END)  # Limpiar el área de texto para mostrar resultados
    if errores:
        text_area.insert(tk.END, "Errores encontrados:\n")
        for error in errores:
            text_area.insert(tk.END, error + "\n")
        
        guardar = messagebox.askyesno(
            "Errores Encontrados", 
            "Errores encontrados. ¿Deseas guardar los errores en un archivo Excel?"
        )
        if guardar:
            guardar_errores_excel(errores)
    else:
        text_area.insert(tk.END, "Datos validados correctamente.\n")
        preguntar_dashboard()

def preguntar_dashboard():
    """Pregunta si se desea crear un dashboard con los datos correctos."""
    crear_dashboard = messagebox.askyesno(
        "Crear Dashboard", 
        "¿Deseas crear un dashboard con los datos obtenidos?"
    )
    if crear_dashboard:
        mostrar_graficos()

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

def mostrar_graficos():
    """Mostrar gráficos de los datos utilizando Matplotlib y Seaborn."""
    try:
        plt.figure(figsize=(10, 5))
        
        # Usar los colores de la empresa en los gráficos
        sns.barplot(x='NOMBRE', y='MONTO', data=data, estimator=sum, color="#F26B2B")
        plt.title('Suma de Montos por Nombre', color="#F26B2B")
        plt.xticks(rotation=45)
        plt.tight_layout()
        
        plt.show()
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo generar el gráfico: {e}")

# Configuración de la ventana principal
ventana = tk.Tk()
ventana.title("Validador de Datos Capitalizarme Excel")
ventana.geometry("600x400")  # Color de fondo de la ventana principal

# Frame principal
frame = tk.Frame(ventana)  # Color de fondo del frame
frame.pack( fill=tk.BOTH, expand=True)

# Título
titulo = tk.Label(frame, text="Validador de Datos Capitalizarme Excel", font=("Helvetica", 16), fg="black")
titulo.pack(pady=10)

# Botón para cargar archivo
btn_cargar = tk.Button(frame, text="Cargar Archivo", command=cargar_archivo, bg="#F26B2B", fg="white", relief="flat")
btn_cargar.pack(pady=10)

# Área de texto para mostrar resultados
text_area = Text(frame, wrap=tk.WORD, width=70, height=20, bg="#ffffff", fg="#000000")
text_area.pack(padx=15, pady=10)

# Iniciar la aplicación
ventana.mainloop()