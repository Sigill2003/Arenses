import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

# Función para exportar los datos a Excel (Subir)
def export_to_excel():
    # Datos de ejemplo que exportaremos a Excel
    data = {
        'Nombre': ['Juan', 'Ana', 'Pedro', 'Luis'],
        'Edad': [23, 30, 22, 28],
        'Ciudad': ['Madrid', 'Barcelona', 'Valencia', 'Sevilla']
    }

    # Crear un DataFrame de pandas
    df = pd.DataFrame(data)

    # Usar cuadro de diálogo para seleccionar dónde guardar el archivo
    file_path = filedialog.asksaveasfilename(defaultextension='.xlsx',
                                             filetypes=[("Excel files", "*.xlsx")],
                                             title="Guardar archivo Excel")

    if file_path:
        try:
            # Guardar el DataFrame como archivo Excel
            df.to_excel(file_path, index=False)
            # Mostrar mensaje de éxito
            messagebox.showinfo("Exportar a Excel", "Datos exportados con éxito!")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo exportar los datos: {str(e)}")
    else:
        messagebox.showwarning("Cancelado", "Exportación cancelada")

# Función para cargar datos desde un archivo Excel (Cargar)
def load_from_excel():
    file_path = filedialog.askopenfilename(defaultextension='.xlsx',
                                           filetypes=[("Excel files", "*.xlsx")],
                                           title="Cargar archivo Excel")
    if file_path:
        try:
            # Cargar el archivo Excel en un DataFrame
            df = pd.read_excel(file_path)
            # Mostrar los datos en el Text widget
            text_widget.delete(1.0, tk.END)  # Limpiar el área de texto
            text_widget.insert(tk.END, df.to_string())  # Mostrar datos
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar los datos: {str(e)}")
    else:
        messagebox.showwarning("Cancelado", "Carga cancelada")

# Función para revisar los datos actuales
def revisar_datos():
    # Datos de ejemplo
    data = {
        'Nombre': ['Juan', 'Ana', 'Pedro', 'Luis'],
        'Edad': [23, 30, 22, 28],
        'Ciudad': ['Madrid', 'Barcelona', 'Valencia', 'Sevilla']
    }

    # Crear un DataFrame de pandas
    df = pd.DataFrame(data)

    # Mostrar los datos en el Text widget
    text_widget.delete(1.0, tk.END)  # Limpiar el área de texto
    text_widget.insert(tk.END, df.to_string())  # Insertar datos en el widget

# Función para salir del programa
def salir():
    if messagebox.askyesno("Salir", "¿Seguro que deseas salir?"):
        root.destroy()

# Función para crear la interfaz gráfica
def create_gui():
    global root, text_widget
    # Crear la ventana principal
    root = tk.Tk()
    root.title("Gestión de Base de Datos")

    # Crear los botones
    export_button = tk.Button(root, text="Subir", command=export_to_excel)
    export_button.pack(pady=10)

    load_button = tk.Button(root, text="Cargar", command=load_from_excel)
    load_button.pack(pady=10)

    review_button = tk.Button(root, text="Revisar", command=revisar_datos)
    review_button.pack(pady=10)
    exit_button = tk.Button(root, text="Salir", command=salir)
    exit_button.pack(pady=10)

    # Crear un widget Text para mostrar los datos
    text_widget = tk.Text(root, height=10, width=50)
    text_widget.pack(pady=10)

    # Ejecutar el bucle principal de Tkinter
    root.mainloop()

# Llamar a la función para crear la interfaz gráfica
create_gui()
