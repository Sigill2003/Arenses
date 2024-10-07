import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re

# Función para validar letras (Nombre)
def validar_letras(valor):
    return bool(re.match(r'^[a-zA-Z\s]+$', str(valor)))

# Función para validar números enteros (ID, Cantidad Disponible, Stock Mínimo)
def validar_entero_positivo(valor):
    return bool(re.match(r'^\d+$', str(valor))) and int(valor) >= 0

# Función para validar los datos del DataFrame
def validar_datos(df):
    errores = []
    for i, fila in df.iterrows():
        if not validar_entero_positivo(fila['ID']):
            errores.append(f"Fila {i + 1}: ID '{fila['ID']}' no es un número entero válido.")
        if not validar_letras(fila['Nombre']):
            errores.append(f"Fila {i + 1}: Nombre '{fila['Nombre']}' contiene caracteres inválidos.")
        if not validar_entero_positivo(fila['Cantidad Disponible']):
            errores.append(f"Fila {i + 1}: Cantidad Disponible '{fila['Cantidad Disponible']}' no es un número entero válido.")
        if not validar_entero_positivo(fila['Stock Mínimo']):
            errores.append(f"Fila {i + 1}: Stock Mínimo '{fila['Stock Mínimo']}' no es un número entero válido.")
    return errores

# Función para cargar la tabla desde un archivo Excel (sin validar)
def load_excel():
    global current_df
    file_path = filedialog.askopenfilename(defaultextension='.xlsx',
                                           filetypes=[("Excel files", "*.xlsx")],
                                           title="Cargar archivo Excel")
    if file_path:
        try:
            # Cargar el archivo Excel en un DataFrame
            df = pd.read_excel(file_path)

            # Verificar si las columnas necesarias están presentes
            required_columns = ['ID', 'Nombre', 'Cantidad Disponible', 'Stock Mínimo']
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"Falta la columna: '{col}' en el archivo Excel.")

            # Guardar el DataFrame cargado
            current_df = df
            # Mostrar los datos en el Treeview
            mostrar_datos_en_treeview(df)

            messagebox.showinfo("Carga Exitosa", "Archivo cargado correctamente. Ahora puedes validarlo.")
        except ValueError as ve:
            messagebox.showerror("Error", str(ve))
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar los datos: {str(e)}")
    else:
        messagebox.showwarning("Cancelado", "Carga cancelada")

# Función para validar los datos después de cargarlos
def validar_excel():
    global current_df
    if current_df is not None:
        # Validar los datos cargados
        errores = validar_datos(current_df)

        if errores:
            # Mostrar los errores si los hay
            messagebox.showerror("Errores encontrados", "\n".join(errores))
        else:
            messagebox.showinfo("Validación Exitosa", "Datos validados correctamente.")
    else:
        messagebox.showwarning("Error", "No hay datos cargados para validar.")

# Función para mostrar los datos en el Treeview
def mostrar_datos_en_treeview(df):
    # Limpiar el Treeview antes de mostrar nuevos datos
    for item in tree.get_children():
        tree.delete(item)

    # Agregar los nuevos datos al Treeview
    for index, row in df.iterrows():
        tree.insert("", tk.END, values=(row['ID'], row['Nombre'], row['Cantidad Disponible'], row['Stock Mínimo']))

# Función para corregir y guardar los datos en un nuevo archivo Excel
def save_corrected_excel():
    global current_df
    if current_df is not None:
        # Obtener datos del Treeview
        data = []
        for row in tree.get_children():
            data.append(tree.item(row)['values'])

        # Crear un nuevo DataFrame
        corrected_df = pd.DataFrame(data, columns=['ID', 'Nombre', 'Cantidad Disponible', 'Stock Mínimo'])

        # Validar antes de guardar
        errores = validar_datos(corrected_df)
        if errores:
            mostrar_errores(errores)
        else:
            try:
                # Guardar el DataFrame corregido
                file_path = filedialog.asksaveasfilename(defaultextension='.xlsx',
                                                         filetypes=[("Excel files", "*.xlsx")],
                                                         title="Guardar archivo Excel corregido")
                if file_path:
                    corrected_df.to_excel(file_path, index=False)
                    messagebox.showinfo("Guardar archivo", "Archivo Excel corregido guardado con éxito.")
                else:
                    messagebox.showwarning("Cancelado", "Guardado cancelado.")
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo guardar los datos: {str(e)}")
    else:
        messagebox.showwarning("Error", "No hay datos cargados para guardar.")

# Función para mostrar los errores y permitir la corrección
def mostrar_errores(errores):
    error_window = tk.Toplevel(root)
    error_window.title("Errores encontrados")

    # Crear un Text widget para mostrar los errores
    error_text = tk.Text(error_window, height=15, width=50)
    error_text.pack(pady=10)

    # Mostrar errores en el Text widget
    error_text.insert(tk.END, "\n".join(errores))
    error_text.config(state=tk.DISABLED)

    # Crear un botón para cerrar la ventana
    close_button = tk.Button(error_window, text="Cerrar", command=error_window.destroy)
    close_button.pack(pady=5)

# Función para salir del programa
def salir():
    if messagebox.askyesno("Salir", "¿Seguro que deseas salir?"):
        root.destroy()

# Función para crear la interfaz gráfica
def create_gui():
    global root, tree, current_df
    current_df = None  # Variable global para almacenar el DataFrame cargado y corregido

    # Crear la ventana principal
    root = tk.Tk()
    root.title("Validación y Corrección de Tabla Excel")

    # Crear los botones
    load_button = tk.Button(root, text="Cargar Excel", command=load_excel)
    load_button.pack(pady=10)

    validate_button = tk.Button(root, text="Validar Excel", command=validar_excel)
    validate_button.pack(pady=10)

    save_button = tk.Button(root, text="Guardar Correcciones", command=save_corrected_excel)
    save_button.pack(pady=10)

    exit_button = tk.Button(root, text="Salir", command=salir)
    exit_button.pack(pady=10)

    # Crear un Treeview para mostrar los datos
    tree = ttk.Treeview(root, columns=('ID', 'Nombre', 'Cantidad Disponible', 'Stock Mínimo'), show='headings')
    tree.heading('ID', text='ID')
    tree.heading('Nombre', text='Nombre')
    tree.heading('Cantidad Disponible', text='Cantidad Disponible')
    tree.heading('Stock Mínimo', text='Stock Mínimo')
    tree.pack(pady=10)

    # Ejecutar el bucle principal de Tkinter
    root.mainloop()

# Llamar a la función para crear la interfaz gráfica
create_gui()
