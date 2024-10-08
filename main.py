import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re


# Función para validar letras (Nombre) y ajustar el formato
def validar_letras_y_formatear(valor):
    valor = str(valor).strip()  # Eliminar espacios en blanco
    if re.match(r'^[a-zA-Z\s]+$', valor):
        # Convertir a formato "Primera Letra Mayúscula"
        return valor.title()  # Convierte "juan" en "Juan", "ANA" en "Ana"
    return None


# Función para validar números enteros (ID, Cantidad Disponible, Stock Mínimo)
def validar_entero_positivo(valor):
    return bool(re.match(r'^\d+$', str(valor))) and int(valor) >= 0


# Función para validar los datos del DataFrame
def validar_datos(df):
    errores = []
    # Verificar IDs duplicados
    id_duplicados = df[df.duplicated(subset=['ID'], keep=False)]
    if not id_duplicados.empty:
        for i, fila in id_duplicados.iterrows():
            errores.append(f"Fila {i + 1}: ID '{fila['ID']}' está duplicado.")

    for i, fila in df.iterrows():
        if not validar_entero_positivo(fila['ID']):
            errores.append(f"Fila {i + 1}: ID '{fila['ID']}' no es un número entero válido.")
        nombre_formateado = validar_letras_y_formatear(fila['Nombre'])
        if nombre_formateado is None:
            errores.append(f"Fila {i + 1}: Nombre '{fila['Nombre']}' contiene caracteres inválidos.")
        else:
            df.at[i, 'Nombre'] = nombre_formateado  # Guardar nombre con formato corregido
        if not validar_entero_positivo(fila['Cantidad Disponible']):
            errores.append(
                f"Fila {i + 1}: Cantidad Disponible '{fila['Cantidad Disponible']}' no es un número entero válido.")
        if not validar_entero_positivo(fila['Stock Mínimo']):
            errores.append(f"Fila {i + 1}: Stock Mínimo '{fila['Stock Mínimo']}' no es un número entero válido.")
    return errores


# Función para cargar la tabla desde un archivo Excel (sin validar)
def load_excel():
    global current_df, original_file_path
    original_file_path = filedialog.askopenfilename(defaultextension='.xlsx',
                                                    filetypes=[("Excel files", "*.xlsx")],
                                                    title="Cargar archivo Excel")
    if original_file_path:
        try:
            # Cargar el archivo Excel en un DataFrame
            df = pd.read_excel(original_file_path)

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
            # Actualizar el Treeview con los nombres formateados
            mostrar_datos_en_treeview(current_df)
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


# Función para permitir la edición de celdas
def editar_celda(event):
    selected_item = tree.selection()[0]  # Obtener el elemento seleccionado
    column = tree.identify_column(event.x)  # Identificar columna
    row = tree.identify_row(event.y)  # Identificar fila

    col_index = int(column.replace("#", "")) - 1
    row_id = tree.item(selected_item, 'values')[0]

    # Crear una ventana de entrada para la edición
    entry_window = tk.Toplevel(root)
    entry_window.title(f"Editar valor en fila {row} columna {column}")

    label = tk.Label(entry_window, text=f"Valor actual: {tree.item(selected_item, 'values')[col_index]}")
    label.pack(pady=5)

    entry = tk.Entry(entry_window)
    entry.pack(pady=5)
    entry.insert(0, tree.item(selected_item, 'values')[col_index])

    def guardar_cambio():
        nuevo_valor = entry.get()

        # Validar y actualizar la celda correspondiente
        if col_index == 0:
            # Validar ID
            if not validar_entero_positivo(nuevo_valor):
                messagebox.showerror("Error", "ID debe ser un número entero positivo.")
                return
            elif int(nuevo_valor) in current_df['ID'].values and int(nuevo_valor) != int(
                    tree.item(selected_item, 'values')[0]):
                messagebox.showerror("Error", f"ID '{nuevo_valor}' ya existe. Debe ser único.")
                return
            else:
                tree.set(selected_item, column=column, value=int(nuevo_valor))
        elif col_index == 1:
            # Validar y formatear Nombre
            nombre_formateado = validar_letras_y_formatear(nuevo_valor)
            if nombre_formateado is None:
                messagebox.showerror("Error", "Nombre solo debe contener letras.")
                return
            else:
                tree.set(selected_item, column=column, value=nombre_formateado)
        elif col_index == 2:
            # Validar Cantidad Disponible
            if not validar_entero_positivo(nuevo_valor):
                messagebox.showerror("Error", "Cantidad Disponible debe ser un número entero positivo.")
                return
            else:
                tree.set(selected_item, column=column, value=int(nuevo_valor))
        elif col_index == 3:
            # Validar Stock Mínimo
            if not validar_entero_positivo(nuevo_valor):
                messagebox.showerror("Error", "Stock Mínimo debe ser un número entero positivo.")
                return
            else:
                tree.set(selected_item, column=column, value=int(nuevo_valor))
        entry_window.destroy()

        # Actualizar el DataFrame actual
        actualizar_current_df()

    save_button = tk.Button(entry_window, text="Guardar", command=guardar_cambio)
    save_button.pack(pady=5)


# Función para agregar una nueva fila
def agregar_fila():
    # Crear una ventana de entrada para la nueva fila
    new_row_window = tk.Toplevel(root)
    new_row_window.title("Agregar nueva fila")

    labels = ["ID", "Nombre", "Cantidad Disponible", "Stock Mínimo"]
    entries = []

    for label_text in labels:
        label = tk.Label(new_row_window, text=label_text)
        label.pack()
        entry = tk.Entry(new_row_window)
        entry.pack()
        entries.append(entry)

    def guardar_fila():
        nueva_fila = [entry.get() for entry in entries]

        # Validar la nueva fila
        # Validar ID
        if not validar_entero_positivo(nueva_fila[0]):
            messagebox.showerror("Error", "ID debe ser un número entero positivo.")
            return
        elif int(nueva_fila[0]) in current_df['ID'].values:
            messagebox.showerror("Error", f"ID '{nueva_fila[0]}' ya existe. Debe ser único.")
            return

        # Validar y formatear Nombre
        nombre_formateado = validar_letras_y_formatear(nueva_fila[1])
        if nombre_formateado is None:
            messagebox.showerror("Error", "Nombre solo debe contener letras.")
            return

        # Validar Cantidad Disponible
        if not validar_entero_positivo(nueva_fila[2]):
            messagebox.showerror("Error", "Cantidad Disponible debe ser un número entero positivo.")
            return

        # Validar Stock Mínimo
        if not validar_entero_positivo(nueva_fila[3]):
            messagebox.showerror("Error", "Stock Mínimo debe ser un número entero positivo.")
            return

        # Agregar la nueva fila al Treeview
        nueva_fila = [int(nueva_fila[0]), nombre_formateado, int(nueva_fila[2]), int(nueva_fila[3])]
        tree.insert("", tk.END, values=nueva_fila)
        new_row_window.destroy()

        # Actualizar el DataFrame actual
        actualizar_current_df()

    save_button = tk.Button(new_row_window, text="Agregar Fila", command=guardar_fila)
    save_button.pack(pady=5)


# Función para actualizar el DataFrame actual desde el Treeview
def actualizar_current_df():
    global current_df
    data = []
    for row in tree.get_children():
        data.append(tree.item(row)['values'])
    current_df = pd.DataFrame(data, columns=['ID', 'Nombre', 'Cantidad Disponible', 'Stock Mínimo'])


# Función para corregir y guardar los datos en un nuevo archivo Excel
def save_corrected_excel():
    global current_df, original_file_path
    if current_df is not None:
        try:
            # Validar antes de guardar
            errores = validar_datos(current_df)
            if errores:
                mostrar_errores(errores)
                return

            # Abrir un cuadro de diálogo para elegir el nombre del archivo corregido
            file_path = filedialog.asksaveasfilename(defaultextension='.xlsx',
                                                     filetypes=[("Excel files", "*.xlsx")],
                                                     title="Guardar archivo Excel corregido")
            if file_path:
                # Verificar si el archivo a guardar es el mismo que el original
                if file_path == original_file_path:
                    overwrite = messagebox.askyesno("Confirmar",
                                                    "Está a punto de sobrescribir el archivo original. ¿Desea continuar?")
                    if not overwrite:
                        messagebox.showinfo("Guardado Cancelado", "El archivo no se ha guardado.")
                        return

                # Guardar el DataFrame corregido en el archivo Excel especificado
                current_df.to_excel(file_path, index=False)
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
    global root, tree, current_df, original_file_path
    current_df = None  # Variable global para almacenar el DataFrame cargado y corregido
    original_file_path = ""

    # Crear la ventana principal
    root = tk.Tk()
    root.title("Validación y Corrección de Tabla Excel")

    # Crear un frame para los botones
    button_frame = tk.Frame(root)
    button_frame.pack(pady=10)

    # Crear los botones en una sola hilera usando grid
    load_button = tk.Button(button_frame, text="Cargar", command=load_excel)
    load_button.grid(row=0, column=0, padx=5)

    validate_button = tk.Button(button_frame, text="Validar", command=validar_excel)
    validate_button.grid(row=0, column=1, padx=5)

    save_button = tk.Button(button_frame, text="Guardar Correcciones", command=save_corrected_excel)
    save_button.grid(row=0, column=2, padx=5)

    add_row_button = tk.Button(button_frame, text="Agregar Fila", command=agregar_fila)
    add_row_button.grid(row=0, column=3, padx=5)

    exit_button = tk.Button(button_frame, text="Salir", command=salir)
    exit_button.grid(row=0, column=4, padx=5)

    # Crear un Treeview para mostrar los datos
    tree = ttk.Treeview(root, columns=('ID', 'Nombre', 'Cantidad Disponible', 'Stock Mínimo'), show='headings')
    tree.heading('ID', text='ID')
    tree.heading('Nombre', text='Nombre')
    tree.heading('Cantidad Disponible', text='Cantidad Disponible')
    tree.heading('Stock Mínimo', text='Stock Mínimo')
    tree.pack(pady=10)

    # Asociar doble clic para editar celdas
    tree.bind('<Double-1>', editar_celda)

    # Ejecutar el bucle principal de Tkinter
    root.mainloop()


# Llamar a la función para crear la interfaz gráfica
create_gui()
