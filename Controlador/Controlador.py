import tkinter as tk

class Vista:
    def __init__(self, controlador):
        self.controlador = controlador
        self.root = tk.Tk()
        self.root.title('MVC - Base de datos')
        self.crear_widgets()

    def crear_widgets(self):
        self.subir_btn = tk.Button(self.root, text="Subir", command=self.controlador.subir)
        self.subir_btn.pack(pady=10)

        self.cargar_btn = tk.Button(self.root, text="Cargar", command=self.controlador.cargar)
        self.cargar_btn.pack(pady=10)

        self.revisar_btn = tk.Button(self.root, text="Revisar", command=self.controlador.revisar)
        self.revisar_btn.pack(pady=10)

        self.salir_btn = tk.Button(self.root, text="Salir", command=self.root.quit)
        self.salir_btn.pack(pady=10)

        self.texto_area = tk.Text(self.root, height=10, width=40)
        self.texto_area.pack(pady=10)

    def mostrar_datos(self, datos):
        self.texto_area.delete(1.0, tk.END)
        for fila in datos:
            self.texto_area.insert(tk.END, f"{fila}\n")

    def iniciar(self):
        self.root.mainloop()
