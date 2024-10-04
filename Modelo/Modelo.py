import sqlite3

class ModeloDB:
    def __init__(self):
        self.conexion = sqlite3.connect('datos.db')
        self.cursor = self.conexion.cursor()
        self.crear_tabla()

    def crear_tabla(self):
        self.cursor.execute('''CREATE TABLE IF NOT EXISTS datos (
                               id INTEGER PRIMARY KEY AUTOINCREMENT,
                               info TEXT)''')
        self.conexion.commit()

    def subir_dato(self, dato):
        self.cursor.execute("INSERT INTO datos (info) VALUES (?)", (dato,))
        self.conexion.commit()

    def cargar_datos(self):
        self.cursor.execute("SELECT * FROM datos")
        return self.cursor.fetchall()

    def cerrar_conexion(self):
        self.conexion.close()
