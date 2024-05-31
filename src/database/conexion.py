from ..utils.resource_path import resource_path
from tkinter import messagebox
import sqlite3 as sql

def get_conexion():
    try:
        conexion = sql.connect(resource_path("src/database/rutas.db"))
        return conexion
    except sql.Error as ex:
        messagebox.showerror("Error", "Error al conectar a la base de datos: " + str(ex))
        return

def get_rutas():
    try:
        query = "SELECT NOMBRE, RUTA FROM RUTAS"
        conexion = get_conexion()
        cursor = conexion.cursor()
        cursor.execute(query)
        resultados = cursor.fetchall()
        return resultados
    except sql.Error as ex:
        messagebox.showerror("Error", "Error al obtener las rutas: " + str(ex))
        return
    finally:
        cursor.close()
        conexion.close()

def set_ruta(ruta, nombre):
    try:
        query = """UPDATE RUTAS SET RUTA = '"""+ruta+"""' WHERE NOMBRE = '"""+nombre+"""'"""
        conexion = get_conexion()
        cursor = conexion.cursor()
        cursor.execute(query)
        conexion.commit()
    except sql.Error as ex:
        messagebox.showerror("Error", "Error al actualizar la ruta: " + str(ex))
        return
    finally:
        cursor.close()
        conexion.close()