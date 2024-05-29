from src.utils.resource_path import resource_path
import sqlite3 as sql

def conexionSQLite():
    try:
        conexion = sql.connect(resource_path("./db.db"))
        return conexion
    except sql.Error as ex:
        print("Error al conectar a la base de datos SQLite:", ex)
        return None

def ejecutar_query(query, parametros=None):
    conexion = conexionSQLite()
    try:
        cursor = conexion.cursor()
        if parametros is None:
            cursor.execute(query)
        else:
            cursor.execute(query, parametros)
        resultados = cursor.fetchall()
        return resultados
    except sql.Error as ex:
        error = "Error al ejecutar la consulta:" + ex
        return error
    finally:
        cursor.close()
        conexion.close