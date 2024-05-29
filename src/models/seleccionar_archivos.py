from src.database.conexion import conexionSQLite
from tkinter import filedialog, messagebox
import os


def seleccionar_base():
    archivo_excel = filedialog.askopenfilename(
        initialdir="/",
        title="Seleccionar archivo BASE",
        filetypes=(("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*"))
    )
    directorio_base = os.path.dirname(archivo_excel)
    base_path = archivo_excel
    resultado_path = directorio_base+"/DEUDAS_VENCIDAS.xlsx"
    
    query1 = ("""UPDATE RUTAS
                SET BASE == '""" + base_path + """'
                WHERE ID == 0""")
    query2 = ("""UPDATE RUTAS
                SET RESULTADO == '""" + resultado_path + """'
                WHERE ID == 0""")
    conexion = conexionSQLite()
    try:
        cursor = conexion.cursor()
        cursor.execute(query1)
        conexion.commit()
        cursor.execute(query2)
        conexion.commit()
    except Exception as ex:
        messagebox.showerror("Error", str(ex))
    finally:
        cursor.close()
        conexion.close

def seleccionar_dacxanalista():
    archivo_excel = filedialog.askopenfilename(
        initialdir="/",
        title="Seleccionar archivo DACxANALISTA",
        filetypes=(("Archivos de Excel", "*.xlsx"), ("Todos los archivos", "*.*"))
    )
    dacxanalista_path = archivo_excel
    
    query = ("""UPDATE RUTAS
                SET DACXANALISTA == '""" + dacxanalista_path + """'
                WHERE ID == 0""")
    conexion = conexionSQLite()
    try:
        cursor = conexion.cursor()
        cursor.execute(query)
        conexion.commit()
    except Exception as ex:
        messagebox.showerror("Error", str(ex))
    finally:
        cursor.close()
        conexion.close