import pandas as pd
import openpyxl
import warnings
import sys
import os
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from openpyxl.utils import get_column_letter
from tkinter import messagebox
from customtkinter import *
from conexion import *

warnings.filterwarnings("ignore")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def main():
    def formatear_excel(excel_file):
        try:
            wb = openpyxl.load_workbook(excel_file)
            ws = wb.active
            ws.title = "DETALLE"
            
            fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
            font_header = Font(name="Calibri", size=11, color="000000", bold=True)
            font_cells = Font(name="Calibri", size=11)
            border = Border(left=Side(style="thin"), 
                            right=Side(style="thin"), 
                            top=Side(style="thin"), 
                            bottom=Side(style="thin"))
            alignment = Alignment(vertical="center")
            
            for row in ws.iter_rows():
                for cell in row:
                    cell.border = border
                    cell.alignment = alignment
                    cell.font = font_cells
                    if cell.row == 1:
                        cell.fill = fill
                        cell.font = font_header
                        cell.alignment = Alignment(horizontal="center")
            
            column_widths = [10.5, 40, 8.5, 23, 13.5, 12, 14]
            for i, column_width in enumerate(column_widths):
                ws.column_dimensions[get_column_letter(i+1)].width = column_width
            
            wb.save(excel_file)
            
        except Exception:
            messagebox.showerror("Error", "Algo salió mal. Por favor, intente nuevamente.")

    def obtener_deudas_vencidas(base_path, dacxanalista_path, resultado_path):
        try:
            df_base = pd.read_excel(base_path)
            df_base = df_base.iloc[:, 3:]
            df_base = df_base.iloc[7:, :]
            df_base = df_base.rename(columns=df_base.iloc[0])
            df_base = df_base[1:]
            df_base = df_base.reset_index(drop=True)
            df_base = df_base.dropna(subset=["Cuenta","ACC"])
            df_base = df_base.reset_index(drop=True)
            df_base = df_base.rename(columns={"     Importe en ML": "Importe"})
            columnas_deseadas = ["ACC", "Cuenta", "Demora", "Importe"]
            df_base = df_base[columnas_deseadas]
            df_base["Demora"] = df_base["Demora"].astype("Int64")
            df_base["Importe"] = df_base["Importe"].astype(float)
            # Condition 1
            df_base["Status"] = df_base["Importe"].apply(lambda x: "DEUDA" if x > 0 else "SALDOS A FAVOR")
            # Condition 2
            df_base["Tipo Deuda"] = df_base["Demora"].apply(lambda x: "CORRIENTE" if x <= 0 else "VENCIDA")
            # Condition 3
            df_base["Saldo Final"] = df_base.apply(lambda row: row["Importe"] if (row["Status"] == "DEUDA" and row["Tipo Deuda"] == "VENCIDA") else (row["Importe"] if row["Status"] == "SALDOS A FAVOR" else "NO"), axis=1)
            df_base = df_base[df_base["Saldo Final"] != "NO"]
            df_base = df_base.sort_values(by=["Cuenta"], ascending=[True])
            df_base = df_base.sort_values(by=["ACC"], ascending=[True])
            df_base = df_base.sort_values(by=["Demora"], ascending=[False])
            df_base = df_base.reset_index(drop=True)
            
            ultima_fila = df_base.shape[0]
            for i in range(ultima_fila):
                if df_base.loc[i, "Status"] == "DEUDA":
                    saldoDeuda = df_base.loc[i, "Saldo Final"]
                    for j in range(ultima_fila):
                        if (
                            df_base.loc[i, "Cuenta"]    == df_base.loc[j, "Cuenta"] and 
                            df_base.loc[i, "ACC"]       == df_base.loc[j, "ACC"]    and 
                            df_base.loc[j, "Status"]    == "SALDOS A FAVOR"
                            ):
                            saldoFavor = df_base.loc[j, "Saldo Final"]
                            montoCompensar = min(saldoDeuda, abs(saldoFavor))
                            df_base.loc[i, "Saldo Final"] = saldoDeuda - montoCompensar
                            df_base.loc[j, "Saldo Final"] = saldoFavor + montoCompensar
                            saldoDeuda = df_base.loc[i, "Saldo Final"]
                            
            df_base = df_base[(df_base["Tipo Deuda"] == "VENCIDA") & (df_base["Status"] == "DEUDA")]
            df_base = df_base.reset_index(drop=True)
            
            grouped_df = df_base.groupby(["Cuenta", "ACC"]).agg({"Demora": "max", "Saldo Final": "sum"})
            
            df_final = grouped_df.reset_index()[["Cuenta", "ACC", "Saldo Final", "Demora"]]
            df_final = df_final.rename(columns={"Cuenta":"Cod Cliente", "ACC":"Área Ctrl", "Saldo Final":"Deuda Vencida", "Demora":"Días Morosidad"})
            
            df_dacxanalista = pd.read_excel(dacxanalista_path, sheet_name="Base_NUEVA")
            
            df_final = df_final.merge(df_dacxanalista[["DEUDOR", "NOMBRE"]], left_on="Cod Cliente", right_on="DEUDOR", how="left")
            df_final = df_final.rename(columns={"NOMBRE": "Razón Social"})
            df_final = df_final.drop(columns=["DEUDOR"])
            
            areas_de_control = {
                "PE01": "Post-Pago",
                "PE02": "Pre-Pago",
                "PE03": "Tiempo Aire",
                "PE04": "Reintegro",
                "PE05": "Reestructura",
                "PE07": "Contado / Administrativas",
                "PE09": "Cargos Admtivos / Otros",
                "PE10": "Sim Card",
                "PE11": "Recarga Prepago",
                "PE12": "Recarga Física",
                "PE13": "Arrendamiento",
                "PE14": "Tel.Fija Inalamb.",
                "PE15": "Prendas",
                "PE16": "DTH",
                "PE17": "HFC"
            }
            df_final["Producto"] = df_final["Área Ctrl"].apply(lambda x: areas_de_control[x])
            df_final["Código Pago"] = "33" + df_final["Área Ctrl"].str[-2:] + df_final["Cod Cliente"].astype(str)
            df_final = df_final[["Cod Cliente", "Razón Social", "Área Ctrl", "Producto", "Deuda Vencida", "Código Pago", "Días Morosidad"]]
            df_final["Deuda Vencida"] = df_final["Deuda Vencida"].astype(float)
            df_final = df_final[df_final["Deuda Vencida"] != 0]
            df_final.to_excel(resultado_path, index=False)
            
            formatear_excel(resultado_path)
            os.startfile(resultado_path)
            
        except Exception:
            messagebox.showerror("Error", "Algo salió mal. Por favor, intente nuevamente.")

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
                    SET BASE == '"""+base_path+"""'
                    WHERE ID == 1""")
        query2 = ("""UPDATE RUTAS
                    SET RESULTADO == '"""+resultado_path+"""'
                    WHERE ID == 1""")
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
                    SET DACXANALISTA == '"""+dacxanalista_path+"""'
                    WHERE ID == 1""")
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

    def ejecutar():
        query = ("""SELECT * FROM RUTAS WHERE ID == 1""")
        try:
            datos = ejecutar_query(query)
            ruta_base = datos[0][1]
            ruta_dacxa = datos[0][2]
            ruta_resultado = datos[0][3]
            if ruta_base == None or ruta_dacxa == None or ruta_resultado == None:
                messagebox.showerror("Error", "Por favor, configure las rutas de los archivos.")
            elif not os.path.exists(ruta_base):
                messagebox.showerror("Error", "No se encontraró el archivo BASE en la ruta especificada.")
            elif not os.path.exists(ruta_dacxa):
                messagebox.showerror("Error", "No se encontraró el archivo DACxANALISTA en la ruta especificada.")
            else:
                obtener_deudas_vencidas(ruta_base, ruta_dacxa, ruta_resultado)
        except Exception as ex:
            messagebox.showerror("Error", str(ex))

    def app():
        app = CTk()
        app.title("DV")
        app.iconbitmap(resource_path("icono.ico"))
        app.resizable(False, False)
        set_appearance_mode("light")
        
        main_frame = CTkFrame(app)
        main_frame.pack_propagate(0)
        main_frame.pack(fill="both", expand=True)
        
        titulo = CTkLabel(main_frame, text="Deudas Vencidas", font=("Calibri",25,"bold"), text_color="black")
        titulo.grid(row=0, column=0, columnspan=2, padx=(20,20), pady=(20, 0), sticky="nsew")
        
        ruta_base = CTkLabel(main_frame, text="BASE", font=("Calibri",15,"bold"), text_color="black")
        ruta_base.grid(row=1, column=0, padx=(20,10), pady=(20, 0), sticky="nsew")
        boton_seleccionar_ruta_base = CTkButton(
            main_frame, text="Seleccionar", fg_color="gray", border_color="black", border_width=2,
            font=("Calibri",15,"bold"), text_color="black", hover_color="red", width=15,
            command=lambda: seleccionar_base())
        boton_seleccionar_ruta_base.grid(row=2, column=0, padx=(20,10), pady=(0, 0), sticky="nsew")
        
        ruta_dacxa = CTkLabel(main_frame, text="DACxANALISTA", font=("Calibri",15,"bold"), text_color="black")
        ruta_dacxa.grid(row=1, column=1, padx=(10,20), pady=(20, 0), sticky="nsew")
        boton_seleccionar_ruta_dacxa = CTkButton(
            main_frame, text="Seleccionar", fg_color="gray", border_color="black", border_width=2,
            font=("Calibri",15,"bold"), text_color="black", hover_color="red", width=15,
            command=lambda: seleccionar_dacxanalista())
        boton_seleccionar_ruta_dacxa.grid(row=2, column=1, padx=(10,20), pady=(0, 0), sticky="nsew")
        
        boton_ejecutar = CTkButton(
            main_frame, text="EJECUTAR", fg_color="gray", border_color="black", border_width=2,
            font=("Calibri",20,"bold"), text_color="black", hover_color="red",
            command=lambda: ejecutar())
        boton_ejecutar.grid(row=3, column=0, columnspan=2, ipady=10, padx=(20,20), pady=(30, 20), sticky="nsew")

        app.mainloop()

    app()

if __name__ == "__main__":
    main()