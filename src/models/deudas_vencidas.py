from src.models.validar_apoyos import Validar_Apoyos
from src.models.formatear_excel import formatear_excel
from datetime import datetime
from tkinter import messagebox
import pandas as pd
import time
import os


class Deudas_Vencidas():
    def __init__(self, rutas):
        fecha_actual = datetime.today().strftime("%d.%m.%Y")
        self.ruta_dacxa = rutas[0]
        self.ruta_celulares = rutas[1]
        self.ruta_hoja = rutas[1] + "/DEUDAS_VENCIDAS_HOJA.xlsx"
        self.ruta_fichero = rutas[1] + "/DEUDAS_VENCIDAS_FICHERO.xlsx"
        self.sap = rutas[1] + "/SAP.xlsx"
        self.ruta_vacaciones = rutas[2] + "/VACACIONES.xlsx"
        self.ruta_apoyos = rutas[2] + "/APOYOS_CCD.xlsx"
        
        if not os.path.exists(rutas[1] + "/REPORTE FINAL"):
            os.makedirs(rutas[1] + "/REPORTE FINAL")
            
        self.ruta_final = rutas[1] + "/REPORTE FINAL/DEUDAS_VENCIDAS_" + fecha_actual + ".xlsx"
        self.ruta_final_apoyos = rutas[1] + "/REPORTE FINAL/DEUDAS_VENCIDAS_APOYOS_" + fecha_actual + ".xlsx"
    
    def exportar_deudores(self):
        df_dacxanalista = pd.read_excel(self.ruta_dacxa, sheet_name="Base_NUEVA")
        df_dacxanalista[["DEUDOR"]].to_excel(self.sap, index=False)
        os.startfile(self.sap)
    
    def fichero_local(self):
        df_base = pd.read_excel(self.ruta_fichero)
        df_base = df_base.iloc[:, 3:]
        df_base = df_base.iloc[7:, :]
        df_base = df_base.drop(df_base.index[1])
        df_base = df_base.iloc[:-3, :]
        df_base.columns = df_base.iloc[0]
        df_base = df_base[1:]
        self.df_base = df_base.rename(columns={"     Importe en ML": "Importe"})
    
    def obtener_deudas_vencidas(self, analista, formato, dias_morosidad, lista_estados, apoyo):
        start = time.time()
        try:
            df_dacxanalista = pd.read_excel(self.ruta_dacxa, sheet_name="Base_NUEVA", usecols=["DEUDOR", "NOMBRE", "ANALISTA_ACT", "ESTADO"])
            df_dacxanalista["DEUDOR"] = df_dacxanalista["DEUDOR"].astype("Int64").astype(str)
            # Verifica si se debe incluir los apoyos
            if apoyo:
                lista_cartera = df_dacxanalista[df_dacxanalista["ANALISTA_ACT"] == analista]["DEUDOR"].tolist()
                validar_apoyos = Validar_Apoyos(self.ruta_dacxa, self.ruta_vacaciones, self.ruta_apoyos, analista)
                lista_apoyos = validar_apoyos.obtener_deudores()
                df_dacxanalista = df_dacxanalista[df_dacxanalista["DEUDOR"].isin(list(set(lista_cartera + lista_apoyos)))]
                ruta_final = self.ruta_final_apoyos
            else:
                if analista != "TODOS":
                    df_dacxanalista = df_dacxanalista[df_dacxanalista["ANALISTA_ACT"] == analista]
                else:
                    df_dacxanalista = df_dacxanalista[df_dacxanalista["ANALISTA_ACT"] != "SIN INFORMACION"]
                ruta_final = self.ruta_final
            # Indica que estados se deben considerar
            df_dacxanalista = df_dacxanalista[df_dacxanalista["ESTADO"].isin(lista_estados)]
            lista_deudores = df_dacxanalista["DEUDOR"].tolist()
            # Selecciona el formato hoja o fichero
            if formato==False:
                self.fichero_local()
            else:
                self.df_base = pd.read_excel(self.ruta_hoja)
                self.df_base = self.df_base.rename(columns={"Importe en ML": "Importe"})
            
            self.df_base = self.df_base[self.df_base["Cuenta"].notna()]
            self.df_base["Cuenta"] = self.df_base["Cuenta"].astype("Int64").astype(str)
            self.df_base.dropna(subset=["ACC","Cuenta"], inplace=True)
            self.df_base = self.df_base.reset_index(drop=True)
            self.df_base = self.df_base[["ACC", "Cuenta", "Demora", "Importe"]]
            self.df_base["Demora"] = self.df_base["Demora"].astype("Int64")
            self.df_base["Importe"] = self.df_base["Importe"].astype(float)
            self.df_base = self.df_base.reset_index(drop=True)
            
            self.df_base = self.df_base[self.df_base["Cuenta"].isin(lista_deudores)]
            
            self.df_base["Status"] = self.df_base["Importe"].apply(lambda x: "DEUDA" if x > 0 else "SALDOS A FAVOR")
            self.df_base["Tipo Deuda"] = self.df_base["Demora"].apply(lambda x: "CORRIENTE" if x <= 0 else "VENCIDA")
            self.df_base["Saldo Final"] = self.df_base.apply(lambda row: row["Importe"] 
                                                    if (row["Status"]=="DEUDA" and row["Tipo Deuda"]=="VENCIDA") 
                                                    else (row["Importe"] if row["Status"]=="SALDOS A FAVOR" else "NO"), axis=1)
            # Elimina los registros menores a "dias_morosidad" donde la columna "Status" sea "DEUDA" y la columna "Tipo Deuda" sea "VENCIDA"
            self.df_base = self.df_base[~(
                (self.df_base["Demora"] < int(dias_morosidad)) & 
                (self.df_base["Status"] == "DEUDA") & 
                (self.df_base["Tipo Deuda"] == "VENCIDA"))]
            
            self.df_base = self.df_base[self.df_base["Saldo Final"] != "NO"]
            self.df_base = self.df_base.sort_values(by=["Cuenta"], ascending=[True])
            self.df_base = self.df_base.sort_values(by=["ACC"], ascending=[True])
            self.df_base = self.df_base.sort_values(by=["Demora"], ascending=[False])
            self.df_base = self.df_base.sort_values(by=["Cuenta"], ascending=[True])
            self.df_base = self.df_base.reset_index(drop=True)
            
            cuentas_verificadas = []
            ultima_fila = self.df_base.shape[0]
            for i in range(ultima_fila):
                cuenta_actual = self.df_base.loc[i, "Cuenta"]
                if cuenta_actual not in cuentas_verificadas:
                    cuentas_verificadas.append(cuenta_actual)
                    inicio = i
                if self.df_base.loc[i, "Status"] == "DEUDA":
                    saldoDeuda = self.df_base.loc[i, "Saldo Final"]
                    rango = self.df_base[self.df_base['Cuenta'] == cuenta_actual].shape[0]
                    for j in range(inicio, inicio+rango):
                        if (
                            self.df_base.loc[j, "Cuenta"] == cuenta_actual and 
                            self.df_base.loc[j, "ACC"] == self.df_base.loc[i, "ACC"] and 
                            self.df_base.loc[j, "Status"] == "SALDOS A FAVOR"
                            ):
                            saldoFavor = self.df_base.loc[j, "Saldo Final"]
                            montoCompensar = min(saldoDeuda, abs(saldoFavor))
                            self.df_base.loc[i, "Saldo Final"] = saldoDeuda - montoCompensar
                            self.df_base.loc[j, "Saldo Final"] = saldoFavor + montoCompensar
                            saldoDeuda = self.df_base.loc[i, "Saldo Final"]
            
            self.df_base = self.df_base[(self.df_base["Status"] == "DEUDA") & (self.df_base["Tipo Deuda"] == "VENCIDA")]
            self.df_base = self.df_base.reset_index(drop=True)
            grouped_df = self.df_base.groupby(["Cuenta", "ACC"]).agg({"Demora": "max", "Saldo Final": "sum"})
            
            df_final = grouped_df.reset_index()[["Cuenta", "ACC", "Saldo Final", "Demora"]]
            df_final = df_final.rename(columns={"Cuenta": "Cod Cliente", "ACC": "Área Ctrl", 
                                                "Saldo Final": "Deuda Vencida", "Demora": "Días Morosidad"})
            df_final = df_final.merge(df_dacxanalista, left_on="Cod Cliente", right_on="DEUDOR", how="left")
            df_final = df_final.rename(columns={"NOMBRE": "Razón Social", "ANALISTA_ACT": "Analista", "ESTADO": "Estado"})
            df_final.dropna(subset=["Razón Social","Analista","Estado"], inplace=True)
            df_final.drop(columns=["DEUDOR"], inplace=True)
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
            df_final = df_final[df_final["Área Ctrl"].isin(areas_de_control.keys())]
            df_final["Producto"] = df_final["Área Ctrl"].apply(lambda x: areas_de_control.get(x))
            df_final["Código Pago"] = "33" + df_final["Área Ctrl"].str[-2:] + df_final["Cod Cliente"].astype(str)
            df_final = df_final[[
                "Cod Cliente", "Razón Social", "Área Ctrl", "Producto", "Deuda Vencida", 
                "Código Pago", "Días Morosidad", "Analista", "Estado"]]
            df_final["Deuda Vencida"] = df_final["Deuda Vencida"].astype(float).round(2)
            df_final = df_final[df_final["Deuda Vencida"] != 0]
            df_final = df_final.sort_values(by=["Área Ctrl"], ascending=[True])
            df_final = df_final.sort_values(by=["Cod Cliente"], ascending=[True])
            df_final = df_final.sort_values(by=["Deuda Vencida"], ascending=[False])
            df_final = df_final.sort_values(by=["Días Morosidad"], ascending=[False])
            df_final = df_final.reset_index(drop=True)
            df_final.to_excel(ruta_final, index=False)
            formatear_excel(ruta_final)
            end = time.time()
            messagebox.showinfo(
                "Éxito", "Registros encontrados: " + str(df_final.shape[0]) + 
                "\nTiempo de ejecución: " + str(round(end-start,2)) + " segundos.")
            os.startfile(ruta_final)
        except Exception as ex:
            messagebox.showerror("Error", "Algo salió mal. Por favor, intente nuevamente.\nDetalles: " + str(ex))
            return None