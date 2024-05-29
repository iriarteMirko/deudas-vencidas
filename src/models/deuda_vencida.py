from src.models.formatear_excel import formatear_excel
from tkinter import messagebox
import pandas as pd
import time
import os


def obtener_deudas_vencidas(base_path, dacxanalista_path, resultado_path, variables, analista):
    start = time.time()
    lista_estados = []
    for var, estado in variables:
        if var.get():
            lista_estados.append(estado)
    
    if len(lista_estados) == 0:
        messagebox.showerror("Error", "Por favor, seleccione al menos un estado.")
        return
    try:
        df_base = pd.read_excel(base_path)
        df_base.dropna(subset=["ACC","Cuenta"], inplace=True)
        df_base = df_base.reset_index(drop=True)
        df_base = df_base.rename(columns={"Importe en ML": "Importe"})
        columnas_deseadas = ["ACC", "Cuenta", "Demora", "Importe"]
        df_base = df_base[columnas_deseadas]
        df_base["Demora"] = df_base["Demora"].astype("Int64")
        df_base["Importe"] = df_base["Importe"].astype(float)
        df_base = df_base.reset_index(drop=True)
        
        df_dacxanalista = pd.read_excel(dacxanalista_path, sheet_name="Base_NUEVA")
        columnas_dacx = ["DEUDOR", "NOMBRE", "ANALISTA_ACT", "ESTADO"]
        df_dacxanalista = df_dacxanalista[columnas_dacx]
        
        if analista != "TODOS":
            df_dacxanalista = df_dacxanalista[df_dacxanalista["ANALISTA_ACT"] == analista]
        else:
            df_dacxanalista = df_dacxanalista[df_dacxanalista["ANALISTA_ACT"] != "SIN INFORMACION"]
        
        if len(lista_estados) > 0:
            df_dacxanalista = df_dacxanalista[df_dacxanalista["ESTADO"].isin(lista_estados)]
        
        lista_cartera = df_dacxanalista["DEUDOR"].tolist()
        df_base = df_base[df_base["Cuenta"].isin(lista_cartera)]
        
        df_base["Status"] = df_base["Importe"].apply(lambda x: "DEUDA" if x > 0 else "SALDOS A FAVOR")
        df_base["Tipo Deuda"] = df_base["Demora"].apply(lambda x: "CORRIENTE" if x <= 0 else "VENCIDA")
        df_base["Saldo Final"] = df_base.apply(lambda row: row["Importe"] 
                                                if (row["Status"]=="DEUDA" and row["Tipo Deuda"]=="VENCIDA") 
                                                else (row["Importe"] if row["Status"]=="SALDOS A FAVOR" else "NO"), 
                                                axis=1)
        df_base = df_base[df_base["Saldo Final"] != "NO"]
        df_base = df_base.sort_values(by=["Cuenta"], ascending=[True])
        df_base = df_base.sort_values(by=["ACC"], ascending=[True])
        df_base = df_base.sort_values(by=["Demora"], ascending=[False])
        df_base = df_base.sort_values(by=["Cuenta"], ascending=[True])
        df_base = df_base.reset_index(drop=True)
        
        cuentas_verificadas = []
        ultima_fila = df_base.shape[0]
        for i in range(ultima_fila):
            cuenta_actual = df_base.loc[i, "Cuenta"]
            if cuenta_actual not in cuentas_verificadas:
                cuentas_verificadas.append(cuenta_actual)
                inicio = i
            if df_base.loc[i, "Status"] == "DEUDA":
                saldoDeuda = df_base.loc[i, "Saldo Final"]
                rango = df_base[df_base['Cuenta'] == cuenta_actual].shape[0]
                for j in range(inicio, inicio+rango):
                    if (
                        df_base.loc[j, "Cuenta"] == cuenta_actual and 
                        df_base.loc[j, "ACC"] == df_base.loc[i, "ACC"] and 
                        df_base.loc[j, "Status"] == "SALDOS A FAVOR"
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
        df_final = df_final.rename(columns={"Cuenta": "Cod Cliente", "ACC": "Área Ctrl", 
                                            "Saldo Final": "Deuda Vencida", "Demora": "Días Morosidad"})
        df_final = df_final.merge(df_dacxanalista[["DEUDOR", "NOMBRE", "ANALISTA_ACT", "ESTADO"]], 
                                    left_on="Cod Cliente", right_on="DEUDOR", how="left")
        df_final = df_final.rename(columns={"NOMBRE": "Razón Social", "ANALISTA_ACT": "Analista", "ESTADO": "Estado"})
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
        df_final = df_final[df_final["Área Ctrl"].isin(areas_de_control.keys())]
        df_final["Producto"] = df_final["Área Ctrl"].apply(lambda x: areas_de_control.get(x))
        df_final["Código Pago"] = "33" + df_final["Área Ctrl"].str[-2:] + df_final["Cod Cliente"].astype(str)
        df_final = df_final[["Cod Cliente", "Razón Social", "Área Ctrl", "Producto", "Deuda Vencida", 
                            "Código Pago", "Días Morosidad", "Analista", "Estado"]]
        df_final["Deuda Vencida"] = df_final["Deuda Vencida"].astype(float).round(2)
        df_final = df_final[df_final["Deuda Vencida"] != 0]
        df_final = df_final.sort_values(by=["Área Ctrl"], ascending=[True])
        df_final = df_final.sort_values(by=["Cod Cliente"], ascending=[True])
        df_final = df_final.sort_values(by=["Deuda Vencida"], ascending=[False])
        df_final = df_final.sort_values(by=["Días Morosidad"], ascending=[False])
        df_final.to_excel(resultado_path, index=False)
        
        formatear_excel(resultado_path)
        end = time.time()
        messagebox.showinfo("Éxito", "Registros encontrados: " + str(df_final.shape[0]) + 
                            "\nTiempo de ejecución: " + str(round(end-start,2)) + " segundos.")
        os.startfile(resultado_path)
    except Exception as ex:
        messagebox.showerror("Error", "Algo salió mal. Por favor, intente nuevamente.\nDetalles: " + str(ex))