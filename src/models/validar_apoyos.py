from datetime import datetime
import pandas as pd

class Validar_Apoyos():
    def __init__(self, ruta_dacx, ruta_vacaciones, ruta_apoyos, analista):
        self.ruta_dacx = ruta_dacx
        self.ruta_vacaciones = ruta_vacaciones
        self.ruta_apoyos = ruta_apoyos
        self.analista = analista
    
    def formato_fecha(self, fecha):
        return datetime.strptime(fecha, "%d.%m.%Y")
    
    def obtener_analistas(self,):
        df_vacaciones = pd.read_excel(self.ruta_vacaciones, sheet_name="VACACIONES")
        df_vacaciones.dropna(inplace=True)
        
        df_vacaciones["FECHA_SALIDA"] = pd.to_datetime(df_vacaciones["FECHA_SALIDA"], format="%d.%m.%Y")
        df_vacaciones["FECHA_RETORNO"] = pd.to_datetime(df_vacaciones["FECHA_RETORNO"], format="%d.%m.%Y")
        
        fecha_hoy = pd.to_datetime(datetime.now().strftime("%d.%m.%Y"), format="%d.%m.%Y")
        
        self.vacaciones = df_vacaciones[
            (df_vacaciones["FECHA_SALIDA"] <= fecha_hoy) & (df_vacaciones["FECHA_RETORNO"] >= fecha_hoy)
        ]["ANALISTA"].tolist()
    
    def depurar_dacxanalista(self,):
        columnas = ["DEUDOR", "NOMBRE", "RUC", "ANALISTA_ACT", "TIPO_DAC", "ESTADO"]
        no_analistas = ["REGION NORTE", "REGION SUR", "WALTER LOPEZ", "SIN INFORMACION"]
        no_estados = ["LIQUIDADO", "PROCESO DE LIQUIDACIÓN"]
        
        df_dacx = pd.read_excel(self.ruta_dacx, sheet_name="Base_NUEVA", usecols=columnas)
        df_dacx.rename(columns={"ANALISTA_ACT": "ANALISTA", "TIPO_DAC": "TIPO"}, inplace=True)
        df_dacx = df_dacx[~df_dacx["ANALISTA"].isin(no_analistas) & ~df_dacx["ESTADO"].isin(no_estados)]
        df_dacx["DEUDOR"] = df_dacx["DEUDOR"].astype("Int64").astype(str)
        df_dacx["RUC"] = df_dacx["RUC"].astype("Int64").astype(str)
        df_dacx.dropna(subset=["DEUDOR"], inplace=True)
        df_dacx.drop_duplicates(subset=["DEUDOR"], inplace=True)
        df_dacx.reset_index(drop=True, inplace=True)
        self.df_dacx = df_dacx
        self.list_analistas = df_dacx["ANALISTA"].unique()
    
    def actualizar_apoyos(self,):
        self.depurar_dacxanalista()
        df_apoyos = pd.read_excel(self.ruta_apoyos, sheet_name="GENERAL", usecols=["DEUDOR", "APOYO1", "APOYO2", "APOYO3"])
        df_apoyos["DEUDOR"] = df_apoyos["DEUDOR"].astype("Int64").astype(str)
        df_apoyos = pd.merge(df_apoyos, self.df_dacx, on="DEUDOR", how="right")
        df_apoyos = df_apoyos[["DEUDOR", "NOMBRE", "ANALISTA", "APOYO1", "APOYO2", "APOYO3", "ESTADO", "TIPO"]]
        df_apoyos.dropna(subset=["DEUDOR"], inplace=True)
        df_apoyos.drop_duplicates(subset=["DEUDOR"], inplace=True)
        df_apoyos.sort_values(by="DEUDOR", inplace=True, ignore_index=True)
        df_apoyos.reset_index(drop=True, inplace=True)
        df_apoyos.to_excel(self.ruta_apoyos, sheet_name="GENERAL", index=False)
        
        with pd.ExcelWriter(self.ruta_apoyos) as writer:
            df_temp = df_apoyos.copy()
            df_apoyos.to_excel(writer, sheet_name="GENERAL", index=False)
            for analista in self.list_analistas:
                df_apoyos = df_temp[df_temp["ANALISTA"] == analista]
                df_apoyos.sort_values(by="DEUDOR", inplace=True, ignore_index=True)
                df_apoyos.reset_index(drop=True, inplace=True)
                df_apoyos.to_excel(writer, sheet_name=analista, index=False)
    
    def obtener_deudores(self,):
        self.obtener_analistas()
        self.actualizar_apoyos()
        columnas = ["DEUDOR", "ANALISTA", "APOYO1", "APOYO2", "APOYO3"]
        df_apoyos = pd.read_excel(self.ruta_apoyos, sheet_name="GENERAL", usecols=columnas)
        df_apoyos["DEUDOR"] = df_apoyos["DEUDOR"].astype("Int64").astype(str)
        df_apoyos.dropna(subset=["DEUDOR"], inplace=True)
        df_apoyos = df_apoyos[df_apoyos["ANALISTA"].isin(self.vacaciones)]
        df_apoyos["APOYO"] = df_apoyos[["APOYO1", "APOYO2", "APOYO3"]].apply(
            lambda x: x[0] if x[0] not in (self.vacaciones) 
            else (x[1] if x[1] not in (self.vacaciones) else x[2]), axis=1)
        df_apoyos = df_apoyos[df_apoyos["APOYO"]==self.analista]
        df_apoyos.reset_index(drop=True, inplace=True)
        return df_apoyos["DEUDOR"].to_list()
