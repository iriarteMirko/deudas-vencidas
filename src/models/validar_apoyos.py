from datetime import datetime
import pandas as pd

class Class_Validar_Apoyos():
    def __init__(self, ruta_vacaciones):
        self.ruta_vacaciones = ruta_vacaciones
    
    def formato_fecha(self, fecha):
        return datetime.strptime(fecha, '%d.%m.%Y')
    
    def obtener_analistas(self,):
        df_vacaciones = pd.read_excel(self.ruta_vacaciones, sheet_name='VACACIONES')
        df_vacaciones.dropna(inplace=True)
        
        dict_analistas = dict(zip(df_vacaciones['ANALISTA'], zip(df_vacaciones['FECHA_SALIDA'], df_vacaciones['FECHA_RETORNO'])))
        
        vacaciones = {}
        fecha_hoy = datetime.now().strftime('%d.%m.%Y')
        
        for analista, fechas in dict_analistas.items():
            if self.formato_fecha(fechas[0]) <= self.formato_fecha(fecha_hoy) <= self.formato_fecha(fechas[1]):
                print(f'{analista} de vacaciones desde {fechas[0]} hasta {fechas[1]}')
                vacaciones.update({analista: fechas[1]})
        
        if vacaciones == {}:
            return []
        else:
            return list(vacaciones.keys())
