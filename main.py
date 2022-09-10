#PANDAS
import pandas as pd
from pandas import ExcelFile

#UTILIDADES
import datetime
import re

#LOGGER 
##FILE 
import logging

formato = ' [%(asctime)s], %(levelname)s, %(message)s, %(filename)s:%(lineno)d'
logging.basicConfig( filename='loggers.log', level=logging.INFO, format=formato)

##CONSOLE 
console = logging.StreamHandler()
console.setLevel(logging.DEBUG)
formatter = logging.Formatter('%(name)-12s: %(levelname)-8s %(message)s')
console.setFormatter(formatter)
logging.getLogger('').addHandler(console)

logger = logging.getLogger(__name__)



class Read_file:
    def __init__(self, url:str):
        self.url = url
        self.list_sheet_name = self.__get_sheet_from_file('Ref_Hoja_Excel')
        self.unidad = self.__get_sheet_from_file('Nemo_Grupo')

    def __get_sheet_from_file(self, column_name):
        self.read_file = pd.read_excel(
            self.url, na_values="Missing", sheet_name=0, header=1)
        self.name_sheet = self.read_file[column_name].unique()
        return self.name_sheet

    def get_list_sheet_name(self):
        return self.list_sheet_name

    def get_unidad(self):
        return self.unidad


class newDF(Read_file):
    def __init__(self, url, sheetname, columns_df):
        logger.info("Inicializando la clase newDF...")
        super().__init__(url)
        self.read_excel_file()
        self.columns_df = columns_df
        logger.info(f"Se crearon las columnas correctamente: {columns_df}")
    
    
    def read_excel_file(self):
        """
        Read File

        - Description:
            - Read excel type files and get the name of the sheets to be able to iterate through each of them.

        - Parameters:
            - path_file:str : path to the location of the file to read.

        - Return:
            - Returns a list with the names of the sheets.
        """
        logger.info('Intentando cargar la hoja')
        try:
            self.sheetname = self.get_list_sheet_name()
            logger.info(f'Se cargo correctamente la Hoja: {self.sheetname}')
            logger.info(f'Leyendo la Hoja: {self.sheetname}')
            self.read_sheet_file = pd.read_excel(
                self.url, na_values="Missing", sheet_name=self.sheetname[0], header=4)
            logger.info(f'Se leyo correctamente la Hoja: {self.sheetname}')            
            return self.read_sheet_file
        except TypeError as e:
            logger.error(f'Ocurrio un error: {e}')

    def __get_data_from_url(self):
        try:
            logger.info(f'Obteniendo informacion del path: {self.url}')
            logger.info(f'Obteniendo la fecha del path...')
            self.date_in_url = re.findall('[0-9]{6}', self.url)
            self.months = self.date_in_url[0][4:6]
            self.years = self.date_in_url[0][:4]
            logger.info('Se obtuvieron exitosamente las fechas.') 
        except TypeError as e:
            logger.error(f'Ocurrio un error: {e}')

    def get_the_current_quarter(self):
        logger.info('Intentando obtener el trimestre')
        self.__get_data_from_url()
        
        try:
            self.date_file = datetime.datetime(
                int(self.years), int(self.months), 1, 0, 0)
            self.date = datetime.datetime(int(self.years), int(
                self.months) - (int(self.months) - 1) % 3, 1, 0, 0)
            self.dates = pd.date_range(
                self.date, periods=3, freq='MS', normalize=False)
            self.dates = [date for date in self.dates if date <= self.date_file]
            return self.dates
        except TypeError as e:
            logger.error(f'Ocurrio un error: {e}')


    def filter_data_by_quarters(self):
        self.get_the_current_quarter()
        try:
            self.df_quaters = self.read_sheet_file[[
                self.read_sheet_file.columns[0], *self.dates]].dropna(subset=[self.read_sheet_file.columns[0]])
            return self.df_quaters
        except:
            return self.read_sheet_file.dropna(subset=[self.read_sheet_file.columns[0]])
            print("no tiene las columnas")

    def create_dataframe(self):
        self.filter_data_by_quarters()
        self.columnR = self.df_quaters.columns[0]
        self.column_list = [
            col for col in self.df_quaters.columns if col != self.columnR]
        self.df = self.df_quaters.melt(
            id_vars=self.columnR, value_vars=self.column_list, value_name="VALOR")
        self.df_var = self.df[self.df.columns[1]]
        self.df['MES'] = self.df_var.dt.month
        self.df['AÑO'] = self.df_var.dt.year
        self.df["INICIO DEL MES"] = self.df_var.dt.day
        self.df['UNIDAD'] = self.unidad[0]
        self.df.drop(self.df.columns[1], inplace=True, axis=1)
        self.df.rename(columns={self.columnR: "R"}, inplace=True)
        self.df = self.df.reindex(columns=self.columns_df)
        return self.df


if __name__ == "__main__":
    path_file = './ReporteN2-ACHI-Real-202103.xlsx'
    logger.info(f'Leyendo el Archivo: {path_file[2:]}' )
    exc_file = Read_file(path_file)
    logger.info('Obteniendo la lista de nombre de hojas')
    num_of_sheet = exc_file.get_list_sheet_name()

    columns = ["AÑO", "MES", "UNIDAD", "R", "VALOR", "INICIO DEL MES"]

    logger.info('Comenzando a Iterar por cada hoja')
    for sheet in range(len(num_of_sheet)):
        logger.info(f'Sheet {sheet+1}/{len(num_of_sheet)}')
        df = newDF(path_file,
                    num_of_sheet[sheet], columns)
        print(df.create_dataframe())
