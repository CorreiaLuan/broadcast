
from datetime import datetime
import xlwings as xw
import pandas as pd
import time
import os


class xlAddin:
    def __init__(self,visible = False,addin_path = None) -> None:
        """
        Initializes the xlAddin object to manage Excel automation via xlwings.

        Parameters:
            visible (bool): Whether the Excel application should be visible.
            addin_path (str): Path to the Excel add-in. Can be set via an environment variable.
        """
        if addin_path is None:
            self.addin_path = os.getenv('BROADCAST_ADDIN_PATH',None)
        else:
            self.addin_path = addin_path
        self.app = None
        self.visible = visible
        if self.addin_path is None:
            raise ValueError("addin_path cannot be None, please provide a value or put it in 'BROADCAST_ADDIN_PATH' environment variable")

    def initiate(self) ->xw.App:
        """
        Initializes an Excel application instance with the given add-in.

        Returns:
            xw.App: The Excel application object.
        """
        try:
            app = xw.App(visible=self.visible,add_book=False)
            app.api.RegisterXLL(self.addin_path) # Registers the Excel add-in.
            self.app = app
            return app  
        except Exception as e:
            print(e)
            app.kill()

    def wait(self,cell:xw.Range,timeout:int = 10):
        """
        Waits for an Excel formula to resolve.

        Parameters:
            cell (xw.Range): The Excel cell or range to monitor.
            timeout (int): Maximum time to wait in seconds.
        """
        count = 0
        start = time.time()
        while True and count < timeout:
            if not isinstance(cell.value,list):
                if cell.value not in ["#BCONN","#VALUE!",None]: break
            else:
                if not any(item in cell.value for item in ["#BCONN","#VALUE!",None]): break
            count = time.time() - start

    def do_things(self,app:xw.App,formula:str,range = False):
        """
        Executes a formula in Excel and retrieves the result.

        Parameters:
            app (xw.App): The Excel application instance.
            formula (str): The formula to execute.
            range (bool): Whether to return the range instead of the value.

        Returns:
            list or xw.Range: Formula results or the Excel range object.
        """
        wb = app.books.add()
        sht:xw.Sheet = wb.sheets[0]
        rng:xw.Range = sht.range("A1")
        if isinstance(formula,list):
          rng = rng.resize(len(formula))  
          formula = [[x] for x in formula]
        rng.formula2 = formula
        self.wait(rng)
        if not range:
            return sht.used_range.value
        return sht.used_range

    def bc(self,ativo:str="",campos:list="",data:dict=None,dataframe = False,date_columns=""):
        """
        Usa a Fórmula do Excel BC para buscar um campo de um Ativo\n
        Ex.: bc("Ibov","ult") retorna a última Cotação do Ibov\n
        Caso o parâmetro data (dict) seja passado, então ativo e campos serão ignorados
        """
        with self.initiate() as app:
            if data is None:
                if isinstance(campos,list):campos = ";".join(campos)
                return self.do_things(app,f"""=BC("{ativo}","{campos}")""")
            old_data = data.copy()
            data = [f"""=BC("{key}","{";".join(value) if isinstance(value,list) else value}")"""  for key,value in data.items()]
            df = self.do_things(app,data)
        if dataframe:
            columns = [subitem for x in old_data.values() for subitem in x.split(";")]
            columns = ['ativo'] + list(dict.fromkeys(columns))
            df = [ [key.upper()] + row  for key,row in zip(old_data.keys(),df)]
            return self.handle_date(
                    pd.DataFrame(
                        data = df,
                        columns = columns
                    ).replace("N/A",None),
                    date_columns
            )
        return df
    
    def handle_date(self,df,date_columns):
        if date_columns == "" and "DRF" in df.columns: date_columns = "DRF"
        if date_columns == "" and "drf" in df.columns: date_columns = "drf"
        if date_columns != "":
            df.loc[:,date_columns] = df[date_columns].apply(self.convert_excel_date)
        return df
    
    def bch(self,ativo:str,campos:list="",data_inicial="",data_final="",parametros_opcionais:list="",date_columns:list=""):
        if isinstance(campos,list):campos = ";".join(campos)
        if isinstance(parametros_opcionais,list):parametros_opcionais = ";".join(parametros_opcionais)
        if parametros_opcionais != "": parametros_opcionais = f',"{parametros_opcionais}"'
        data_inicial = self.datetime_to_excel_date(data_inicial)
        data_final = self.datetime_to_excel_date(data_final)
        with self.initiate() as app:
            result = self.do_things(
                app,
                f"""=BCH("{ativo}","{campos}",{data_inicial},{data_final}{parametros_opcionais})""",
                True
            )
            columns = campos.split(";")
            df:pd.DataFrame = result.options(
                pd.DataFrame,
                header = 0,
                index = False,
                expand = 'table',
                dates = datetime.date
            ).value
            df.columns = columns
            df.insert(0,"ativo",ativo.upper())
            if "drf" in campos and campos not in date_columns:
                if date_columns == "":
                    date_columns = "drf"
                elif isinstance(date_columns,str):
                    date_columns = [date_columns,"drf"]
                elif isinstance(date_columns,list):
                    date_columns.append("drf")
            return self.handle_date(df,date_columns)
    
    def convert_excel_date(self,date):
        initial_date = datetime(1900,1,1).toordinal()
        if isinstance(date,tuple):
            return (None if pd.isna(dt) else datetime.fromordinal(initial_date + int(dt) -2).date() for dt in date)
        return None if pd.isna(date) else datetime.fromordinal(initial_date + int(date) -2).date()

    def datetime_to_excel_date(self,dt: str) -> float:
        """
        Converts a str date object to an Excel date (serial number).
        
        Parameters:
            dt (str): The date string to convert on format yyyy-mm-dd.
        
        Returns:
            float: The Excel date as a serial number.
        """
        # Excel's base date
        dt = datetime.strptime(dt,'%Y-%m-%d')
        excel_base_date = datetime(1899, 12, 30)
        delta = dt - excel_base_date
        return delta.days + delta.seconds / 86400  # days + fractional day


if __name__ == '__main__':
    broad = xlAddin()
    px = broad.bch(ativo="dapk35",campos=["ult;drf"],data_inicial="2024-11-15",data_final="2024-11-26")
    print(px)