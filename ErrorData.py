import pandas as pd


class ErrorInformation:

    def __init__(self, path=""):
        self.path = path

    def invoke_pandas(self):
        df = pd.read_excel(self.path)
        ErrorInfo = ""
        for index, row in df.iterrows():
            ErrorInfo += row["型号"] + ","
        return ErrorInfo


ErrorInformation(r"C:\Users\windo\Desktop\JUNIP - PFMLN210078 - 檢-出单错误.xlsx").invoke_pandas()
