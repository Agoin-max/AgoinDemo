import pandas as pd
import re


class Place:

    def __init__(self, path=""):
        self.path = path

    def invoke_pandas(self):
        df = pd.read_excel(self.path, header=1, dtype=str)
        for index, row in df.iterrows():
            if row["产地"].find("x") > 0:
                parttern = re.compile(r"\d+", re.S)
                List_Info = re.findall(parttern, row["产地"])
                if "\n" in row["产地"]:
                    pass
                if " " in row["产地"]:
                    place = row["产地"].rsplit(" ", 1)[1].strip().split("x")[1]
                    df.loc[index, ["产地", "数量", "件数", "净重", "毛重"]] = [row["产地"].rsplit(" ", 1)[0].strip().split("x")[1],
                                                                     List_Info[0],
                                                                     float(List_Info[0]) / float(row["数量"]) * float(
                                                                         row["件数"]),
                                                                     float(List_Info[0]) / float(row["数量"]) * float(
                                                                         row["净重"]),
                                                                     float(List_Info[0]) / float(row["数量"]) * float(
                                                                         row["毛重"])]
                    Data = {"商品类型": row["商品类型"],
                            "商品小类": row["商品小类"],
                            "品牌": row["品牌"],
                            "型号": row["型号"],
                            "商品描述": row["商品描述"],
                            "产地": place,
                            "单位": row["单位"],
                            "数量": List_Info[1],
                            "报关单价": row["报关单价"],
                            "件数": float(List_Info[1]) / float(row["数量"]) * float(row["件数"]),
                            "净重": float(List_Info[1]) / float(row["数量"]) * float(row["净重"]),
                            "毛重": float(List_Info[1]) / float(row["数量"]) * float(row["毛重"]),
                            "税号": row["税号"],
                            "SKU": row["SKU"],
                            "供应商": row["供应商"],
                            "期票天数": row["期票天数"],
                            "对应的采购": row["对应的采购"],
                            "料号": row["料号"],
                            "托盘数": row["托盘数"],
                            "箱号": row["箱号"],
                            "备注": row["备注"],
                            "收款方": row["收款方"],
                            "支付方式": row["支付方式"],
                            "关务品名": row["关务品名"]
                            }
                    df1 = df.loc[:index]
                    df2 = df.loc[index + 1:]
                    NewInsert = pd.DataFrame(Data, index=[9])
                    df1 = df1.append(NewInsert)
                    df = df1.append(df2, ignore_index=True)

        df.to_excel(r"C:\Users\windo\Desktop\PFSH21A0024- 20,870p 8A Y210226 -出单1.xlsx", index=False, startrow=1)


ins = Place(r"C:\Users\windo\Desktop\PFSH21A0024- 20,870p 8A Y210226 -出单.xlsx")
ins.invoke_pandas()
