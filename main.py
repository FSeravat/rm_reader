from modules import excel
import os
import pandas

try:
    files = list(filter(lambda x:os.path.splitext(x)[1]==".xlsx" and os.path.splitext(x)[0]!="resultado",os.listdir()))
    data = {"RM":[],"Código":[],"Descrição":[],"DN":[],"Código PDMS":[],"Código SAP":[]}
    qtdFiles = files.__len__()
    count = 1
    print(str(qtdFiles)+" arquivos encontrados")

    for file in files:
        print("Abrindo arquivo: "+file+" "+str(count)+"/"+str(qtdFiles))
        doc = excel.Excel(file)
        print(doc.data)
        data["RM"].extend(doc.data["RM"])
        data["Código"].extend(doc.data["Código"])
        data["Descrição"].extend(doc.data["Descrição"])
        data["DN"].extend(doc.data["DN"])
        data["Código PDMS"].extend(doc.data["Código PDMS"])
        data["Código SAP"].extend(doc.data["Código SAP"])
        count+=1

    df = pandas.DataFrame(data)
    df.to_excel("resultado.xlsx", index=False)
except Exception as e:
    print(e)