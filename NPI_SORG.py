import pandas as pd
import xlwings as xw
import time
sorg = pd.read_csv(r"C:\Users\liewkmbr\Python_Projects\SORG_ZZ.csv")
sorglist = list(sorg["Material No"])
dclist = list(sorg["Sales Org"])



indx = pd.MultiIndex.from_product([sorglist,dclist], names = ["sorglist","dclist"])
indx1 = indx.dropna()
#print(indx1)  

#xlwings to capture pandas dataframe into excel
wb = xw.Book(r"C:\Users\liewkmbr\Python_Projects\sorg.xlsx")
sht1 = wb.sheets["Sheet1"]
sht1.range("A1:B1000").value = indx1
wb.save()
wb.close()

#using pandas to clean up data and get ready for the SAP update format
sorgzz = pd.read_excel(r"C:\Users\liewkmbr\Python_Projects\sorg.xlsx")
sorgzzclean = sorgzz.dropna()
sorgzzclean.to_csv(r"C:\Users\liewkmbr\Python_Projects\sorg.csv",index = False)


#time.sleep(30)

#sorgzz = pd.read_excel(r"C:\Users\liewkmbr\Python_Projects\sorg.xlsx")
#print (sorgzz)

#index2 = index1.reset_index(inplace=True)
#options(dropna=True).
#sorg_update = index1.to_csv(r"C:\Users\liewkmbr\Downloads\sorg_update.csv")

#print(index1)    
