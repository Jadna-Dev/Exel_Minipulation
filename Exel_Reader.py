import pandas as pd

#Pandas Is The Library That We Use To Minipulate Data From The Exel Spreadsheet

#👇Open The Exel Sheet
data = pd.read_excel("./Data.xlsx")

#You Can Specify The Sheet Index using the "sheet_name" attr,
#Also You Can Specify The Data Type Of Each Column Using The "dtype" attr as Shown Below 
data1 = pd.read_excel("./Data.xlsx",sheet_name = 0,dtype={'item_no': int,})


#This Function Is The Defualt Untouched (Takes   First SpreadSheet By Defualt )
def defualtSheet():
    return data

#This Function returns The First Row in the spreadsheet as indexes (Column Title)
def getColumns():
    return data.columns

#This Function returns Row Data as Lists
def getRawData():
    return data.values.view()

#Examples:

#This Function returns The Total Items In Stock
def getTotalStock():
    j = data.values.view()
    total = []
    for items in j:
        tempt = int(int(items[3])-int(items[4]))
        total.append([items[0],items[1],tempt])
    return total

#This Function returns The Total Sales Per Item
def getTotalSalesPerItem():
    j = data.values.view()
    total = []
    for items in j:
        #You can round the number as shown below
        tempt = round(float(float(items[2]) * float(items[4])),3)
        
        
        total.append([items[0],items[1],tempt])
    return total

#This Function returns The Total Sales (all items)
def getTotalSales():
    j = data.values.view()
    total = float(0)
    for items in getTotalSalesPerItem():
        total = total + float(items[2])
        
        #You can format the amount to be human readable as shown below
    return f'{total:,}'
        
    

print("This Is The defualtSheet :\n ",defualtSheet())
print("\n")
print("This Is The getColumns :\n ",getColumns())
print("\n")
print("This Is The getRawData :\n ",getRawData())
print("\n")
print("This Is The getTotalStock :\n ",getTotalStock())
print("\n")
print("This Is The getTotalSalesPerItem :\n ",getTotalSalesPerItem())
print("\n")
print("This Is The getTotalSales :\n ",getTotalSales())