import pandas as pd
pd.options.mode.chained_assignment = None


def businessCategory():
    pareto = pd.read_excel("1. BusinessCategory.xlsx", sheet_name="PARETO")
    rbu = pd.read_excel("1. BusinessCategory.xlsx", sheet_name="RBU")
    customer = pd.read_excel("1. BusinessCategory.xlsx", sheet_name="CUSTOMER")
    conversion = pd.read_excel("1. BusinessCategory.xlsx", sheet_name="CONVERSION")
    salescode = pd.read_excel("1. BusinessCategory.xlsx",sheet_name="SALESCODE")
    return pareto, rbu, customer, conversion, salescode



def sales_fy22():
    fy22 = pd.read_excel("22. SALES_FY22.xlsx")
    fy22['AccRbuItem']  = fy22['AccNo'].astype(str) + "_" + fy22['RBU'].astype(str) + "_" + fy22['ItemNumber'].astype(str)
    fy22 = fy22[['AccRbuItem','Jul_FY22','Aug_FY22','Sep_FY22','Oct_FY22','Nov_FY22','Dec_FY22','Jan_FY22',"Feb_FY22"]]
    return fy22


def sales_fy21():
    fy21 = pd.read_excel("21. SALES_FY21.xlsx")
    fy21['AccRbuItem']  = fy21['AccNo'].astype(str) + "_" + fy21['RBU'].astype(str) + "_" + fy21['ItemNumber'].astype(str)
    fy21 = fy21[['AccRbuItem','Jul_FY21','Aug_FY21','Sep_FY21','Oct_FY21','Nov_FY21','Dec_FY21','Jan_FY21','Feb_FY21','Mar_FY21','Apr_FY21','May_FY21','Jun_FY21']]
    return fy21

def sales_fy20():
    fy20 = pd.read_excel("20. SALES_FY20.xlsx")
    fy20['AccRbuItem']  = fy20['AccNo'].astype(str) + "_" + fy20['RBU'].astype(str) + "_" + fy20['ItemNumber'].astype(str)
    fy20 = fy20[['AccRbuItem','Jul_FY20','Aug_FY20','Sep_FY20','Oct_FY20','Nov_FY20','Dec_FY20','Jan_FY20','Feb_FY20','Mar_FY20','Apr_FY20','May_FY20','Jun_FY20']]
    return fy20


def sales_fy19():
    fy19 = pd.read_excel("19. SALES_FY19.xlsx")
    fy19['AccRbuItem']  = fy19['AccNo'].astype(str) + "_" + fy19['RBU'].astype(str) + "_" + fy19['ItemNumber'].astype(str)
    fy19 = fy19[['AccRbuItem','Jul_FY19','Aug_FY19','Sep_FY19','Oct_FY19','Nov_FY19','Dec_FY19','Jan_FY19','Feb_FY19','Mar_FY19','Apr_FY19','May_FY19','Jun_FY19']]
    return fy19

def sales_fy18():
    fy18 = pd.read_excel("18. SALES_FY18.xlsx")
    fy18['AccRbuItem']  = fy18['AccNo'].astype(str) + "_" + fy18['RBU'].astype(str) + "_" + fy18['ItemNumber'].astype(str)
    fy18 = fy18[['AccRbuItem','Jul_FY18','Aug_FY18','Sep_FY18','Oct_FY18','Nov_FY18','Dec_FY18','Jan_FY18','Feb_FY18','Mar_FY18','Apr_FY18','May_FY18','Jun_FY18']]
    return fy18


def sales_fy17():
    fy17 = pd.read_excel("17. SALES_FY17.xlsx")
    fy17['AccRbuItem']  = fy17['AccNo'].astype(str) + "_" + fy17['RBU'].astype(str) + "_" + fy17['ItemNumber'].astype(str)
    fy17 = fy17[['AccRbuItem','Jul_FY17','Aug_FY17','Sep_FY17','Oct_FY17','Nov_FY17','Dec_FY17','Jan_FY17','Feb_FY17','Mar_FY17','Apr_FY17','May_FY17','Jun_FY17']]
    return fy17

def sales_fy16():
    fy16 = pd.read_excel("16. SALES_FY16.xlsx")
    fy16['AccRbuItem']  = fy16['AccNo'].astype(str) + "_" + fy16['RBU'].astype(str) + "_" + fy16['ItemNumber'].astype(str)
    fy16 = fy16[['AccRbuItem','Jul_FY16','Aug_FY16','Sep_FY16','Oct_FY16','Nov_FY16','Dec_FY16','Jan_FY16','Feb_FY16','Mar_FY16','Apr_FY16','May_FY16','Jun_FY16']]
    return fy16

def sales_fy15():
    fy15 = pd.read_excel("15. SALES_FY15.xlsx")
    fy15['AccRbuItem']  = fy15['AccNo'].astype(str) + "_" + fy15['RBU'].astype(str) + "_" + fy15['ItemNumber'].astype(str)
    fy15 = fy15[['AccRbuItem','Jul_FY15','Aug_FY15','Sep_FY15','Oct_FY15','Nov_FY15','Dec_FY15','Jan_FY15','Feb_FY15','Mar_FY15','Apr_FY15','May_FY15','Jun_FY15']]
    return fy15

def sales_fy14():
    fy14 = pd.read_excel("14. SALES_FY14.xlsx")
    fy14['AccRbuItem']  = fy14['AccNo'].astype(str) + "_" + fy14['RBU'].astype(str) + "_" + fy14['ItemNumber'].astype(str)
    fy14 = fy14[['AccRbuItem','Jul_FY14','Aug_FY14','Sep_FY14','Oct_FY14','Nov_FY14','Dec_FY14','Jan_FY14','Feb_FY14','Mar_FY14','Apr_FY14','May_FY14','Jun_FY14']]
    return fy14

def sales_fy13():
    fy13 = pd.read_excel("13. SALES_FY13.xlsx")
    fy13['AccRbuItem']  = fy13['AccNo'].astype(str) + "_" + fy13['RBU'].astype(str) + "_" + fy13['ItemNumber'].astype(str)
    fy13 = fy13[['AccRbuItem','Jul_FY13','Aug_FY13','Sep_FY13','Oct_FY13','Nov_FY13','Dec_FY13','Jan_FY13','Feb_FY13','Mar_FY13','Apr_FY13','May_FY13','Jun_FY13']]
    return fy13


pareto, rbu, customer, conversion, salescode = businessCategory()

fy22 = sales_fy22()
fy21 = sales_fy21()
fy20 = sales_fy20()
fy19 = sales_fy19()
fy18 = sales_fy18()
fy17 = sales_fy17()
fy16 = sales_fy16()
fy15 = sales_fy15()
fy14 = sales_fy14()
fy13 = sales_fy13()




df = pd.merge(fy13,fy14, on="AccRbuItem", how="outer")
df = pd.merge(df,fy15, on="AccRbuItem", how="outer")
df = pd.merge(df,fy16, on="AccRbuItem", how="outer")
df = pd.merge(df,fy17, on="AccRbuItem", how="outer")
df = pd.merge(df,fy18, on="AccRbuItem", how="outer")
df = pd.merge(df,fy19, on="AccRbuItem", how="outer")
df = pd.merge(df,fy20, on="AccRbuItem", how="outer")
df = pd.merge(df,fy21, on="AccRbuItem", how="outer")
df = pd.merge(df,fy22, on="AccRbuItem", how="outer")




def addAttributes(df,pareto, rbu, customer, conversion, salescode):
    df[['AccNo', 'RBU','ItemNumber']] = df["AccRbuItem"].apply(lambda x: pd.Series(str(x).split("_")))

    df['RBU'] = df['RBU'].astype(int)

    df = pd.merge(df,rbu, on="RBU", how="left")
    df = pd.merge(df, pareto[["ItemNumber","Sls Cd3","Sls Cd4","ABC 1 Sls"]], on ='ItemNumber', how='left')
    df = df.fillna(0)

    df['GroupAcc'] = df['Group'].astype(str)+df['AccNo'].astype(str)
    customer['GroupAcc'] = customer['Group'].astype(str)+customer['AccNo'].astype(str)

    df = pd.merge(df,customer[["GroupAcc","AccountName","Customer"]], on="GroupAcc", how="left")

    df['Customer'] = df['Customer'].fillna("Others")
    df = df.fillna(0)


    df = pd.merge(df,conversion[["ItemNumber","Conversion Factor"]], on="ItemNumber", how="left")

    df['Conversion Factor'] = df['Conversion Factor'].fillna(1)

    df.iloc[:,1:-13]=df.iloc[:,1:-13].multiply(df.iloc[:,-1], axis=0)
    
    df.drop_duplicates(subset='AccRbuItem', keep="last", inplace=True)
    

    return df


def groupby_rbu(df):
    rbu = df.drop(['AccNo', 'ItemNumber','Group','Area',"Sls Cd3",'Sls Cd4','ABC 1 Sls','GroupAcc','AccountName','Customer','Conversion Factor','AccRbuItem'], axis=1)
    rbu['Type_RBU'] = rbu['Type'].astype(str)+"_"+rbu['RBU'].astype(str)
    rbu = rbu.drop(['Type','RBU'], axis=1)
    rbu = rbu[[rbu.columns[-1]] + list(rbu.columns[:-1])]
    rbu1 = rbu.pivot_table(index='Type_RBU',aggfunc='sum')
    rbu1 = rbu1.reset_index()
    
    rbu1 = rbu1.reindex(columns=rbu.columns)

    return rbu1

df = addAttributes(df,pareto, rbu, customer, conversion, salescode)
rbu = groupby_rbu(df)
print('Successfully')









