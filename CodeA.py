import pandas as pd
pd.options.mode.chained_assignment = None


def businessCategory():
    pareto = pd.read_excel("1. BusinessCategory.xlsx", sheet_name="PARETO")
    rbu = pd.read_excel("1. BusinessCategory.xlsx", sheet_name="RBU")
    customer = pd.read_excel("1. BusinessCategory.xlsx", sheet_name="CUSTOMER")
    conversion = pd.read_excel("1. BusinessCategory.xlsx", sheet_name="CONVERSION")
    salescode = pd.read_excel("1. BusinessCategory.xlsx",sheet_name="SALESCODE")
    return pareto, rbu, customer, conversion, salescode


def sales_fy21():
    fy21 = pd.read_excel("3. SALES_FY21.xlsx")
    fy21['AccRbuItem']  = fy21['AccNo'].astype(str) + "_" + fy21['RBU'].astype(str) + "_" + fy21['ItemNumber'].astype(str)
    fy21 = fy21[['AccRbuItem','Jul_FY21','Aug_FY21','Sep_FY21','Oct_FY21','Nov_FY21','Dec_FY21','Jan_FY21','Feb_FY21','Mar_FY21','Apr_FY21','May_FY21','Jun_FY21']]
    return fy21

def sales_fy22():
    fy22 = pd.read_excel("4. SALES_FY22.xlsx")
    fy22['AccRbuItem']  = fy22['AccNo'].astype(str) + "_" + fy22['RBU'].astype(str) + "_" + fy22['ItemNumber'].astype(str)
    fy22 = fy22[['AccRbuItem','Jul_FY22','Aug_FY22','Sep_FY22','Oct_FY22','Nov_FY22','Dec_FY22','Jan_FY22']]
    return fy22

def budget():
    budget = pd.read_excel("2. BUDGET.xlsx")
    budget['AccRbuItem']  = budget['AccNo'].astype(str) + "_" + budget['RBU'].astype(str) + "_" + budget['ItemNumber'].astype(str)
    budget = budget[["AccRbuItem",'Jul_FY22','Aug_FY22','Sep_FY22','Oct_FY22','Nov_FY22','Dec_FY22','Jan_FY22','Feb_FY22','Mar_FY22','Apr_FY22','May_FY22','Jun_FY22']]
    budget['QTR1_B22'] = budget['Jul_FY22']+budget['Aug_FY22']+budget['Sep_FY22']
    budget['QTR2_B22'] = budget['Oct_FY22']+budget['Nov_FY22']+budget['Dec_FY22']
    budget['QTR3_B22'] = budget['Jan_FY22']+budget['Feb_FY22']+budget['Mar_FY22']
    budget['QTR4_B22'] = budget['Apr_FY22']+budget['May_FY22']+budget['Jun_FY22']
    budget =  budget[["AccRbuItem",'QTR1_B22','QTR2_B22','QTR3_B22','QTR4_B22']]
    return budget

def openOrders():
    openOrder = pd.read_excel("5. OPEN_ORDERS.xlsx")
    openOrder = openOrder[["RBU","AccNo","ItemNumber","OrderQty","Salescode"]]
    openOrder['AccRbuItem'] =  openOrder['AccNo'].astype(str) + "_" + openOrder['RBU'].astype(str) + "_" + openOrder['ItemNumber'].astype(str)
    return openOrder




def merge(fy21,fy22,budget,openOrder):
    df = pd.merge(fy21,fy22, on="AccRbuItem", how="outer")
    df = df.fillna(0)
    df['QTR1_FY21'] = df['Jul_FY21']+df['Aug_FY21']+df['Sep_FY21']
    df['QTR2_FY21'] = df['Oct_FY21']+df['Nov_FY21']+df['Dec_FY21']
    df['QTR3_FY21'] = df['Jan_FY21']+df['Feb_FY21']+df['Mar_FY21']
    df['QTR4_FY21'] = df['Apr_FY21']+df['May_FY21']+df['Jun_FY21']
    df['QTR1_FY22'] = df['Jul_FY22']+df['Aug_FY22']+df['Sep_FY22']
    df['QTR2_FY22'] = df['Oct_FY22']+df['Nov_FY22']+df['Dec_FY22']

    df = pd.merge(df,budget, on='AccRbuItem', how='left')
    df = pd.merge(df,openOrder[["AccRbuItem","OrderQty"]],on='AccRbuItem', how='left')
    df = df.fillna(0)

    return df

    
    
    
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

    
    
                  
    return df
    

pareto, rbu, customer, conversion, salescode = businessCategory()
fy21 = sales_fy21()
fy22 = sales_fy22()
budget = budget()
openOrder = openOrders()

df = merge(fy21,fy22,budget,openOrder)
df = addAttributes(df,pareto, rbu, customer, conversion, salescode)

df.to_excel("df.xlsx")
