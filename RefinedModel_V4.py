import pandas as pd
import time
import numpy as np
pd.options.mode.chained_assignment = None
from datetime import timedelta,datetime,date


def load_df():

    df =  pd.read_excel("2. Local_Forecast.xlsx", sheet_name = "SALES12M")
    mtd = pd.read_excel("2. Local_Forecast.xlsx", sheet_name = "MTD")
    #FC =  pd.read_excel("2. Local_Forecast.xlsx", sheet_name = "FORECAST")
    pareto = pd.read_excel("1. BusinessCategory.xlsx", sheet_name = "PARETO")
    rbuArea = pd.read_excel("1. BusinessCategory.xlsx", sheet_name = "RBU")
    #soh = pd.read_excel("3. SOH_ORDERS_WO.xlsx", sheet_name = "SOH")
    orders = pd.read_excel("3. SOH_ORDERS_WO.xlsx", sheet_name = "OPEN_ORDERS")
    wo = pd.read_excel("3. SOH_ORDERS_WO.xlsx", sheet_name = "WO")
    sit = pd.read_excel("3. SOH_ORDERS_WO.xlsx",sheet_name = "SIT")
    sohb = pd.read_excel("3. SOH_ORDERS_WO.xlsx", sheet_name = "SOH_B")
    con = pd.read_excel("1. BusinessCategory.xlsx", sheet_name = "CONVERSION")
        
    return df,pareto,orders,wo,sit,sohb,mtd,rbuArea,con

def orders_prep(orders):
    orders.rename(columns={"BU\nHeader":"RBU","Item\nNumber":"ItemNumber","Qty\nOrder/\nTransaction":"Ordered_Qty","Qty\nBackorder/\nHeld":"BackOrder_Qty",'Branch/\nPlant':"Branch"}, inplace=True)
    bo = orders[["RBU","ItemNumber","BackOrder_Qty","Branch"]]
    
    orders = orders[["RBU","ItemNumber","Ordered_Qty","BackOrder_Qty"]]
    orders['RBU_ItemNumber'] = orders['RBU'].astype(str) +"_"+ orders['ItemNumber'].astype(str)
    orders['Ordered_Qty'] = orders['Ordered_Qty'].astype(int)
    orders['BackOrder_Qty'] = orders['BackOrder_Qty'].astype(int)
    orders = orders[['RBU_ItemNumber',"Ordered_Qty","BackOrder_Qty"]]
    orders = orders.pivot_table(index = "RBU_ItemNumber", aggfunc = "sum")
    orders = orders.reset_index()



    bo['Branch'] = bo['Branch'].str.rsplit(' ',expand=True)[0]
    bo = bo[(bo["Branch"]=="164")|(bo["Branch"]=="166")|(bo["Branch"]=="168")]
    bo["RBU_ItemNumber"] = bo["RBU"].astype(str) + "_" + bo["ItemNumber"].astype(str)
    bo = bo.pivot_table(index="RBU_ItemNumber", columns="Branch", values="BackOrder_Qty",aggfunc=np.sum).rename_axis(index=None, columns=None)
    bo = bo.fillna(0)
    bo = bo.round(0)
    bo.reset_index(inplace=True)
    bo.rename(columns = {"index":"RBU_ItemNumber","164":"BO_164","166":"BO_166","168":"BO_168"}, inplace = True)

    return orders,bo

def wo_prep(wo):
    wo['[SCHEDULED DATE]'] = pd.to_datetime(wo['[SCHEDULED DATE]'])
    wo['Year'] = wo['[SCHEDULED DATE]'].dt.year
    wo['Year'] = wo['Year'].astype(int)
    wo['Month'] = wo['[SCHEDULED DATE]'].dt.month
    wo['Month'] = wo['Month'].astype(int)
    wo['weekofyear'] = wo['[SCHEDULED DATE]'].dt.isocalendar().week
    wo['weekofyear'] = wo['weekofyear'].astype(int)
    wo['Today'] = date.today()
    wo['TodayW'] = wo['Today'].apply(lambda x:x.isocalendar()[1])
    wo['TodayY'] = pd.to_datetime(wo['Today'])
    wo['TodayY'] =  wo['TodayY'].dt.year

    conditions = [
        (wo['Year'] < wo['TodayY']),
        (wo['Year'] == wo['TodayY']) & (wo['weekofyear'] < wo['TodayW']),
        (wo['Year'] == wo['TodayY']) & (wo['weekofyear'] == wo['TodayW']),
        (wo['Year'] == wo['TodayY']) & (wo['weekofyear'] == wo['TodayW']+1),
        (wo['Year'] == wo['TodayY']) & (wo['weekofyear'] == wo['TodayW']+2),
        (wo['Year'] == wo['TodayY']) & (wo['weekofyear'] == wo['TodayW']+3),
        (wo['Year'] == wo['TodayY']) & (wo['weekofyear'] == wo['TodayW']+4),
        (wo['Year'] == wo['TodayY']) & (wo['weekofyear'] > wo['TodayW']+4),
        (wo['Year'] > wo['TodayY']) & (wo['weekofyear'] == wo['TodayW']+1),
        (wo['Year'] > wo['TodayY']) & (wo['weekofyear'] == wo['TodayW']+2),
        (wo['Year'] > wo['TodayY']) & (wo['weekofyear'] == wo['TodayW']+3),
        (wo['Year'] > wo['TodayY']) & (wo['weekofyear'] == wo['TodayW']+4),
        (wo['Year'] > wo['TodayY']) & (wo['weekofyear'] > wo['TodayW']+4)]

    values = ['ByPass', 'ByPass', 'Current', 'Curr+1','Curr+2','Curr+3','Curr+4','Curr+5','Curr+1','Curr+2','Curr+3','Curr+4','Curr+5']

    
    wo['WO_Status'] = np.select(conditions, values)

    cols = ['ByPass', 'Current', 'Curr+1','Curr+2','Curr+3','Curr+4','Curr+5']

    for i in range(7):

        condition = [wo['WO_Status']== cols[i]]
        val = [wo['[QTY OUTSTANDING]']]

        wo[cols[i]] = np.select(condition,val)
        
    wo = wo.fillna(0)

    wo.rename(columns = {"ITEM NUMBER":"ItemNumber"}, inplace = True)
    wo['ItemNumber'] = wo['ItemNumber'].astype(str)
    wo1 = wo[["ItemNumber",'ByPass', 'Current','Curr+1','Curr+2','Curr+3','Curr+4','Curr+5']]
    wo = wo1.pivot_table(index = "ItemNumber", aggfunc= "sum")
    wo = wo.reset_index()
    wo = wo.reindex(columns=wo1.columns)
    wo['InMonth']= wo['Curr+1']+wo['Curr+2']+wo['Curr+3']+wo['Curr+4']
    
    return wo

def sit_prep(sit):
    sit = sit[["Item Number","P/O Bus Unit","ST Shipped Qty"]]
    sit['P/O Bus Unit'] = sit['P/O Bus Unit'].astype(int)
    sit = sit[(sit["P/O Bus Unit"]==164)|(sit["P/O Bus Unit"]==166)|(sit["P/O Bus Unit"]==168)]
    sit = sit.pivot_table(index="Item Number", columns="P/O Bus Unit", values="ST Shipped Qty",aggfunc=np.sum).rename_axis(index=None, columns=None)
    sit = sit.fillna(0)
    sit = sit.round(0)
    sit.reset_index(inplace=True)
    sit.rename(columns = {"index":"ItemNumber",164:"ST_164",166:"ST_166",168:"ST_168"}, inplace = True)

    return sit


def sohb_prep(sohb):
    sohb = sohb[["Branch/\nPlant","Location","Item\nNumber","Available"]]
    sohb.rename(columns = {"Branch/\nPlant":"Branch","Item\nNumber":"Item"}, inplace = True)

    hq = sohb[(sohb['Location']=="H")|(sohb['Location']=="Q")|(sohb['Location']=="INQ")]

    fact = sohb[(sohb['Branch']=="104")|(sohb['Branch']=="205")|(sohb['Branch']=="595")|(sohb['Branch']=="699")|(sohb['Branch']=="801")]
    dis = sohb[(sohb['Branch']=="161")|(sohb['Branch']=="262")|(sohb['Branch']=="696")|(sohb['Branch']=="803")]

    sohb = sohb.set_index("Location")
    
    sohb = sohb.drop(['H','Q','INQ'])

    sohb.reset_index(inplace=True)

    sohb = sohb[(sohb["Branch"]=="164")|(sohb["Branch"]=="166")|(sohb["Branch"]=="168")]

    sohb['ItemNumber'] = sohb['Item'].str.rsplit(' ',expand=True)[0]

    sohb = sohb.pivot_table(index="ItemNumber", columns="Branch", values="Available",aggfunc=np.sum).rename_axis(index=None, columns=None)
    sohb = sohb.fillna(0)
    sohb = sohb.round(0)
    sohb.reset_index(inplace=True)

    sohb.rename(columns = {"index":"ItemNumber","164":"B_164","166":"B_166","168":"B_168"}, inplace = True)

    return sohb, hq, fact, dis

def qaulityHold(hq):
    hq = hq[(hq["Branch"]=="164")|(hq["Branch"]=="166")|(hq["Branch"]=="168")]
    hq['ItemNumber'] = hq['Item'].str.rsplit(' ',expand=True)[0]
    hq = hq.pivot_table(index="ItemNumber", columns="Branch", values="Available",aggfunc=np.sum).rename_axis(index=None, columns=None)
    hq = hq.fillna(0)
    hq = hq.round(0)
    hq.reset_index(inplace=True)

    hq.rename(columns = {"index":"ItemNumber","164":"HQ_164","166":"HQ_166","168":"HQ_168"}, inplace = True)

    return hq


def Distribute(dis):
    dis =dis[dis['Available'] !=0]
    dis[['BR', 'Dest','idd']] = dis["Location"].apply(lambda x: pd.Series(str(x).split("|")))
    dis = dis[(dis['Dest']=="164")|(dis['Dest']=="166")|(dis['Dest']=="168")]
    dis  = dis[["Dest","Item","Available"]]
    dis['ItemNumber'] = dis['Item'].str.rsplit(' ',expand=True)[0]
    dis = dis.pivot_table(index="ItemNumber", columns="Dest", values="Available",aggfunc=np.sum).rename_axis(index=None, columns=None)
    dis = dis.fillna(0)
    dis = dis.round(0)
    dis.reset_index(inplace=True)
    dis.rename(columns = {"index":"ItemNumber","164":"D_164","166":"D_166","168":"D_168"}, inplace = True)
    return dis

def factory(fact):
    fact =fact[fact['Available'] !=0]
    fact[['BR', 'Dest','idd']] = fact["Location"].apply(lambda x: pd.Series(str(x).split("|")))
    fact = fact[(fact['Dest']=="164")|(fact['Dest']=="166")|(fact['Dest']=="168")]
    fact  = fact[["Dest","Item","Available"]]
    fact['ItemNumber'] = fact['Item'].str.rsplit(' ',expand=True)[0]
    fact = fact.pivot_table(index="ItemNumber", columns="Dest", values="Available",aggfunc=np.sum).rename_axis(index=None, columns=None)
    fact = fact.fillna(0)
    fact = fact.round(0)
    fact.reset_index(inplace=True)
    fact.rename(columns = {"index":"ItemNumber","164":"F_164","166":"F_166","168":"F_168"}, inplace = True)
    return fact


def calculate_avg(df):
    
    df['12mAVG']= df.iloc[:,4:16].mean(axis=1)
    df['6mAVG']= df.iloc[:,10:16].mean(axis=1)
    df['3mAVG']= df.iloc[:,13:16].mean(axis=1)
    df["Max"] = df[["12mAVG","6mAVG", "3mAVG"]].max(axis=1)
    Max = df
    Max = Max.round(0)


    return Max

def getMTD(mtd):
    mtd['RBU_ItemNumber'] = mtd['CATEGORY CODE 01'].astype(str)+"_"+mtd['Item Number'].astype(str)
    mtd['MTD'] = mtd.iloc[:,4:16].sum(axis=1)
    mtd = mtd[["RBU_ItemNumber","MTD"]]
    mtd = mtd.pivot_table(index="RBU_ItemNumber", values="MTD",aggfunc=np.sum).rename_axis(index=None, columns=None)
    mtd = mtd.fillna(0)
    mtd = mtd.round(0)
    mtd.reset_index(inplace=True)
    mtd.rename(columns = {"index":"RBU_ItemNumber"}, inplace = True)
    return mtd

def concat1(Max,orders,wo,sit,pareto,hq,dis,bo,sohb,fact,mtd,rbuArea,con):

    rbuArea['RBU'] = rbuArea['RBU'].astype(str)
    Max["RBU_ItemNumber"] = Max['RBU'].astype(str)+"_"+Max['ItemNumber'].astype(str)
    Max = pd.merge(Max,mtd, on = "RBU_ItemNumber", how="outer")
    Max[['RBU', 'ItemNumber']] = Max["RBU_ItemNumber"].apply(lambda x: pd.Series(str(x).split("_")))

    
    Max = pd.merge(Max, rbuArea[["RBU","Group","GROUP_AREA","Type"]], on = "RBU", how = "left")
    
    Max = pd.merge(Max, pareto[["ItemNumber","Sls Cd3","Sls Cd4","ABC 1 Sls"]], on ='ItemNumber', how='left')
    Max = pd.merge(Max,orders, on = "RBU_ItemNumber", how="left")
    Max = pd.merge(Max,bo, on = "RBU_ItemNumber", how="left")
    #Max = pd.merge(Max,soh, on = 'ItemNumber', how="left")
    Max = pd.merge(Max,sohb, on = 'ItemNumber', how="left")
    Max = pd.merge(Max,hq, on = "ItemNumber", how="left")
    Max = pd.merge(Max,dis, on = "ItemNumber", how='left')
    Max = pd.merge(Max, fact, on = "ItemNumber",how='left')


    Max = pd.merge(Max, sit, on = "ItemNumber",how="left")
    Max = pd.merge(Max, wo, on = "ItemNumber", how="left")
    #Max = pd.merge(Max, cost, on = "ItemNumber", how='left')
    result = Max.fillna(0)


    result = pd.merge(result,con[["ItemNumber","Conversion Factor"]], on="ItemNumber", how="left")

    result['Conversion Factor'] = result['Conversion Factor'].fillna(1)

    result.iloc[:,3:20]=result.iloc[:,3:20].multiply(result.iloc[:,-1], axis=0)
    result.iloc[:,26:54]=result.iloc[:,26:54].multiply(result.iloc[:,-1], axis=0)

    
    
    return result


def reason(result):
    branch = ['164',"166",'168']
    
    for i in branch:
        def stock(row):
            if row["B_"+i.format(i)]> 0:
                return "Stock/"
            else:
                return "NoStock/"
        def bo(row):
            if row['BO_'+i.format(i)]> 0:
                return "BO/"
            else:
                return ""
        def hq(row):
            if row['HQ_'+i.format(i)]> 0:
                return "HQ/"
            else:
                return ""
        def depo(row):
            if row["D_"+i.format(i)] > 0:
                return "Depo/"
            else:
                return ""
        def fact(row):
            if row["F_"+i.format(i)] > 0:
                return "Fact/"
            else:
                return ""
        def sit(row):
            if row["ST_"+i.format(i)] >0:
                return "SIT/"
            else:
                return ""
        def bypass(row):
            if row['ByPass'] > 0:
                return "ByPassWO/"
            else:
                return ""
        def current(row):
            if row['Current']>0:
                return "CurrentWO/"
            else:
                return ""
        def inMonth(row):
            if row['InMonth']>0:
                return "InMonthWO/"
            else:
                return ""
        def forwardDue(row):
            if row['Curr+5']>0:
                return "ForwardDueWO/"
            else:
                return ""

        
                
        result['R_STOCK'+i.format(i)] = result.apply( lambda row : stock(row), axis = 1)
        result['R_BO'+i.format(i)] = result.apply( lambda row : bo(row), axis = 1)
        result['R_HQ'+i.format(i)] = result.apply( lambda row : hq(row), axis = 1)
        result['R_DEPO'+i.format(i)] = result.apply( lambda row : depo(row), axis = 1)
        result['R_SIT'+i.format(i)] = result.apply( lambda row : sit(row), axis = 1)
        result['R_BYPASS'+i.format(i)] = result.apply( lambda row : bypass(row), axis = 1)
        result['R_CURRENT'+i.format(i)] = result.apply( lambda row : current(row), axis = 1)
        result['R_INMONTH'+i.format(i)] = result.apply( lambda row : inMonth(row), axis = 1)
        result['R_FORWARDDUE'+i.format(i)] = result.apply( lambda row : forwardDue(row), axis = 1)
        result['R_FACT'+i.format(i)] = result.apply( lambda row : fact(row), axis = 1)
        
        result['Reason_'+i.format(i)] = result['R_STOCK'+i.format(i)].astype(str)+result['R_BO'+i.format(i)].astype(str)+result['R_HQ'+i.format(i)].astype(str)+ result['R_SIT'+i.format(i)].astype(str)
        result['Reason_'+i.format(i)] = result['Reason_'+i.format(i)].astype(str)+result['R_DEPO'+i.format(i)].astype(str)+ result['R_FACT'+i.format(i)].astype(str) +result['R_BYPASS'+i.format(i)].astype(str)
        result['Reason_'+i.format(i)] = result['Reason_'+i.format(i)].astype(str)+result['R_CURRENT'+i.format(i)].astype(str)+result['R_INMONTH'+i.format(i)].astype(str)
        result['Reason_'+i.format(i)] = result['Reason_'+i.format(i)].astype(str)+result['R_FORWARDDUE'+i.format(i)].astype(str)

        result = result.drop(['R_STOCK'+i.format(i),'R_FACT'+i.format(i) ,'R_BO'+i.format(i),'R_HQ'+i.format(i),'R_DEPO'+i.format(i),'R_SIT'+i.format(i),'R_BYPASS'+i.format(i),'R_CURRENT'+i.format(i),'R_INMONTH'+i.format(i),'R_FORWARDDUE'+i.format(i)], axis=1)
        

    return result

def executeONE():
    df,pareto,orders,wo,sit,sohb,mtd,rbuArea,con = load_df()
    orders,bo = orders_prep(orders)
    wo = wo_prep(wo)
    sit = sit_prep(sit)
    sohb,hq,fact,dis = sohb_prep(sohb)
    hq = qaulityHold(hq)
    dis = Distribute(dis)
    fact = factory(fact)
    Max = calculate_avg(df)
    mtd = getMTD(mtd)
    result = concat1(Max,orders,wo,sit,pareto,hq,dis,bo,sohb,fact,mtd,rbuArea,con)
    result = reason(result)

    return result
result = executeONE()
with pd.ExcelWriter('Demand&Supply_BO.xlsx') as writer:
    result.to_excel(writer, sheet_name='Data',index=False)










