import pandas as pd
import time
import numpy as np
pd.options.mode.chained_assignment = None
from datetime import timedelta,datetime,date
from alive_progress import alive_bar


def load_df():
    try:
        df =  pd.read_excel("2. Local_Forecast.xlsx", sheet_name = "SALES")
        FC =  pd.read_excel("2. Local_Forecast.xlsx", sheet_name = "FORECAST")
        pareto = pd.read_excel("1. BusinessCategory.xlsx", sheet_name = "PARETO")
        soh = pd.read_excel("3. SOH_ORDERS_WO.xlsx", sheet_name = "SOH")
        orders = pd.read_excel("3. SOH_ORDERS_WO.xlsx", sheet_name = "OPEN_ORDERS")
        wo = pd.read_excel("3. SOH_ORDERS_WO.xlsx", sheet_name = "WO")
        sit = pd.read_excel("3. SOH_ORDERS_WO.xlsx",sheet_name = "SIT")
        sohb = pd.read_excel("3. SOH_ORDERS_WO.xlsx", sheet_name = "SOH_B")
        
        message = True
    except:
        df = 0
        FC = 0
        pareto = 0
        soh = 0 
        wo = 0
        orders = 0
        message = False
    return df,FC,pareto,soh,orders,wo,message,sit,sohb



def prepare_files(soh,orders):
    try:
        drop = soh[(soh['SALES DIV']=="P10")|(soh['SALES DIV']=="S10")|(soh['SALES DIV']=="S20")|(soh['SALES DIV']=="S50")].index
        soh.drop(drop , inplace=True)
        soh['SALES DIV'] = soh['SALES DIV'].astype(int)
        soh['QTY ON HAND'] = soh['QTY ON HAND'].str.replace(",","").astype(int)
        soh['QTY ON HAND'] = soh['QTY ON HAND'].astype(int)
        soh = soh[['ITEM NUMBER','QTY ON HAND']]
        soh.rename(columns = {"ITEM NUMBER":"ItemNumber","QTY ON HAND":"soh"}, inplace=True)
        soh = soh.pivot_table(index = "ItemNumber", aggfunc = "sum")
        soh = soh.reset_index()

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

        message = True
    except:
        wo = 0
        soh = 0
        orders =0
        message=False

    return soh,orders,message,bo

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
    wo = wo[["ItemNumber",'ByPass', 'Current','Curr+1','Curr+2','Curr+3','Curr+4','Curr+5']]
    wo = wo.pivot_table(index = "ItemNumber", aggfunc= "sum")
    wo = wo.reset_index()
    

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

    return sohb
    

def calculate_avg(df):
    try:
        df['6mAVG']= df.iloc[:,3:9].mean(axis=1)
        df['3mAVG']= df.iloc[:,6:-1].mean(axis=1)
        df["Max"] = df[["6mAVG", "3mAVG"]].max(axis=1)
        Max = df
        #Max = df[['ItemNumber',"RBU",'Group',"6mAVG","3mAVG",'Max']]
        Max['Max'] = Max['Max']/4
        message=True
    except:
        message=False
        Max = 0
    return Max, message

def concat1(Max,soh,orders,wo,sit,pareto):
    Max["RBU_ItemNumber"] = Max['RBU'].astype(str)+"_"+Max['ItemNumber'].astype(str)
    Max = pd.merge(Max, pareto[["ItemNumber","Sls Cd3","Sls Cd4","ABC 1 Sls"]], on ='ItemNumber', how='left')
    Max = pd.merge(Max,soh, on = 'ItemNumber', how="left")
    Max = pd.merge(Max,sohb, on = 'ItemNumber', how="left")
    Max = pd.merge(Max,bo, on = "RBU_ItemNumber", how="left")
    Max = pd.merge(Max,orders, on = "RBU_ItemNumber", how="left")
    Max = pd.merge(Max, sit, on = "ItemNumber",how="left")
    Max = pd.merge(Max, wo, on = "ItemNumber", how="left")
    Max = Max.fillna(0)
    return Max
     

def separate(Max):
    try:
        GUD_LOCAL = Max[Max['Group']=='GUD_LOCAL']
        FRAM_LOCAL = Max[Max['Group']=='FRAM_LOCAL']
        SAF_LOCAL = Max[Max['Group']=='SAF_LOCAL']
        MPT_LOCAL = Max[Max['Group']=='MPT_LOCAL']
        FMO_LOCAL = Max[Max['Group']=='FMO_LOCAL']

        ALL = [GUD_LOCAL,FRAM_LOCAL,SAF_LOCAL,MPT_LOCAL,FMO_LOCAL]
        name = ['GUD_LOCAL','FRAM_LOCAL','SAF_LOCAL','MPT_LOCAL','FMO_LOCAL']
        message=True
    except:
        message=False
    return ALL,name,message
   

def forecast(ALL,name,FC):
    try:
        na = len(ALL)
        for t in range(na):
            for i in range(24):
                i = str(i)
                ALL[t]['F'+i.format(str(i))] = FC.loc[FC['Forecast_P']==name[t]]["F"+i.format(str(i))].values*ALL[t]['Max']

        result = pd.concat(ALL)
        message=True
    except:
        message=False
    return result,message

def concat2(result):
    #Max = pd.merge(Max,df[['ItemNumber','Group']], on="ItemNumber", how='left')

    result.iloc[result['Sls Cd3'].isin(["DIS", 980]), 35:59] = 0

    group = result.drop(['ItemNumber',"RBU"], axis = 1)
    group = group.drop(group.iloc[:,1:33],axis=1)
    

    group1 = group.pivot_table(index='Group',aggfunc='sum')
    group1 = group1.reset_index()
    group1 = group1.reindex(columns=group.columns)
    

    result = result.round(0)
    group1 = group1.round(0)
    return result,group1

def reason(result):

    def problem(row):

        #164 Reason
        if row['B_164'] <0 and (row['BO_164']*-1 > row['B_164'] or row['BO_164']*-1 == row['B_164']) and row['ST_164'] > 0 and row['ByPass'] == 0:
            return "S3| @164/Demand (BO)/supply in ST"
        elif row['B_164'] <0 and (row['BO_164']*-1 > row['B_164'] or row['BO_164']*-1 == row['B_164']) and row['ByPass'] > 0 and row['ST_164'] == 0:
            return "S5| @164/Demand (BO)/supply in ByPass"
        elif row['B_164'] <0 and (row['BO_164']*-1 > row['B_164'] or row['BO_164']*-1 == row['B_164']) and row['ST_164'] > 0 and row['ByPass'] > 0:
            return "S3| @164/Demand (BO)/supply in ST and ByPass"

        elif row['B_164'] <0 and (row['BO_164']*-1 > row['B_164'] or row['BO_164']*-1 == row['B_164']) and row['Current']+row['Curr+1']+row['Curr+2']+row['Curr+3']+row['Curr+4'] > 0 and row['Curr+5']==0:
            return "S4| @164/Demand (BO)/Supply WO Sched within 5WKS"
        
        elif row['B_164'] <0 and (row['BO_164']*-1 > row['B_164'] or row['BO_164']*-1 == row['B_164']) and row['Current']+row['Curr+1']+row['Curr+2']+row['Curr+3']+row['Curr+4'] == 0 and row['Curr+5']>0:
            return "S4| @164/Demand (BO)/Supply WO Sched after 5WKS"

        elif row['B_164'] <0 and (row['BO_164']*-1 > row['B_164'] or row['BO_164']*-1 == row['B_164']) and row['ST_164'] == 0 and row['Current']+row['Curr+1']+row['Curr+2']+row['Curr+3']+row['Curr+4'] + row['Curr+5'] == 0:
            return "S5| @164/Demand (BO)/No Supply insight"
        else:
            return "NA"

        

    result['Reason'] = result.apply( lambda row : problem(row), axis = 1)

    return result
a = 1
for i in range(a):
    with alive_bar(i, title="Forecasting", theme='smooth') as bar:
        df,FC,pareto,soh,orders,wo,message,sit,sohb = load_df()
        if message == False:
            print("Error: File failed to load")
            time.sleep(7)
        else:
            print("File loaded successfully")
            soh,orders,message,bo = prepare_files(soh,orders)
            wo = wo_prep(wo)
            sit = sit_prep(sit)
            sohb = sohb_prep(sohb)
            if message == False:
                print("Error: File preparation failed, please check files")
                time.sleep(7)
            else:
                Max,message = calculate_avg(df)
                if message==False:
                    print("Error: Averages Calculation Failed, check months columns")
                    time.sleep(7)
                else:
                    print("Averages calculated Successfully")
                    Max = concat1(Max,soh,orders,wo,sit,pareto)
                    ALL,name,message= separate(Max)
                    if message==False:
                        print('Error: Separation Failed')
                        time.sleep(7)
                    else:
                        print("File separation successfully")
                        result,message = forecast(ALL,name,FC)
                        if message == False:
                            print("Error: Forecasting failed")
                            time.sleep(7)
                        else:
                            print("Foresting was successful")
                            result,group1 = concat2(result)
                            result = reason(result)
                            try:
                                with pd.ExcelWriter('FORECAST_NUMBERS.xlsx') as writer:
                                    group1.to_excel(writer, sheet_name='Group',index=False)
                                    result.to_excel(writer, sheet_name='ALL',index=False)
                                message=True
                            except:
                                message=False

                            if message == False:
                                print("Error: Cannot write into an Open File")
                                time.sleep(7)
                            else: 
                                print("Forecast File successfully Exported!")
                                time.sleep(5)
        bar()

