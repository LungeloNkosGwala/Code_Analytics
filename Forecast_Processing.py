import pandas as pd
import time
pd.options.mode.chained_assignment = None
from alive_progress import alive_bar


def load_df():
    try:
        df =  pd.read_excel("2. Local_Forecast.xlsx", sheet_name = "SALES")
        FC =  pd.read_excel("2. Local_Forecast.xlsx", sheet_name = "FORECAST")
        pareto = pd.read_excel("1. BusinessCategory.xlsx", sheet_name = "PARETO")
        soh = pd.read_excel("3. SOH_ORDERS_WO.xlsx", sheet_name = "SOH")
        orders = pd.read_excel("3. SOH_ORDERS_WO.xlsx", sheet_name = "OPEN_ORDERS")
        wo = pd.read_excel("3. SOH_ORDERS_WO.xlsx", sheet_name = "WO")
        message = True
    except:
        df = 0
        FC = 0
        pareto = 0
        soh = 0 
        wo = 0
        orders = 0
        message = False
    return df,FC,pareto,soh,orders,wo,message



def prepare_files(soh,orders,wo):
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

        orders.rename(columns={"BU\nHeader":"RBU","Item\nNumber":"ItemNumber","Qty\nOrder/\nTransaction":"Ordered_Qty","Qty\nBackorder/\nHeld":"BackOrder_Qty"}, inplace=True)
        orders = orders[["RBU","ItemNumber","Ordered_Qty","BackOrder_Qty"]]
        orders['RBU_ItemNumber'] = orders['RBU'].astype(str) +"_"+ orders['ItemNumber'].astype(str)
        orders['Ordered_Qty'] = orders['Ordered_Qty'].astype(int)
        orders['BackOrder_Qty'] = orders['BackOrder_Qty'].astype(int)


        orders = orders[['RBU_ItemNumber',"Ordered_Qty","BackOrder_Qty"]]
        orders = orders.pivot_table(index = "RBU_ItemNumber", aggfunc = "sum")
        orders = orders.reset_index()

        wo = wo[['ITEM NUMBER','[QTY OUTSTANDING]']]
        wo.rename(columns = {"ITEM NUMBER":"ItemNumber","[QTY OUTSTANDING]":"WO_Outstanding_Qty"}, inplace = True)
        wo['WO_Outstanding_Qty'] = wo['WO_Outstanding_Qty'].astype(int)
        wo = wo.pivot_table(index = "ItemNumber", aggfunc= "sum")
        wo = wo.reset_index()
        message = True
    except:
        wo = 0
        soh = 0
        orders =0
        message=False

    return soh,orders,wo,message


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

def concat1(Max,soh,orders,wo):
    Max["RBU_ItemNumber"] = Max['RBU'].astype(str)+"_"+Max['ItemNumber'].astype(str)
    Max =  pd.merge(Max,soh, on = 'ItemNumber', how="left")
    Max = pd.merge(Max,orders, on = "RBU_ItemNumber", how="left")
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
        message=True
    except:
        message=False
    return ALL,message

def concat2(ALL,pareto):
    #Max = pd.merge(Max,df[['ItemNumber','Group']], on="ItemNumber", how='left')
    result = pd.concat(ALL)
    group = result.drop(result.iloc[:,3:9],axis=1)
    group = group.drop(['ItemNumber','Max',"6mAVG","3mAVG","BackOrder_Qty","WO_Outstanding_Qty","Ordered_Qty","soh","RBU_ItemNumber","RBU"], axis=1)
    #group = group[[group.columns[-1]] + list(group.columns[:-1])]
    group1 = group.pivot_table(index='Group',aggfunc='sum')
    group1 = group1.reset_index()
    group1 = group1.reindex(columns=group.columns)
    
    result = pd.merge(result, pareto[["ItemNumber","Sls Cd3","Sls Cd4","ABC 1 Sls"]], on ='ItemNumber', how='left')
    result = result.round(0)
    group1 = group1.round(0)
    return result,group1
a = 1
for i in range(a):
    with alive_bar(i, title="Forecasting", theme='smooth') as bar:
        df,FC,pareto,soh,orders,wo,message = load_df()
        if message == False:
            print("Error: File failed to load")
            time.sleep(7)
        else:
            print("File loaded successfully")
            soh,orders,wo,message = prepare_files(soh,orders,wo)
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
                    Max = concat1(Max,soh,orders,wo)
                    ALL,name,message= separate(Max)
                    if message==False:
                        print('Error: Separation Failed')
                        time.sleep(7)
                    else:
                        print("File separation successfully")
                        ALL,message = forecast(ALL,name,FC)
                        if message == False:
                            print("Error: Forecasting failed")
                            time.sleep(7)
                        else:
                            print("Foresting was successful")
                            result,group1 = concat2(ALL,pareto)
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

