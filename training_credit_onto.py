import pandas as pd
import glob
import os

directories = ["log_o","log_sap"]
oracle_items = ["Customer Name","Customer PO", "Product Family", "Region", "Order Quantity", "Line Total(USD)","Promise Date"]
sap_items = ["Sold-To Party Name", "Order Quantity (Item)","Net Value (Item)","Delivery Date"]
def load_data():
    to_return = []
    for i in directories:
        paths= glob.glob(os.path.join(i, "*xlsx"))
        if not paths:
            raise FileNotFoundError(i, "File not found")
        for path in paths:
            xl = pd.ExcelFile(path)
            for sheet in xl.sheet_names: 
                df = pd.read_excel(path, sheet_name=sheet)
        to_return.append(df)
    return to_return

def get_credits(data):
    log_o, log_sap = data[0],data[1]
    credits = log_o.drop(log_o[log_o['Part Number'] != '778752'].index, inplace=False)
    training_items = credits["Customer PO"].tolist()
    training_orders = log_o[log_o['Customer PO'].isin(training_items)]

    log_sap["Material"] = log_sap["Material"].astype(str)
    credits_sap = log_sap.drop(log_sap[log_sap['Material'] != '778752'].index, inplace=False)


    cleaned_orders = training_orders.loc[~(training_orders['Part Number'] == '778752')].filter(oracle_items)
    cleaned_orders["Quarter"] = pd.to_datetime(cleaned_orders["Promise Date"],dayfirst=True).dt.quarter
    cleaned_orders["Promise Date"] = (pd.to_datetime(cleaned_orders["Promise Date"],dayfirst=True)).dt.date

    cleaned_orders_sap = credits_sap.filter(sap_items)
    cleaned_orders_sap["Quarter"] = pd.to_datetime(cleaned_orders_sap["Delivery Date"],dayfirst=True).dt.quarter
    cleaned_orders_sap["Delivery Date"] = pd.to_datetime(cleaned_orders_sap["Delivery Date"],dayfirst=True).dt.date

    with pd.ExcelWriter("training_items.xlsx") as writer:
        cleaned_orders.to_excel(writer, sheet_name="oracle", index=False)
        cleaned_orders_sap.to_excel(writer, sheet_name="sap", index=False)

get_credits(load_data())



