import pandas as pd
# import xlsxwriter
import glob
import os

directories = ["log_o", "log_sap"]
oracle_items = ["Customer Name", "Customer PO", "Product Family",
                "Part Number", "Region", "Order Quantity", "Line Total(USD)", "Promise Date"]
sap_items = ["Sold-To Party Name",
             "Order Quantity (Item)","Material", "Net Value (Item)", "Delivery Date"]


def load_data():
    to_return = []
    for i in directories:
        paths = glob.glob(os.path.join(i, "*xlsx"))
        if not paths:
            raise FileNotFoundError(i, "File not found")
        for path in paths:
            xl = pd.ExcelFile(path)
            df = pd.read_excel(path, sheet_name=xl.sheet_names[0])
        to_return.append(df)
    return to_return


def get_credits(data):
    log_o = data[0]
    if len(data) > 0:
        log_sap = data[1]
    credits = log_o.drop(
        log_o[log_o['Part Number'] != '778752'].index, inplace=False)
    training_items = credits["Customer PO"].tolist()
    training_orders = log_o[log_o['Customer PO'].isin(training_items)]

    print(log_sap)
    log_sap["Material"] = log_sap["Material"].astype(str)
    # credits_sap = log_sap
    # credits_sap = log_sap.drop(log_sap[log_sap['Material'] != '778752'].index, inplace=False)

    # cleaned_orders = training_orders.loc[~(training_orders['Part Number'] == '778752')].filter(oracle_items)
    cleaned_orders = training_orders[oracle_items]
    cleaned_orders["Quarter"] = pd.to_datetime(
        cleaned_orders["Promise Date"], dayfirst=True).dt.quarter
    cleaned_orders["Promise Date"] = (pd.to_datetime(
        cleaned_orders["Promise Date"], dayfirst=True)).dt.date
    cleaned_orders = cleaned_orders[
        (cleaned_orders['Line Total(USD)'] >= 1) |
        (cleaned_orders['Part Number'] == '778752')
    ]

    cleaned_orders_sap = log_sap[sap_items]
    cleaned_orders_sap["Quarter"] = pd.to_datetime(
        cleaned_orders_sap["Delivery Date"], dayfirst=True).dt.quarter
    cleaned_orders_sap["Delivery Date"] = pd.to_datetime(
        cleaned_orders_sap["Delivery Date"], dayfirst=True).dt.date
    
    cleaned_orders.loc[cleaned_orders['Part Number'] == '778752', 'Total Credit Price'] = cleaned_orders['Order Quantity'] * 7500
    cleaned_orders_sap['Total Credit Price'] = cleaned_orders_sap['Order Quantity (Item)'] * 7500



    with pd.ExcelWriter("training_items.xlsx") as writer:
        cleaned_orders.to_excel(writer, sheet_name="oracle", index=False)
        # writer.sheets["oracle"].autofilter(0,0,cleaned_orders.shape[0],cleaned_orders.shape[1])
        cleaned_orders_sap.to_excel(writer, sheet_name="sap", index=False)
        # writer.sheets["sap"].autofilter(0,0,cleaned_orders_sap.shape[0],cleaned_orders_sap.shape[1])

get_credits(load_data())
