""""
Corina Elisii
2022-08-23

GOAL:
Here is where we enter the THD details to make our program work
"""
import os
import glob

import pandas as pd
import CHub_Walk as cw


def download_ord_dtls(df, pack_dir):
    cw.dnwld_order_csv("A: THD Undelivered")
    # RESULT_CSV IS THE MOST RECENT DOWNLOADS FILE AND STARTS AS ROW 6
    result_csv = pd.read_csv(max(glob.glob(pack_dir+'*.csv'), key=os.path.getctime), skiprows=6)
    result_csv = result_csv.drop_duplicates()  # DROP DUPLICATE ROWS
    # MAKE SURE BOTH JOIN COLUMNS ARE CAST AS INT
    result_csv['PO Number'] = result_csv['PO Number'].astype(int)
    df['PO Number'] = df['PO Number'].astype(int)
    thd_ord_dtl = df.merge(result_csv, on='PO Number', how='left')
    # CLEAN DF
    thd_ord_dtl.drop(columns=['Total', 'Sending Partner', 'Order Date_x'])  # DROP UNNECESSARY COLUMNS
    thd_ord_dtl.rename(columns={"Order Date_y": "Order Date"})  # RENAME ORDER DATE COLUMN
    thd_ord_dtl['BillTo Name'] = thd_ord_dtl['BillTo Name'].str.upper()
    thd_ord_dtl['ShipTo Name'] = thd_ord_dtl['ShipTo Name'].str.upper()
    thd_ord_dtl['ShipTo Address1'] = thd_ord_dtl['ShipTo Address1'].str.upper()
    thd_ord_dtl['ShipTo City'] = thd_ord_dtl['ShipTo City'].str.upper()
    thd_ord_dtl['ShipTo First Name'] = thd_ord_dtl['ShipTo First Name'].str.upper()
    # os.remove(pack_dir, max(glob.glob(pack_dir+'*.csv'))) # ADD IN FUTURE TO DELETE THE DOWNLOADED PS FILE
    return thd_ord_dtl

def get_thd_carrier(df):
    file_path = "Y:/User/Order Entry User/CTM Household Appliances Orders Processed/The Home Depot Inc/Updated Carrier List.xlsx"
    thd_carrier = pd.read_excel(file_path, sheet_name='Collect LTL - State to State')
    thd_carrier = thd_carrier.rename(columns={"Origin State": "Ship From"})  # RENAME SHIP FROM COLUMN
    thd_carrier = thd_carrier.rename(columns={"Destination State": "ShipTo State"})  # RENAME SHIP TO COLUMN
    thd_carrier.drop(columns=['Origin Country', 'Destination Country'])  # DROP UNNECESSARY COLUMNS
    df["State Pair"] = df["Ship From"] + "_" + df["ShipTo State"]
    thd_ord_dtl = pd.merge(df, thd_carrier, on="State Pair", how='left')
    return thd_ord_dtl