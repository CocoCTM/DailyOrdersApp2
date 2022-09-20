""""
Corina Elisii
2022-08-23

GOAL:
Here is where we enter the Lowe's details to make our program work
"""
import os
import glob

import pandas as pd
import CHub_Walk as cw



def download_ord_dtls(df, pack_dir):
    cw.dnwld_order_csv("B: LOWES Undelivered")
    # RESULT_CSV IS THE MOST RECENT DOWNLOADS FILE AND STARTS AS ROW 6
    result_csv = pd.read_csv(max(glob.glob(pack_dir+'*.csv'), key=os.path.getctime), skiprows=6)
    result_csv = result_csv.drop_duplicates()  # DROP DUPLICATE ROWS
    # MAKE SURE BOTH JOIN COLUMNS ARE CAST AS INT
    result_csv['PO Number'] = result_csv['PO Number'].astype(int)
    df['PO Number'] = df['PO Number'].astype(int)
    lowes_ord_dtl = df.merge(result_csv, on='PO Number', how='left')
    # CLEAN DF
    lowes_ord_dtl.drop(columns=['Total', 'Sending Partner', 'Order Date_x'])  # DROP UNNECESSARY COLUMNS
    lowes_ord_dtl.rename(columns={"Order Date_y": "Order Date"})  # RENAME ORDER DATE COLUMN
    # lowes_ord_dtl["ShipTo Address Rate Class"] = lowes_ord_dtl["ShipTo Address Rate Class"].fillna('LOWE\'S STORE') #FILL NA WITH "LOWE'S STORE"
    lowes_ord_dtl['BillTo Name'] = lowes_ord_dtl['BillTo Name'].str.upper()
    lowes_ord_dtl['ShipTo Name'] = lowes_ord_dtl['ShipTo Name'].str.upper()
    lowes_ord_dtl['ShipTo Address1'] = lowes_ord_dtl['ShipTo Address1'].str.upper()
    lowes_ord_dtl['ShipTo City'] = lowes_ord_dtl['ShipTo City'].str.upper()
    lowes_ord_dtl['ShipTo First Name'] = lowes_ord_dtl['ShipTo First Name'].str.upper()
    # lowes_ord_dtl["ShipTo Day Phone"] = lowes_ord_dtl["ShipTo Day Phone"].fillna("LOWE'S STORE"+lowes_ord_dtl['ShipTo Customer Number'])  # FILL NA WITH "LOWE'S STORE"
    # os.remove(pack_dir, max(glob.glob(pack_dir+'*.csv'))) # ADD IN FUTURE TO DELETE THE DOWNLOADED PS FILE
    return lowes_ord_dtl


def get_lowes_file_path():
    file_path = "Y:/User/Order Entry User/CTM Household Appliances Orders Processed/Lowe's/NEW - PROCESS"
    return file_path

def get_lowes_carrier(df):
    file_path = "Y:/User/Order Entry User/CTM Household Appliances Orders Processed/Lowe's/LOWES MATRIX - 2022-07-01.xlsx"
    lowes_carrier = pd.read_excel(file_path, sheet_name='SOS LTL RM Lookup Tool')
    lowes_carrier = lowes_carrier.rename(columns={"Origin St": "Ship From"})  # RENAME SHIP FROM COLUMN
    lowes_carrier = lowes_carrier.rename(columns={"Dest. St": "ShipTo State"}) # RENAME SHIP TO COLUMN
    df["State Pair"] = df["Ship From"] + "_" + df["ShipTo State"]
    lowes_ord_dtl = pd.merge(df, lowes_carrier, on="State Pair", how='left')
    return lowes_ord_dtl