""""
Corina Elisii
2022-08-23

GOAL:
Here is where we enter the BestBuy details to make our program work
"""
import os
import glob

import pandas as pd
import CHub_Walk as cw


def download_ord_dtls(df, pack_dir):
    cw.dnwld_order_csv("C: BB Undelivered")
    # RESULT_CSV IS THE MOST RECENT DOWNLOADS FILE AND STARTS AS ROW 6
    result_csv = pd.read_csv(max(glob.glob(pack_dir+'*.csv'), key=os.path.getctime), skiprows=6)
    result_csv = result_csv.drop_duplicates()  # DROP DUPLICATE ROWS
    # MAKE SURE BOTH JOIN COLUMNS ARE CAST AS INT
    #result_csv['PO Number'] = result_csv['PO Number'].astype(int)
    #df['PO Number'] = df['PO Number'].astype(int)
    bb_ord_dtl = df.merge(result_csv, on='PO Number', how='left')
    # CLEAN DF
    bb_ord_dtl.drop(columns=['Total', 'Sending Partner', 'Order Date_x' ])  # DROP UNNECESSARY COLUMNS
    bb_ord_dtl.rename(columns={"Order Date_y": "Order Date"})  # RENAME ORDER DATE COLUMN
    bb_ord_dtl['BillTo Name'] = bb_ord_dtl['BillTo Name'].str.upper()
    bb_ord_dtl['ShipTo Name'] = bb_ord_dtl['ShipTo Name'].str.upper()
    bb_ord_dtl['ShipTo Address1'] = bb_ord_dtl['ShipTo Address1'].str.upper()
    bb_ord_dtl['ShipTo City'] = bb_ord_dtl['ShipTo City'].str.upper()
    bb_ord_dtl['ShipTo First Name'] = bb_ord_dtl['ShipTo First Name'].str.upper()
    # os.remove(pack_dir, max(glob.glob(pack_dir+'*.csv'))) # ADD IN FUTURE TO DELETE THE DOWNLOADED PS FILE
    return bb_ord_dtl
