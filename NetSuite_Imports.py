""""
Corina Elisii
2022-08-29

GOAL:
Here is where we enter the NetSuite Imports
"""

import pandas as pd


def get_cust_id(cust_name):
    ns_cust_dir = 'Y:/User/Order Entry User/NETSUITE IMPORTS/NETSUITE CUSTOMER'
    result_csv = pd.read_csv(ns_cust_dir + "/CustomerList.csv")
    result_csv = result_csv[['Internal ID', 'Name', 'Category', 'Terms']]
    int_id = result_csv.loc[result_csv['Name'] == cust_name, 'Internal ID'].item()
    if pd.isna(int_id):
        return 000
    else:
        return int_id

def get_item_id(df):
    ns_cust_dir = 'Y:/User/Order Entry User/NETSUITE IMPORTS/NETSUITE ITEMS'
    result_csv = pd.read_csv(ns_cust_dir + "/ItemsChild.csv")
    result_csv = result_csv.rename(columns={"Name": "Vendor SKU", "Internal ID": "Item ID"})
    result_csv = result_csv[['Item ID', 'Vendor SKU']]
    merge_csv = df.merge(result_csv, on='Vendor SKU', how='left')
    return merge_csv

def get_ship_from(df):
    ship_from_file = "Y:/User/Order Entry User/CTM Household Appliances Orders Processed/OrderRouting_InventoryLookup.xlsx"
    ship_from = pd.read_excel(ship_from_file, sheet_name="Warehouse Routing", skiprows=5)
    ship_from = ship_from[['Internal ID', 'Location', 'Location ID', 'Location State']]
    ship_from = ship_from.rename(columns={"Location State": "Ship From", "Internal ID": "Item ID"})
    merge_csv = df.merge(ship_from, on='Item ID', how='left')
    calc_ship_from = merge_csv.groupby('PO Number')['Ship From'].nunique().to_frame() # Get the number of Ship From locations per PO# and cast as DF
    calc_ship_from = calc_ship_from.rename(columns={"Ship From": "Ship From#"})
    merge_csv = merge_csv.merge(calc_ship_from, on='PO Number', how='left')
    merge_csv['Ship From'][merge_csv['Ship From#'] > 1] = "INPUT MANUALLY"  # IF PO# has > 1 location - remove ship from
    merge_csv = merge_csv.drop(columns=['Ship From#'])
    return merge_csv