# import modules
import pandas as pd
import glob

# path of the folder
path = r"Y:/User/Order Entry User/CTM Household Appliances Orders Processed/The Home Depot Inc/FILES/Daily Order/"

# reading all the excel files
filenames = glob.glob(path + "\\THD_*.xlsx")
print('File names:', filenames)

# initializing empty data frame
finalexcelsheet = pd.DataFrame()

# to iterate excel file one by one
# inside the folder
for file in filenames:
    # combining multiple excel worksheets
    # into single data frames
    df = pd.read_excel(file, engine='openpyxl')
    # appending excel files one by one
    finalexcelsheet = finalexcelsheet.append(
        df, ignore_index=True)

# to print the combined data
print('Final Sheet:')


finalexcelsheet.to_excel(path+"appended.xlsx", index=False)

