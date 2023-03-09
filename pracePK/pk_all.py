# coding: utf-8
import glob
import pandas as pd

# set directory path
directory_path = "./"

# get a list of all Excel files in the directory
excel_files = glob.glob(directory_path + "*.xlsx")

# initialize an empty list to store dataframes
dfs = []

# loop through each Excel file
for file in excel_files:
    # loop through each sheet in the workbook
    for sheet_name in pd.read_excel(file, sheet_name=None, skiprows=1):
        # read data from columns D, G, H, I, K, L, M, N, O
        df = pd.read_excel(file, sheet_name=sheet_name, usecols="D,G,H,I,K,L,M,N,O", skiprows=1)
        # add a column with the sheet name
        df["Sheet_Name"] = sheet_name
        # append dataframe to the list
        dfs.append(df)

# concatenate all dataframes into one
result_df = pd.concat(dfs, ignore_index=True)
