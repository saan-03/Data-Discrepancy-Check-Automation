# Importing libraries
import pandas as pd
import numpy as np 
import openpyxl as op
import warnings
import os

# Common dataframes needed across daily comparisons
KbDockCC_file = r"J:\$PC Materials\Overseas\09. Projects\OS Modernization\File Comparison Tool\KB-Dock-Container Code File.xlsm"
FA_cons_file = r"J:\$PC Materials\Overseas\09. Projects\OS Modernization\File Comparison Tool\240108_TMMC_FA_Consolidated.xlsx"
df3 = pd.read_excel(KbDockCC_file)
df4 = pd.read_excel(FA_cons_file, sheet_name='TNSO Layout')
df3c = df3.copy()
df4c = df4.copy()

# Creating a directory to store comparison files
# New directory is created every day based on the date
date_time = pd.Timestamp.now()
folder_name = date_time.strftime("%Y-%m-%d")
folder_path = "j:\\$PC Materials\\Overseas\\09. Projects\\OS Modernization\\File Comparison Tool\\Tests\\" + folder_name

# Checks if directory exists, creates one if does not exist
if not os.path.exists(folder_path):
    os.makedirs(folder_path)

# FUNCTIONS FOR DATA CLEANING ----------------------------------------------------------------------------------------------------------------
# van_format(Dataframe, col_name) is a function to format vanning data for TMMK Final Order df2c
# From M/D/YYYY -> YYMMDD or based on if the month or day is first in the new format
def van_format(df, column_name):
    lead_zero = 2

    # splits into 3 columns
    date_series = df[column_name].str.split('-', expand=True)

    # combining together
    df[column_name] = date_series[0].str[2:] + date_series[1].str.zfill(lead_zero) + date_series[2].str.zfill(lead_zero)
    df[column_name] = df[column_name].astype(int)

# remove_spaces(Dataframe, Column Name, Delimeter, Number of Splits Needed) is a 
# function that removes trailing spaces in a df
def remove_spaces(df,column,separator,split):
    series = df[column].str.split(separator, n=split, expand=True)
    df[column] = series[0]

# part_format(Dataframe, Column Name, Delimeter) is a function that removes separator
# and trailing whitespaces to format part number
def part_format(df,column):
    df['new_PART'] = df[column].str[:5] + '-' + df[column].str[5:10] + '-' + df[column].str[10:]
# -------------------------------------------------------------------------------------------------------------------------------------------

# DATA CLEANING FOR DF3C
df3c.rename(columns = {'PART NUMBER':'PARTNO','Dest Code':'DEST_CODE','Container Code':'CONT_CODE',
                          'ORDRLOT':'ORDER_LOT'}, inplace=True)
df3c.drop(['DESCRIPTION','TM','CLS','CRTCL','USER MOD','Key?','VENSHR','LCC','DK'], axis=1, inplace=True)
df3c.drop_duplicates(inplace=True) # Dropping duplicates for merging

# applying formatting functions
# Changing whitespace to empty strings
df3c['KANBAN'] = df3c['KANBAN'].replace('\\s+', '', regex=True)

# Changing order of columns
df3c_order = ['KANBAN','CONT_CODE','DEST_CODE','PARTNO','ORDER_LOT']
df3c = df3c.reindex(columns=df3c_order)

# DATA CLEANING FOR DF4C
df4c.drop(['DATAID','BLANK','BLANK2','KOUKU'], axis=1, inplace=True)
df4c.drop(index=0,inplace=True)    # Dropping first row of df4c

# Renaming all columns
df4c_new_columns = ['DEST_CODE','YYMM', 'PARTNO', 'LOT_SIZE']
df4c_new_columns += ['SCH {:02}'.format(i) for i in range(1, 32)] + ['SCH TTL']
df4c_new_columns += ['MAX {:02}'.format(i) for i in range(1, 32)] + ['MAX TTL']
df4c_new_columns += ['MIN {:02}'.format(i) for i in range(1, 32)] + ['MIN TTL']
df4c.columns = df4c_new_columns

# Changing datatypes
df4c['YYMM'] = df4c['YYMM'].astype(int)
change_dtypes_columns = df4c.columns[-97:]      # changing last 97 columns to an int
df4c[change_dtypes_columns] = df4c[change_dtypes_columns].apply(pd.to_numeric, downcast='integer')
pd.set_option('future.no_silent_downcasting', True)

warnings.simplefilter(action='ignore', category=pd.errors.PerformanceWarning)
warnings.simplefilter(action='ignore', category=pd.errors.ParserWarning)
# MAIN FUNCTIONS ----------------------------------------------------------------------------------------------------------------------------

# row_diff(name of dataframe1, name of datframe2, dataframe1, dataframe2) outputs whether the dataframes
# have the same number of rows, or if a dataframe has more row than the other.
def row_diff(name_1, name_2, df1, df2):
    count_1 = df1.shape[0]
    count_2 = df2.shape[0]
    diff = count_1 - count_2

    if (diff == 0):
        print("The number of rows in each file matches.")
    elif (diff > 0):
        print("There are", diff, "more rows in", name_1)
    else:
        print("There are", -(diff), "more rows in", name_2)

# merged_row_diff(name of dataframe1, name of datafram2, left_merged_df, right_merged_df, inner/common merged_df)
# outputs the count of common rows between each dataframe, and the count of unmatched rows in each dataframe
# unmatched rows are those that do not have a match to any entry in the other dataframe
def merged_row_diff(name_1, name_2, left, right, inner):
    count_1 = left.shape[0]
    count_2 = right.shape[0]
    common = inner.shape[0]

    if ((count_1 == 0) and (count_2 == 0)):
        print("There are 0 unmatched rows between both files")
    else:
        print("There are", common, "common rows in", name_1, "and", name_2)
        print("There are", count_1, "unmatched rows in", name_1)
        print("There are", count_2, "unmatched rows in", name_2)
        print("\n")

def orderplan_disc_test(merged_df, store_disc_df, ims_col_name, osp_col_name, ims_col_name2, osp_col_name2):
        for row in range(len(merged_df)):
            if ((merged_df[ims_col_name].iloc[row] != merged_df[osp_col_name].iloc[row]) or 
                (merged_df[ims_col_name2].iloc[row] != merged_df[osp_col_name2].iloc[row])):
                store_disc_df.loc[len(store_disc_df)] = merged_df.iloc[row]
            else:
                continue

# Function to replace NaNs with their appropriate container codes
def insert_cont_code(merged_df, df2c):
    # boolean mask to determine if CONT_CODE has NaN value
    mask_nan = merged_df["CONT_CODE"].isna()

    conditions = [
        (mask_nan & (merged_df["DEST_CODE"] == "102A"), "C"),
        (mask_nan & (merged_df["DEST_CODE"] == "102K"), "H"),
        (mask_nan & (merged_df["DEST_CODE"] == "102W"), "E"),
        (mask_nan & (merged_df["DEST_CODE"] == "102B"), "V"),
        (mask_nan & (merged_df["DEST_CODE"] == "102S"), "T")
    ]

    # Fills CONT_CODE with appropriate value if condition is determined to be True (i.e. is NaN)
    for condition, value in conditions:
        merged_df.loc[condition, "CONT_CODE"] = value

    # Fills rest of the NaNs with "G or B" for 102C
    merged_df.fillna({"CONT_CODE": "G or B"}, inplace=True)

    # Boolean mask to determine which rows have "G or B" CONT_CODE
    mask = merged_df['CONT_CODE'].isin(['G or B'])

    for index, _ in merged_df[mask].iterrows():

        temp_df_G = merged_df.copy()
        temp_df_B = merged_df.copy()

        temp_df_G.at[index, 'CONT_CODE'] = 'G'
        temp_df_B.at[index, 'CONT_CODE'] = 'B'

        # Flag to check if either modified row exists in df2c
        found_in_df2c = False
        # Iterate over rows in df2c to check if the modified row exists
        for _, row_df2c in df2c.iterrows(): # index of df2c not being used
            if (row_df2c[['PARTNO', 'CONT_CODE']] == temp_df_G.loc[index, ['PARTNO', 'CONT_CODE']]).all():
                merged_df.at[index, 'CONT_CODE'] = 'G'
                found_in_df2c = True
                break
            elif (row_df2c[['PARTNO', 'CONT_CODE']] == temp_df_B.loc[index, ['PARTNO', 'CONT_CODE']]).all():
                merged_df.at[index, 'CONT_CODE'] = 'B'
                found_in_df2c = True
                break

        # If neither modified row exists in df2c, revert back to original "G or B"
        if not found_in_df2c:
            merged_df.at[index, 'CONT_CODE'] = "G or B"

# Updates cont_code for certain part numbers instead of manually doing it
def update_cont_code(df: pd.DataFrame) -> pd.DataFrame:
    part_numbers_to_change = [
        "35151-78010-00",
        "161B0-47010-00",
        "82715-78270-00",
        "88718-78030-00"
    ]
    mask_nan = df["PARTNO"].isin(part_numbers_to_change)

    conditions = [
        (mask_nan & (df["DEST_CODE"] == "102A"), "C"),
        (mask_nan & (df["DEST_CODE"] == "102C"), "B")
    ]

    for condition, value in conditions:
        df.loc[condition, "CONT_CODE"] = value


# merge_disc_test(dataframe of common rows, discrepancy_entry_dataframe, ims_col_name, osp_col_name) checks whether 
# the value in the ims_col_name matches the value in osp_col_name for each row in merged_df. If it does not match,
# those rows are added to the disc_df to portray the discrepancy.
def merge_disc_test(merged_df, store_disc_df, ims_col_name, osp_col_name):
        for row in range(len(merged_df)):
            if merged_df[ims_col_name].iloc[row] != merged_df[osp_col_name].iloc[row]:
                store_disc_df.loc[len(store_disc_df)] = merged_df.iloc[row]
            else:
                continue

def ttl_disc_test(merged_df, store_disc_df, start_ims_col, start_osp_col):
    for row in range(len(merged_df)):
        ims_row_vals = []
        osp_row_vals = []
        if ((merged_df['SCH TTL_ims'].iloc[row] != merged_df['SCH TTL_osp'].iloc[row]) or
            (merged_df['MAX TTL_ims'].iloc[row] != merged_df['MAX TTL_osp'].iloc[row]) or
            (merged_df['MIN TTL_ims'].iloc[row] != merged_df['MIN TTL_osp'].iloc[row])):
            # appending common column vals to ims and osp val lists
            dest_code_val = merged_df.iloc[row, 0]
            yymm_val = merged_df.iloc[row, 1]
            partno_val = merged_df.iloc[row, 2]
            lot_size_val = merged_df.iloc[row, 3]
            ims_row_vals.extend(['_ims', dest_code_val, partno_val, yymm_val, lot_size_val])
            osp_row_vals.extend(['_osp', dest_code_val, partno_val, yymm_val, lot_size_val])
            
            start_ims_index = merged_df.columns.get_loc(start_ims_col)
            start_osp_index = merged_df.columns.get_loc(start_osp_col)
            end_index = merged_df.shape[1]
            while (start_ims_index != end_index):
                ims_col_val = merged_df.iloc[row, start_ims_index]
                osp_col_val = merged_df.iloc[row, start_osp_index]
                if ims_col_val != osp_col_val:
                    ims_row_vals.append(ims_col_val)
                    osp_row_vals.append(osp_col_val)
                else:
                    ims_row_vals.append("ok")
                    osp_row_vals.append("ok")
                
                start_ims_index += 1
                start_osp_index += 1
                
            store_disc_df.loc[len(store_disc_df)] = ims_row_vals
            store_disc_df.loc[len(store_disc_df)] = osp_row_vals

        else:
            continue

# forecast_disc_test(merged_df, store_disc_df) parses through all columns with _ims and _osp suffixes to compare columns with
# same names. If values are different, the differences and the row are both added into disc_df
def forecast_disc_test(merged_df, store_disc_df, start_ims_col, start_osp_col):
    for row in range(len(merged_df)):
        count = 0
        ims_row_vals = []
        osp_row_vals = []
    # appending common column vals to ims and osp val lists
        partno_val = merged_df.iloc[row, 0]
        n_vann_val = merged_df.iloc[row, 1]
        cc_val = merged_df.iloc[row, 2]
        order_lot_val = merged_df.iloc[row, 3]
        dest_osp_val = merged_df.iloc[row, 4]
        dest_ims_val = merged_df.iloc[row, 5]

        ims_row_vals.extend(['_ims', partno_val, n_vann_val, cc_val, order_lot_val, dest_ims_val])
        osp_row_vals.extend(['_osp', partno_val, n_vann_val, cc_val, order_lot_val, dest_osp_val])
        
        start_ims_index = merged_df.columns.get_loc(start_ims_col)
        start_osp_index = merged_df.columns.get_loc(start_osp_col)
        end_index = merged_df.shape[1]
        while (start_ims_index != end_index):
            ims_col_val = merged_df.iloc[row, start_ims_index]
            osp_col_val = merged_df.iloc[row, start_osp_index]
            if ims_col_val != osp_col_val:
                ims_row_vals.append(ims_col_val)
                osp_row_vals.append(osp_col_val)
            else:
                count += 1
                ims_row_vals.append("ok")
                osp_row_vals.append("ok")
            
            start_ims_index += 1
            start_osp_index += 1
        if (count == 130):
            continue
        else:
            store_disc_df.loc[len(store_disc_df)] = ims_row_vals
            store_disc_df.loc[len(store_disc_df)] = osp_row_vals

# disc_report(num, df1, df2, nameofdf1, nameodfdf2, left_merge, right_merge, inner_merge, discrepancy_df) is a fucntion
# that calls onto the comparison functions to determine the discrepancy and observations
def disc_report(num, df1, df2, name_1, name_2, left, right, inner, disc_df):
    if (num == 1):
        row_diff(name_1, name_2, df1, df2)
        merged_row_diff(name_1, name_2, left, right, inner)
        merge_disc_test(inner, disc_df, 'FINAL_ORDER_ims', 'FINAL_ORDER_osp')
    elif (num == 2):
        row_diff(name_1, name_2, df1, df2)
        merged_row_diff(name_1, name_2, left, right, inner)
        forecast_disc_test(inner, disc_df, 'N-1 LOTS_ims', 'N-1 LOTS_osp')
    elif (num == 3):
        row_diff(name_1, name_2, df1, df2)
        merged_row_diff(name_1, name_2, left, right, inner)
        ttl_disc_test(inner, disc_df,'SCH 01_ims','SCH 01_osp')
    else:
        row_diff(name_1, name_2, df1, df2)
        merged_row_diff(name_1, name_2, left, right, inner)
        orderplan_disc_test(inner, disc_df,'START C/O_ims','START C/O_osp','BASIC REQUIREMENT_ims','BASIC REQUIREMENT_osp')


# COMPARISON FUNCTIONS

# comparison_num_1() performs the comparison between the MF Final Order and OSP Final Order files and exports the 
# discrepancy results into csv files.
def comparison_num_1(df3c, folder_path):
    MF_Final_Order_txt = r'\\t21ftp01\OSPARTS\OS_BATCH_RPTS\L3OS06.csv'
    OSP_Final_Order = r'j:\$PC Materials\Overseas\09. Projects\OS Modernization\File Comparison Tool\OSP Daily Data Files\Final_Order_TMMC_20240425_220129.xlsx'
    df1 = pd.read_csv(MF_Final_Order_txt)
    df2 = pd.read_excel(OSP_Final_Order)
    df1c = df1.copy()
    df2c = df2.copy()

    # Changing column names, datatypes, dropping unwanted columns and duplicates
    df1c.rename(columns = {'VANNING DATE':'VANNING_DATE', 'FINAL ORDER (LOTS)                  ':'FINAL_ORDER',
                       'PART NUMBER':'PARTNO','VANNING DATE':'VANNING_DATE','DEST CODE':'DEST_CODE','ORDER LOT':'ORDER_LOT'}, inplace=True)
    df2c.rename(columns = {'DEST CD':'DEST_CODE','CONT CD':'CONT_CODE',
                            'PART#':'PARTNO','VANNING DATE':'VANNING_DATE',
                            'FINAL ORDER (LOTS)':'FINAL_ORDER','ORDER LOT':'ORDER_LOT'}, inplace=True)
        
    df1c.drop(['DATA ID','T/M'], axis=1, inplace=True)
    df2c.drop(['T/M','DATA ID','ORDER TYPE'], axis=1, inplace=True)
    df1c.drop_duplicates(inplace=True)
    df2c.drop_duplicates(inplace=True)
    df2c['VANNING_DATE'] = df2c['VANNING_DATE'].astype(str)

    # Formatting VANNING_DATE for df1c
    date_series = df1c['VANNING_DATE'].str.split('/', expand=True)
    df1c['VANNING_DATE'] = date_series[0] + date_series[1].str.zfill(2) + date_series[2].str.zfill(2)
    df1c['VANNING_DATE'] = df1c['VANNING_DATE'].astype(int)

    # Changing Kanban whitespaces to empty strings
    df1c['KANBAN'] = df1c['KANBAN'].replace('\\s+', 0, regex=True)

    # applying formatting functions
    part_format(df2c,'PARTNO') 
    van_format(df2c,'VANNING_DATE')      # Formatting Vanning_Date for df2c

    # dropping PARTNO to make the delimeter separated column new partno column
    df2c.drop(['PARTNO'], axis=1, inplace=True)
    df2c.rename(columns = {'new_PART':'PARTNO'}, inplace=True)

    # JOINING df1c and df3c
    df1_merge = pd.merge(df1c, df3c[['CONT_CODE','DEST_CODE','PARTNO','ORDER_LOT']], on=['DEST_CODE','PARTNO','ORDER_LOT'], how='left')
    # Calling function for Cont_code
    insert_cont_code(df1_merge, df2c)

    # Updating cont_code for certain partnumbers
    update_cont_code(df1_merge)

    # MERGING dfs
    df2_merge = pd.merge(df2c, df1_merge, on=['VANNING_DATE','PARTNO','CONT_CODE'], how='inner',suffixes=['_osp','_ims'])
    df2_leftmerge = pd.merge(df2c, df1_merge, on=['VANNING_DATE','PARTNO','CONT_CODE'], how='left', suffixes=['_osp','_ims'], indicator=True).query('_merge=="left_only"').drop(columns='_merge')
    df2_rightmerge = pd.merge(df2c, df1_merge, on=['VANNING_DATE','PARTNO','CONT_CODE'], how='right', suffixes=['_osp','_ims'], indicator=True).query('_merge=="right_only"').drop(columns='_merge')

    # Cleaning up merged dfs
    df2_leftmerge.drop(['KANBAN_ims','FINAL_ORDER_ims', 'ORDER_LOT_ims'], axis=1, inplace=True)
    df2_rightmerge.drop(['KANBAN_osp','FINAL_ORDER_osp', 'ORDER_LOT_osp'], axis=1, inplace=True)

    # Reordering columns
    rightmerge_order = ['PARTNO','VANNING_DATE','CONT_CODE','DEST_CODE_ims','KANBAN_ims','ORDER_LOT_ims','FINAL_ORDER_ims']
    df2_rightmerge = df2_rightmerge.reindex(columns=rightmerge_order)

    leftmerge_order = ['PARTNO','VANNING_DATE','CONT_CODE','DEST_CODE_osp','KANBAN_osp','ORDER_LOT_osp','FINAL_ORDER_osp']
    df2_leftmerge = df2_leftmerge.reindex(columns=leftmerge_order)

    df2_merge_order = ['PARTNO','VANNING_DATE','CONT_CODE','DEST_CODE_ims','DEST_CODE_osp','KANBAN_ims','KANBAN_osp',
                    'ORDER_LOT_ims','ORDER_LOT_osp','FINAL_ORDER_ims','FINAL_ORDER_osp']
    df2_merge = df2_merge.reindex(columns=df2_merge_order)

    # CALLING FUNCTIONS FOR COMPARISON
    print("\nCOMPARISON BETWEEN MF FINAL ORDER AND OSP FINAL ORDER ----------------------------------------------")
    wrong_values_df = pd.DataFrame(columns=df2_merge.columns) # dataframes to store discrepancy
    disc_report(1, df1c, df2c, "MF Final Order", "OSP Final Order", df2_leftmerge, df2_rightmerge, df2_merge, wrong_values_df)

    # Creting dfs to add text details
    text_1_df = pd.DataFrame({'text_column': ['IMS Final Order: Below are additional entries only present in the IMS format',' ']})
    text_2_df = pd.DataFrame({'text_column': ['OSP Final Order: Below are additional entries only present in the OSP format',' ']})
    text_3_df = pd.DataFrame({'text_column': ['Below are entries with incorrect Final Order Values',' ']})

    # Exporting files
    question = """Would you like to export the discrepancy report as .csv files?
    Enter Y for yes and N for no: """
    export_input = input(question)

    if (export_input == "Y" or export_input == "y"):
        file_path_1 = folder_path + "\\1.1.IMS_FinalOrder_Unmatched_Entries.csv"
        file_path_2 =  folder_path + "\\1.2.OSP_FinalOrder_Unmatched_Entries.csv"
        file_path_3 = folder_path + "\\1.3.FinalOrder_Incorrect__Values.csv"

        text_1_df.to_csv(file_path_1, header=None, index=False)
        df2_rightmerge.to_csv(file_path_1, mode='a', index=False)

        text_2_df.to_csv(file_path_2, header=None, index=False)
        df2_leftmerge.to_csv(file_path_2, mode='a', index=False)

        text_3_df.to_csv(file_path_3, header=None, index=False)
        wrong_values_df.to_csv(file_path_3,mode='a', index=False)
    else:
        print("")

# comparison_num_2() performs the comparison between the MF Forecast and OSP Forecast files and exports the 
# discrepancy results into csv files.
def comparison_num_2(df3c, folder_path):
    forecast = r'\\t21ftp01\OSPARTS\FORECAST.CSV'
    tmmc_forecast = r"j:\$PC Materials\Overseas\09. Projects\OS Modernization\File Comparison Tool\OSP Daily Data Files\TMMC_pxp_forecast_report_20240425_221219.xlsx"
    df6 = pd.read_csv(forecast, index_col=False)
    df7 = pd.read_excel(tmmc_forecast)
    df6c = df6.copy()
    df7c = df7.copy()

    # Dropping
    df6c.drop(['CC'], axis=1, inplace=True)
    df7c.drop(['N MONTH CO LOTS',], axis=1, inplace=True)
    df7c.rename(columns = {'DEST':'DEST_CODE','PART NUMBER':'PARTNO','CC':'CONT_CODE'}, inplace=True)

    # Renaming columns to have same names
    df_new_cols = ['ID','DEST_CODE','PARTNO','OT','ORDER_LOT','N-1 LOTS','N LOTS','N+1 LOTS','N+2 LOTS','N VANN','PRF','N-1 MONTH CO LOTS']
    df_new_cols += ['N-1 D{:02}'.format(i) for i in range(1, 32)]
    df_new_cols += ['N D{:02}'.format(i) for i in range(1, 32)]
    df_new_cols += ['N+1 D{:02}'.format(i) for i in range(1, 32)]
    df_new_cols += ['N+2 D{:02}'.format(i) for i in range(1, 32)]
    df6c.columns = df_new_cols

    df_new_cols = ['ID','DEST_CODE','PARTNO','OT','ORDER_LOT','N-1 LOTS','N LOTS','N+1 LOTS','N+2 LOTS','N VANN','CONT_CODE','PRF','KANBAN','N-1 MONTH CO LOTS']
    df_new_cols += ['N-1 D{:02}'.format(i) for i in range(1, 32)]
    df_new_cols += ['N D{:02}'.format(i) for i in range(1, 32)]
    df_new_cols += ['N+1 D{:02}'.format(i) for i in range(1, 32)]
    df_new_cols += ['N+2 D{:02}'.format(i) for i in range(1, 32)]
    df7c.columns = df_new_cols

    df6c.drop_duplicates(inplace=True)
    df7c.drop_duplicates(inplace=True)
    
    # formatting part no
    part_format(df6c,'PARTNO') 
    df6c.drop(['PARTNO'], axis=1, inplace=True)
    df6c.rename(columns = {'new_PART':'PARTNO'}, inplace=True)

    part_format(df7c,'PARTNO') 
    df7c.drop(['PARTNO'], axis=1, inplace=True)
    df7c.rename(columns = {'new_PART':'PARTNO'}, inplace=True)

    column_to_move = df6c.pop("PARTNO")
    df6c.insert(2, "PARTNO", column_to_move)
    column_to_move = df7c.pop("PARTNO")
    df7c.insert(2, "PARTNO", column_to_move)

    df3c['KANBAN'] = df3c['KANBAN'].replace('', 0, regex=True) # replacing empty spaces with 0s
    # Merging to get CC
    df6_merge = pd.merge(df6c, df3c[['DEST_CODE','PARTNO','ORDER_LOT','CONT_CODE','KANBAN']], on=['DEST_CODE','PARTNO','ORDER_LOT'], how='left')
    insert_cont_code(df6_merge, df7c)
    update_cont_code(df6_merge)

    # Positioning the merged df similar to df7 for later comparison
    column_to_move = df6_merge.pop("CONT_CODE")
    df6_merge.insert(10, "CONT_CODE", column_to_move)
    df6_merge.fillna({'CONT_CODE':' '}, inplace=True)
    df6_merge.drop(['KANBAN'], axis=1, inplace=True)
    df7c.drop(['KANBAN'], axis=1, inplace=True)
    df6_merge.drop_duplicates(inplace=True)
    df7c.drop_duplicates(inplace=True)

    df7_merge = pd.merge(df7c, df6_merge, on=['PARTNO','N VANN','CONT_CODE','ORDER_LOT'], how='inner', suffixes=['_osp','_ims'])
    df7_leftmerge = pd.merge(df7c, df6_merge, on=['ID','OT','PARTNO','N VANN','CONT_CODE','ORDER_LOT'], how='left', suffixes=['_osp','_ims'], indicator=True).query('_merge=="left_only"').drop(columns='_merge')
    df7_rightmerge = pd.merge(df7c, df6_merge, on=['ID','OT','PARTNO','N VANN','CONT_CODE','ORDER_LOT'], how='right', suffixes=['_osp','_ims'], indicator=True).query('_merge=="right_only"').drop(columns='_merge')
    df7_merge.drop_duplicates(inplace=True)
    # Dropping columns
    df7_merge.drop(['ID_ims','ID_osp','OT_ims','OT_osp'], axis=1, inplace=True)

    column_to_move = df7_merge.pop("N VANN")
    df7_merge.insert(2, "N VANN", column_to_move)
    column_to_move = df7_merge.pop("CONT_CODE")
    df7_merge.insert(3, "CONT_CODE", column_to_move)
    column_to_move = df7_merge.pop("DEST_CODE_osp")
    df7_merge.insert(4, "DEST_CODE_osp", column_to_move)
    column_to_move = df7_merge.pop("DEST_CODE_ims")
    df7_merge.insert(5, "DEST_CODE_ims", column_to_move)

    # df to store discrepancy
    incorrect_val_df = pd.DataFrame(columns=df6_merge.columns)
    incorrect_val_df.drop(['OT','ID'], axis=1, inplace=True)
    incorrect_val_df['FORMAT'] = None
    move_col = incorrect_val_df.pop("FORMAT")
    incorrect_val_df.insert(0, "FORMAT", move_col)
    column_to_move = incorrect_val_df.pop("N VANN")
    incorrect_val_df.insert(3, "N VANN", column_to_move)
    column_to_move = incorrect_val_df.pop("CONT_CODE")
    incorrect_val_df.insert(4, "CONT_CODE", column_to_move)
    column_to_move = incorrect_val_df.pop("DEST_CODE")
    incorrect_val_df.insert(5, "DEST_CODE", column_to_move)

    print("\nCOMPARISON BETWEEN MF FORECAST AND OSP PxP FORECAST ----------------------------------------------")
    disc_report(2, df6c, df7c, "MF Forecast", "OSP Forecast", df7_leftmerge, df7_rightmerge, df7_merge, incorrect_val_df)

    # Cleaning merged dfs
    df7_leftmerge.drop(df7_leftmerge.iloc[:,137:], axis=1, inplace=True)
    df7_rightmerge.drop(['DEST_CODE_osp'],axis=1,inplace=True)
    df7_rightmerge.drop(df7_rightmerge.iloc[:,4:8], axis=1, inplace=True)
    df7_rightmerge.drop(df7_rightmerge.iloc[:,6:132], axis=1, inplace=True)

    # Filter df7_rightmerge (IMS unmatched) to exclude columns with N to N+2 sum <=1
    df7_rightmerge['Sum N-1 to N+2'] = df7_rightmerge['N-1 LOTS_ims'] + df7_rightmerge['N LOTS_ims'] + df7_rightmerge['N+1 LOTS_ims'] + df7_rightmerge['N+2 LOTS_ims']
    filtered_df7_rightmerge = df7_rightmerge[df7_rightmerge["Sum N-1 to N+2"] > 1]

    # dropping if any duplicates toget actual discrepancy
    incorrect_val_df.drop_duplicates(subset=incorrect_val_df.columns.difference(['FORMAT']), keep=False, inplace=True)

    # Creting dfs to add text details
    text_1_df = pd.DataFrame({'text_column': ['IMS Forecast: Below are additional entries only present in the IMS format',' ']})
    text_2_df = pd.DataFrame({'text_column': [' ','OSP Forecast: Below are additional entries only present in the OSP format',' ']})
    text_3_df = pd.DataFrame({'text_column': ['Observations column includes details about incorrect values for each entry','Note: Does not compare Dest Code Values but they are still provided',' ']})

    # Exporting files
    question = """Would you like to export the discrepancy report as .csv files?

    Enter Y for yes and N for no: """
    export_input = input(question)

    if (export_input == "Y" or export_input == "y"):
        file_path_1 = folder_path + "\\2.1.IMS_Forecast_Unmatched_Entries.csv"
        file_path_2 = folder_path + '\\2.2.OSP_Forecast_Unmatched_Entries.csv'
        file_path_3 = folder_path + '\\2.3.Forecast_Incorrect__Values.csv'
        
        text_1_df.to_csv(file_path_1, header=None, index=False)
        filtered_df7_rightmerge.to_csv(file_path_1, mode='a', index=False)

        text_2_df.to_csv(file_path_2, header=None, index=False)
        df7_leftmerge.to_csv(file_path_2, mode='a', index=False)

        text_3_df.to_csv(file_path_3, header=None, index=False)
        incorrect_val_df.to_csv(file_path_3, mode='a', index=False)
    else:
        print("")
    
# comparison_num_3() performs the comparison between the MF L3OS15 and OSP FA_consolidated files and exports the 
# discrepancy results into csv files.
def comparison_num_3(df4c, folder_path):
    L3OS15_file = "C:\\Users\\sharmas3\\Downloads\\Projects and Tasks Docs\\Discrepancy Reports\\L3OS15_11-22-2023_14-31-30.xlsm"
    df5 = pd.read_excel(L3OS15_file)
    df5c = df5.copy()

    # DATA CLEANING
    # Changing column names, datatypes, dropping unwanted columns and duplicates
    df5c.drop(['Unnamed: 100'], axis=1, inplace=True)
    df5c.rename(columns = {'PART NUMBER':'PARTNO','LOT SIZE':'LOT_SIZE','MONTH':'YYMM'}, inplace=True)
    df5c.drop_duplicates(inplace=True)
    # Applying formatting functions
    part_format(df5c,'PARTNO','-')

    # Formatting date to change from YYYYMM to YYMM 
    df5c['YYMM'] = df5c['YYMM'].astype(str).str[2:]
    # changing back to an int
    df5c['YYMM'] = df5c['YYMM'].astype(int)

    # Grouping df5c by part number
    df5_group = df5c.groupby(['PARTNO','YYMM','LOT_SIZE']).sum().reset_index()
    # Merging dataframes to get common and unmatched entries
    common_entries_df = pd.merge(df4c, df5_group, on=['PARTNO','YYMM','LOT_SIZE'], how='inner', suffixes=['_osp','_ims'])
    osp_leftmerge_df = pd.merge(df4c, df5_group, on=['PARTNO','YYMM','LOT_SIZE'], how='left', suffixes=['_osp','_ims'], indicator=True).query('_merge=="left_only"').drop(columns='_merge')
    ims_rightmerge_df = pd.merge(df4c, df5_group, on=['PARTNO','YYMM','LOT_SIZE'], how='right', suffixes=['_osp','_ims'], indicator=True).query('_merge=="right_only"').drop(columns='_merge')
    common_entries_df.drop_duplicates(inplace=True)

    # dropping duplicates
    df5_group.drop_duplicates(inplace=True)
    common_entries_df.drop_duplicates(inplace=True)

    # dataframes to store discrepancy
    wrong_totals_df = pd.DataFrame(columns=df5_group.columns)
    wrong_totals_df['FORMAT'] = None
    wrong_totals_df.rename(columns={'DEST CD':'DEST_CODE'}, inplace=True)
    move_col = wrong_totals_df.pop("DEST_CODE")
    wrong_totals_df.insert(0, "DEST_CODE", move_col)
    move_col = wrong_totals_df.pop("FORMAT")
    wrong_totals_df.insert(0, "FORMAT", move_col)

    # CALLING FUNCTION FOR COMPARISON
    print("\nCOMPARISON BETWEEN MF L3OS15 AND OSP FA_CONSOLIDATED ----------------------------------------------")
    disc_report(3, df5c, df4c, "MF L3OS15", "OSP FAConsolidated", osp_leftmerge_df, ims_rightmerge_df, common_entries_df, wrong_totals_df)

    # Cleaning up merged dfs
    osp_leftmerge_df.drop(osp_leftmerge_df.iloc[:,100:], axis=1, inplace=True)
    ims_rightmerge_df.drop(ims_rightmerge_df.iloc[:,4:100], axis=1, inplace=True)
    ims_rightmerge_df.drop(['DEST_CODE'], axis=1, inplace=True)

    # Creating dataframes to add text details
    text_1_df = pd.DataFrame({'text_column': ['IMS FA_consol Format: Below are additional entries only present in the IMS format',' ']})
    text_2_df = pd.DataFrame({'text_column': [' ','OSP TNSO Format: Below are additional entries only present in the OSP format',' ']})
    text_3_df = pd.DataFrame({'text_column': ['Observations column includes details about incorrect values for each entry', ' ']})

    # Exporting files
    question = """Would you like to export the discrepancy report as .csv files?
    
    Enter Y for yes and N for no: """
    export_input = input(question)

    if (export_input == "Y" or export_input == "y"):
        file_path_1 = folder_path + "\\3.1.IMS_FAconsol_Unmatched_Entries.csv"
        file_path_2 = folder_path + "\\3.2.OSP_FAconsol_Unmatched_Entries.csv"
        file_path_3 = folder_path + "\\3.3.FAconsol_Incorrect_Totals.csv"
        
        text_1_df.to_csv(file_path_1, header=None, index=False)
        ims_rightmerge_df.to_csv(file_path_1, mode='a', index=False)

        text_2_df.to_csv(file_path_2, header=None, index=False)
        osp_leftmerge_df.to_csv(file_path_2, mode='a', index=False)

        text_3_df.to_csv(file_path_3, header=None, index=False)
        wrong_totals_df.to_csv(file_path_3, mode='a', index=False)
    else:
        print("")

# comparison_num_4() performs the comparison between the MF L3OS01 and OSP Order Plan files and exports the 
# discrepancy results into csv files.
def comparison_num_4(df3c, folder_path):
    MF_OrderPlan = r'j:\$PC Materials\Overseas\09. Projects\OS Modernization\File Comparison Tool\IMS Order Plan Files\L3OS01_04-24-2024_13-38-36 5-9.CSV'
    OSP_OrderPlan = r'j:\$PC Materials\Overseas\09. Projects\OS Modernization\File Comparison Tool\OSP Daily Data Files\OrderForecast_20240424_113002_export_5-9.xlsx'
    df8 = pd.read_csv(MF_OrderPlan)
    df9 = pd.read_excel(OSP_OrderPlan)
    df8c = df8.copy()
    df9c = df9.copy()

    # dropping columns
    df8c.drop(['SYS ADJUSTMENT','FINAL LOT','FINAL C/O','TEAM MEMBER','KANBAN','PLAN'],axis=1, inplace=True)
    df9c.drop(['PROD DT','PART DESCRIPTION','QPC','DOCK','SPC','LIFE CYCLE','ORDER TYPE','NV RQMT','SYS ADJ',
            'BOA','MAN ADJ','ORDER/FORECAST','BO QTY','KVC SHIP','ADJ REASON','UPDATED DT','USER','KANBAN','ORDER PLAN'],axis=1, inplace=True)

    # renaming columns
    df8c.rename(columns={'DESTINATION CODE':'DEST_CODE','PART NUMBER':'PARTNO','VANNING DATE':'VAN_DATE',
                        'ORDER LOT':'ORDER_LOT','F/A MIN                                              ':'F/A MIN'},inplace=True)
    df9c.rename(columns={'PART NO':'PARTNO','VAN DT':'VAN_DATE','RQMT':'BASIC REQUIREMENT',
                        'C/O':'START C/O','FA MIN':'F/A MIN','FA MAX':'F/A MAX','CC':'CONT_CODE'},inplace=True)

    # formatting columns
    df8c['VAN_DATE'] = df8c['VAN_DATE'].astype(str)
    date_series = df8c['VAN_DATE'].str.split('/', expand=True)
    df8c['VAN_DATE'] = date_series[0] + date_series[1].str.zfill(2) + date_series[2].str.zfill(2)
    df8c['VAN_DATE'] = df8c['VAN_DATE'].astype(int)

    # applying formatting functions
    part_format(df9c,'PARTNO') 
    df9c['VAN_DATE'] = df9c['VAN_DATE'].astype(str)
    van_format(df9c,'VAN_DATE')

    df9c.drop(['PARTNO'], axis=1, inplace=True)
    df9c.rename(columns = {'new_PART':'PARTNO'}, inplace=True)

    # joining to get cc
    df8_merge = pd.merge(df8c, df3c[['CONT_CODE','DEST_CODE','PARTNO','ORDER_LOT']], on=['DEST_CODE','PARTNO','ORDER_LOT'], how='left')
    insert_cont_code(df8_merge, df9c)
    update_cont_code(df8_merge)

    # merging dataframes
    df9_merge = pd.merge(df9c, df8_merge, on=['VAN_DATE','PARTNO','CONT_CODE'], how='inner',suffixes=['_osp','_ims'])
    df9_leftmerge = pd.merge(df9c, df8_merge, on=['VAN_DATE','PARTNO','CONT_CODE'], how='left', suffixes=['_osp','_ims'], indicator=True).query('_merge=="left_only"').drop(columns='_merge')
    df9_rightmerge = pd.merge(df9c, df8_merge, on=['VAN_DATE','PARTNO','CONT_CODE'], how='right', suffixes=['_osp','_ims'], indicator=True).query('_merge=="right_only"').drop(columns='_merge')

    # cleaning up merged df and reordering columns
    df9_leftmerge.drop(['DEST_CODE','ORDER_LOT','START C/O_ims','BASIC REQUIREMENT_ims','F/A MAX_ims','F/A MIN_ims'], axis=1,inplace=True)
    df9_rightmerge.drop(['START C/O_osp','BASIC REQUIREMENT_osp','F/A MAX_osp','F/A MIN_osp'], axis=1,inplace=True)
    df9_merge.drop(['F/A MAX_ims','F/A MIN_ims','F/A MAX_osp','F/A MIN_osp'],axis=1, inplace=True)

    df9_merge_order = ['DEST_CODE','PARTNO','CONT_CODE','VAN_DATE','ORDER_LOT','START C/O_ims','START C/O_osp','BASIC REQUIREMENT_ims','BASIC REQUIREMENT_osp']
    df9_merge = df9_merge.reindex(columns=df9_merge_order)

    # dataframe to store discrepancy
    unmatched_values_df = pd.DataFrame(columns=df9_merge.columns)

    # calling comparison function
    print("\nCOMPARISON BETWEEN MF ORDER PLAN AND OSP ORDER PLAN----------------------------------------------")
    disc_report(4, df8c, df9c, "MF Order Plan", "OSP Order Plan", df9_leftmerge, df9_rightmerge, df9_merge, unmatched_values_df)

    # exporting dataframes
    # Creting dfs to add text details
    text_1_df = pd.DataFrame({'text_column': ['IMS Order Plan: Below are additional entries only present in the IMS format',' ']})
    text_2_df = pd.DataFrame({'text_column': ['OSP Order Plan: Below are additional entries only present in the OSP format',' ']})
    text_3_df = pd.DataFrame({'text_column': ['Below are entries with incorrect START C/O and BASIC REQUIREMENT Values',' ']})

    # Exporting files
    question = """Would you like to export the discrepancy report as .csv files?
    Enter Y for yes and N for no: """
    export_input = input(question)

    if (export_input == "Y" or export_input == "y"):
        file_path_1 = folder_path + "\\4.1.IMS_OrderPlan_Unmatched_Entries.csv"
        file_path_2 = folder_path + "\\4.2.OSP_OrderPlan_Unmatched_Entries.csv"
        file_path_3 = folder_path + "\\4.3.OrderPlan_Incorrect__Values.csv"

        text_1_df.to_csv(file_path_1, header=None, index=False)
        df9_rightmerge.to_csv(file_path_1, mode='a', index=False)

        text_2_df.to_csv(file_path_2, header=None, index=False)
        df9_leftmerge.to_csv(file_path_2, mode='a', index=False)

        text_3_df.to_csv(file_path_3, header=None, index=False)
        unmatched_values_df.to_csv(file_path_3,mode='a', index=False)
    else:
        print("")


def cumulative_sum(df3c, folder_path):
    # OSP FILES
    Current_OSP_OrderPlan = r'j:\$PC Materials\Overseas\09. Projects\OS Modernization\File Comparison Tool\OSP Daily Data Files\OrderForecast_20240424_113002_export_5-9.xlsx'
    Before_OSP_OrderPlan = r'j:\$PC Materials\Overseas\09. Projects\OS Modernization\File Comparison Tool\Copy of Cumulative Files\OrderPlan_Cumulative_Sum 5-8.csv'

    df8 = pd.read_excel(Current_OSP_OrderPlan)
    # df9 = pd.read_excel(Before_OSP_OrderPlan)
    df9 = pd.read_csv(Before_OSP_OrderPlan)
    df8c = df8.copy()
    df9c = df9.copy()

    # IMS FILES
    Current_IMS_OrderPlan = r'j:\$PC Materials\Overseas\09. Projects\OS Modernization\File Comparison Tool\IMS Order Plan Files\L3OS01_04-24-2024_13-38-36 5-9.CSV'
    Before_IMS_OrderPlan = r'j:\$PC Materials\Overseas\09. Projects\OS Modernization\File Comparison Tool\Copy of Cumulative Files\OrderPlan_Cumulative_Sum 5-8.csv'
    df10 = pd.read_csv(Current_IMS_OrderPlan)
    df11 = pd.read_csv(Before_IMS_OrderPlan)
    df10c = df10.copy()
    df11c = df11.copy()

    # USER INPUT TO DETERMINE IF THE FILE IS THE FIRST OF THE MONTH
    question2 = """Is the Vanning Date of one of the files the first of the month?
        Enter Y for yes and N for no: """
    first_month_file = input(question2)

    if (first_month_file == "Y" or first_month_file == "y"):
        # OSP
        df9c['PREV_OSP_SUM'] = df9c.loc[:, 'RQMT']
        df9c.drop(['PROD DT','PART DESCRIPTION','QPC','SPC','LIFE CYCLE','ORDER TYPE','NV RQMT','SYS ADJ','VAN DT','RQMT',
            'BOA','MAN ADJ','ORDER/FORECAST','BO QTY','KVC SHIP','ADJ REASON','UPDATED DT','USER','ORDER PLAN','FA MIN','FA MAX','KANBAN'],axis=1, inplace=True)
        df9c.rename(columns={'PART NO':'PARTNO','CC':'CONT_CODE'},inplace=True)
        part_format(df9c,'PARTNO')
        df9c.drop(['PARTNO','C/O'], axis=1, inplace=True)
        df9c.rename(columns = {'new_PART':'PARTNO'}, inplace=True)

        # IMS
        df11c['PREV_IMS_SUM'] = df11c.loc[:, 'BASIC REQUIREMENT']
        df11c.drop(['START C/O','SYS ADJUSTMENT','FINAL LOT','FINAL C/O','TEAM MEMBER','VANNING DATE','PLAN',
                    'BASIC REQUIREMENT','F/A MIN                                              ','F/A MAX','KANBAN'],axis=1, inplace=True)
        df11c.rename(columns={'DESTINATION CODE':'DEST_CODE','PART NUMBER':'PARTNO','ORDER LOT':'ORDER_LOT'},inplace=True) 
        # merging to get cc
        df11c = pd.merge(df11c, df3c[['CONT_CODE','DEST_CODE','PARTNO','ORDER_LOT']], on=['DEST_CODE','PARTNO','ORDER_LOT'], how='left')
        insert_cont_code(df11c, df9c)
        update_cont_code(df11c)

    else:
        df9c.rename(columns={'CUMULATIVE_SUM_OSP':'PREV_OSP_SUM'}, inplace=True)
        df11c.rename(columns={'CUMULATIVE_SUM_IMS':'PREV_IMS_SUM'}, inplace=True)
        df9c.drop(['DEST_CODE','ORDER_LOT'],axis=1, inplace=True)
        df11c.drop(['DOCK'],axis=1, inplace=True)

    # dropping columns
    df8c.drop(['PROD DT','PART DESCRIPTION','QPC','SPC','LIFE CYCLE','ORDER TYPE','NV RQMT','SYS ADJ','C/O',
            'BOA','MAN ADJ','ORDER/FORECAST','BO QTY','KVC SHIP','ADJ REASON','UPDATED DT','USER','ORDER PLAN','FA MIN','FA MAX','KANBAN'],axis=1, inplace=True)
    df10c.drop(['START C/O','SYS ADJUSTMENT','VANNING DATE','FINAL LOT','FINAL C/O','TEAM MEMBER','PLAN','F/A MIN                                              ','F/A MAX','KANBAN'],axis=1, inplace=True)

    # renaming columns
    df8c.rename(columns={'PART NO':'PARTNO','VAN DT':'VAN_DATE','RQMT':'BASIC_REQUIREMENT','CC':'CONT_CODE'},inplace=True)
    df10c.rename(columns={'DESTINATION CODE':'DEST_CODE','PART NUMBER':'PARTNO','ORDER LOT':'ORDER_LOT'},inplace=True) 

    # fill nans in BR values with 0
    df8c.fillna({'BASIC_REQUIREMENT':0}, inplace=True)
    df10c.fillna({'BASIC REQUIREMENT':0}, inplace=True)
    df9c.fillna({'PREV_OSP_SUM':0}, inplace=True)
    df11c.fillna({'PREV_IMS_SUM':0}, inplace=True)

    # applying formatting functions
    df8c['VAN_DATE'] = df8c['VAN_DATE'].astype(str)
    van_format(df8c,'VAN_DATE')
    part_format(df8c,'PARTNO')
    df8c.drop(['PARTNO'], axis=1, inplace=True)
    df8c.rename(columns = {'new_PART':'PARTNO'}, inplace=True)

    # joining to get cc in df10c
    df10_merge = pd.merge(df10c, df3c[['CONT_CODE','DEST_CODE','PARTNO','ORDER_LOT']], on=['DEST_CODE','PARTNO','ORDER_LOT'], how='left')
    insert_cont_code(df10_merge, df8c)
    update_cont_code(df10_merge)

    # merging dataframes
    merge_ims = pd.merge(df11c, df10_merge, on=['PARTNO','CONT_CODE','ORDER_LOT'], how='outer', suffixes=['__before','__'])
    merge_osp = pd.merge(df9c[['DOCK','PARTNO','CONT_CODE','PREV_OSP_SUM']], df8c[['DOCK','PARTNO','CONT_CODE','BASIC_REQUIREMENT']], 
                        on=['DOCK','PARTNO','CONT_CODE'], how='outer',suffixes=['_before','_'])

    # dropping unwantd columns
    merge_ims.drop(['DEST_CODE__before'], axis=1, inplace=True)

    # grouping BR for cum sum
    merge_ims_group = merge_ims.groupby(['PARTNO','CONT_CODE','ORDER_LOT']).sum().reset_index()
    merge_osp_group = merge_osp.groupby(['PARTNO','DOCK','CONT_CODE']).sum().reset_index()

    # merging osp and ims dataframes
    cum_sum_df = pd.merge(merge_ims_group, merge_osp_group, on=['PARTNO','CONT_CODE'], how='outer',suffixes=['_before','_'])

    current_van_date = str(df8c['VAN_DATE'].iloc[0])

    cum_sum_df['CUMULATIVE_SUM_OSP'] = cum_sum_df['PREV_OSP_SUM'] + cum_sum_df['BASIC_REQUIREMENT']
    cum_sum_df['CUMULATIVE_SUM_IMS'] = cum_sum_df['PREV_IMS_SUM'] + cum_sum_df['BASIC REQUIREMENT']
    cum_sum_df[current_van_date] = cum_sum_df['CUMULATIVE_SUM_OSP'] - cum_sum_df['CUMULATIVE_SUM_IMS']

    # dropping unwanted columns
    cum_sum_df.drop(['PREV_OSP_SUM','PREV_IMS_SUM','BASIC REQUIREMENT','BASIC_REQUIREMENT'],axis=1,inplace=True)
    cum_sum_df.rename(columns={'DEST_CODE__':'DEST_CODE'},inplace=True)

    # order is in reverse of what we need so that these columns can be inserted at the very beginning
    reindex_order = ['CUMULATIVE_SUM_OSP','CUMULATIVE_SUM_IMS','ORDER_LOT','DOCK','CONT_CODE','PARTNO','DEST_CODE']
    selected_columns = [cum_sum_df.pop(col) for col in reindex_order]

    for col, col_name in zip(selected_columns, reindex_order):
        cum_sum_df.insert(0, col_name, col)

    question = """Would you like to export the cumulative sum report as .csv files?
        Enter Y for yes and N for no: """
    export_input = input(question)

    # Exporting
    if (export_input == "Y" or export_input == "y"):
        file_path = folder_path + "\\OrderPlan_Cumulative_Sum.csv\\" + current_van_date
        copy_file_path = r"j:\$PC Materials\Overseas\09. Projects\OS Modernization\File Comparison Tool\Copy of Cumulative Files" + "\\OrderPlan_Cumulative_Sum.csv\\" + current_van_date
        cum_sum_df.to_csv(file_path, index=False)
        cum_sum_df.to_csv(copy_file_path, index=False)
    else:
        print("")

# comparison() is a function which requests user input to determine which files need to be compared.
# User is prompted for an input until 0 is entered.
def comparison():

    question = """\nPick a number listed below for the files you wish to compare:\n 
    1. MF Final Order(L3OS06) and OSP Final Order
    2. MF Forecast AND OSP Forecast
    3. MF L3OS15 AND OSP FA_Consolidated
    4. MF Order Plan(L3OS01) AND OSP Order Plan
    5. Cumulative Sum of Order Plan Files\n
    Enter 0 to quit

    Enter Number: """
    user_request = int(input(question))

    while (user_request != 0):
        if (user_request < 0 or user_request > 5):
            print("Invalid Input")
        elif (user_request == 1):
            comparison_num_1(df3c, folder_path)
        elif (user_request == 2):
            comparison_num_2(df3c, folder_path)
        elif (user_request == 3):
            comparison_num_3(df4c, folder_path)
        elif (user_request == 4):
            comparison_num_4(df3c, folder_path)
        else:
            cumulative_sum(df3c, folder_path)
        
        user_request = int(input(question))

comparison()