import os
import pandas as pd

# set the path to the folder containing the Excel files
folder_path = ""

# get a list of all the Excel files in the folder
excel_files = [f for f in os.listdir(folder_path) if not f.startswith("~$")]

# create an empty dataframe to store the combined data
combined_data = pd.DataFrame()

# loop through each Excel file and append its data to the combined dataframe
for file in excel_files:
    file_path = "/".join([folder_path, file])
    plant_name_col = pd.read_excel(
        file_path, sheet_name="", usecols="", header=None, nrows=1
    )
    #replace plant_name with specific variable
    plant_name = plant_name_col.values[0][0].strip()
    #replace date_col with specific variable

    date_col = pd.read_excel(
        file_path,
        sheet_name="",
        skiprows=1,
        nrows=1,
        usecols="",
        header=None,
    )
    date = date_col.values[0][0]

    # select the "BreakDown_Details" sheet and get the desired columns
    details_sheet = pd.read_excel(
        file_path,
        sheet_name="",
        skiprows=3,
        nrows=250,
        # usecols=" #define columns, for example C:K,N:S",
    )

    details_sheet = details_sheet[details_sheet["BreakDown"].notna()]
    # select the columns from C4 to S4 and append them to the combined dataframe
    columns = details_sheet.columns
    df_subset = details_sheet[columns]
    # combine the data into a single dataframe
    df_subset.insert(0, "", plant_name)
    df_subset.insert(1, "", date)
    combined_data = combined_data.append(df_subset)

# write the combined data to a new Excel file
combined_file_path = "combined_data.xlsx"
combined_data.to_excel(combined_file_path, index=False)
