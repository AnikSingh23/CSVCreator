import pandas as pd
import os
import warnings

# Suppress specific warning
warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

# Provinces, sectors, and source folder
Prov = ["ab", "atl", "bc", "can", "mb", "nb", "nl", "ns", "on", "pe", "qc", "sk", "ter"]
Sec = ["com", "res", "tra", "ind", "ind template"]
source_folder = r'C:\Users\anik1\Desktop\Work\LEAP\leap-canada all scenarios_sperry et al._2023-03-16'

# Variables for row and column settings (can be adjusted as needed, Note you must remove 2 from the row that is why the minus 2 is there)
row_years = 10 - 2  # Row index for the years (common for both AGR and other files)
row_values = 12 - 2  # Row index for the values (common for both AGR and other files)
agr_name_row = 6 - 2  # Row index for the name in AGR.xlsx
default_name_row = 7 - 2  # Row index for the name in other files

# CSV file name (can be changed here)
csv_file_name = "Total_energy_end_use_data.csv"  # Output CSV file name

file_paths = []

# Nested for loop to create file paths
for i in Prov:
    for j in Sec:
        file_path = os.path.join(source_folder, i, i + " " + j + ".xlsx")
        file_paths.append(file_path)

# Placeholder for the combined data
combined_data = []
all_years = set()  # Use a set to collect all unique years


# Function to read data from excel with error handling
def read_excel_data(file_path, table, row_years, row_values, name_row, is_agr=False):
    try:
        df = pd.read_excel(file_path, sheet_name=table)
        years_data = df.iloc[row_years, 2:].values  # Skip the first two columns (0 and 1)
        values = df.iloc[row_values, 2:].values
        name = df.iloc[name_row, 0]  # Get the name from the specified row, column 1
        print(name)
        return years_data, values, name
    except FileNotFoundError:
        print(f"File not found: {file_path}. Skipping...")
        return None, None, None
    except Exception as e:
        print(f"Error reading file {file_path}: {e}. Skipping...")
        return None, None, None


# Loop through file paths and save data for each province and sector
for file_path in file_paths:
    province = file_path.split("\\")[-2]  # Extract province from file path
    main_sector = file_path.split(" ")[-1].replace(".xlsx", "")

    # Exclude "can ind" files
    if province == "can" and "ind" in main_sector.lower():
        print(f"Skipping {file_path} as it is in a different format.")
        continue

    print(file_path)

    # Opening file for all sectors (Table 1)
    years_data, energy_end_use, name = read_excel_data(file_path, "Table 1", row_years, row_values, name_row=default_name_row)

    if years_data is None or energy_end_use is None:
        # Skip if the file was not found or another error occurred
        continue

    # Add years to the set to ensure we capture all unique years
    all_years.update(years_data)

    # Prepare row with province, sector name, and name from the appropriate row
    sector_name = f"{province}_{main_sector}_{name}"
    row = [sector_name] + list(energy_end_use)
    combined_data.append((years_data, row))  # Save both years and row for processing later

    # For "ind", we take data from multiple tables (Table 3 to Table 12)
    if "ind" in main_sector.lower():
        for table_num in range(3, 13):
            years_data, ind_values, table_name = read_excel_data(file_path, f"Table {table_num}", row_years, row_values, name_row=default_name_row)
            if ind_values is None:
                continue
            # Create a separate sector_name for each table, including the name in the specified row
            ind_sector_name = f"{province}_ind_{table_name}"
            ind_row = [ind_sector_name] + list(ind_values)
            combined_data.append((years_data, ind_row))

# Handle the special case of AGR.xlsx
agr_file_path = os.path.join(source_folder, "can", "AGR.xlsx")
agr_tables = ["Table 5", "Table 5-A", "Table 5-B", "Table 5-C", "Table 6", "Table 7", "Table 8", "Table 9", "Table 10", "Table 11", "Table 11-A", "Table 11-B"]

for table in agr_tables:
    # Pass the `is_agr=True` flag and use the `agr_name_row` for AGR.xlsx
    years_data, agr_values, table_name = read_excel_data(agr_file_path, table, row_years, row_values, name_row=agr_name_row, is_agr=True)
    if agr_values is None:
        continue
    # Create a unique name for each AGR table
    agr_sector_name = f"can_agr_{table_name}"
    agr_row = [agr_sector_name] + list(agr_values)
    combined_data.append((years_data, agr_row))

    # Add years to the set to ensure we capture all unique years
    all_years.update(years_data)

# Create a sorted list of all unique years
all_years = sorted(list(all_years))

# Now create the combined DataFrame
final_data = []
for years_data, row in combined_data:
    # Create a dictionary of year -> value for each row
    data_dict = {year: value for year, value in zip(years_data, row[1:])}

    # Ensure the row has values for all years, and fill missing values with "-"
    final_row = [row[0]] + [data_dict.get(year, "-") for year in all_years]
    final_data.append(final_row)

# Write the combined data to a CSV file
if final_data:
    years_column = ["Sector"] + all_years  # Use all unique years collected
    combined_df = pd.DataFrame(final_data, columns=years_column)
    combined_df.to_csv(csv_file_name, index=False)  # Use the variable for CSV file name
    print(f"CSV file '{csv_file_name}' created successfully.")
else:
    print("No data available to write to the CSV file.")
