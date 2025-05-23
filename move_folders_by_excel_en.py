"""
Script Description:
This script moves folders based on information in an Excel patient data table.
It supports configuring multiple move tasks, each specifying a different Excel file, sheet, matching column, source folder, and destination folder.
It can also filter patients based on the value in a specific Excel column.
Folder names are assumed to be in the format "ID-Name".
Log information will be saved to the move_file.log file.
The terminal will only print information about successfully moved folders and task summaries.

Dependencies: pandas, openpyxl
"""

import os
import shutil
import pandas as pd

# Log file path
LOG_FILE = 'move_file.log'

# Open log file
try:
    log_file_handle = open(LOG_FILE, 'w', encoding='utf-8')
except IOError as e:
    print(f"Error: Unable to open log file {LOG_FILE} for writing: {e}")
    log_file_handle = None # If opening fails, set handle to None and don't try to write later

def log_message(message):
    """Writes a message to the log file if the file handle is valid."""
    if log_file_handle:
        try:
            log_file_handle.write(message + '\n')
        except Exception as e:
            # Even if writing to log fails, don't interrupt the main process, just print error to terminal
            print(f"Error: Failed to write to log file: {e}")


# ========== Configuration Section ==========
# move_tasks is a list, each element representing an independent folder moving task.
# Each dictionary element corresponds to a set of configuration items:
# 'excel_path': Full path to the Excel file to read.
# 'sheet_name': The sheet in the Excel file to read. Can be a sheet name string or a 0-based sheet index (integer).
# 'name_col': The column in the Excel file containing data for matching folder names (ID or Name). Can be a column name string or a 0-based column index (integer).
# 'header': The row (0-based index) to use as the column headers. For example, if headers are on the second row of Excel, set header to 1.
# 'source_path': The path to the source folder containing folders to be moved.
# 'destination_path': The path to the destination folder for moved folders. Can use relative paths, relative to the script's current working directory.
# 'filter_col': (Optional) The column to use for filtering based on its value. Can be a column name string or a 0-based column index (integer). If no filtering is needed, this can be omitted.
# 'filter_value': (Optional) The target value to filter by. Only rows where the value in filter_col equals this value will be used for matching folders. If no filtering is needed, this can be omitted.
#
# Please modify the following move_tasks configuration list according to your actual needs:
move_tasks = [
    {
        'excel_path': '/path/to/your/excel_file_1.xlsx',
        'sheet_name': 0,
        'name_col': 'PatientID', # Example: column name 'PatientID'
        'header': 0,
        'source_path': '/path/to/your/source_folder_1',
        'destination_path': './processed_data_1',
        'filter_col': 'Diagnosis', # Example: column name 'Diagnosis'
        'filter_value': 'UA'
    },
    {
        'excel_path': '/path/to/your/excel_file_2.xlsx',
        'sheet_name': 'Sheet1', # Example: sheet name 'Sheet1'
        'name_col': 2,          # Example: column index 2 (3rd column)
        'header': 1,
        'source_path': '/path/to/your/source_folder_2',
        'destination_path': './processed_data_2',
        # filter_col and filter_value are optional
    },
    # Add more tasks as needed, following the structure above.
    # Remember to use either column names (strings) or column indices (integers, 0-based) consistently
    # for name_col and filter_col within each task.
]
# ========== Configuration Section End ==========


# Iterate through each folder moving task
for task in move_tasks:
    # Extract configuration information from the current task dictionary
    excel_path = task['excel_path']
    sheet_name = task['sheet_name']
    name_col = task['name_col']
    header = task['header']
    source_path = task['source_path']
    destination_path = task['destination_path']
    # Use .get() method to get optional filter configurations, None if not present
    filter_col = task.get('filter_col') # Get filter column name or index
    filter_value = task.get('filter_value') # Get target filter value

    log_message(f"\n--- Starting task: Excel file '{excel_path}', Sheet '{sheet_name}', Source folder '{source_path}', Destination folder '{destination_path}' ---")
    if filter_col is not None:
        log_message(f"  Filter condition: Column '{filter_col}' has value '{filter_value}'")
    else:
        log_message("  No filter condition, using all data in the specified column for matching.")


    # Build usecols list for pd.read_excel to read only necessary columns for efficiency
    # usecols parameter requires a list of column names or column indices
    use_cols_list = [name_col]
    # If filter column is set and is different from the name column, add it to the list of columns to read
    # Ensure element types in the usecols list are consistent
    if filter_col is not None and filter_col != name_col:
         # Try to convert filter_col and name_col to integers, if fails keep original type (might be string column names)
        try:
            int_name_col = int(name_col)
            int_filter_col = int(filter_col)
            use_cols_list = [int_name_col]
            if int_filter_col != int_name_col: # Check again to avoid adding twice
                 use_cols_list.append(int_filter_col)
        except (ValueError, TypeError):
            # If conversion fails, it means column names were used, ensure usecols is a list of strings
             use_cols_list = [str(name_col)]
             if filter_col is not None and str(filter_col) != str(name_col): # Check again to avoid adding twice
                 use_cols_list.append(str(filter_col))


    # Read data from the Excel file
    try:
        # Read data for the specified sheet, header, and columns
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header, usecols=use_cols_list)
        log_message(f"Successfully read Excel file '{excel_path}' Sheet '{sheet_name}'.")
        log_message(f"Successfully read Excel file, Sheet '{sheet_name}' has {len(df.columns)} columns")
        log_message(f"Column names: {df.columns.tolist()}")

        #log_message("DataFrame before filtering (first 5 rows):") # Debug printing, uncomment if needed
        #log_message(df.head().to_string()) # Debug printing
        if filter_col is not None and df.shape[1] > 1:
            #log_message(f"First 5 values of filter column '{df.columns[1]}':") # Debug printing
            #log_message(str(df.iloc[:5, 1].tolist())) # Debug printing
            log_message(f"Target filter value: '{filter_value}' (Type: {type(filter_value)})")


    except FileNotFoundError:
        log_message(f"Error: Excel file not found - {excel_path}, skipping current task.")
        print(f"Error: Excel file not found - {excel_path}, skipping current task. See log file {LOG_FILE} for details.")
        continue # File not found, skip current task and continue to the next
    except ValueError as e:
        # Catch read failures due to incorrect parameters like sheet_name, header, usecols
        log_message(f"Failed to read Excel file: {excel_path} - {e}")
        log_message(f"Please check if sheet_name ({sheet_name}), header ({header}), name_col ({name_col}), and filter_col ({filter_col}) are correct in the task configuration.")
        print(f"Failed to read Excel file: {excel_path} - {e}. See log file {LOG_FILE} for more details.")
        continue # Read failed, skip current task and continue to the next
    except Exception as e:
        # Catch other possible read errors
        log_message(f"An unknown error occurred while reading Excel file: {excel_path} - {e}, skipping current task.")
        print(f"An unknown error occurred while reading Excel file: {excel_path} - {e}. See log file {LOG_FILE} for more details.")
        continue


    # Based on whether a filter condition is set, get the list of names/IDs for matching folders
    names_list = [] # Initialize list of names/IDs
    if filter_col is not None:
        # === Perform filtering operation ===
        # Ensure DataFrame has at least two columns, first for matching, second for filtering
        if df.shape[1] < 2:
             log_message(f"Error: Sheet '{sheet_name}' has insufficient columns ({df.shape[1]} columns) for filtering. Please check if name_col ({name_col}) and filter_col ({filter_col}) are correct in the task configuration. Skipping current task.")
             print(f"Error: Sheet '{sheet_name}' has insufficient columns for filtering. See log file {LOG_FILE} for details.")
             continue

        # Filter rows based on the filter condition and extract data from the name/ID column
        try:
            # Use iloc[:, 1] to access the second column (index 1) of the new DataFrame, which corresponds to the original filter_col data
            # Use iloc[:, 0] to access the first column (index 0) of the new DataFrame, which corresponds to the original name_col data
            # Convert filter column values to string and strip whitespace for better matching robustness
            df_filtered = df.loc[df.iloc[:, 1].astype(str).str.strip() == str(filter_value).strip(), df.columns[0]]
            # Convert filtered data to a string list and drop empty values
            names_list = df_filtered.dropna().astype(str).tolist()
            log_message(f"Obtained {len(names_list)} names/IDs for moving based on filter condition.")

            #log_message("Filtered DataFrame (first 5 rows):") # Debug printing, uncomment if needed
            #log_message(df_filtered.head().to_string()) # Debug printing
            #log_message(f"names_list obtained from filtered results (first 5): {names_list[:5]}") # Debug printing


        except Exception as e:
             # Catch potential errors during filtering
             log_message(f"An error occurred while getting the list of names from file '{excel_path}' Sheet '{sheet_name}' based on filter condition: {e}, skipping current task.")
             print(f"An error occurred while getting the list of names based on filter condition. See log file {LOG_FILE} for details.")
             continue

    else:
        # === No filtering, use all data in the specified column ===
        # Ensure DataFrame has at least one column for matching
        if df.shape[1] < 1:
             log_message(f"Error: Sheet '{sheet_name}' has no columns, unable to get matching data. Please check if name_col ({name_col}) is correct in the task configuration. Skipping current task.")
             print(f"Error: Sheet '{sheet_name}' has no columns for matching data. See log file {LOG_FILE} for details.")
             continue
        # Get all data from the first column (index 0) of the new DataFrame
        names_list = df.iloc[:, 0].dropna().astype(str).tolist()
        log_message(f"No filter condition set, obtained {len(names_list)} names/IDs for moving.")


    # Remove the '001-' prefix from items in names_list (if present)
    processed_names_list = []
    removed_prefix_count = 0
    for name in names_list:
        # Ensure it's a string before processing
        if isinstance(name, str) and name.strip().lower().startswith('001-'):
            processed_name = name.strip()[len('001-'):]
            processed_names_list.append(processed_name)
            removed_prefix_count += 1
            log_message(f"Removed prefix '001-': Original '{name}' -> Processed '{processed_name}'")
        else:
            processed_names_list.append(name)
    names_list = processed_names_list # Update names_list with the processed list
    if removed_prefix_count > 0:
        log_message(f"Successfully removed '001-' prefix from {removed_prefix_count} items.")


    # === Iterate through the source folder to find and move matching folders ===
    log_message(f"Starting search for matching folders in source folder '{source_path}'...")
    moved_count = 0 # Count of successfully moved folders
    skipped_count = 0 # Count of skipped folders (format incorrect or destination exists or move failed)
    not_matched_count = 0 # Count of folders not found in the Excel list

    # Check if source folder exists
    if not os.path.exists(source_path):
        log_message(f"Error: Source folder '{source_path}' does not exist, unable to perform move operation, skipping current task.")
        print(f"Error: Source folder '{source_path}' does not exist. See log file {LOG_FILE} for details.")
        continue # Source folder does not exist, skip current task

    # Iterate through all files and folders in the source folder
    for entry_name in os.listdir(source_path):
        source_entry_path = os.path.join(source_path, entry_name)
        # Only process directories
        if os.path.isdir(source_entry_path):
            folder_name = entry_name # Folder name

            # Check folder name format, assumed to be "ID-Name"
            if '-' in folder_name:
                folder_name_split = folder_name.split('-', 1) # Split only once, in case name contains '-'
                # Ensure there are two parts after splitting
                if len(folder_name_split) == 2:
                    id_in_folder = folder_name_split[0].strip() # Get ID part and strip whitespace
                    name_in_folder = folder_name_split[1].strip() # Get Name part and strip whitespace

                    # Check if the value from Excel (from names_list) matches the ID or Name part of the folder
                    # Convert both Excel value and folder name parts to lowercase for case-insensitive matching
                    matched_excel_value = None # To record the matched Excel value
                    for excel_value in names_list:
                        # Convert Excel value to string, strip whitespace, and convert to lowercase
                        processed_excel_value = str(excel_value).strip().lower()
                        # Convert folder ID and Name to lowercase
                        lower_id_in_folder = id_in_folder.lower()
                        lower_name_in_folder = name_in_folder.lower()

                        if processed_excel_value == lower_id_in_folder or processed_excel_value == lower_name_in_folder:
                            matched_excel_value = excel_value # Record the original Excel value
                            break # Found a match, break the inner loop

                    if matched_excel_value is not None:
                        # Construct the full path for the destination folder
                        target_folder_path = os.path.join(destination_path, folder_name)

                        # Check if the destination parent folder exists, create if not
                        target_parent_dir = os.path.dirname(target_folder_path)
                        if not os.path.exists(target_parent_dir):
                            try:
                                os.makedirs(target_parent_dir)
                                log_message(f"Created destination parent folder: {target_parent_dir}")
                            except OSError as e:
                                log_message(f"Error: Unable to create destination parent folder {target_parent_dir}: {e}, unable to move folder {folder_name}")
                                print(f"Error: Unable to create destination parent folder {target_parent_dir}, unable to move folder {folder_name}. See log file {LOG_FILE} for details.")
                                skipped_count += 1 # Unable to create destination folder, count as skipped
                                continue # Skip current folder

                        # Check if destination folder already exists to avoid duplicate moves or overwriting
                        if os.path.exists(target_folder_path):
                             log_message(f"Destination folder {target_folder_path} already exists, skipping moving folder {folder_name}")
                             print(f"Destination folder {target_folder_path} already exists, skipping moving folder {folder_name}. See log file {LOG_FILE} for details.")
                             skipped_count += 1 # Destination folder already exists, count as skipped
                             continue # Skip current folder

                        # Move the folder
                        try:
                            shutil.move(source_entry_path, target_folder_path)
                            print(f"Successfully moved folder '{folder_name}' to '{target_folder_path}' (Matched Excel value: '{matched_excel_value}')")
                            moved_count += 1 # Successfully moved, increment count
                            log_message(f"Successfully moved folder '{folder_name}' to '{target_folder_path}' (Matched Excel value: '{matched_excel_value}')")
                        except Exception as e:
                            log_message(f"Failed to move folder '{folder_name}' to '{target_folder_path}': {e}")
                            print(f"Failed to move folder '{folder_name}' to '{target_folder_path}': {e}. See log file {LOG_FILE} for details.")
                            skipped_count += 1 # Move failed, count as skipped

                    else:
                        # Folder name matches format but no match found in the Excel list
                        log_message(f"Folder '{folder_name}' (ID: '{id_in_folder}', Name: '{name_in_folder}') not found in the Excel list ({excel_path}, Sheet '{sheet_name}', Column '{name_col}'), skipping")
                        not_matched_count += 1 # Not matched, increment count

                else:
                    # Folder name contains '-' but split into more/less than 2 parts
                    log_message(f"Folder '{folder_name}' format incorrect (expected 'ID-Name', actual split parts count is not 2), skipping")
                    skipped_count += 1 # Incorrect format, count as skipped
            else:
                # Folder name does not contain '-'
                log_message(f"Folder '{folder_name}' does not contain '-', skipping")
                skipped_count += 1 # Does not contain '-', count as skipped
        # If it's a file, skip it
        # else:
        #     log_message(f"'{entry_name}' is a file, skipping") # Uncomment to log skipped files

    log_message(f"--- Task processing complete: Successfully moved {moved_count} folders, skipped {skipped_count} folders (incorrect format, destination exists, or move failed), {not_matched_count} folders not matched in Excel. ---")
    print(f"Task processing complete: Successfully moved {moved_count} folders. See log file {LOG_FILE} for details.")

print("\nAll tasks processed.")

# Close log file
if log_file_handle:
    log_file_handle.close() 