# Folder Mover by Excel Configuration

This Python script moves folders based on criteria specified in an Excel file. It is useful for organizing data directories according to external metadata.

## Features

-   Configurable tasks for moving folders based on different Excel files and sheets.
-   Filters rows in the Excel file based on a specified column and value.
-   Matches folder names (assuming 'ID-Name' format) against a list generated from the Excel file.
-   Automatically removes the '001-' prefix from matching IDs if present in the Excel data.
-   Logs detailed information about the process (skipped folders, errors) to a log file (`move_file.log`).
-   Prints only successful folder moves to the console for a cleaner output.

## Prerequisites

-   Python 3.x
-   pandas library (`pip install pandas`)
-   openpyxl library (`pip install openpyxl`) - required for reading `.xlsx` files

## Installation

1.  Clone this repository or download the script files.
2.  Navigate to the script directory in your terminal.
3.  Install the required libraries:

    ```bash
    pip install -r requirements.txt
    ```

## Usage

1.  **Configure `move_tasks`**: Open `move_folders_by_excel.py` and modify the `move_tasks` list. Each dictionary in the list represents a moving task with specific configurations:

    -   `excel_path`: Full path to the Excel file.
    -   `sheet_name`: Sheet name (string) or index (integer, 0-based).
    -   `name_col`: Column name (string) or index (integer, 0-based) containing the folder name/ID to match.
    -   `header`: Header row index (integer, 0-based).
    -   `source_path`: Path to the source folder containing folders to be moved.
    -   `destination_path`: Path to the destination folder. Can be relative to the script's working directory.
    -   `filter_col` (Optional): Column name (string) or index (integer, 0-based) for filtering.
    -   `filter_value` (Optional): Value to filter by in `filter_col`.

    Example configuration (replace with your actual paths and details):

    ```python
    move_tasks = [
        {
            'excel_path': '/path/to/your/excel_file.xlsx',
            'sheet_name': 0,  # First sheet
            'name_col': 'PatientID', # Or column index like 0
            'header': 0,            # First row as header
            'source_path': '/path/to/your/source_folder',
            'destination_path': './processed_data',
            'filter_col': 'Diagnosis', # Or column index like 5
            'filter_value': 'UA'
        },
        # Add more tasks as needed
    ]
    ```

2.  **Run the script**: Execute the script from your terminal:

    ```bash
    python move_folders_by_excel.py
    ```

## Logging

Detailed logs, including skipped folders and errors, are saved to `move_file.log` in the same directory as the script. 