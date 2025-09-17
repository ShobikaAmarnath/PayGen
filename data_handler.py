# data_handler.py
import pandas as pd
from column_config import COLUMN_MAP

def load_employee_data(filepath):
    """Loads employee data from the specified Excel file."""
    try:
        df = pd.read_excel(filepath, header=3)
        # Clean up column names
        rename_dict = {}
        excel_columns_cleaned = {col.strip(): col for col in df.columns}

        for internal_name, possible_names in COLUMN_MAP.items():
            for excel_name in possible_names:
                if excel_name in excel_columns_cleaned:
                    rename_dict[excel_columns_cleaned[excel_name]] = internal_name
                    break # Found a match, move to the next internal name
        
        df.rename(columns=rename_dict, inplace=True)
        return df
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        return None