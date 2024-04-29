import pandas as pd
from openpyxl import load_workbook  # noqa: F401
from openpyxl.utils import get_column_letter  # noqa: F401
import re
import ast

cn_pattern = r"CN=([^,]+)"


def extract_ad_data(input_data, the_attributes):
    data = []
    for entry in input_data:
        row = {}
        for attr in the_attributes:
            try:
                if attr == "userAccountControl":
                    # Column names and flag values
                    flag_map = {
                        "Account_Disabled": 0x2,
                        "Pwd_Not_Required": 0x20,
                        "Pwd_Cant_Change": 0x40,
                        "Pwd_Doesnt_Expire": 0x10000,
                    }
                    # Set default "N/A" for all flag columns
                    for column_name in flag_map.keys():
                        row[column_name] = "N/A"
                    # Check for set flags
                    for column_name, flag_value in flag_map.items():
                        if column_name == "Account_Disabled":
                            row[column_name] = "Yes" if is_user_enabled(entry) else "No"
                        elif column_name == "Pwd_Not_Required":
                            row[column_name] = "Yes" if is_password_not_required(entry) else "No"
                        elif column_name == "Pwd_Cant_Change":
                            row[column_name] = "Yes" if is_password_cant_change(entry) else "No"
                        elif column_name == "Pwd_Doesnt_Expire":
                            row[column_name] = "Yes" if is_password_doesnt_expire(entry) else "No"
                elif attr == "pwdLastSet" or attr == "modifyTimeStamp" or attr == "accountExpires" or attr == "createTimeStamp":
                    if attr == "modifyTimeStamp":
                        attr_name = "Date_Modified"
                    else:
                        attr_name = attr
                    value = str(entry[attr].values[0])[:10]
                    if value == "1601-01-01" or value == "9999-12-31":
                        value = "Never"
                elif attr == "member":  # Attributes that contain multiple items must be defined here.
                    attr_name = attr
                    old_value = entry[attr].values
                    new_data = []
                    for item in old_value:
                        new_value = re.search(cn_pattern, item)
                        if new_value:
                            new_data.append(new_value.group(1))
                    if not new_data:
                        value = ""
                    else:
                        value = new_data
                elif attr == "msDS-parentdistname":
                    attr_name = "Parent_Location"
                    value = entry[attr].values[0]
                else:
                    attr_name = attr
                    value = entry[attr].values[0]
                row[attr_name] = value
            except (AttributeError, IndexError):  # Handle both missing attributes and empty value lists
                row[attr_name] = ""
        data.append(row)
    return data


def extract_and_format_ad_data(input_data, attributes):
    """Extracts desired attributes and creates a DataFrame."""
    data = extract_ad_data(input_data, attributes)  # Reuse your existing function
    return pd.DataFrame(data)


def search_and_extract_data(base_dn, search_filter, attributes, the_connection):
    """Extracts required data."""
    the_connection.search(base_dn, search_filter, attributes=attributes)
    entries = the_connection.entries
    return extract_and_format_ad_data(entries, attributes)


def is_user_enabled(entry):
    uac = entry["userAccountControl"][0]  # Get the userAccountControl value
    return uac & 0x2  #  Check if the flag is NOT set


def is_password_not_required(entry):
    uac = entry["userAccountControl"][0]
    return uac & 0x20


def is_password_cant_change(entry):
    uac = entry["userAccountControl"][0]
    return uac & 0x40


def is_password_doesnt_expire(entry):
    uac = entry["userAccountControl"][0]
    return uac & 0x10000


def create_csv(input_data, filename):
    input_data.to_csv(filename, index=False)


def create_xlsx(input_data, filename):
    excel_writer = pd.ExcelWriter(filename)
    for item in input_data:
        input_data[item] = input_data[item].reset_index(drop=True)
        input_data[item].to_excel(excel_writer, sheet_name=item, index=False)

    # Auto-size columns
    for sheet_name in excel_writer.sheets:
        worksheet = excel_writer.sheets[sheet_name]
        # Auto-size columns
        for col_num, col in enumerate(worksheet.iter_cols()):
            max_col_width = 0
            for cell in col:
                max_col_width = max(max_col_width, len(str(cell.value)))
            max_col_width = min(max_col_width + 5, 30)  # Cap at 30 characters (+2 padding)
            worksheet.column_dimensions[get_column_letter(col_num + 1)].width = max_col_width
        # Apply text wrapping
        for row in worksheet.iter_rows(min_row=2):  # Start from the second row (header is row 1)
            for cell in row:
                new_alignment = cell.alignment.copy()  # Create a copy
                new_alignment.wrap_text = True  # Modify wrap_text on the copy
                cell.alignment = new_alignment

                # Newline insertion (conditional)
                # print(cell.value, type(cell.value))
                if cell.value.startswith("[") and cell.value.endswith("]"):
                    print("Cell Value:", cell.value, "Type:", type(cell.value))
                    list_version = ast.literal_eval(cell.value)  # Convert to list
                    cell.value = "\n".join(list_version)  # Convert back to string with newlines

    excel_writer._save()


def create_all_reports(extracted_data, config):
    if config["export_type"] in ["CSV", "BOTH"]:
        for item in extracted_data:
            create_csv(extracted_data[item], f"{item}_report.csv")

    if config["export_type"] in ["XLSX", "BOTH"]:
        create_xlsx(extracted_data, "ad_report.xlsx")
