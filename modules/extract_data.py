"""This module contains various functions related to extracting and/or formatting data from Active Directory, as well as generating the external files."""

# Standard library imports
import re

# Public library imports
import pandas as pd

# Personal library imports
# Module imports
# Initialized variables
CN_PATTERN = r"CN=([^,]+)"
UAC_FLAG_MAP = {
    "Account_Disabled": 0x2,
    "Pwd_Not_Required": 0x20,
    "Pwd_Cant_Change": 0x40,
    "Pwd_Doesnt_Expire": 0x10000,
}
GROUP_FLAG_MAP = {
    "Security": 0x80000000,
    "Global": 0x00000002,
    "Domain": 0x00000004,
    "Universal": 0x00000008,
}


def extract_ad_data(input_data, the_attributes):
    """Extracts all requested data from Active Directory."""
    try:
        data = []
        for entry in input_data:
            row = {}
            for attr in the_attributes:
                try:
                    if attr == "userAccountControl":
                        row.update(extract_user_account_control(entry, UAC_FLAG_MAP))
                    elif attr == "groupType":
                        row["groupType"] = extract_group_type(entry[attr].values[0])
                    elif attr in ["pwdLastSet", "modifyTimeStamp", "accountExpires", "createTimeStamp", "lastLogon"]:
                        attr_name = attr if attr != "modifyTimeStamp" else "Date_Modified"
                        row[attr_name] = extract_timestamp(entry[attr].values[0])
                    elif attr in {"member", "memberOf"}:
                        row[attr] = extract_member(entry[attr].values)
                    elif attr == "msDS-parentdistname":
                        row["Parent_Location"] = entry[attr].values[0]
                    else:
                        row[attr] = entry[attr].values[0]
                except (AttributeError, IndexError):
                    row[attr] = ""
            data.append(row)
        return data
    except Exception:
        raise


def extract_group_type(group_type_value):
    """Helper function for groupType attribute. Read the bitmask and returns correct values."""
    try:
        group_types = []
        if not group_type_value & GROUP_FLAG_MAP["Security"]:
            group_types.append("Distribution")
        for key, value in GROUP_FLAG_MAP.items():
            if group_type_value & value:
                group_types.append(key)
        return "\n".join(group_types)
    except Exception:
        raise


def extract_user_account_control(entry, flags):
    """Helper function for user_account_control attribute. Splits the user account control flags into seperate columns and returns a dictionary."""
    try:
        col = {column_name: "N/A" for column_name in flags.keys()}
        if is_user_enabled(entry):
            col["Account_Disabled"] = "Yes"
        else:
            col["Account_Disabled"] = "No"
        if is_password_not_required(entry):
            col["Pwd_Not_Required"] = "Yes"
        else:
            col["Pwd_Not_Required"] = "No"
        if is_password_cant_change(entry):
            col["Pwd_Cant_Change"] = "Yes"
        else:
            col["Pwd_Cant_Change"] = "No"
        if is_password_doesnt_expire(entry):
            col["Pwd_Doesnt_Expire"] = "Yes"
        else:
            col["Pwd_Doesnt_Expire"] = "No"
        return col
    except Exception:
        raise


def extract_timestamp(value):
    """Helper function to convert datetime attributes into a date. Also converts 1601-01-01 and 999-12-31 into "Never" when they show up."""
    try:
        if str(value)[:10] in ["1601-01-01", "9999-12-31"]:
            return "Never"
        return str(value)[:10]
    except Exception:
        raise


def extract_member(value):
    """Helper function to convert member attribute into a list, or blank string if appropiate."""
    try:
        new_data = []
        for item in value:
            new_value = re.search(CN_PATTERN, item)
            if new_value:
                new_data.append(new_value.group(1))
        return new_data if new_data else ""
    except Exception:
        raise


def extract_and_format_ad_data(input_data, attributes):
    """Extracts desired attributes and creates a DataFrame."""
    try:
        data = extract_ad_data(input_data, attributes)
        return pd.DataFrame(data)
    except Exception:
        raise


def search_and_extract_data(base_dn, search_filter, attributes, the_connection):
    """Extracts required data."""
    try:
        the_connection.search(base_dn, search_filter, attributes=attributes)
        entries = the_connection.entries
        return extract_and_format_ad_data(entries, attributes)
    except Exception:
        raise


def is_user_enabled(entry):
    """Checks if the account is enabled."""
    try:
        uac = entry["userAccountControl"][0]
        return uac & 0x2
    except Exception:
        raise


def is_password_not_required(entry):
    """Checks if the account requires a password."""
    try:
        uac = entry["userAccountControl"][0]
        return uac & 0x20
    except Exception:
        raise


def is_password_cant_change(entry):
    """Checks if the account cannot change its password."""
    try:
        uac = entry["userAccountControl"][0]
        return uac & 0x40
    except Exception:
        raise


def is_password_doesnt_expire(entry):
    """Checks if the account password does not expire."""
    try:
        uac = entry["userAccountControl"][0]
        return uac & 0x10000
    except Exception:
        raise
