"""This application extracts data from Active Directory and exports to CSV and/or Excel files.
It uses the Pandalibs library for YAML configuration and printing dictionarys in a readable manner when necessary."""

# Standard library imports
import os

# Public library imports
# Personal library imports
from pandalibs.pprint_nosort import pp
from pandalibs.yaml_importer import get_configuration_data

# Module imports
from modules.connect_ad import connect_to_ad
from modules.create_files import create_all_reports
from modules.event_logs import user_lockout_events
from modules.extract_data import search_and_extract_data

# Initialized variables
if not os.path.exists("config"):
    config = get_configuration_data(up_a_level=False, subfolder="")
else:
    config = get_configuration_data(up_a_level=False)
SEARCH_FILTER_USERS = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=*))"
SEARCH_FILTER_COMPS = "(&(objectCategory=computer)(objectClass=user)(sAMAccountName=*))"
SEARCH_FILTER_GROUPS = "(&(objectCategory=group)(objectClass=group)(sAMAccountName=*))"
extracted_data = {}


# Print out the current configuraion if debugging is set.
if config["DEBUG"]:
    print("\nCurrent configuration.\n")
    pp(config)
    print("\nEnd current configuration.\n")

# Connect to Active Directory
connection = connect_to_ad(
    config["server"],
    config["domain"],
    config["username"],
    config["password"],
)

# Extract data
extracted_data["user_data"] = search_and_extract_data(
    config["base_dn"],
    SEARCH_FILTER_USERS,
    config["user_attributes"],
    connection,
)
extracted_data["comp_data"] = search_and_extract_data(
    config["base_dn"],
    SEARCH_FILTER_COMPS,
    config["comp_attributes"],
    connection,
)
extracted_data["group_data"] = search_and_extract_data(
    config["base_dn"],
    SEARCH_FILTER_GROUPS,
    config["group_attributes"],
    connection,
)
extracted_data["lockout_data"] = user_lockout_events(
    domain=config["domain"],
    server=config["server"],
    username=config["username"],
    password=config["password"],
    query=config["lockouts_query"],
)

# Export data to files.
create_all_reports(extracted_data, config)
