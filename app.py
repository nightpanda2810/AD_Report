# Public library imports

# Personal library imports
from pandalibs.yaml_importer import get_configuration_data
from pandalibs.pprint_nosort import pp


# Module imports
from modules.extract_data import search_and_extract_data, create_all_reports
from modules.connect_ad import connect_to_ad

# Initialized variables
config = get_configuration_data(up_a_level=False)
search_filter_users = "(&(objectCategory=person)(objectClass=user)(sAMAccountName=*))"  # Find all users
search_filter_comps = "(&(objectCategory=computer)(objectClass=user)(sAMAccountName=*))"  # Find all computers
search_filter_grups = "(&(objectCategory=group)(sAMAccountName=*))"  # Find all groups


# Remainder of the code

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
extracted_data = {}
extracted_data["user_data"] = search_and_extract_data(config["base_dn"], search_filter_users, config["required_attributes"], connection)
extracted_data["comp_data"] = search_and_extract_data(config["base_dn"], search_filter_comps, config["required_attributes"], connection)
extracted_data["grup_data"] = search_and_extract_data(config["base_dn"], search_filter_grups, config["group_attributes"], connection)

# Export data to files.
create_all_reports(extracted_data, config)
