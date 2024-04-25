# Public library imports


# Personal library imports
from pandalibs.yaml_importer import get_configuration_data
from pandalibs.pprint_nosort import pp


# Module imports


# Initialized variables
config = get_configuration_data(up_a_level=False)


# Remainder of the code
pp(config)
