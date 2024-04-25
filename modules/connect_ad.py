# Public library imports
import ldap3 as ad

# Personal library imports


# Module imports

# Initialized variables

# Connect to active directory


def connect_to_ad(server, domain, username, password):
    server = ad.Server(f"{server}.{domain}", get_info=ad.ALL)
    conn = ad.Connection(server, user=(f"{username}@{domain}"), password=password, auto_bind=True)
    return conn
