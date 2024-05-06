"""This module is purely to connect to Active Directory."""

# Public library imports
import ldap3 as ad

# Personal library imports


# Module imports
# Initialized variables


# Connect to active directory
def connect_to_ad(server, domain, username, password):
    try:
        """Connects to Active Directory. Requires server name, domain name, username, and password."""
        server = ad.Server(f"{server}.{domain}", get_info=ad.ALL)
        conn = ad.Connection(server, user=(f"{username}@{domain}"), password=password, auto_bind=True)
        return conn
    except Exception:
        raise
