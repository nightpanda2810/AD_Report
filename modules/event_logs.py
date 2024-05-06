"""This module contains various functions related to pulling event log details."""

# Standard library imports
from datetime import datetime
import pytz

# Public library imports
import pandas as pd
import wmi

# Personal library imports

# Module imports
# Initialized variables


def user_lockout_events(domain, server, username, password, query):
    """Extract user lockout events from logs."""
    event_conn = wmi.WMI(computer=f"{server}.{domain}", user=username, password=password)
    events = event_conn.query(query)
    lockout_events = []
    event_number = 0
    for event in events:
        event_number += 1
        dt_object = datetime.strptime(event.TimeGenerated[:-11], "%Y%m%d%H%M%S")
        original_timezone = pytz.utc
        target_timezone = pytz.timezone("US/Eastern")
        dt_object_in_target_tz = original_timezone.localize(dt_object).astimezone(target_timezone)
        date_utc = dt_object.strftime("%Y-%m-%d")
        time_utc = dt_object.strftime("%H:%M:%S %Z")
        date_target = dt_object_in_target_tz.strftime("%Y-%m-%d")
        time_target = dt_object_in_target_tz.strftime("%H:%M:%S %Z")
        lockout_events.append(
            {
                "User": event.InsertionStrings[0],
                "Event ID": event.EventCode,
                "Date UTC": date_utc,
                "Time UTC": time_utc,
                f"Date {target_timezone}": date_target,
                f"Time {target_timezone}": time_target,
            }
        )
    return pd.DataFrame(lockout_events)
