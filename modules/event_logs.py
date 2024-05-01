import wmi
from datetime import datetime
import pytz

domain_controller = "panda-dc-001.dc.wesslercomputers.com"  # Replace with the actual domain controller
username = "administrator"
password = "$zeddHARRYperrin27!"

connection = wmi.WMI(computer=domain_controller, user=username, password=password)
wql_query = "SELECT * FROM Win32_NTLogEvent WHERE Logfile='Security' AND EventCode='4740'"
events = connection.query(wql_query)


for event in events:
    dt_object = datetime.strptime(event.TimeGenerated[:-11], "%Y%m%d%H%M%S")
    original_timezone = pytz.utc
    target_timezone = pytz.timezone("US/Eastern")
    dt_object_in_target_tz = original_timezone.localize(dt_object).astimezone(target_timezone)
    readable_format = dt_object_in_target_tz.strftime("%Y-%m-%d %H:%M:%S %Z")
    print(
        "Target User:",
        (event.InsertionStrings[0]),
        "Event ID:",
        event.EventCode,
        "Time Generated UTC:",
        dt_object,
        f"Time Generated {target_timezone}:",
        readable_format,
    )


def user_lockout_events(domain, server, username, password, query):
    event_conn = wmi.WMI(computer=f"{server}.{domain}", user=username, password=password)
    events = event_conn.query(query)
    for event in events:
        pass
