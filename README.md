# Active Directory Report
Simple script to gather the folling data. Exports to CSV for each, or single Excel file.
- User data
- Computer data
- Group data
- Recent account lockouts

## Requirements
- Python 3.11 (untested with 3.12 or above)
    - ldap3
    - openpyxl
    - pandalibs
    - pandas
    - wmi

## Known issues or limitations
- Last logon may not be accurate.
    - This is a known limitation of the Active Directory when targeting lastLogon. This attribute does not replicate across domain controllers.
        - LastLogonTimestamp is replicated, however it can update even if an actual login was not performed, so I did not target this.
        - At a later update I will attempt to query each domain controller in the domain and pick the latest one, however this attribute should still be considered a "best effort" piece of data.
- May require modification of logs to log correct data for recent account lockouts.
    - In my test domain, I have the following group policies configured.
        - Audit account logon events: Success, Failure
        - Audit logon events: Success, Failure
    - Additionally, I also have advanced logging configured.
        - Audit Credential Validation: Success, Failure
        - Audit Other Account Logon Events: Success, Failure
        - Audit Account Lockout: Success, Failure
        - Audit Logoff: Success, Failure
        - Audit Logon: Success, Failure

## Usage
1. Install poetry. 
    - Or  review pyproject.toml [tool.poetry.dependencies] and install them with your preferred method.
    - pip install poetry
2. Install base requirements. 
    - Or use your preferred virtual environment method.
    - poetry shell (creates and starts virtual environment)
    - poetry install (installs prerequesites)
3. Create config.yaml. See example in config_example.yaml
4. Run.
    - poetry run app.py
    - or whichever method you prefer