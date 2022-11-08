Manage Power BI service (the cloud thing at http://app.powerbi.com) using PowerShell.

All activities except authentication are performed via [PBI REST API](https://learn.microsoft.com/en-us/rest/api/power-bi/). Authentication is done by MicrosoftPowerBIMgmt PS Module.

As a start, we collect some statistics:
- Workspaces
- Users
- Reports
- Datasets
- Orphanated datasets (not used by any report)

As the next steps:
- Users (add, remove, change roles)
- Visuals (create programmatically, format, rename columns etc.)
- Datasets (upload, refresh)