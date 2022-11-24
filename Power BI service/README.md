Manage Power BI service (the cloud thing at http://app.powerbi.com) using PowerShell.

All activities except authentication are performed via [PBI REST API](https://learn.microsoft.com/en-us/rest/api/power-bi/). Authentication is done by MicrosoftPowerBIMgmt PS Module.

# Get Power BI Service statistics

## Use case

In a large corporate environment, you have a company-wide PBI service subscription. Users work with it. The service ends up with thousands of workspaces, reports and datasets. Then the CIO asks you, the PBI subject matter expert: the PBI environment has become messy, can you clean it up?

Ok, you think, the approach should be the following:
- Get all reports uploaded to PBI service
- Check when the uploaded reports have been created and who uses them
- Then send a mail to the users asking "do you still need this report?"
- Then work on the deletion

## Solution

PowerShell to the rescue.

[This script](./power%20bi%20service%20-%20get%20workspaces%20users%20reports%20datasets.ps1) helps you perform the first step: collect the statistics about your PBI service:
- Workspaces
- Reports
- Datasets
- Users

The results are then exported to Excel.

![how the Excel looks](https://i.imgur.com/bRz1oDL.png)

## To be aware of

This script does not use [Power BI Admin APIs](https://learn.microsoft.com/en-us/rest/api/power-bi/admin/reports-get-report-users-as-admin) via [PBI service principal](https://powerbi.microsoft.com/en-us/blog/use-power-bi-api-with-service-principal-preview/). Admin APIs would be perfect for this use case of course. However, the reality is that the PBI service in the company I work for is managed on corporate level by the global IT. These folks are located in a different country than the subsidiary that I am in. It is difficult to find the right contact person at "the Global". They are very busy. Or they have no clue. So I work with whatever is available - that is, without Admin APIs.

Moreover, in my case, users can be assigned not via workspaces, but via direct access to reports.

![direct access to Power BI report](https://i.imgur.com/iF4z2A3.png)

This dictates the way to get the report users:
1. As the first step, get workspace users including workspace admins. These users are automatically added to every dataset in this workspace.
2. Next, we get the dataset users. These users will include both the workspace admins from the point above, and the users who have been granted direct access to the report.
3. Now that you have two sets of users, subtract admin users from dataset users - and you will get the users with direct access. 👉 These are going to be the folks to whom you are going to send the mail.

