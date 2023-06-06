M365 User Details

Update (in bulk) your user's Microsoft 365 Azure AD Details using CSV files.
Also creates CSV reports that export existing details.

M365UserDetailsReport.ps1
Use the report tool to export your users' telephone numbers, etc properties.
![ReportRun](https://github.com/ITAutomator/M365UserDetails/assets/135157036/fd3094e1-aecc-40a4-82e7-968e503e66a0)

![ExcelReport](https://github.com/ITAutomator/M365UserDetails/assets/135157036/cd2078a0-71d3-449a-b41d-e3de71687c21)

Then create a M365UserDetailsUpdate.csv with the rows (users) and columns (properties) you want to update. 
If you want to update a property, enter the new value in the CSV.
If you want to leave a property as-is, just leave it blank in the CSV.
If you want to clear a property, use the keyword <clear> in the CSV (include the angle brackets in the text).
  
  
![ExcelUpdate](https://github.com/ITAutomator/M365UserDetails/assets/135157036/b9067404-3ea0-47e2-873f-0d6a85d33870)


M365UserDetailsUpdate.ps1
![UpdateRun](https://github.com/ITAutomator/M365UserDetails/assets/135157036/f7016c97-61ae-4100-91c3-7dbbd8760f6c)

  
  
