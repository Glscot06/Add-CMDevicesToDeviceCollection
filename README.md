# Add-CMDevicesToDeviceCollection
## Overview
The Add-ToDeviceCollection PowerShell function allows administrators to efficiently add devices to an SCCM (System Center Configuration Manager) device collection. This script supports various input methods, such as individual computer names, CSV files, comma-separated lists, and queries. It also has error-handling mechanisms and can create device collections if they don't exist when using the -Force parameter.

### Functionality
*Adding a Single Computer*
<br>
Add-ToDeviceCollection -SCCMServer "MySCCMServer" -SiteCode "ABC" -ComputerName "PC01" -CollectionName "MyCollection"

*Adding Multiple Computers from a Text File*
<br>
Add-ToDeviceCollection -SCCMServer "MySCCMServer" -SiteCode "ABC" -TextFilePath "C:\Computers.txt" -CollectionName "MyCollection"

*Adding Multiple Computers from a CSV File*
<br>
Add-ToDeviceCollection -SCCMServer "MySCCMServer" -SiteCode "ABC" -CSVFilePath "C:\Computers.csv" -CSVColumnName "ComputerName" -CollectionName "MyCollection"

*Adding Multiple Computers from a Comma-Separated List*
<br>
Add-ToDeviceCollection -SCCMServer "MySCCMServer" -SiteCode "ABC" -CommaSeparatedList "PC01,PC02,PC03" -CollectionName "MyCollection"

*Adding a Query-Based Membership Rule*
<br>
Add-ToDeviceCollection -SCCMServer "MySCCMServer" -SiteCode "ABC" -Query "SELECT * FROM SMS_R_System WHERE SMS_R_System.Name LIKE 'PC%'" -QueryName "PCs" -CollectionName "MyCollection"

*Creating a Collection if it Doesn't Exist*
<br>
Add-ToDeviceCollection -SCCMServer "MySCCMServer" -SiteCode "ABC" -ComputerName "PC01" -CollectionName "MyCollection" -Force -LimitingCollection "All Systems" -RefreshType "Both"


### Error Handling
The script includes comprehensive error-handling to ensure that missing or incorrect parameters are flagged.
If a device collection is not found and the -Force parameter is not used, the script will not proceed and will prompt the user to specify -Force if creation of the collection is desired.
Contributing
Contributions are welcome! Please submit issues or pull requests on GitHub.
