# Test-SPOFileActionsAsServicePrincipal

To ensure proper permission configuration and demonstrate full functionality for uploading files to SharePoint Online and Download files from SharePoint Online, multiple scripts have been developed.
Other file types besides text files are working with this script. The script has been successfully tested with .xlsx files, for example.

##  Prerequisites

The following components and information are required for script execution and to ensure full functionality:

- PowerShell 7: Necessary for utilizing specific parameters within the scripts.
- App Information (TenantID, AppID, ClientSecret): These enable authentication and confirm correct execution with the intended identity.
- Site URL: Required for determining the SiteID.
- Access to the Site or the Global Reader role: Necessary for retrieving the SiteID.
- Document Library Name: This is needed to identify the correct DriveID from the retrieved drives."

##  Description

### Get SiteId

There are two easy methods to determine the SiteID:

Method 1 (for site members):

Simply append _api/site/id to the site's URL. The displayed Edm.Guid is the required SiteID.

![image](https://github.com/user-attachments/assets/16358fec-0197-4eb6-85ba-67d1e2da21d2)

Method 2 (for users with Global Reader rights):

Search for the desired site in the SharePoint Admin Portal. The SiteID can be found at the end of the URL for that site after when editing Site Settings.

![image](https://github.com/user-attachments/assets/9c1043d3-ef58-4883-bf79-8d03dcf13df5)


### Adjusting configuration file

Execution of the first script requires the configuration section 'ServicePrincipal' to be populated with the appropriate values and the 'SiteId' entry to be adjusted. The 'DriveId' will be determined in the next step.

### Get driveId
The script 01_Get-SPOSiteDriveIDs.ps1 outputs all drives and their corresponding IDs for the site specified in the configuration file. The DriveID of the desired drive should be entered into the configuration file.

Example output
![image](https://github.com/user-attachments/assets/c7f3642f-d6b7-4639-aad4-e6382ce2e493)



### Uploading a File

The script 02_Test-SPOFileUploadAsServicePrincipal.ps1 uploads the file test_upload.txt from the files subfolder to SharePoint Online using the configuration file. 
The file will be displayed in SharePoint Online as test.txt.

Example output
![image](https://github.com/user-attachments/assets/1924f43d-d1ad-4400-9544-8ea92cc90205)


### Downloading a File

The script 03_Test-SPOFileDownloadAsServicePrincipal.ps1 downloads the file test.txt from SharePoint Online into the files subfolder. 
The downloaded file will be saved in the file system as test_download.txt to avoid confusion with files of the same name.

Example output
![image](https://github.com/user-attachments/assets/8e7a62fc-1381-43a9-b2c8-7725ca7d139d)


### Downloading a File

99_Test-SingleCall.ps1 can be used to make single API calls. When troubleshooting a 3rd party app, the call can be read out via the Microsoft Graph Activity logs and than executed with this script to check the result.
