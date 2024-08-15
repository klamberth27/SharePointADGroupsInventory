# SharePoint and Azure AD Group Inventory Script

This PowerShell script retrieves SharePoint Online group details and Azure AD group information, processes the data, and exports the results to CSV files.

## Prerequisites

- **PowerShell**: Ensure you have PowerShell installed.
- **SharePoint Online Management Shell**: The script uses the `Microsoft.Online.Sharepoint.PowerShell` module. The script will automatically install it if it's not already installed.
- **Azure AD Module**: The script uses the `AzureAD` module. The script will automatically install it if it's not already installed.
- **Permissions**: The script requires an account with SharePoint Online Admin and Azure AD Admin roles for authentication.

## Instructions

### Edit the Script

1. Open the script file `SharePointADGroupsInventory.ps1`.
2. Locate the following line:
   ```powershell
   $AdminURL = "https://m365x84490777-admin.sharepoint.com/"
   ```
3. Replace `"https://m365x84490777-admin.sharepoint.com/"` with your own SharePoint tenant admin URL. It should look something like this:
   ```powershell
   $AdminURL = "https://your-tenant-admin.sharepoint.com/"
   ```

### Prepare the CSV File

1. Ensure you have a `Sites.csv` file in the same directory as the script. The CSV file should have the following structure:
   ```plaintext
   SiteURL
   https://your-tenant.sharepoint.com/sites/site1
   https://your-tenant.sharepoint.com/sites/site2
   ```
2. Sample file uploaded
   
### Run the Script

1. Open PowerShell as an administrator.
2. Navigate to the directory containing the script.
3. Run the script:
   ```powershell
   .\SharePointADGroupsInventory.ps1
   ```
4. Authenticate using an account with the necessary permissions when prompted.

The script will connect to your SharePoint tenant, retrieve group details, check for Azure AD group GUIDs, and export the results to CSV files.

### Check the Output

- The results will be saved in two CSV files:
  - `SharePoint_AD_Groups_report.csv`: Contains details about all SharePoint AD groups.
  - `AD_Groups_Details.csv`: Contains detailed information about the associated Azure AD groups.

## Troubleshooting

- **No CSV file to read**: Ensure the `Sites.csv` file is present and correctly formatted.
- **Permission Issues**: Ensure you have the necessary permissions to connect to the SharePoint tenant and Azure AD.
- **Module Installation**: If the script fails to install the required modules, try manually installing them:
  ```powershell
  Install-Module -Name Microsoft.Online.Sharepoint.PowerShell -Force
  Install-Module -Name AzureAD -Force
  ```

## Additional Notes

- This script is designed to be run in an environment with access to the SharePoint Online Management Shell and Azure AD module.
- The script handles errors and provides prompts for user authentication when connecting to SharePoint and Azure AD.

---

This `README.md` provides clear instructions for users on how to configure, run, and troubleshoot the `SharePointADGroupsInventory.ps1` script.
