<#
.SYNOPSIS
This script retrieves SharePoint Online group details and Azure AD group information, processes the data, and exports the results to CSV files.

.DESCRIPTION
The script first ensures that the necessary PowerShell modules (`Microsoft.Online.SharePoint.PowerShell` and `AzureAD`) are installed and imported. It connects to SharePoint Online using admin credentials, reads site URLs from a CSV file, and collects group details for each site. The script identifies unique Azure AD group GUIDs from the SharePoint groups and, if any are found, connects to Azure AD to retrieve additional details about these groups. The collected data is then exported to separate CSV files for SharePoint groups and Azure AD groups. Finally, the script disconnects from both SharePoint Online and Azure AD.

The script automates the process of inventorying AD groups in SharePoint sites, ensuring that group data is consolidated and exported efficiently.

.NOTES
- The script requires an account with SharePoint Online Admin and Azure AD Admin roles for authentication.
- The `Sites.csv` input file should be placed in the same directory as the script.
- The output CSV files will be saved in the same directory as the scrip.
- Ensure that PowerShell execution policies allow the script to run if you encounter any issues.

.AUTHOR
SubjectData

.EXAMPLE
.\SharePointADGroupsInventory.ps1
This will run the script in the current directory, processing the 'Sites.csv' file, connecting to both SharePoint Online and Azure AD, and generating 'SharePoint_AD_Groups_report.csv' and 'AD_Groups_Details.csv' with the group details.
#>

$MicrosoftOnline = "Microsoft.Online.Sharepoint.PowerShell"

# Check if the module is already installed
if (-not(Get-Module -Name $MicrosoftOnline)) {
    # Install the module
    Install-Module -Name $MicrosoftOnline -Force
}

Import-Module $MicrosoftOnline -Force

$myDir = Split-Path -Parent $MyInvocation.MyCommand.Path

$XLloc = "$myDir\"
$ReportsPath = "$myDir\"

# SharePoint section
$AdminURL = "https://m365x84490777-admin.sharepoint.com/"
$CSVPath = $XLloc + "Sites.csv"

Connect-SPOService -Url $AdminURL 

$GroupsData = @()

Function Get-ADGroups() {
    param(
        $AllGroups = $(throw "No Group parameter value"),
        $CurrentURL = $(throw "No Site URL parameter value")
    )

    $localGroupsData = @()
    try {
        $i = 0
        foreach ($Group in $AllGroups) {
            if ($AllGroups[$i].Users.Count -gt '0') {
                for ($user = 0; $user -lt $AllGroups[$i].Users.Count; $user++) {
                    $GUID = $AllGroups[$i].Users[$user]
                    if ($GUID -match "^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$") {
                        $details = @{
                            SharePointSiteUrl = $CurrentURL
                            SiteDisplayName = $AllGroups[$i].Title
                            GroupGUID = $GUID
                            Roles = $AllGroups[$i].Roles -join ", "
                        }
                        $localGroupsData += New-Object PSObject -Property $details   
                    }
                }
            }
            $i++
        }

        $RootADGroups = Get-SPOUser -Site $sitesImport[$siteCount].SiteURL | Where-Object { $_.LoginName -match "^(\{){0,1}[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}(\}){0,1}$" }
        for ($ADGroup = 0; $ADGroup -lt $RootADGroups.Count; $ADGroup++) {
            $details = @{
                SharePointSiteUrl = $CurrentURL
                SiteDisplayName = $RootADGroups[$ADGroup].DisplayName
                GroupGUID = $RootADGroups[$ADGroup].LoginName
                Roles = $RootADGroups[$ADGroup].Roles -join ", "                     
            }
            $localGroupsData += New-Object PSObject -Property $details   
        }
    } catch {
        Write-Host "Error processing site - $CurrentURL" 
    }
    
    return $localGroupsData # Return the accumulated data
}

Write-Host "Connected" -ForegroundColor Green

$bool = $false
$sitesImport = Import-Csv $CSVPath

if ($sitesImport.Count -gt '0' -or $sitesImport.SiteURL.ToString() -ne "") {
    for ($siteCount = 0; $siteCount -lt $sitesImport.Count; $siteCount++) {
        Write-Host $sitesImport[$siteCount].SiteURL
        $Groups = Get-SPOSiteGroup -Site $sitesImport[$siteCount].SiteURL
        $GroupsData += Get-ADGroups -AllGroups $Groups -CurrentURL $sitesImport[$siteCount].SiteURL
        $bool = $true
    }

    if (!($bool)) {
        Write-Host $sitesImport.SiteURL
        $Groups = Get-SPOSiteGroup -Site $sitesImport.SiteURL
        $GroupsData += Get-ADGroups -AllGroups $Groups -CurrentURL $sitesImport.SiteURL
    }
}

# Remove duplicates based on GroupGUID
$UniqueGroupsData = $GroupsData | Select-Object GroupGUID -Unique

# Proceed with Azure AD processing only if there are unique GroupGUIDs
if ($UniqueGroupsData.Count -gt 0) {

    Write-Host "AD groups found. Connecting to Azure AD to retrieve group details..." -ForegroundColor Cyan

    $AzureADModule = "AzureAD"

    # Check if the module is already installed
    if (-not(Get-Module -Name $AzureADModule -ListAvailable)) {
        # Install the module
        Install-Module -Name $AzureADModule -Force
    }

    # Import the module
    Import-Module $AzureADModule -Force


    # Connect to Azure AD
    $null = Connect-AzureAD

    # Define an array to hold the group details
    $GroupDetails = @()

    $AllGroupTypes = Get-AzureADMSGroup -All $true | Select-Object Id, DisplayName, GroupTypes, MailEnabled, SecurityEnabled

    # Loop through each unique GroupGUID and find the group details
    foreach ($GroupGuid in $UniqueGroupsData) {
        try {
            $group = Get-AzureADGroup -ObjectId $GroupGuid.GroupGUID
            $objGroupType = $AllGroupTypes | Where-Object { $_.Id -eq $GroupGuid.GroupGUID }

            if ($group) {
                try {
                    $owners = Get-AzureADGroupOwner -ObjectId $group.ObjectId | Select-Object -ExpandProperty UserPrincipalName -ErrorAction SilentlyContinue
                    $members = Get-AzureADGroupMember -ObjectId $group.ObjectId | Select-Object -ExpandProperty UserPrincipalName -ErrorAction SilentlyContinue
                } catch {}

                # Determine the group type
                if ($objGroupType.GroupTypes -contains "Unified" -and $objGroupType.MailEnabled -ne $true) {
                    $groupType = "Mail-enabled Microsoft 365 Group"
                } elseif ($objGroupType.GroupTypes -contains "Unified" -and $objGroupType.MailEnabled -ne $false) {
                    $groupType = "Microsoft 365 Group"
                } elseif ($objGroupType.SecurityEnabled -eq $true -and $objGroupType.MailEnabled -eq $true) {
                    $groupType = "Mail-enabled Security Group"
                } elseif ($objGroupType.SecurityEnabled -eq $true -and $objGroupType.MailEnabled -eq $false) {
                    $groupType = "Security Group"
                } elseif ($objGroupType.MailEnabled -eq $true -and $objGroupType.SecurityEnabled -eq $false) {
                    $groupType = "Distribution List"
                } else {
                    $groupType = "Unknown Group Type"
                }

                # Prepare the group details
                $groupDetail = @{
                    GroupGUID    = $GroupGuid.GroupGUID
                    GroupName    = $group.DisplayName
                    GroupType    = $groupType
                    OwnersCount  = $owners.Count
                    MembersCount = $members.Count
                    Owners       = $owners -join "; "
                    Members      = $members -join "; "
                }

                # Add to the GroupDetails array
                $GroupDetails += New-Object PSObject -Property $groupDetail
            } else {
                Write-Host "Group with GUID $($GroupGuid.GroupGUID) not found." -ForegroundColor Yellow
            }
        } catch {
            continue
        }
    }

    # Export Azure AD group details
    $GroupDetails | Export-Csv -Path "$($XLloc)\AD_Groups_Details.csv" -NoTypeInformation

    # Disconnect from Azure AD
    Disconnect-AzureAD
}

# Export SharePoint group data
$GroupsData | Export-Csv -Path "$($XLloc)\SharePoint_AD_Groups_report.csv" -NoTypeInformation

Disconnect-SPOService

