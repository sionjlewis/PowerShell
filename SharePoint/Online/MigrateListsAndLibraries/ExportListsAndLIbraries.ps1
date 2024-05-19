# =====================================================================================================
# SharePoint Online: Export Site Columns, Content Types, Listas and Libraries to CSV files.
# -----------------------------------------------------------------------------------------------------
# Created By:    Siôn Lewis (www.sjlewis.com)
# Created Date:  18/05/2024
# Modified By:   Siôn Lewis (www.sjlewis.com)
# Modified Date: 18/05/2024
# -----------------------------------------------------------------------------------------------------
# Prerequisites: This script uses PnP Management Shell, see the link for configuration instructions:
# https://learn.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-pnppowershell
# -----------------------------------------------------------------------------------------------------
# Instructions:  Update the variables at the top of the script before running within your environment.
# =====================================================================================================


# -----------------------------------------------------------------------------------------------------
# Update these variables.
# -----------------------------------------------------------------------------------------------------
[string]$orgName            = "sourcedemo";                                 # Tenant name or sub-domain
[string]$sitePath           = "/sites/Demo";                                # Example: "/sites/Teamwork"
[string]$projectRootPath    = "C:\PowerShell\SharePoint\Online\SPOAssets";  # Path to the scripts folder.
[double]$slowSPCalls 		= 0;                                            # Default to 0, unless the server is thowing remote server errors.
[string[]]$arrFieldGroups   = @(
    "Demo Site Columns",
    "Another Group of Site Columns"
);											                                # Restrict the columns retrived by their groups.
[string[]]$arrCTGroupNames  = @(
    "Demo Content Types",
    "Another Group of Content Types"
);											                                # Restrict the content types retrived by their groups.
[string[]]$arrList          = @(
    "Demo List",
    "Another List"
);											                                # Add the name of Lists and Libraries to export.
[bool]$openInEdge = $true;                                                  # Set to false to use Google Chrome.
# ------------------------------------------------------------------------------------------------------
# Use these properties to manage existing Content Types and
# set them both to false to export the Content Types' details.
[bool]$manageCTName = $false;
[bool]$reviewCTName = $false;
# -----------------------------------------------------------------------------------------------------
# Leave these variables with default values.
# -----------------------------------------------------------------------------------------------------
[string]$csvSiteColumnsFile = (".\Assets\{0}-CSV01-SiteColumns.csv" -f $orgName.ToUpper());
[string]$csvContentTypesFile = (".\Assets\{0}-CSV02-ContentTypes.csv" -f $orgName.ToUpper());
[string]$csvContentTypesFieldsFile = (".\Assets\{0}-CSV03-ContentTypesFields.csv" -f $orgName.ToUpper());
[string]$siteTemplatePath = (".\Assets\{0}-PnPSiteTemplate.xml" -f $orgName.ToUpper());
# -----------------------------------------------------------------------------------------------------


Set-Location $projectRootPath;

Write-Host "Export script starting...`n" -ForegroundColor White;

.\SPO-GET-SiteColumns.ps1 -OrgName $orgName -SitePath $sitePath -ProjectRootPath $projectRootPath -SlowSPCalls $slowSPCalls -ArrFieldGroups $arrFieldGroups -CsvSiteColumnsFile $csvSiteColumnsFile;
.\SPO-GET-ContentTypes.ps1 -OrgName $orgName -SitePath $sitePath -ProjectRootPath $projectRootPath -SlowSPCalls $slowSPCalls -ArrCTGroupNames $arrCTGroupNames -OpenInEdge $openInEdge -ManageCTName $manageCTName -ReviewCTName $reviewCTName -CsvContentTypesFile $csvContentTypesFile -CsvContentTypesFieldsFile $csvContentTypesFieldsFile; 
.\SPO-GET-SiteTemplate.ps1 -OrgName $orgName -SitePath $sitePath -ProjectRootPath $projectRootPath -ArrList $arrList -SiteTemplatePath $siteTemplatePath;

Write-Host "Export script complete`n" -ForegroundColor Green;