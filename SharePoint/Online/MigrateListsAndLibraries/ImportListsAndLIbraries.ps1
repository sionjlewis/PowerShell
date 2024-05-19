# =====================================================================================================
# SharePoint Online: Import Site Columns, Content Types, Listas and Libraries to CSV files.
# -----------------------------------------------------------------------------------------------------
# Created By:    Siôn Lewis (www.sjlewis.com)
# Modified Date: 18/05/2024
# Modified By:   Siôn Lewis (www.sjlewis.com)
# Modified Date: 19/05/2024
# -----------------------------------------------------------------------------------------------------
# Prerequisites: This script uses PnP Management Shell, see the link for configuration instructions:
# https://learn.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-pnppowershell
# -----------------------------------------------------------------------------------------------------
# Instructions:  Update the variables at the top of the script before running within your environment.
# =====================================================================================================


# -----------------------------------------------------------------------------------------------------
# Update these variables.
# -----------------------------------------------------------------------------------------------------
[string]$orgName            = "destination";                                # Tenant name or sub-domain
[string]$sitePath           = "/sites/Demo";                                # Example: "/sites/Teamwork"
[string]$projectRootPath    = "C:\PowerShell\SharePoint\Online\SPOAssets";  # Path to the scripts folder.
[double]$slowSPCalls 		= 0;                                            # Default to 0, unless the server is thowing remote server errors.
[string]$sourceName         = "sourcedemo";									# Source tenant name or sub-domain.
# -----------------------------------------------------------------------------------------------------
# Leave these variables with default values.
# -----------------------------------------------------------------------------------------------------
[string]$csvSiteColumnsFile = (".\Assets\{0}-CSV01-SiteColumns.csv" -f $sourceName.ToUpper());
[string]$csvContentTypesFile = (".\Assets\{0}-CSV02-ContentTypes.csv" -f $sourceName.ToUpper());
[string]$csvContentTypesFieldsFile = (".\Assets\{0}-CSV03-ContentTypesFields.csv" -f $sourceName.ToUpper());
[string]$siteTemplatePath = (".\Assets\{0}-PnPSiteTemplate.xml" -f $sourceName.ToUpper());
# -----------------------------------------------------------------------------------------------------


Set-Location $projectRootPath;

Write-Host "Import script starting...`n" -ForegroundColor White;

.\SPO-SET-SiteColumns.ps1 -OrgName $orgName -SitePath $sitePath -ProjectRootPath $projectRootPath -SlowSPCalls $slowSPCalls -SourceName $sourceName -CsvSiteColumnsFile $csvSiteColumnsFile;
.\SPO-SET-ContentTypes.ps1 -OrgName $orgName -SitePath $sitePath -ProjectRootPath $projectRootPath -SourceName $sourceName -CsvContentTypesFile $csvContentTypesFile -CsvContentTypesFieldsFile $csvContentTypesFieldsFile;
.\SPO-SET-SiteTemplate.ps1 -OrgName $orgName -SitePath $sitePath -ProjectRootPath $projectRootPath -SourceName $sourceName -SiteTemplatePath $siteTemplatePath;

Write-Host "Import script complete`n" -ForegroundColor Green;