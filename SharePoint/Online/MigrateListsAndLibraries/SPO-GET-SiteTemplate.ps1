# =====================================================================================================
# SharePoint Online: Export a Site Template to an XML file.
# -----------------------------------------------------------------------------------------------------
# Created By:    Siôn Lewis (www.sjlewis.com)
# Modified By:   Siôn Lewis (www.sjlewis.com)
# Modified Date: 18/05/2024
# -----------------------------------------------------------------------------------------------------
# Prerequisites: This script uses PnP Management Shell, see the link for configuration instructions:
# https://learn.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-pnppowershell
# =====================================================================================================

param(
     [Parameter(Mandatory)]
     [string]$OrgName,
     [Parameter(Mandatory)]
     [string]$SitePath,
     [Parameter(Mandatory)]
     [string]$ProjectRootPath,
     [Parameter()]
     [string[]]$ArrList,
     [Parameter(Mandatory)]
     [string]$SiteTemplatePath
 )

 [string]$orgName = $OrgName;
 [string]$sitePath = $SitePath;
 [string]$projectRootPath = $ProjectRootPath;
 [string[]]$arrList = $ArrList;
 [string]$siteTemplatePath = $SiteTemplatePath;


<#
# -----------------------------------------------------------------------------------------------------
# Variables for debugging.
# -----------------------------------------------------------------------------------------------------
[string]$orgName            = "sourcedemo";                                 # Tenant name or sub-domain
[string]$sitePath           = "/sites/Demo";                                # Example: "/sites/Teamwork"
[string]$projectRootPath    = "C:\PowerShell\SharePoint\Online\SPOAssets";  # Path to the scripts folder.                                          # Default to 0, unless the server is thowing remote server errors.
[string[]]$arrList          = @(
    "Demo List",
    "Another List"
);											                                # Add the name of Lists and Libraries to export.
# -----------------------------------------------------------------------------------------------------
# Default values.
# -----------------------------------------------------------------------------------------------------
[string]$siteTemplatePath = (".\Assets\{0}-PnPSiteTemplate.xml" -f $orgName.ToUpper());
# -----------------------------------------------------------------------------------------------------
#>


Set-Location $projectRootPath;

[string]$siteUrl = ("https://{0}.sharepoint.com{1}" -f $orgName, $sitePath);

Write-Host ("Connecting to: {0}`n" -f $siteUrl);
Connect-PnPOnline -Url $siteUrl -Interactive;

Write-Host ("`tGetting Site Template: '{0}'.`n" -f $sitePath) -ForegroundColor White;
#Set-PnPTraceLog -On -Level Debug;
Get-PnPSiteTemplate -Handlers Lists -ListsToExtract $arrList -Out $siteTemplatePath;
#Set-PnPTraceLog -Off;
Write-Host ("`tSite Template: '{0}' has been processed.`n" -f $sitePath) -ForegroundColor Green;

Disconnect-PnPOnline;
Write-Host ("Disconnected from: {0}`n" -f $siteUrl) -ForegroundColor Green;
