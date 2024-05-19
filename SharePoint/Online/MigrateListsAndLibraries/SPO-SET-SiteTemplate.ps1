# =====================================================================================================
# SharePoint Online: Import a Site Template from the XML file.
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
     [Parameter(Mandatory)]
     [string]$SourceName,
     [Parameter(Mandatory)]
     [string]$SiteTemplatePath
 )

 [string]$orgName = $OrgName;
 [string]$sitePath = $SitePath;
 [string]$projectRootPath = $ProjectRootPath;
 [string[]]$sourceName = $SourceName;
 [string]$siteTemplatePath = $SiteTemplatePath;


<#
# -----------------------------------------------------------------------------------------------------
# Variables for debugging.
# -----------------------------------------------------------------------------------------------------
[string]$orgName            = "destination";                                # Tenant name or sub-domain
[string]$sitePath           = "/sites/Demo";                                # Example: "/sites/Teamwork"
[string]$projectRootPath    = "C:\PowerShell\SharePoint\Online\SPOAssets";  # Path to the scripts folder.                                          # Default to 0, unless the server is thowing remote server errors.
[string]$sourceName         = "sourcedemo";									# Source tenant name or sub-domain.
# -----------------------------------------------------------------------------------------------------
# Default values.
# -----------------------------------------------------------------------------------------------------
[string]$siteTemplatePath = (".\Assets\{0}-PnPSiteTemplate.xml" -f $sourceName.ToUpper());
# -----------------------------------------------------------------------------------------------------
#>


Set-Location $projectRootPath;

[string]$siteUrl = ("https://{0}.sharepoint.com{1}" -f $orgName, $sitePath);

Write-Host ("Connecting to SGM Site: {0}" -f $siteUrl);
Connect-PnPOnline -Url $siteUrl -Interactive;

# Enable Scripting.
Write-Host ("Updating NoScriptSite for: {0}" -f $siteUrl);
Set-PnPSite -NoScriptSite $false;
Write-Host "Done`n" -ForegroundColor Green;

# Apply PnP Site Template.
Write-Host ("Applying PnP Site Template to SGM: {0}" -f $siteUrl);
Invoke-PnPSiteTemplate -Path $siteTemplatePath -Handlers Lists;
Write-Host "Done`n" -ForegroundColor Green;

Disconnect-PnPOnline;
Write-Host ("Disconnected from (SGM): {0}`n" -f $siteSiteRegtUrl) -ForegroundColor Green;
