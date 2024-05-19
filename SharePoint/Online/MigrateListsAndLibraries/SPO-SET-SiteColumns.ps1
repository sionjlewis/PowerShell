# =====================================================================================================
# SharePoint Online: Import Site Columns from the CSV file.
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
     [double]$SlowSPCalls = 0,
     [Parameter()]
     [string[]]$SourceName,
     [Parameter(Mandatory)]
     [string[]]$CsvSiteColumnsFile
 )

 [string]$orgName = $OrgName;
 [string]$sitePath = $SitePath;
 [string]$projectRootPath = $ProjectRootPath
 [double]$slowSPCalls = $SlowSPCalls;
 [string[]]$sourceName = $SourceName;
 [string]$csvSiteColumnsFile = $CsvSiteColumnsFile


<#
# -----------------------------------------------------------------------------------------------------
# Variables for debugging.
# -----------------------------------------------------------------------------------------------------
[string]$orgName            = "destination";                                # Tenant name or sub-domain
[string]$sitePath           = "/sites/Demo";                                # Example: "/sites/Teamwork"
[string]$projectRootPath    = "C:\PowerShell\SharePoint\Online\SPOAssets";  # Path to the scripts folder.
[double]$slowSPCalls 		= 0;                                            # Default to 0, unless the server is thowing remote server errors.
[string]$sourceName         = "sourcedemo";									# Source tenant name or sub-domain.
# -----------------------------------------------------------------------------------------------------
# Default values.
# -----------------------------------------------------------------------------------------------------
[string]$csvSiteColumnsFile = (".\Assets\{0}-CSV01-SiteColumns.csv" -f $sourceName.ToUpper());
# -----------------------------------------------------------------------------------------------------
#>


Set-Location $projectRootPath;

[string]$siteUrl = ("https://{0}.sharepoint.com{1}" -f $orgName, $sitePath);

Write-Host ("Connecting to: {0}" -f $siteUrl);
Connect-PnPOnline -Url $siteUrl -Interactive;

$dataSet = Import-Csv $csvSiteColumnsFile -Delimiter ",";
foreach ($row in $dataSet) {
	
    try {

		# Declared a local variable to ensure that the value is processed as a string. If we don't do this then 
		# the 'Add-PnPFieldFromXml' cmd throws an exemption: Object reference not set to an instance of an object.
		[string]$Local:xml = $row.FieldXml;
		[string]$Local:name = $row.InternalName;
        
		if ([string]::IsNullOrWhiteSpace($Local:xml) -eq $false -and [string]::IsNullOrWhiteSpace($Local:name) -eq $false) {
		
			Write-Host ("`tAdding field: '{0}', FieldXml: {1}`n" -f $Local:name, $Local:xml) -NoNewline -ForegroundColor Gray;
			Add-PnPFieldFromXml -FieldXml $Local:xml -ErrorAction Stop;
			Write-Host "`tDone`n" -ForegroundColor DarkGreen;
			if ($null -ne $slowSPCalls) {
				Start-Sleep -Seconds $slowSPCalls;
			}

		}
		elseif ([string]::IsNullOrWhiteSpace($Local:name) -eq $true) {

			Write-Host "`tSkip row`n" -ForegroundColor DarkGray;	

		}
		else {

			Write-Error -Message "`tError: please set a value for FieldXml within the CSV file.`n" -Category NotImplemented;
		
        }
	}
	catch {

		$errorMessage = $_.Exception.Message;
		if($errorMessage.Contains("A duplicate field name") -eq $true) {

			Write-Host ("`tWarning: {0}`n" -f $errorMessage) -ForegroundColor DarkYellow;

		} else {

			$failedItem = $_.Exception.ItemName;
			Write-Host ("`tError adding Field: {0} using XML. {1}, {2}`n" -f $row.InternalName, $errorMessage, $failedItem) -ForegroundColor DarkRed;
		
        }
	}
}

Disconnect-PnPOnline;
Write-Host ("Disconnected from: {0}`n" -f $siteUrl) -ForegroundColor Green;
