# =====================================================================================================
# SharePoint Online: Export Site Columns to a CSV file.
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
     [string[]]$ArrFieldGroups,
     [Parameter(Mandatory)]
     [string[]]$CsvSiteColumnsFile
 )

 [string]$orgName = $OrgName;
 [string]$sitePath = $SitePath;
 [string]$projectRootPath = $ProjectRootPath
 [double]$slowSPCalls = $SlowSPCalls;
 [string[]]$arrFieldGroups = $ArrFieldGroups;
 [string]$csvSiteColumnsFile = $CsvSiteColumnsFile


<#
# -----------------------------------------------------------------------------------------------------
# Variables for debugging.
# -----------------------------------------------------------------------------------------------------
[string]$orgName            = "sourcedemo";                                 # Tenant name or sub-domain
[string]$sitePath           = "/sites/Demo";                                # Example: "/sites/Teamwork"
[string]$projectRootPath    = "C:\PowerShell\SharePoint\Online\SPOAssets";  # Path to the scripts folder.
[double]$slowSPCalls 		= 0;                                            # Default to 0, unless the server is thowing remote server errors.
[string[]]$arrFieldGroups   = @(
    "Demo Site Columns",
    "Another Group of Site Columns"
);											                                # Restrict the columns retrived by their groups.
# -----------------------------------------------------------------------------------------------------
# Default values.
# -----------------------------------------------------------------------------------------------------
[string]$csvSiteColumnsFile = (".\Assets\{0}-CSV01-SiteColumns.csv" -f $orgName.ToUpper());
# -----------------------------------------------------------------------------------------------------
#>


#
# Returns a field from a list or site as XML.
# ------------------------------------------------------------------------------------------------------
# https://pnp.github.io/powershell/cmdlets/Get-PnPField.html
# ------------------------------------------------------------------------------------------------------
# Pre-connected to SharePoint Online:
# Get-FieldAsXML -WebUrl $webUrl -InternalName "demoDesc";
# Get-FieldAsXML -WebUrl $webUrl -InternalName "demoDesc" -ListName "Project Demo";
#
function Get-FieldAsXML {
    param (
        [string]$WebUrl, 
        [string]$InternalName, 
        [string]$ListName = ""
    )
    # The field will look for a Site Column...
    Write-Host "Getting a Site Column..." -ForegroundColor White;

    try {
        #Connect-PnPOnline -Url $WebUrl -Interactive;

        if ([string]::IsNullOrEmpty($ListName)) {
            $col = Get-PnPField -Identity $InternalName;
        }
        else {
            $col = Get-PnPField -Identity $InternalName -List $ListName;
        }
	
        $xml = $col.SchemaXml;
        Write-Host ("Site Column: '{0}' has been located." -f $InternalName) -ForegroundColor Green;
        Write-Host ("XML: `t{1}`n" -f $InternalName, $xml) -ForegroundColor Yellow;

        return $xml;
    }
    catch {
        $errorMessage = $_.Exception.Message;
        $failedItem = $_.Exception.ItemName;
        Write-Host ("Error getting a Site Column: '{0}' from List: {2}. {3}, {4}`n" -f $InternalName, $ListName, $errorMessage, $failedItem) -ForegroundColor Red;
    }
    finally {
        #Disconnect-PnPOnline;
    }
}

#
# Clean XML data and saves it to a file.
# ------------------------------------------------------------------------------------------------------
# Add-OutputToFile -Path $xmlOuputPath -InternalName "demoDesc" -Value $tmp01;
#
function Add-OutputToCSVFile {
    param (
        [string]$Path,
		[string]$InternalName,
        [string]$Value
    )
	
    [string]$xmlContent = $Value;

    # Removes SourceID, for example: SourceID=\"{00000000-0000-0000-0000-000000000000}\"
    $xmlContent = $xmlContent -replace '(SourceID=\"\{).*(\}\")', '';

    # Removes Version, for example: Version=\"12\"
    $xmlContent = $xmlContent -replace 'Version=\"\d+\"', '';

    # Removes Version, for example: AllowDeletion=\"TRUE\"
    $xmlContent = $xmlContent -replace 'AllowDeletion=\".*\"', '';
	
	# Escape double quotes.
    $xmlContent = $xmlContent -replace '"', '""';

    # Remove double spaces.
    $xmlContent = $xmlContent -replace '  ', '';
	
	# Create CSV format: InternalName,FieldXml.
    $xmlContent = ("{0},`"{1}`"" -f $InternalName, $xmlContent);
    
    Add-Content -Path $Path -Value $xmlContent;
}


# =====================================================================================================


Set-Location $projectRootPath;

[string]$siteUrl = ("https://{0}.sharepoint.com{1}" -f $orgName, $sitePath);

Write-Host ("Connecting to: {0}" -f $siteUrl);
Connect-PnPOnline -Url $siteUrl -Interactive -ForceAuthentication;

# Add CSV Headings.
Add-Content -Path $csvSiteColumnsFile -Value "InternalName,FieldXml";

foreach ($fieldGroup in $arrFieldGroups) {

    [Object[]]$arrFields = Get-PnPField -Group $fieldGroup;
    if ($null -ne $slowSPCalls) {
        Start-Sleep -Seconds $slowSPCalls;
    }

    # Add an empty line to divide the different groups of fields.
    Add-Content -Path $csvSiteColumnsFile -Value ",";

    foreach ($field in $arrFields) {

        $fieldXML = Get-FieldAsXML -WebUrl $siteUrl -InternalName $field.InternalName;
        if ($null -ne $slowSPCalls) {
            Start-Sleep -Seconds $slowSPCalls;
        }
        
        if ([string]::IsNullOrWhiteSpace($fieldXML) -eq $false) {
            Add-OutputToCSVFile -Path $csvSiteColumnsFile -InternalName $field.InternalName -Value $fieldXML;
            if ($null -ne $slowSPCalls) {
                Start-Sleep -Seconds $slowSPCalls;
            }
        } else {
            Write-Host "Try running the following code manually:`n" -ForegroundColor Magenta;
            Write-Host ("`t`$fieldXML = Get-FieldAsXML -WebUrl {0} -InternalName {1};" -f $siteUrl, $field.InternalName) -ForegroundColor Yellow;
            Write-Host ("`tAdd-OutputToCSVFile -Path {0} -InternalName {1} -Value `$fieldXML;`n" -f $csvSiteColumnsFile, $field.InternalName) -ForegroundColor Yellow;
            Write-Host ("TIP: Add a break point here to keep the context within the foreach loop...`n") -ForegroundColor Magenta;
        }
    }
}

Disconnect-PnPOnline;
Write-Host ("Disconnected from: {0}`n" -f $siteUrl) -ForegroundColor Green;
