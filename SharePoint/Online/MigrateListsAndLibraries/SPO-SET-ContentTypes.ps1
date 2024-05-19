# =====================================================================================================
# SharePoint Online: Import Content Types and Columns (Fields) from the CSV files.
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
     [string]$CsvContentTypesFile,
     [Parameter(Mandatory)]
     [string]$CsvContentTypesFieldsFile
 )

 [string]$orgName = $OrgName;
 [string]$sitePath = $SitePath;
 [string]$projectRootPath = $ProjectRootPath
 [string]$sourceName = $SourceName;
 [string]$csvContentTypesFile = $CsvContentTypesFile;
 [string]$csvContentTypesFieldsFile = $CsvContentTypesFieldsFile;


<#
# -----------------------------------------------------------------------------------------------------
# Variables for debugging.
# -----------------------------------------------------------------------------------------------------
[string]$orgName            = "sourcedemo";                                 # Tenant name or sub-domain
[string]$sitePath           = "/sites/Demo";                                # Example: "/sites/Teamwork"
[string]$projectRootPath    = "C:\PowerShell\SharePoint\Online\SPOAssets";  # Path to the scripts folder.
[string]$sourceName         = "sourcedemo";									# Source tenant name or sub-domain.
# -----------------------------------------------------------------------------------------------------
# Default values.
# -----------------------------------------------------------------------------------------------------
[string]$csvContentTypesFile = (".\Assets\{0}-CSV02-ContentTypes.csv" -f $sourceName.ToUpper());
[string]$csvContentTypesFieldsFile = (".\Assets\{0}-CSV03-ContentTypesFields.csv" -f $sourceName.ToUpper());
# -----------------------------------------------------------------------------------------------------
#>


#
# Adds a new Content Type.
# ------------------------------------------------------------------------------------------------------
# https://pnp.github.io/powershell/cmdlets/Add-PnPContentType.html
# -Name: 				**Required** Specify the name of the new content type 
# -Description: 		Specifies the description of the new content type
# -Group: 				Specifies the group of the new content type
# -ParentCTName: 		Specifies the name of the parent of the new content type
# -ContentTypeId: 		If specified, in the format of 0x0100233af432334r434343f32f3, will create a content type with the specific ID
# ------------------------------------------------------------------------------------------------------
# Add-ContentType -Name "SL Demo CT" -Description "My Demo Content Type" -Group "SL Content Types" -ParentCTName "Item";
#
function Add-ContentType {
    param (
        [string]$Name, 
        [string]$Description, 
        [string]$Group, 
        [string]$ParentCTName, 
        [string]$ContentTypeId
    )

    try {
        Write-Host ("Creating Content Type '{0}'..." -f $Name) -ForegroundColor White;
        if ([string]::IsNullOrWhiteSpace($ContentTypeId) -eq $false) {
            Add-PnPContentType -Name $Name -Description $Description -Group $Group -ContentTypeId $ContentTypeId -ErrorAction Stop;
        }
        else {
            # https://pnp.github.io/powershell/cmdlets/Get-PnPContentType.html
            Write-Host "`tGetting Parent Content Type..." -ForegroundColor Gray;
            $Local:parentContentType = Get-PnPContentType -Identity $ParentCTName -InSiteHierarchy;
            Add-PnPContentType -Name $Name -Description $Description -Group $Group -ParentContentType $Local:parentContentType -ErrorAction Stop;
        }
        Write-Host ("Content Type has been created under the '{1}' group." -f $Name, $Group) -ForegroundColor Green;
    }
    catch {
        $errorMessage = $_.Exception.Message;
		if($errorMessage.Contains("A duplicate content type") -eq $true) {
			Write-Host ("`tWarning: {0}`n" -f $errorMessage) -ForegroundColor DarkYellow;
		} else {
			$failedItem = $_.Exception.ItemName;
            Write-Host ("Error creating Content Type: '{0}' under: '{1}' group. {2}, {3}" -f $Name, $Group, $errorMessage, $failedItem) -ForegroundColor Red;
		}
    }
}

#
# Adds an existing Site Column to a Content Type.
# ------------------------------------------------------------------------------------------------------
# https://pnp.github.io/powershell/cmdlets/Add-PnPFieldToContentType.html
# ------------------------------------------------------------------------------------------------------
# Pre-connected to SharePoint Online:
# Add-FieldToContentType -InternalFieldName "SLDemoield" -ContentTypeName "SL Demo CT";
# Add-FieldToContentType -InternalFieldName "SLDemoField" -ContentTypeName "SL Demo CT" -Hidden -Required;
#
function Add-FieldToContentType {
    param (
        [Parameter(Mandatory = $true)][string]$InternalFieldName, 
        [Parameter(Mandatory = $true)][string]$ContentTypeName, 
        [switch]$Hidden = $false, 
        [switch]$Required = $false
    )
	
    try {
        Write-Host "Adding Site Column to Content Type..." -ForegroundColor White;

        [string]$cmd = ("Add-PnPFieldToContentType -Field `"{0}`" -ContentType `"{1}`"" -f $InternalFieldName, $ContentTypeName);
        if ($Hidden) {
            $cmd = ("{0} -Hidden" -f $cmd);
        }
        if ($Required) {
            $cmd = ("{0} -Required" -f $cmd);
        }
        $cmd = ("{0} -ErrorAction Stop;" -f $cmd);
        
        Write-Host $cmd -ForegroundColor Gray;
        Invoke-Expression -Command $cmd;
				
        Write-Host ("Site Column: '{0}' has been added to '{1}' Content Type." -f $InternalFieldName, $ContentTypeName) -ForegroundColor Green;
    }
    catch {
        $errorMessage = $_.Exception.Message;
        $failedItem = $_.Exception.ItemName;
        Write-Host ("Error adding Site Column: '{0}' to Content Type: '{1}'. {2}, {3}" -f $InternalFieldName, $ContentTypeName, $errorMessage, $failedItem) -ForegroundColor Red;
    }
}


# ======================================================================================================


Set-Location $projectRootPath;

[string]$siteUrl = ("https://{0}.sharepoint.com{1}" -f $orgName, $sitePath);

Write-Host ("Connecting to: {0}" -f $siteUrl);
Connect-PnPOnline -Url $siteUrl -Interactive;

if([string]::IsNullOrWhiteSpace($csvContentTypesFile) -eq $false) {

	$dataSet = Import-Csv $csvContentTypesFile -Delimiter ",";
	
	ForEach ($row in $dataSet) {
		
        if ([string]::IsNullOrWhiteSpace($row.Name) -eq $false -and [string]::IsNullOrWhiteSpace($row.ParentCTName) -eq $false) {
			
            Add-ContentType -Name $row.Name -Description $row.Description -Group $row.Group -ParentCTName $row.ParentCTName -ContentTypeId $row.ContentTypeId;
		
        }
	}
}

Write-Host "`n ---- `n";

if([string]::IsNullOrWhiteSpace($csvContentTypesFieldsFile) -eq $false) {

	$dataSet = Import-Csv $csvContentTypesFieldsFile -Delimiter ",";
	
	ForEach ($row in $dataSet) {
		
        if ([string]::IsNullOrWhiteSpace($row.ContentTypeName) -eq $false -and [string]::IsNullOrWhiteSpace($row.InternalFieldName) -eq $false) {
			
            [bool]$Local:hidden = $false;
			if ([string]::IsNullOrWhiteSpace($row.Hidden) -eq $false -and $row.Hidden.ToLower() -eq "true") {
				$Local:hidden = $true;
			}
			
            [bool]$Local:required = $false;
			if ([string]::IsNullOrWhiteSpace($row.Required) -eq $false -and $row.Required.ToLower() -eq "true") {
				$Local:required = $true;
			}
			
            Add-FieldToContentType -ContentTypeName $row.ContentTypeName -InternalFieldName $row.InternalFieldName -Hidden:$Local:hidden -Required:$Local:required;
		}
	}
}

Disconnect-PnPOnline;
Write-Host ("Disconnected from: {0}`n" -f $siteUrl) -ForegroundColor Green;
