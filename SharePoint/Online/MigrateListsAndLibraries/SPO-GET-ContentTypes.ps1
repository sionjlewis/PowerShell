# =====================================================================================================
# SharePoint Online: Export Content Types and Columns (Fields) to CSV files.
# -----------------------------------------------------------------------------------------------------
# Created By:    Siôn Lewis (www.sjlewis.com)
# Modified By:   Siôn Lewis (www.sjlewis.com)
# Modified Date: 18/05/2024
# -----------------------------------------------------------------------------------------------------
# Prerequisites: This script uses PnP Management Shell, see the link for configuration instructions:
# https://learn.microsoft.com/en-us/sharepoint/dev/declarative-customization/site-design-pnppowershell
# -----------------------------------------------------------------------------------------------------
# Instructions:  Update the variables at the top of the script before running within your environment.
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
     [string[]]$ArrCTGroupNames,
     [Parameter()]
     [bool]$OpenInEdge = $true,
     [Parameter()]
     [bool]$ManageCTName = $false,
     [Parameter()]
     [bool]$ReviewCTName = $false,
     [Parameter(Mandatory)]
     [string]$CsvContentTypesFile,
     [Parameter(Mandatory)]
     [string]$CsvContentTypesFieldsFile
 )

 [string]$orgName = $OrgName;
 [string]$sitePath = $SitePath;
 [string]$projectRootPath = $ProjectRootPath
 [double]$slowSPCalls = $SlowSPCalls;
 [string[]]$arrCTGroupNames = $ArrCTGroupNames;
 [bool]$openInEdge = $OpenInEdge;
 [bool]$manageCTName = $ManageCTName;
 [bool]$reviewCTName = $ReviewCTName;
 [string]$csvContentTypesFile = $CsvContentTypesFile;
 [string]$csvContentTypesFieldsFile = $CsvContentTypesFieldsFile;


<#
# -----------------------------------------------------------------------------------------------------
# Variables for debugging.
# -----------------------------------------------------------------------------------------------------
[string]$orgName            = "sourcedemo";                                 # Tenant name or sub-domain
[string]$sitePath           = "/sites/Demo";                                # Example: "/sites/Teamwork"
[string]$projectRootPath    = "C:\PowerShell\SharePoint\Online\SPOAssets";  # Path to the scripts folder.
[double]$slowSPCalls 		= 0;                                            # Default to 0, unless the server is thowing remote server errors.
[string[]]$arrCTGroupNames  = @(
    "Demo Content Types",
    "Another Group of Content Types"
);											                                # Restrict the content types retrived by their groups.
[bool]$openInEdge = $true;                                                  # Set to false to use Google Chrome.
# ------------------------------------------------------------------------------------------------------
# Use these properties to manage existing Content Types and
# set them both to false to export the Content Types' details.
[bool]$manageCTName = $false;
[bool]$reviewCTName = $false;
# -----------------------------------------------------------------------------------------------------
# Ddefault values.
# -----------------------------------------------------------------------------------------------------
[string]$csvContentTypesFile = (".\Assets\{0}-CSV02-ContentTypes.csv" -f $orgName.ToUpper());
[string]$csvContentTypesFieldsFile = (".\Assets\{0}-CSV03-ContentTypesFields.csv" -f $orgName.ToUpper());
# -----------------------------------------------------------------------------------------------------
#>


#
# Get an array of unique content type group names.
# -----------------------------------------------------------------------------------------------------
# Get-ContentTypeGroups -ContentTypes $cTypes -CTGroupNames $CTGroupNames;
#
function Get-ContentTypeGroups {
    param (
        [object]$ContentTypes, 
        [string]$CTGroupNames
    )
    $ctGroupArray = New-Object System.Collections.ArrayList;

    foreach ($ctype in $ContentTypes) {
        
        # TIP: To match on the start of a group name, change the '-eq "$CTGroupNames"' to '-like "*$CTGroupNames*"'.
        if ($ctype.Group -eq "$CTGroupNames" -and !$ctGroupArray.Contains($ctype.Group)) {
            
            # Note: We needed to add "| Out-Null" to avoid the returning unwanted rows... 
            $ctGroupArray.Add($ctype.Group) | Out-Null;
        }
    }

    return $ctGroupArray;
}

#
# Open a Chrome browser window. Requires Google Chrome to be locally installed.
# -----------------------------------------------------------------------------------------------------
# Open-BrowserWindow -PageUrl $url -OpenInEdge $true;
#
function Open-BrowserWindow([string]$PageUrl, [bool]$OpenInEdge) {
    if(OpenInEdge){
        Start-Process -FilePath msedge.exe -ArgumentList $PageUrl;
    } else {
        Start-Process -FilePath Chrome -ArgumentList $PageUrl;
    }
}

#
# Output Content Type to CSV file.
# -----------------------------------------------------------------------------------------------------
# Add-OutputContentTypeToCsvFile -Path $csvContentTypesFile -ContentType $ctype;
#
function Add-OutputContentTypeToCsvFile {
    param (
        [string]$Path,
        [object]$ContentType
    )

    $ctParent = $ctype.Parent;
    $ctx.Load($ctParent);
    $ctx.ExecuteQuery();

    #"Name,Group,ParentCTName,ContentTypeId,Description"
    [string]$csvRow = ("{0},{1},{2},{3},{4}" -f $ContentType.Name, $ContentType.Group, $ctParent.Name, $ContentType.Id.StringValue, $ContentType.Description);
    
    Add-Content -Path $Path -Value $csvRow;
}

#
# Output Content Type Fields to CSV file.
# -----------------------------------------------------------------------------------------------------
# Add-OutputContentTypeFieldsToCsvFile -Path $csvContentTypesFieldsFile -ContentType $ctype;
#
function Add-OutputContentTypeFieldsToCsvFile {
    param (
        [string]$Path,
        [object]$ContentType
    )

    $ctFields = $ContentType.Fields;
    $ctx.Load($ctFields);
    $ctx.ExecuteQuery();

    foreach ($field in $ctFields) {
	
		if($field.InternalName -ne "Title" -or $field.InternalName -ne "ContentType") {

			#"ContentTypeName,InternalFieldName,Hidden,Required"
			[string]$csvRow = ("{0},{1},{2},{3}" -f $ContentType.Name, $field.InternalName, $field.Hidden, $field.Required);
		
			Add-Content -Path $Path -Value $csvRow;
		}
    }
}


# =====================================================================================================


Set-Location $projectRootPath;

[string]$siteUrl = ("https://{0}.sharepoint.com{1}" -f $orgName, $sitePath);

Write-Host ("Connecting to: {0}" -f $siteUrl);
# Note: We won't be using '-ForceAuthentication' as this is the 2nd script to be run...
Connect-PnPOnline -Url $siteUrl -Interactive; # -ForceAuthentication;

try {
    Write-Host "Processing Content Type Groups(s)..." -ForegroundColor White;

    $cnn = Get-PnPConnection;
    $ctx = $cnn.Context;
    $ctx.ExecuteQuery();

    if ([string]::IsNullOrWhiteSpace($ctx.TraceCorrelationId) -eq $false) { 
        $site = $ctx.Site;
        $ctx.Load($site);
        $ctx.ExecuteQuery();

        $rootWeb = $site.RootWeb;
        $ctx.Load($rootWeb);
        $ctx.ExecuteQuery();

        $cTypes = $rootWeb.ContentTypes;
        $ctx.Load($cTypes);
        $ctx.ExecuteQuery();

        [System.Collections.ArrayList]$ctGroupArray = [System.Collections.ArrayList]@();
        foreach($CTGroupNames in $arrCTGroupNames) {
            ([array](Get-ContentTypeGroups -ContentTypes $cTypes -CTGroupNames $CTGroupNames)).ForEach({$ctGroupArray.Add($_)});
        }
        if($null -ne $ctGroupArray -and $ctGroupArray.Count -gt 1) {
            $ctGroupArray.Sort();
        }

        if ($manageCTName -eq $true -or $reviewCTName -eq $true) {
            # Wait for the Admin User to respond with any key.  
            Write-Host ("`n`nAre you ready to process the '{0}' group of content types?" -f $group) -ForegroundColor Yellow;
            Read-Host "Press Enter to continue";
        } else {
            # Add headings to the CSV files.
            [string]$ctFileHeader = "Name,Group,ParentCTName,ContentTypeId,Description"
            [string]$ctFileSpace = ",,,,"
            Add-Content -Path $csvContentTypesFile -Value $ctFileHeader;
            Add-Content -Path $csvContentTypesFile -Value $ctFileSpace;

            [string]$fileFileHeader = "ContentTypeName,InternalFieldName,Hidden,Required"
            [string]$fileFileSpace = ",,,"
            Add-Content -Path $csvContentTypesFieldsFile -Value $fileFileHeader;
        }

        foreach ($group in $ctGroupArray) {
            foreach ($ctype in $cTypes) {
                if ($ctype.Group -eq $group) {
                    Write-Host ("`t- Content Type: '{0}' Group: {1}`n" -f $ctype.Name, $ctype.Group) -ForegroundColor White;
                    
					if ($manageCTName -eq $true) {
						# Open a browser window to update the content type's name or description.
						$siteSettingPage = ("{0}/_layouts/15/ctypedit.aspx?ctype={1}" -f $siteUrl, $ctype.Id.ToString());
                    } 
					elseif ($reviewCTName -eq $true) {
						# Open a browser window to review the content type.
						$siteSettingPage = ("{0}/_layouts/15/ManageContentType.aspx?ctype={1}" -f $siteUrl, $ctype.Id.ToString());
                    }
					
					if ($manageCTName -eq $true -or $reviewCTName -eq $true) {
						Open-BrowserWindow -PageUrl $siteSettingPage -OpenInEdge $openInEdge;
						Start-Sleep -Milliseconds 500;
					} else {
                        # Populate the Content Types CSV file.
                        Add-OutputContentTypeToCsvFile -Path $csvContentTypesFile -ContentType $ctype;
                    
                        # Populate the Content Types' Fields CSV file.
                        Add-Content -Path $csvContentTypesFieldsFile -Value $fileFileSpace;
                        Add-OutputContentTypeFieldsToCsvFile -Path $csvContentTypesFieldsFile -ContentType $ctype;

                        if ($null -ne $slowSPCalls) {
                            Start-Sleep -Seconds $slowSPCalls;
                        }
                    }
                }
            }
        }

        Write-Host ("Finished opening the Manage CT Publishing Pages for groups starting with: '{0}'." -f $CTGroupNames) -ForegroundColor Green;
    }
    else {
        Write-Host "Unable to get the SharePoint Online (CSOM) client context." -ForegroundColor Red;
    }
}
catch {
    $errorMessage = $_.Exception.Message;
    $failedItem = $_.Exception.ItemName;
    $groupCSV = $ctGroupArray -join ",";
    Write-Host ("Error opening the Content Type's Manage Publishing Page for group(s): '{0}'. {1} , {2}" -f $groupCSV, $errorMessage, $failedItem) -ForegroundColor Red;
}
finally {
    # Disconnect from SharePoint Online.
    $ctx.TraceCorrelationId = "";
    $ctx.Dispose();

    # Disconnect from SharePoint Online.
    Disconnect-PnPOnline;
	Write-Host ("Disconnected from: {0}`n" -f $siteUrl) -ForegroundColor Green;
}
