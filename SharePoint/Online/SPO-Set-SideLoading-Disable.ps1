# ================================================
# First dowload SharePoint Online management Shell
# http://www.microsoft.com/en-gb/download/details.aspx?id=35588
# ================================================

# Add SharePoint Snapin to PowerShell            
#if((Get-PSSnapin | Where {$_.Name -eq "Microsoft.SharePoint.PowerShell"}) -eq $null)             
#{
#    Add-PSSnapin Microsoft.SharePoint.PowerShell            
#}


$programFiles = [environment]::getfolderpath("programfiles")
add-type -Path $programFiles'\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.dll'


Write-Host 'To DISABLE SharePoint app sideLoading, enter Site Url, username and password'
 

$siteurl = Read-Host 'Site Url'
$username = Read-Host "User Name"
$password = Read-Host -AsSecureString 'Password'
 
# Add default values if User enters nothing...
if ($siteurl -eq '') 
{
    $siteurl = 'https://tenant.sharepoint.com/sites/mysite'
    $username = 'account@tenant.onmicrosoft.com'
    $password = ConvertTo-SecureString -String 'mypassword!' 
        -AsPlainText -Force
}

$outfilepath = $siteurl -replace ':', '_' -replace '/', '_'
 
try 
{
    [Microsoft.SharePoint.Client.ClientContext]$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteurl)
    [Microsoft.SharePoint.Client.SharePointOnlineCredentials]$spocreds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($username, $password)
    $ctx.Credentials = $spocreds
    $sideLoadingEnabled = [Microsoft.SharePoint.Client.appcatalog]::IsAppSideloadingEnabled($ctx);
    $ctx.ExecuteQuery()
    
    if($sideLoadingEnabled.value -eq $false)
    {
        Write-Host -ForegroundColor Green 'SideLoading is alreday DISABLED on site' $siteurl
    }
    else
    {
        Write-Host -ForegroundColor Yellow 'DISABLING SideLoading feature on site' $siteurl
        $sideLoadingGuid = new-object System.Guid "AE3A1339-61F5-4f8f-81A7-ABD2DA956A7D"
        $site = $ctx.Site;
        $site.Features.Remove($sideLoadingGuid, $true);
        $ctx.ExecuteQuery();
        Write-Host -ForegroundColor Green 'SideLoading disabled on site' $siteurl
    }
} 
catch 
{
    Write-Host -ForegroundColor Red 'Error encountered when trying to disable side loading features' $siteurl, ':' $Error[0].ToString();
}
