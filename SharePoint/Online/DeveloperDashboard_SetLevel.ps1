# ========================================================================
# Connect to SharePoint Online with the SharePoint Online Management Shell
# ------------------------------------------------------------------------
# First install the following:
# 1. Windows Management Framework 3.0
#    http://www.microsoft.com/en-us/download/details.aspx?id=34595
# 2. SharePoint Online Management Shell
#    http://www.microsoft.com/en-us/download/details.aspx?id=35588
# 3. Manage Windows Azure AD using Windows PowerShell
#    http://technet.microsoft.com/en-us/library/jj151815.aspx#bkmk_installmodule
# 4. Run the script below, however your accont will need to be a tenant 
#    administration  to avoid the following error:
#    Connect-SPOService : Current site is not a tenant administration site...
# ========================================================================


# Add SharePoint Snapin to PowerShell            
if((Get-PSSnapin | Where {$_.Name -eq "Microsoft.SharePoint.PowerShell"}) -eq $null)             
{
    Add-PSSnapin Microsoft.SharePoint.PowerShell            
}


$userName = "accountname@sponlinsite.onmicrosoft.com"
$password = "[PASSWORD]"
$siteCollectionUrl = "https://sponlinsite.sharepoint.com/sites/[SITE_COLLECTION]"
$siteAdminUrl = "https://sponlinsite-admin.sharepoint.com/"
$securePassword = ConvertTo-SecureString $password –AsPlainText –force
$O365Credential = New-Object System.Management.Automation.PsCredential($username, $securePassword)


Connect-SPOService –url  $siteCollectionUrl –Credential $O365Credential

    $content = ([Microsoft.SharePoint.Administration.SPWebService]::ContentService)
    $appsetting = $content.DeveloperDashboardSettings
    # Comment code in and out as required... ------------------------------------------------------------
    #$appsetting.DisplayLevel = [Microsoft.SharePoint.Administration.SPDeveloperDashboardLevel]::Off
    #$appsetting.DisplayLevel = [Microsoft.SharePoint.Administration.SPDeveloperDashboardLevel]::OnDemand
    $appsetting.DisplayLevel = [Microsoft.SharePoint.Administration.SPDeveloperDashboardLevel]::On
    # ---------------------------------------------------------------------------------------------------
    $appsetting.Update() 

Disconnect-SPOService
