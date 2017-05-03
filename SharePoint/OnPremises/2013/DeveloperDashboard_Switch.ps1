# ======================================================================
# To see the Developer Dashboard, the user browsing the page must have 
# the AddAndCustomizePages permission (by default, site collection 
# admins and users in the Owner group have this permission).
# ======================================================================

$dash = [Microsoft.SharePoint.Administration.SPWebService]::ContentService.DeveloperDashboardSettings;
$dash.DisplayLevel = 'OnDemand';
# Comment in the relevant action On|Off
#$dash.DisplayLevel = 'On';
#$dash.DisplayLevel = 'Off';
$dash.TraceEnabled = $true;
$dash.Update();
