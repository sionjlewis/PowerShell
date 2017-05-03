Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue;

$svc = [Microsoft.SharePoint.Administration.SPWebService]::ContentService
$dds = $svc.DeveloperDashboardSettings
$dds.DisplayLevel = "Off"
$dds.Update()
