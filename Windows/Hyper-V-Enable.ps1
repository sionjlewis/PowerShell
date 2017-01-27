# =========================
# Enable Microsoft Hyper-V
# http://www.SJLewis.com
# =========================

$obj = Get-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V

If ($obj.State -eq 'Disable') {
    Write-Host ("`tWindowsOptionalFeature: {0} is being enabled..." -f $obj.FeatureName) -ForegroundColor Green
    $state = Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V –All
    
    If ($state.RestartNeeded -eq $true) {
        Write-Host "`tPlease restart your system..." -ForegroundColor Yellow
    } Else {
        Write-Host "`tNo need to restart your system..." -ForegroundColor Green
    }
} Else {
    Write-Host ("`tWindowsOptionalFeature: {0}`n`tState: {1}`n" -f $obj.FeatureName, $obj.State) -ForegroundColor DarkYellow
}