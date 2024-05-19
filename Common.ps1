# =====================================================================================================
# Common functions...
# -----------------------------------------------------------------------------------------------------
# Modified By:   Siôn Lewis (www.sjlewis.com)
# Modified Date: 18/05/2024
# =====================================================================================================



#
# Reads the input of of either [y] Yes | [n] No and returns either $true or $false.
# -----------------------------------------------------------------------------------------------------
# Read-Confirmation -Message "Are you sure you want to proceed? [y] Yes | [n] No";
#
function Read-Confirmation {
    param (
        [string]$Message = "Are you sure you want to proceed? [y] Yes | [n] No"
    )
    
    [bool]$output = $false;
    [string]$confirmation = Read-Host $Message;
    if ($confirmation -eq 'y') {
        $output = $true;
    }
    elseif ($confirmation -eq 'n') {
        $output = $false;
    }
    else {
        Write-Host "Please enter either [y] Yes or [n] No."
        $output = Read-Confirmation -Message $Message;
    }
    return $output;
}