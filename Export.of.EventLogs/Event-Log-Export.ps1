# This Script checks if there's a folder titled Devices's HOSTNAME in F-Drive, if not, it will create it, then it will do the same for a subfolder titled "LOGS", afterwards the current Security Logs will be exported in LOGS.
# We can replace any other Log List name with "Security" to have a current export of that.

# ©Payam Avarwand
############################################################



$date = get-date -f yyyyMMdd
$Drive= "F:\"
$PATH ="$Drive$env:COMPUTERNAME"

if (!(Test-Path $PATH)){
New-Item -Path $PATH -ItemType Directory
}


$logs ="LOGS"
$PATH_Log ="$PATH\$logs"
if (!(Test-Path $PATH_Log)){
New-Item -Path $PATH_Log -ItemType Directory
}

wevtutil epl Security $PATH_Log\Security_$date.evtx







# Created_by Payam.Avarwand
# Payam_avar@yahoo.com


