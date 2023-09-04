# This Script checks if there's a folder titled Devices's HOSTNAME in F-Drive, if not, it will be created, and it will check and do for a subfolder titled "LOGS" inside it, afterwards the current Security Logs will be exported in LOGS.
# We can replace any other Log List name with "Security" to have a current export of that.

# ©Payam Avarwand
############################################################

$date = get-date -f yyyyMMdd
$path= "F:\"
$PATH_host ="$path$env:COMPUTERNAME"


if (!(Test-Path $PATH_host)){
New-Item -Path $PATH_host -ItemType Directory
}
$logs ="LOGS"
$PATH_log ="$PATH_host\$logs"
if (!(Test-Path $PATH_log)){
New-Item -Path $PATH_log -ItemType Directory
}

wevtutil epl Security $PATH_log\Security_$date.evtx


# Created_by Payam.Avarwand
# Payam_avar@yahoo.com


