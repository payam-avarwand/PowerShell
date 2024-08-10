# We use this script to update our backup of GPOs
# This script deletes the path where the last GPOs were exported and exports the current GPOs there instead.
############################################################



# Import GPO PS-Modul
import-module grouppolicy

$ExprtFolder = "\\srv\Current-GPOs\"
Get-ChildItem -Path $ExprtFolder -Include * | Remove-Item -Recurse -Force

$GPO = Get-GPO -All


# each GPO will be exported to a named folder
foreach ($GPO1 in $GPO) {
	$Path=$ExprtFolder+$GPO1.Displayname
	New-Item -ItemType directory -Path $Path
	Backup-GPO -Guid $GPO1.id -Path $Path
}




# Created_by Payam.Avarwand
# Payam_avar@yahoo.com

