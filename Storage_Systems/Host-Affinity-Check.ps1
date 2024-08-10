# In Fujitsu storage systems, there's a security Feature called "Host Affinity" to control the Host access to LUNs.
# It will be at least 1 "Host Affinity" required to create every single connection between the Storage and a Host.
# Assuming we use just one "Host Affinity" per Host Connection, now we have to make sure, that there is no 2 or more Host Affinities on every Fujitsu Storages with the same name.
# Assuming we have 10 Fujitsu Storage Systems, and we already have exported the host affinity lists of each Storage to a separate CSV file.

# This script:
## creates an Excel file with 2 Worksheets
## importes every time a csv-List of Host Affinities to the Excel file
## checks the repetition of each individual row



# Check if the Excel Module is available (installed)
###################################################################################################################################################
if (!(Get-Module ImportExcel -ListAvailable)) {
    Write-Host "Installing ImportExcel module..."
    Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
} else {
    Write-Host "ImportExcel module is already installed ;)"
}



# Define the general variables
###################################################################################################################################################
$D = Get-Date -Format 'dd-MM-yyyy'
$UserProfile = [Environment]::GetFolderPath("UserProfile")
$i=1





# Create the Check-loop
###################################################################################################################################################
while ($i -lt 11)
{


# Variables
################################################################
$SERVER="FJS-0$i"

$input_CSV_File="c:\Storage\Host Affinity\$SERVER.csv"
$Output_Path = "$UserProfile\Desktop\"
$OutPut_File = "$SERVER-HA-Check-$D.txt"

$excelFilePath = "c:\Storage\Host Affinity\$SERVER.xlsx"
$CheckListData_Address = "c:\Storage\Host Affinity\$SERVER.xlsx"

$sheetName1 = "sheet1"
$sheetName2 = "sheet2"
$ColumnName_in_SourceData = 'Hosts'
$ColumnName_in_CheckList = 'Hosts'



# Reference to the Microsoft.Office.Interop.Excel assembly:
################################################################
[Reflection.Assembly]::LoadWithPartialName('Microsoft.Office.Interop.Excel') | Out-Null


# Check/Create the Excel-WorkBook:
################################################################
if (-Not (Test-Path -Path $excelFilePath)) {
    Write-Host "File '$excelFilePath' does not exist. Creating it now."
    $excelApp = New-Object -ComObject Excel.Application
    $excelApp.Visible = $false
    $workbook = $excelApp.Workbooks.Add()

    # Create Sheet1
    $Worksheet1 = $workbook.Worksheets.Item(1)
    $Worksheet1.Name = $sheetName1

    # Create Sheet2 and shift it after Sheet1
    $Worksheet2 = $workbook.Worksheets.Add()
    $Worksheet2.Name = $sheetName2
    $Worksheet1.Move($Worksheet2)


    $workbook.SaveAs($excelFilePath)
    $excelApp.Quit()

} else {

    Write-Host "File '$excelFilePath' already exists."
}


# Check/Create WorkSheets and Add the Title:
################################################################
function Ensure-WS-Exists {
    param (
        [string]$Path,
        [string]$SheetName
    )

    try {
        # Attempt to import data from the specified sheet
        $excelData = Import-Excel -Path $Path -SheetName $SheetName -ErrorAction Stop
        Write-Host "Sheet '$SheetName' found."
    } catch {
        Write-Host "Sheet '$SheetName' doesn't exist, will be created."
        
        # Open the workbook
        $workbook = Open-ExcelPackage -Path $Path
        
        # Add the worksheet if it does not exist
        $worksheet = Add-Worksheet -ExcelPackage $workbook -WorksheetName $SheetName
        
        # Save and close the workbook after adding the worksheet
        try {
            Close-ExcelPackage -ExcelPackage $workbook
            Write-Host "Workbook saved successfully."
        } catch {
            Write-Host "Error saving workbook: $_"
        }
    }
}



$sheets = @($sheetName1, $sheetName2)
foreach ($sheetName in $sheets) {
    Ensure-WS-Exists -Path $excelFilePath -SheetName $sheetName

    # Open the workbook to reach Worksheets:
    $workbook = Open-ExcelPackage -Path $excelFilePath
    $worksheet = $workbook.Workbook.Worksheets[$sheetName]

    # Add the title to Worksheet:
    if ($worksheet -eq $null) {
        Write-Host "Failed to retrieve the worksheet '$sheetName'."
    } else {
        # Set the title in the cell [A1]:
        $worksheet.Cells.Item(1, 1).Value = "Host"
        
        # Save and close the workbook:
        try {
            Close-ExcelPackage -ExcelPackage $workbook
            Write-Host "Cell A1 in '$sheetName' is set to 'Host' and workbook saved successfully."
        } catch {
            Write-Host "Error saving workbook after setting cell A1: $_"
        }
    }
}



# Import the CSV-File and transfer its content to the Excel-File:
################################################################
$csvData = Import-Csv -Path $input_CSV_File
$workbook = Open-ExcelPackage -Path $excelFilePath


# Function to write data to a specified worksheet

function Write-Data-To-Worksheet {
    param (
        [string]$SheetName,
        [array]$Data
    )
    
    # Get the worksheet by name
    $worksheet = $workbook.Workbook.Worksheets[$SheetName]

    if ($worksheet -eq $null) {
        Write-Host "Worksheet '$SheetName' does not exist in the workbook."
        return
    }
    

    # Write data to the worksheet
    $rowIndex = 2 # Start writing from row 2 to keep row 1 for the header
    foreach ($row in $Data) {
        $colIndex = 1
        foreach ($col in $row.PSObject.Properties) {
            $worksheet.Cells[$rowIndex, $colIndex].Value = $col.Value
            $colIndex++
        }
        $rowIndex++
    }
    
    Write-Host "Data written to '$SheetName'."
}







foreach ($sheetName in $sheets) {
    Write-Data-To-Worksheet -SheetName $sheetName -Data $csvData
}

# Save and close the workbook
Close-ExcelPackage -ExcelPackage $workbook

Write-Host "Workbook updated and saved successfully."





# Import the Excel-File to check and compare the records:
################################################################

# Read_the_main_worksheet
$MainData = Import-Excel $excelFilePath -Sheet $sheetName1

# Read_the_checking_WorkSheet
$CheckList = Import-Excel $CheckListData_Address -Sheet $sheetName2

# Extract_the_list_of_values_from_the Check_Worksheet (it's a one-column list)
$listValues = $CheckList | ForEach-Object { $_.$ColumnName_in_CheckList }



# Initialize_a_hashtable_to_keep_count_of_occurrences
$matchCount = @{}


# Iterate_through_each_row_in_the_"Sheet1"_worksheet
foreach ($row in $MainData) {
    $firstColumnValue = $row.$ColumnName_in_SourceData 
    if ($null -ne $firstColumnValue -and $listValues -contains $firstColumnValue) {
        if ($matchCount.ContainsKey($firstColumnValue)) {
            $matchCount[$firstColumnValue]++
            }
        else {
            $matchCount[$firstColumnValue] = 1
            }
        }
}




# Create the output Text-file
################################################################
Set-Content -Path "$Output_Path\$OutPut_File" -Value ("------Host Affinities with more than 1 Record auf $SERVER------")

foreach ($key in $matchCount.Keys) {
    $k2=$matchCount[$key]
    if ($k2 -gt 1){
        #Write-Host "Value $key : " $k2 " time/s ---------------------------"
        Add-Content -Path "$Output_Path\$OutPut_File" -Value ("$key : $k2 time/s  ---------------------------")
        }
        #else{
        #Write-Host "Value $key : " $k2 " time/s"
        #Add-Content -Path "$Output_Path\$OutPut_File" -Value ("$key : $k2 time/s")
        #}

}



# Remove the excel file and go to the next loop
################################################################
Remove-Item -Path $excelFilePath -Force
Start-Sleep -Seconds 1

$i++
}



# Close the Excel module to finish the program
###################################################################################################################################################
Remove-Module ImportExcel




# Created_by Payam.Avarwand
# Payam_avar@yahoo.com
# 01.08.2024