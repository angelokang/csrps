Install-Module -Name ImportExcel -Scope CurrentUser


# Set the path to your Excel file
$excelFilePath = "C:\users\kangah\Downloads\Copy of Copy of DABP  DPPS Staffing List as of 3-16-23.xlsx"

# Set the name or index of the sheet to import data from (zero-based index)
#$sheetIndex = 0
$sheetName = "DPPS" # Replace with the name of the 2nd sheet
# Import the data from the 2nd sheet of the Excel file
#$excelData = Import-Excel -Path $excelFilePath -WorksheetIndex $sheetIndex
$excelData = Import-Excel -Path $excelFilePath -WorksheetName $sheetName

# Retrieve the 3rd column data
$thirdColumnIndex = 2 # Index of the 3rd column (zero-based)
$thirdColumnName = ($excelData | Get-Member -MemberType NoteProperty)[$thirdColumnIndex].Name
$thirdColumnData = $excelData | Select-Object -Property $thirdColumnName

# Output the 3rd column data
$thirdColumnData
