Install-Module -Name ImportExcel -Scope CurrentUser
# Import the ImportExcel module
Import-Module ImportExcel

# Set the path to your Excel file
$excelFilePath = "C:\users\kangah\Downloads\Copy of Copy of DABP  DPPS Staffing List as of 3-16-23.xlsx"

# Import the data from the Excel file
$excelData = Import-Excel -Path $excelFilePath

# Retrieve the 3rd column data
$thirdColumnData = $excelData | Select-Object -Property "Employee Name"

# Output the 3rd column data
$thirdColumnData

$outputData = $thirdColumnData
# Remove the header lines and split the output into an array of lines
$employeeLines = ($outputData -split "`n" | Select-Object -Skip 2).Trim()
$employees = $employeeLines | ForEach-Object {
    # Split each line by comma and trim whitespace
    $splitData = $_ -split ', ' | ForEach-Object { $_.Trim() }

    # Create a custom object with separated lastname and firstname
    [PSCustomObject]@{
        Lastname  = $splitData[0]
        Firstname = $splitData[1]
    }
}

# Output the employees with separated lastname and firstname
$employees


$fullnames = $thirdColumnData
foreach($name in $fullnames){
    $splitname = $name -split ', ' |ForEach-Object {$_.Trim()}

    $surname = $splitname[0]
    $givenname = $splitname[1]
    $user = Get-ADUser -Filter{(givenname -eq $givenname) -and (surname -eq $surname)}
    $user.SamAccountName
}
