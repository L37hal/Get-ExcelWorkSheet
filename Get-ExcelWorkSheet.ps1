<#

.SYNOPSIS
This is a simple Powershell script to get an exel worksheet and convert it to a powershell array

.DESCRIPTION
The script will get an xlsx file from the parameters or open a file dialouge if none is provided.
It will then get the worksheet from the parameters or prompt the user if none is provided.
It will then get the header alignment from the parameters or prompt the user if none is provided.

.EXAMPLE
.\Convert-ExcelToWord.ps1
Please select the worksheet: 
1) Sheet1 
2) Sheet2 
: 1
Please select the Header: 
1) Row 1
2) Column A
: 1
header1 header2 header3
------- ------- -------
data1   data2   data3  
data 1  data 2  data 3 
.EXAMPLE
.\Get-ExcelWorkSheet.ps1 -initialDirectory "C:\"
Please select the worksheet: 
1) Sheet1 
2) Sheet2 
: 1
Please select the Header: 
1) Row 1
2) Column A
: 1
header1 header2 header3
------- ------- -------
data1   data2   data3  
data 1  data 2  data 3 
.EXAMPLE
.\Convert-ExcelToWord.ps1 -File ".\testfile.xlsx" -WorkSheet "Sheet1" -Header "Row"
header1 header2 header3
------- ------- -------
data1   data2   data3  
data 1  data 2  data 3 

.NOTES
  Author:   Leigh Butterworth
  Version:  1.0

.LINK
https://github.com/L37hal/Get-ExcelWorkSheet

#>

Param(
    [parameter(Mandatory=$false)][string]$initialDirectory = "C:\",
    [parameter(Mandatory=$false)][string]$File,
    [parameter(Mandatory=$false)][string]$WorkSheet,
    [parameter(Mandatory=$false)][string]$Header
) # End Param()

Set-Location $PSScriptRoot

Function Get-WorkBook($initialDirectory)
{  
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "All files (*.*)| *.*"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
} # end function Get-WorkBook

Function Get-WorkSheets()
{
 for ($i = 1; $i -le $WorkBook.Sheets.Count; $i++)
 {
  $WorkSheet = $WorkBook.Sheets.Item($i)
  $WorkSheet.name
 }
} # end function Get-WorkSheets

Function Get-WorkSheet()
{
 Clear-Host
 $text = "Please select the worksheet: `n"
 $text += "`n"
 for ($i = 0; $i -le $WorkSheets.Count-1; $i++)
 {
  $WorkSheet = $WorkSheets.Item($i)
  $num = $i + 1
  $text += "$num) $WorkSheet `n"
 }
 $text += "`n"
 $Selection = ((Read-Host $text))-1
 $WorkSheet = $WorkSheets.GetValue($Selection)
 $WorkSheet
} # end function Get-WorkSheet

Function Get-Header()
{
 Clear-Host
 $text = "Please select the Header: `n"
 $text += "`n"
 $text += "1) Row 1`n"
 $text += "2) Column A`n"
 $text += "`n"
 $Header = Read-Host $text
 if ($Header -eq 1)
 {
  $Header = "Row"
 }
 ElseIf ($Header -eq 2)
 {
  $Header = "Col"
 }
 $Header
} # end function Get-Header


# *** Entry Point to Script ***

# If the user hasn't included a -File
if (!$File) 
{
 $File = Get-WorkBook -initialDirectory $initialDirectory
}

# Create the excel process
$Excel = New-Object -ComObject Excel.Application
# Make it invisible
$Excel.Visible = $False
# Open the WorkBook
$WorkBook = $Excel.WorkBooks.open($File)

# If the user hasn't included a -WorkSheet
if (!$WorkSheet)
{
 # Get the WorkSheets in the WorkBook
 $WorkSheets = Get-WorkSheets
 # Ask the user for a WorkSheet and get its name
 $WorkSheet = Get-WorkSheet
}
# Set the worksheet as a variable
$WorkSheet = $WorkBook.sheets.item($WorkSheet)

# Get the data width and height
$colMax = ($WorkSheet.UsedRange.Rows).count 
$rowMax = ($WorkSheet.UsedRange.Columns).count

# If the user hasn't included a -Header
if (!$Header)
{
 $Header = Get-Header
}

# Create the array Object
$data = @()            

# If the Header is the first Row
if ($Header -eq "Row") 
{
 # For each row
 For ($e = 2; $e -le $colMax; $e++)
 {
  # Create a Custom PS Object to store the Row data
  $row = New-Object -TypeName PSObject
  # For the header in each Column
  For ($i = 1; $i -le $rowMax; $i++)
  {
   $row | Add-Member -Type NoteProperty -Name $WorkSheet.Cells.Item(1, $i).Value2 -Value $WorkSheet.Cells.Item($e, $i).Value2
  }
  $data += $row
 }
}
# If the Header is the first Column
ElseIf ($Header -eq "Col")
{
 # For each Column
 For ($e = 2; $e -le $rowMax; $e++)
 {
  # Create a Custom PS Object to store the Column data
  $col = New-Object -TypeName PSObject
  # For the header in each Row
  For ($i = 1; $i -le $colMax; $i++)
  {
   # Add the Column data
   $col | Add-Member -Type NoteProperty -Name $WorkSheet.Cells.Item($i, 1).Value2 -Value $WorkSheet.Cells.Item($i, $e).Value2
  }
  # Add the Column data to the Powershell array
  $data += $col
 }
}

# Close Excel
$Excel.Workbooks.Close()
$Excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
Remove-Variable Excel

# Return the data
$data