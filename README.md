# Get-ExcelWorkSheet
This is a simple Powershell script to get an exel worksheet and convert it to a powershell array

# DESCRIPTION
The script will get an xlsx file from the parameters or open a file dialouge if none is provided.
It will then get the worksheet from the parameters or prompt the user if none is provided.
It will then get the header alignment from the parameters or prompt the user if none is provided.

# EXAMPLE
./Get-ExcelWorkSheet.ps1
./Get-ExcelWorkSheet.ps1 -initialDirectory "C:\"
./Get-ExcelWorkSheet.ps1 -File "C:\example.xlsx" -WorkSheet "Sheet1" -Header "Row"

# NOTES
  Author:   Leigh Butterworth
  Version:  1.0

# LINK
https://github.com/L37hal/Get-ExcelWorkSheet
