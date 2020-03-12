<#
.SYNOPSIS
    This script checks a list (Excel Document) of files to see if those files exist in a given
    folder.

.DESCRIPTION
    This script allows the input of an excel document that has a list of filenames in one column,
    as well as a directory that SHOULD have those same files. It will iterate through the Excel
    document to find every filename in that column, then search the directory to ensure that file
    exists. It also adds a column next to the list of filenames and flags it TRUE if the file exists
    in the directory, and FALSE if it does not.   

.NOTES
    File Name          : ExcelFileListValidator.ps1
    Author             : David R. Lenz (david.r.lenz@protonmail.com)
    Prerequisite       : Powershell V2, Microsoft Excel
    Date Created       : March 10th, 2020

.VARIABLES_TO_DEFINE
    $FileListPath
        This is to be the fully-qualified path to the Excel Document you wish to query.
        If this is left blank, the script will automatically use the current working directory.

        Example:  $FileListPath = ""
        Example2: $FileListPath = "C:\Users\DavidLenz\Documents\"

    $FileListFile
        The Filename for the Excel Document to be queried. This file must exist in the path defined
        in $FileListPath.

        Example:  $FileListFile = "TestFileList2.xlsx"

    $FileListSheetName
        This is the name of the Worksheet in the Excel Document to be queried. If this is left blank,
        it will be defaulted to "Sheet1"

        Example:  $FileListSheetName = ""
        Example2: $FileListSheetName = "DocumentLocations"

    $FileListColumnName
        This is the column header (the value in Row 1) of the column that contains the filenames.
        This script assumes each worksheet will have a column header, so if there isn't one, please
        create a dummy one in the excel file, and put the value of that column header here.

        Example:  $FileListColumnName = "filenames"
        Example2: $FileListColumnName = "dummy"

    $FilePath
        This is the fully-qualified path that contains all of the files you want to validate. If this
        is left blank, it will then be defaulted to the current working directory

        Example:  $FilePath = ""
        Example2: $FilePath = "C:\Users\DavidLenz\Documents\"

.OUTPUT
    The output of this script is a file with the following naming convention:
        ExcelFileListValidator_YYYYMMDD-hhmmss_out.xlsx

#>

# USER DEFINED VARIABLES HERE, ONLY CHANGE THESE VALUES
$FileListPath = ""
$FileListFile = "TestFileListLong.xlsx"
$FileListSheetName = "Documents"
$FileListColumnName = "Doc_Filename"

$FilePath = ""
# END USER DEFINED VARIABLES

$NewColumnName = 'FileExists'

# Get DateTime stamp for the output filename
$StartDateTime = Get-Date -Format yyyymmdd-HHmmss

# Define Output Filename
$OutputFile = "ExcelFileListValidator_" + $StartDateTime + "_out.xlsx"

# When user doesn't define the location of the FileList, use the current working directory
if($FileListPath -eq "")
    {
        $FileListPath = Get-Location | Select-Object -ExpandProperty Path
        $FileListPath = $FileListPath + "\"
    }

# When user doesn't define the location of the directory of files to check, use the current working directory
if($FilePath -eq "")
    {
        $FilePath = Get-Location | Select-Object -ExpandProperty Path
        $FilePath = $FilePath + "\"
    }

# When user doesn't define the Excel Worksheet name, default to 'Sheet1'
if($FileListSheetName -eq "")
    {
        $FileListSheetName = "Sheet1"
    }

Write-Host "Opening Excel Directory       " -ForegroundColor Green -NoNewLine
Write-Host $FileListPath -ForegroundColor Yellow
Write-Host "Opening Excel Document        " -ForegroundColor Green -NoNewLine
Write-Host $FileListFile -ForegroundColor Yellow

# Start Excel / Define Excel Object
Try
{
    $xls = New-Object -ComObject Excel.Application -ErrorAction Stop
}
Catch
{
    Write-Host "Failed to start Microsoft Excel. Please validate you currently have a licensed copy installed on this machine" -ForegroundColor Red
}

# Check to see if Excel Document exists at that file location, stores boolean result of that check to $IfExcelDocExists
$FQFilename = $FileListPath + $FileListFile
$IfExcelDocExists = Test-Path -LiteralPath $FQFilename -PathType Leaf

# Open Excel Document
If ($IfExcelDocExists)
{
    $wb = $xls.Workbooks.Open($FileListPath + $FileListFile)
    Write-Host "Excel Document Opened" -ForegroundColor Green
}
Else
{
    Write-Host "Excel Document failed to open. Ensure the file path and name are correct, and you do not have the document open in another instance." -ForegroundColor Red
    Write-Host "Terminating script with errors" -ForegroundColor Red
    $xls.Quit()
    Exit
}

# Open Excel Worksheet

Write-Host "Opening Worksheet             " -ForegroundColor Green -NoNewLine
Write-Host $FileListSheetName -ForegroundColor Yellow

$IfWorksheetExists = $wb.Worksheets | where {$_.name -eq $FileListSheetName}
if($IfWorksheetExists)
{
    $ws = $wb.Sheets.Item($FileListSheetName)
    Write-Host "Worksheet Opened" -ForegroundColor Green
}
else
{
    Write-Host "Worksheet does not exist in Workbook" -ForegroundColor Red
    Write-Host "Terminating script with errors" -ForegroundColor Red
    $xls.Quit()
    Exit
}

# Below commented line outputs name of worksheet to validate it was found in the excel workbook
# Uncomment for debugging

#$ws.Name | Write-Host

Write-Host "Finding Column in Worksheet   " -ForegroundColor Green -NoNewLine
Write-Host $FileListColumnName -ForegroundColor Yellow

# Using the $FileListColumn variable, locate the column that contains the filenames
$Range = $ws.Range("A1:Z1")
$ColumnRange = $Range.find($FileListColumnName)

if(!$ColumnRange)
{
    Write-Host "Column does not exist in Workbook" -ForegroundColor Red
    Write-Host "Terminating script with errors" -ForegroundColor Red
    $xls.Quit()
    Exit
}
else
{
    Write-Host "Column Found" -ForegroundColor Green

    $ColumnNum = $ColumnRange.Column

    # Below commented line outputs the column index number to validate it was found in the excel worksheet
    # Uncomment for debugging
    #$ColumnNum | Write-Host

    # Insert new column to the right of the filename column and name the column
    $NewColumn = $ColumnNum + 1
    $NextColumnRange = $ws.cells.item(1,$NewColumn).EntireColumn
    $NextColumnRange.Insert(-4161) | Out-Null
    $ws.Cells.Item(1,$NewColumn) = $NewColumnName

    Write-Host "New Column Added              " -ForegroundColor Green -NoNewline
    Write-Host "$NewColumnName" -ForegroundColor Yellow

    # Find the number of rows in the worksheet
    $RowCount = $ws.UsedRange.Rows.Count

    
    Write-Host "Beginning Filename Validation " -ForegroundColor Green -NoNewline

    # Begin the loop that goes through each row in the column
    foreach ( $RowNum in 2..$RowCount )
    {
        $Progress = [int](($RowNum/$RowCount)*100)
        $CurrentOperation = [string]$Progress+'% Row '+$RowNum+"/"+$RowCount
        #Write-Host $Progress
        Write-Progress -Activity "Validating Filenames" -Status "Progress:" -PercentComplete $Progress -CurrentOperation $CurrentOperation
        # Get the value stored in the iterative cell
        $Filename = $ws.Cells.Item($RowNum, $ColumnNum).Value()
        # Get the fully-qualified pathname for the filename in that cell
        $FQFilename = $FilePath + $Filename

        # Below commented out code-block used for debugging. Will output the name of the filename for all matches
        <#if( Test-Path -LiteralPath $FQFilename -PathType Leaf )
        {
            $Filename | Write-Host
        }#>

        # Check to see if the filename exists, and store the boolean result of that check to $IfExists
        $IfExists = Test-Path -LiteralPath $FQFilename -PathType Leaf
        # Write $IfExists to the newly inserted cell to the right
        $ws.Cells.Item($RowNum,$NewColumn) = $IfExists
    }

    Write-Host "Completed" -ForegroundColor Yellow
    
    Write-Host "Saving New Excel Document     " -ForegroundColor Green -NoNewline
    Write-Host "$OutputFile" -ForegroundColor Yellow

    # Save the Excel Document with the new column
    $ws.SaveAs($FileListPath + $OutputFile)

    Write-Host "New Excel Document Saved" -ForegroundColor Green
    Write-Host "Process Completed - No Errors" -ForegroundColor Green
    

    # Quit this Excel Object
    $xls.Quit()

    # Clean up the variables
    Remove-Variable xls
}