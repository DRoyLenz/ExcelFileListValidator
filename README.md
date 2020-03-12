# ExcelFileListValidator
Compares a list of filenames in an Excel spreadsheet against the corresponding files on the filesystem


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
