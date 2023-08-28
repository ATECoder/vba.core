# ----------------------------------------------------------------------
# Open workbook
#
# PURPOSE: Opens a workbook with little flickers.
#
# CALLING SCRIPT:
#
# PowerShell.EXE -file <script> -RelativePath <relative-path> -BookName <excel-file-name>
#
# EXAMPLE:
#
# PowerShell.EXE -file open.workbook.ps1 -RelativePath ".\" -BookName "cc.isr.core.xlsm"
#
# ----------------------------------------------------------------------

# Must be the first statement in your script (not counting comments)
# 
param(

    # relative path
    [string]$RelativePath = ".\",

    [string]$BookName = "cc.isr.test.fx.xlsm"
) 


# ----------------------------------------------------------------------
# FUNCTIONS

Function LogInfo($message)
{
    Write-Host $message -ForegroundColor Gray
}

Function LogEmptyLine()
{
    echo ""
}

Function HasSuiteFlag($flags, $flag)
{
    Return ($flags -band $flag) -eq $flag
}

# END FUNCTIONS
# ----------------------------------------------------------------------


# ----------------------------------------------------------------------
# SCRIPT ENTRY POINT


$DEBUG = $true

if ( $DEBUG ) { LogInfo ( "Relative Path: '" + $RelativePath + "', Book Name: '" + $BookName + "'"  ) }

# Build paths

Try {

    $CWD = (Resolve-Path .\).Path

    $BUILD_DIRECTORY = [IO.Path]::Combine($CWD, $RelativePath)
    # $BUILD_DIRECTORY = $CWD
    $BUILD_DIRECTORY = (Resolve-Path $BUILD_DIRECTORY).Path

    $FILENAME = [IO.Path]::Combine($BUILD_DIRECTORY, $BookName)

} Catch {

    echo $_.Exception.Message
    LogEmptyLine
    $z = Read-Host "Press enter to exit"
    exit -1
    return
}

LogInfo ( "Collecting tests " + $FILENAME )

$missing = [System.Reflection.Missing]::Value

# start Excel
$excel = New-Object -ComObject Excel.Application

$book = $excel.Workbooks.Open($FILENAME, $missing, $true)

IF ( $DEBUG ) { LogInfo ( "Opened " + $book.Name ) }  

# select the active workbook
$book = $excel.ActiveWorkbook

IF ( $DEBUG ) { LogInfo ( "Active workbook " + $book.Name ) }  

LogInfo( "project loaded. Script will close in 5 seconds" )
Start-Sleep -Seconds 5
# $z = Read-Host "Press enter to exit"

exit 0

