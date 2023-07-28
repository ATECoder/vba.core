# ----------------------------------------------------------------------
# Localize
#
# PURPOSE: open and save all workbooks from the bin folder thus localizing their references.
#
# CALLING SCRIPT:
#
#  ."Localize.ps1"
#
# ----------------------------------------------------------------------

# ----------------------------------------------------------------------
# VARIABLES

$CWD = (Resolve-Path .\).Path
$BUILD_DIRECTORY = [IO.Path]::Combine($CWD, "..\..\bin\demo")
$BUILD_DIRECTORY = (Resolve-Path $BUILD_DIRECTORY).Path
$XL_FILE_FORMAT_MACRO_ENABLED = 52

# END VARIABLES
# ----------------------------------------------------------------------

# ----------------------------------------------------------------------
# FUNCTIONS

Function LogInfo($message)
{
    Write-Host $message -ForegroundColor Gray
}

Function LogError($message)
{
    Write-Host $message -ForegroundColor Red
}

Function LogEmptyLine()
{
    echo ""
}

# END FUNCTIONS
# ----------------------------------------------------------------------


# ----------------------------------------------------------------------
# SCRIPT ENTRY POINT

$DEBUG = $true

# declare Excel
$excel = New-Object -ComObject Excel.Application
$excel.DisplayAlerts = $false;

$missing = [System.Reflection.Missing]::Value

$src = [IO.Path]::Combine($BUILD_DIRECTORY, "cc.isr.core.io.xlsm")
LogInfo( "opening " + $src)
$book = $excel.Workbooks.Open($src, $missing, $false, $missing, $missing, $missing, $true)
$io_book = $excel.ActiveWorkbook
LogInfo ( "Opened " + $io_book.Name )

$src = [IO.Path]::Combine($BUILD_DIRECTORY, "cc.isr.core.xlsm") 
LogInfo( "opening " + $src)
$book = $excel.Workbooks.Open($src, $missing, $false, $missing, $missing, $missing, $true)
$core_book = $excel.ActiveWorkbook
LogInfo ( "Opened " + $core_book.Name )

$src = [IO.Path]::Combine($BUILD_DIRECTORY, "cc.isr.core.demo.xlsm") 
LogInfo( "opening " + $src)
$book = $excel.Workbooks.Open($src, $missing, $false, $missing, $missing, $missing, $true)
$demo_book = $excel.ActiveWorkbook
LogInfo ( "Opened " + $demo_book.Name )

LogInfo ( "saving and closing " + $demo_book.Name )
$demo_book.Close($true)

LogInfo ( "saving and closing " + $core_book.Name )
$core_book.Close($true)

LogInfo ( "saving and closing " + $io_book.Name )
$io_book.Close($true)

LogInfo( "project localized" )
$z = Read-Host "Press enter to exit"

$excel.Quit()

# https://stackoverflow.com/questions/27798567/excel-save-and-close-after-run
if ( $DEBUG ) { LogInfo "finalize." }
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

if ( $DEBUG ) { LogInfo "Release COM Objects." }
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($book) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($io_book) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($core_book) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($demo_book) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# LogInfo "Disposing Excel."
Remove-Variable -Name excel;

exit 0

