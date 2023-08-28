# ----------------------------------------------------------------------
# Localize
#
# PURPOSE: open and save all workbooks from the bin folder thus localizing their references.
#
# CALLING SCRIPT:
#
#  ."open.workbook.ps1"
#
# ----------------------------------------------------------------------

# ----------------------------------------------------------------------
# VARIABLES

$CWD = (Resolve-Path .\).Path
$Bdir = $CWD
$Bdir = (Resolve-Path $Bdir).Path
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
$excel.EnableEvents = $false;

$missing = [System.Reflection.Missing]::Value

$UpdateLinks = $missing
$ReadOnly = $true
$Format = $missing
$Password = $missing
$WriteReservedPassword = $missing
$IgnoreReadOnlyDisplay = $true

$src = "C:\my\lib\vba\core\core\src\io\cc.isr.core.io.xlsm"
LogInfo( "opening " + $src)
$book = $excel.Workbooks.Open($src, $missing, $true, $missing, $missing, $missing, $true)
LogInfo ( "Opened " + $excel.ActiveWorkbook.Name )

$src = "C:\my\lib\vba\core\core\src\core\cc.isr.core.xlsm"
LogInfo( "opening " + $src)
$book = $excel.Workbooks.Open($src, $missing, $true, $missing, $missing, $missing, $true)
LogInfo ( "Opened " + $excel.ActiveWorkbook.Name )

$excel.EnableEvents = $true;

$src = "C:\my\lib\vba\core\core\src\test\cc.isr.core.test.xlsm"
LogInfo( "opening " + $src)
$book = $excel.Workbooks.Open($src, $missing, $false, $missing, $missing, $missing, $true)
LogInfo ( "Opened " + $excel.ActiveWorkbook.Name )

LogInfo( "project loaded. Script will close in 5 seconds" )
Start-Sleep -Seconds 5
# $z = Read-Host "Press enter to exit"

exit 0
