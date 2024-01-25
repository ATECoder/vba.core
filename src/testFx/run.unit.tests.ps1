# ----------------------------------------------------------------------
# Run Unit Tests
#
# PURPOSE: Executes the unit tests defined in the provided workbook.
#
# CALLING SCRIPT:
#
# PowerShell.EXE -file <script> -RelativePath <relative-path> -BookName <excel-file-name>
#
# EXAMPLE:
#
# PowerShell.EXE -file cc.isr.core.test.ps1 -RelativePath ".\" -BookName "cc.isr.core.test.xlsm"
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

Function LogSummary($title, $passed, $failed, $inconclusive)
{
    Write-Host $title -NoNewline
    Write-Host "$passed passed" -ForegroundColor Green -NoNewline

    if ($failed -gt 0)
    {
        Write-Host ", " -NoNewline
        Write-Host "${failed} failed" -ForegroundColor Red -NoNewline
    }

    if ($inconclusive -gt 0)
    {
        Write-Host ", " -NoNewline
        Write-Host "${$inconclusive} inconclusive" -ForegroundColor Yellow -NoNewline
    }

    $total = $passed + $failed + $inconclusive
    Write-Host ", $total total"
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

$modules = $book.VBProject.VBComponents;
$suites = @{}
$suiteFlags = @{}

$SUITE_FLAG_BEFORE_ALL = 1
$SUITE_FLAG_AFTER_ALL = 2
$SUITE_FLAG_BEFORE_EACH = 4
$SUITE_FLAG_AFTER_EACH = 8

$testCount = 0

For ($moduleIndex = 0; $moduleIndex -lt $modules.Count; $moduleIndex++)
{
    # vba interop seems to be written in VB6 => indices start at 1
    $module = $modules.Item($moduleIndex + 1)

    If (!$module.Name.EndsWith("Tests"))
    {
        Continue
    }

    $code = $module.CodeModule.Lines(1, $module.CodeModule.CountOfLines)
    $lines = $code.Split("`r`n")

    $suiteFlags[$module.Name] = 0

    ForEach ($line in $lines)
    {
        If ($line.StartsWith("Public Function Test"))
        {
            If (!$suites.ContainsKey($module.Name))
            {
                $suites[$module.Name] = [System.Collections.ArrayList]@()
            }

            $testName = $line.Split(" ")[2].Trim("()")
            $_ = $suites[$module.Name].Add($testName)
            $testCount += 1
        }

        If ($line.StartsWith("Public Sub BeforeAll()"))
        {
            $suiteFlags[$module.Name] = $suiteFlags[$module.Name] -bor $SUITE_FLAG_BEFORE_ALL
        }

        If ($line.StartsWith("Public Sub AfterAll()"))
        {
            $suiteFlags[$module.Name] = $suiteFlags[$module.Name] -bor $SUITE_FLAG_AFTER_ALL
        }

        If ($line.StartsWith("Public Sub BeforeEach()"))
        {
            $suiteFlags[$module.Name] = $suiteFlags[$module.Name] -bor $SUITE_FLAG_BEFORE_EACH
        }

        If ($line.StartsWith("Public Sub AfterEach()"))
        {
            $suiteFlags[$module.Name] = $suiteFlags[$module.Name] -bor $SUITE_FLAG_AFTER_EACH
        }
    }
}

LogInfo "Found $($suites.Count) suites with $testCount tests in $($book.Name)"
LogEmptyLine

$passedSuites = 0
$passedTests = 0
$failedTests = 0
$inconclusiveTests = 0

ForEach ($suite in $suites.Keys)
{
    $successful = $true
    $tests = $suites[$suite]
    $flags = $suiteFlags[$suite]

    $title = $suite + " (" + $tests.Count + " tests)"
    echo $title

    $hasBeforeAll = HasSuiteFlag $flags $SUITE_FLAG_BEFORE_ALL
    $hasAfterAll = HasSuiteFlag $flags $SUITE_FLAG_AFTER_ALL
    $hasBeforeEach = HasSuiteFlag $flags $SUITE_FLAG_BEFORE_EACH
    $hasAfterEach = HasSuiteFlag $flags $SUITE_FLAG_AFTER_EACH

    If ($hasBeforeAll)
    {
        $excel.Run($suite + "." + "BeforeAll")
    }

    ForEach ($test in $tests)
    {
        If ($hasBeforeEach)
        {
            $excel.Run($suite + "." + "BeforeEach")
        }

        $result = $excel.Run($suite + "." + $test)
		
        Write-Host "  " -NoNewline

        if ($result.AssertInconclusive)
        {
            $inconclusiveTests += 1
            Write-Host " MOOT " -BackgroundColor Brown -ForegroundColor White -NoNewline
        }
        elseif ($result.AssertSuccessful)
        {
            $passedTests += 1
            Write-Host " PASS " -BackgroundColor Green -ForegroundColor White -NoNewline
        }
        else
        {
            $failedTests += 1
            $successful = $false
            Write-Host " FAIL " -BackgroundColor Red -ForegroundColor White -NoNewline
        }

        Write-Host " $test" -NoNewline

        if ($result.AssertSuccessful)
        {
            Write-Host ": $($result.AssertMessage)" -ForegroundColor Gray
        }
        else
        {
            LogEmptyLine
            LogEmptyLine
            if ($result.AssertInconclusive)
            {
                Write-Host $result.AssertMessage -ForegroundColor white
            }
            else
            {
                Write-Host $result.AssertMessage -ForegroundColor Red
            }
            LogEmptyLine
        }

        If ($hasAfterEach)
        {
            $excel.Run($suite + "." + "AfterEach")
        }
    }

    If ($hasAfterAll)
    {
        $excel.Run($suite + "." + "AfterAll")
    }

    if ($successful)
    {
        $passedSuites += 1
    }

    echo ""
}

{ LogInfo "Closing the active workbook." }
$book.Close()

if ( $DEBUG ) { LogInfo "Closing workbooks." }
$excel.Workbooks.Close()

if ( $DEBUG ) { LogInfo "Quiting Excel." }
$excel.Quit()

if ( $DEBUG ) { LogEmptyLine }

LogSummary "Test suites: " $passedSuites ($suites.Count - $passedSuites)
LogSummary "Tests:       " $passedTests $failedTests $inconclusiveTests
LogEmptyLine
if ( $DEBUG ) { LogInfo "Ran all test suites." }
LogEmptyLine
$z = Read-Host "Press enter to exit"

# https://stackoverflow.com/questions/27798567/excel-save-and-close-after-run
if ( $DEBUG ) { LogInfo "finalize." }
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

if ( $DEBUG ) { LogInfo "Release COM Objects." }
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($modules) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($book) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# LogInfo "Disposing Excel."
Remove-Variable -Name excel;

exit 0

