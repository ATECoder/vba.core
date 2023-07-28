# ----------------------------------------------------------------------
# deploy
#
# PURPOSE: copy a release version of this top level workbook and its 
#          referenced workbooks to the bin folder for deployment.
#
# CALLING SCRIPT:
#
#  ."deploy.ps1"
#
# ----------------------------------------------------------------------

# ----------------------------------------------------------------------
# VARIABLES

$CWD = (Resolve-Path .\).Path

# create the bin directory if new 
$BUILD_DIRECTORY = [IO.Path]::Combine($CWD, "..\..\bin")
$BUILD_DIRECTORY = (Resolve-Path $BUILD_DIRECTORY).Path
MkDir -Force $BUILD_DIRECTORY > $null

# create the build directory if new 
$BUILD_DIRECTORY = [IO.Path]::Combine($BUILD_DIRECTORY, "demo")
MkDir -Force $BUILD_DIRECTORY > $null

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

function AwaitKeyPress()
{
    # this does not work: getting an exception
    # Exception calling "ReadKey" with "1" argument(s): "Cannot read keys when either application does not have a console or when console input has been redirected from a 
    # file. Try Console.Read."
    loginfo( "Press any key" )
	do{ $x = [console]::ReadKey() } while( $x.Key -ne "" )	
}

# -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  
# Summary:  Copies a specified file to the build directory.
#
# Parameters:
# 
# SOURCE           - the path of the source file
# BUILD_DIRECTORY  - the path of the build directory
# -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  -  
function CopyToBuildDirectory( $sourcePath )
{

	Try { 

        $path = (Resolve-Path $sourcePath).Path

		LogInfo( "coping " + $path + " to " + $BUILD_DIRECTORY )
		copy-item $path -destination $BUILD_DIRECTORY
		return $true

	}
	Catch {

		LogError( $_.Exception.Message )
        $z = Read-Host "Press enter to exit: "        
		return $false
	}

}

# END FUNCTIONS
# ----------------------------------------------------------------------


# ----------------------------------------------------------------------
# SCRIPT ENTRY POINT


# Copy all workbooks to the build directory

$src = [IO.Path]::Combine($CWD, "..\io\cc.isr.core.io.xlsm")
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit 1 }

$src = [IO.Path]::Combine($CWD, "..\io\readme.md")
$dest = [IO.Path]::Combine($BUILD_DIRECTORY, "cc.isr.core.io.readme.md") 
LogInfo( "coping " + $src + " to " + $dest )
copy-item $src -Destination $dest

$src = [IO.Path]::Combine($CWD, "..\core\cc.isr.core.xlsm")
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit 1 }

$src = [IO.Path]::Combine($CWD, "..\core\readme.md")
$dest = [IO.Path]::Combine($BUILD_DIRECTORY, "cc.isr.core.readme.md") 
LogInfo( "coping " + $src + " to " + $dest )
copy-item $src -Destination $dest

$src = [IO.Path]::Combine($CWD, "cc.isr.core.demo.xlsm")
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit 1 }

$src = [IO.Path]::Combine($CWD, "cc.isr.core.demo.testing.md")
if ( -Not( CopyToBuildDirectory ( $src  ) ) ) { exit 1 }

$src = [IO.Path]::Combine($CWD, "readme.md")
$dest = [IO.Path]::Combine($BUILD_DIRECTORY, "cc.isr.core.demo.readme.md") 
LogInfo( "coping " + $src + " to " + $dest )
copy-item $src -Destination $dest

LogInfo( "project deployed" )
$z = Read-Host "Press enter to exit"
exit 0
