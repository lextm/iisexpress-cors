#Requires -RunAsAdministrator

[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $fileName,

    [Parameter()]
    [string]
    $msiFile
)

$program32 = "${env:ProgramFiles(x86)}\IIS Express"
$program64 = "${env:ProgramFiles}\IIS Express"

$source32 = '\SysWOW64\inetsrv'
$source64 = '\System32\inetsrv'

$schema = '\config\schema\cors_schema.xml'
$module = '\iiscors.dll'
$sectionName = 'cors'
$globalModuleName = 'CorsModule'
$globalModuleFileName = "%IIS_BIN%$module"
$tempDirPattern = 'IISCORS-{0:x}\'
$useLessMsi = $true # use lessmsi to extract contents so that IIS doesn't need to be installed.

function AddSchemaFiles([string]$sourceDir) {

    $schema32 = $program32 + $schema
    $schema64 = $program64 + $schema
    $source = $sourceDir + $schema
    if (Test-Path $source) {
        Copy-Item $source -Destination $schema32
        Write-Host 'Added schema 32 bit.'

        Copy-Item $source -Destination $schema64
        Write-Host 'Added schema 64 bit.'
    } else {
        Write-Warning "Cannot find the original schema: $source"
    }
}

function AddModuleFiles([string]$sourceDir) {
    $module32 = $program32 + $module
    $module64 = $program64 + $module

    $sourceModule32 = $sourceDir + $source32 + $module
    $sourceModule64 = $sourceDir + $source64 + $module

    if (Test-Path $sourceModule32) {
        Copy-Item $sourceModule32 -Destination $module32
        Write-Host 'Added module 32 bit.'
    } else {
        Write-Warning "Cannot find module 32 bit: $sourceModule32"
    }

    if (Test-Path $sourceModule64) {
        Copy-Item $sourceModule64 -Destination $module64
        Write-Host 'Added module 64 bit.'
    } else {
        Write-Warning "Cannot find module 64 bit: $sourceModule64"
    }
}

function PatchConfigFile([string]$source) {
    if (Test-Path $source) {
        [xml]$xmlDoc = Get-Content $source
        $existing = $xmlDoc.SelectSingleNode("/configuration/configSections/sectionGroup[@name=`"system.webServer`"]/section[@name=`"$sectionName`"]")
        if ($existing) {
            Write-Host "Section $sectionName already registered."
        } else {
            $parent = $xmlDoc.SelectSingleNode('/configuration/configSections/sectionGroup[@name="system.webServer"]')
            if ($parent) {
                $newSection = $parent.section[0].Clone()
                $newSection.name = $sectionName
                $newSection.overrideModeDefault = "Allow"
                $parent.AppendChild($newSection) | Out-Null
                $xmlDoc.Save($source)
                Write-Host "Added section $sectionName."
            } else {
                Write-Warning 'Invalid config file.'
            }
        }

        $global = $xmlDoc.SelectSingleNode("/configuration/system.webServer/globalModules/add[@name=`"$globalModuleName`"]")
        if ($global) {
            Write-Host 'Global module already registered.'
        } else {
            $parent = $xmlDoc.SelectSingleNode('/configuration/system.webServer/globalModules')
            if ($parent) {
                $newModule = $parent.add[0].Clone()
                $newModule.name = $globalModuleName
                $newModule.image = $globalModuleFileName
                $parent.AppendChild($newModule) | Out-Null
                $xmlDoc.Save($source)
                Write-Host 'Added global module.'
            } else {
                Write-Warning 'Invalid config file.'
            }
        }

        $module = $xmlDoc.SelectSingleNode("/configuration/location[@path=`"`"]/system.webServer/modules/add[@name=`"$globalModuleName`"]")
        if ($module) {
            Write-Host 'Module already registered.'
        } else {
            $parent = $xmlDoc.SelectSingleNode('/configuration/location[@path=""]/system.webServer/modules')
            if ($parent) {
                $newModule = $parent.add[0].Clone()
                $newModule.name = $globalModuleName
                $newModule.lockItem = 'true'
                $parent.AppendChild($newModule) | Out-Null
                $xmlDoc.Save($source)
                Write-Host 'Added module.'
            } else {
                Write-Warning 'Invalid config file.'
            }
        }
    } else {
        Write-Warning 'Cannot find config file.'
    }
}

if ($msiFile) {
    $msiFile = Resolve-Path $msiFile
    if (!(Test-Path $msiFile)) {
        Write-Error "Cannot find MSI package $msiFile. Exit."
        exit 1
    }

    $tempPath = [System.IO.Path]::GetTempPath()
    $tempDirName = $tempDirPattern -f (Get-Random)
    $tempDirPath = Join-Path $tempPath $tempDirName
    Write-Debug "Extract files from $msiFile to $tempDirPath."
    if ($useLessMsi) {
        try {
            & lessmsi x "$msiFile" "$tempDirPath" > $null 2>&1
            if ($LASTEXITCODE -ne 0) {
                Write-Host "lessmsi command encountered an error with exit code: $LASTEXITCODE"
                exit 1
            }
        } catch {
            Write-Host "lessmsi command encountered an error: $($_.Exception.Message)"
            exit 1
        }

        $tempDirPath = Join-Path $tempDirPath 'SourceDir' # required for lessmsi
    } else {
        & msiexec /a "$msiFile" /qn TARGETDIR="$tempDirPath"
    }
    $schemaSource = Join-Path $tempDirPath 'inetsrv'
    $moduleSource = $tempDirPath
} else {
    $schemaSource = ${env:windir} + $source64
    $moduleSource = ${env:windir}
}

if ($fileName) {
    Write-Host "Configure $fileName."
    PatchConfigFile($fileName)
} else {
    Write-Host 'Configure all steps and default config file.'
    AddSchemaFiles($schemaSource)
    AddModuleFiles($moduleSource)
    PatchConfigFile([Environment]::GetFolderPath("MyDocuments") + "\IISExpress\config\applicationHost.config")
}

if ($msiFile) {
    Remove-Item $tempDirPath -Recurse
}

Write-Host 'All done.'
