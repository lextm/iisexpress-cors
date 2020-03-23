[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $fileName
)

$program32 = "${env:ProgramFiles(x86)}\IIS Express"
$program64 = "${env:ProgramFiles}\IIS Express"

$source32 = "${env:windir}\SysWOW64\inetsrv"
$source64 = "${env:windir}\System32\inetsrv"

$schema = '\config\schema\cors_schema.xml'
$module = '\iiscors.dll'

function AddSchemaFiles {

    $schema32 = $program32 + $schema
    $schema64 = $program64 + $schema
    $source = $source64 + $schema
    if (Test-Path $source) {
        Copy-Item $source -Destination $schema32
        Write-Host 'Added schema 32 bit.'

        Copy-Item $source -Destination $schema64
        Write-Host 'Added schema 64 bit.'
    } else {
        Write-Warning 'Cannot find the original schema.'
    }
}

function AddModuleFiles {
    $module32 = $program32 + $module
    $module64 = $program64 + $module

    $sourceModule32 = $source32 + $module
    $sourceModule64 = $source64 + $module

    if (Test-Path $sourceModule32) {
        Copy-Item $sourceModule32 -Destination $module32
        Write-Host 'Added module 32 bit.'
    } else {
        Write-Warning 'Cannot find module 32 bit.'
    }

    if (Test-Path $sourceModule64) {
        Copy-Item $sourceModule64 -Destination $module64
        Write-Host 'Added module 64 bit.'
    } else {
        Write-Warning 'Cannot find module 64 bit.'
    }
}

function PatchConfigFile([string]$source) {
    if (Test-Path $source) {
        [xml]$xmlDoc = Get-Content $source
        $existing = $xmlDoc.SelectSingleNode('/configuration/configSections/sectionGroup[@name="system.webServer"]/section[@name="cors"]')
        if ($existing) {
            Write-Host 'Section cors already registered.'
        } else {
            $parent = $xmlDoc.SelectSingleNode('/configuration/configSections/sectionGroup[@name="system.webServer"]')
            if ($parent) {
                $newSection = $parent.section[0].Clone()
                $newSection.name = 'cors'
                $newSection.overrideModeDefault = "Allow"
                $parent.AppendChild($newSection) | Out-Null
                $xmlDoc.Save($source)
                Write-Host 'Added section cors.'
            } else {
                Write-Warning 'Invalid config file.'
            }
        }

        $global = $xmlDoc.SelectSingleNode('/configuration/system.webServer/globalModules/add[@name="CorsModule"]')
        if ($global) {
            Write-Host 'Global module already registered.'
        } else {
            $parent = $xmlDoc.SelectSingleNode('/configuration/system.webServer/globalModules')
            if ($parent) {
                $newModule = $parent.add[0].Clone()
                $newModule.name = 'CorsModule'
                $newModule.image = '%IIS_BIN%\iiscors.dll'
                $parent.AppendChild($newModule) | Out-Null
                $xmlDoc.Save($source)
                Write-Host 'Added global module.'
            } else {
                Write-Warning 'Invalid config file.'
            }
        }

        $module = $xmlDoc.SelectSingleNode('/configuration/location[@path=""]/system.webServer/modules/add[@name="CorsModule"]')
        if ($module) {
            Write-Host 'Module already registered.'
        } else {
            $parent = $xmlDoc.SelectSingleNode('/configuration/location[@path=""]/system.webServer/modules')
            if ($parent) {
                $newModule = $parent.add[0].Clone()
                $newModule.name = 'CorsModule'
                $newModule.lockItem = 'true'
                $parent.AppendChild($newModule) | Out-Null
                $xmlDoc.Save($source)
                Write-Host 'Added module.'
            } else {
                Write-Warning 'Invalid config file.'
            }
        }
    } else {
        Write-Host 'Cannot find config file.'
    }
}

if ($fileName) {
    Write-Host "Configure $fileName."
    PatchConfigFile($fileName)
} else {
    Write-Host 'Configure all steps and default config file.'
    AddSchemaFiles
    AddModuleFiles
    PatchConfigFile([Environment]::GetFolderPath("MyDocuments") + "\IISExpress\config\applicationHost.config")
}

Write-Host 'All done.'
