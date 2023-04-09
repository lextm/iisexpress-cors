#Requires -RunAsAdministrator

[CmdletBinding()]
param (
    [Parameter()]
    [string]
    $fileName
)

$program32 = "${env:ProgramFiles(x86)}\IIS Express"
$program64 = "${env:ProgramFiles}\IIS Express"

$schema = '\config\schema\cors_schema.xml'
$module = '\iiscors.dll'
$sectionName = 'cors'
$globalModuleName = 'CorsModule'

function RemoveSchemaFiles {
    $schema32 = $program32 + $schema
    $schema64 = $program64 + $schema
    if (Test-Path $schema32) {
        Remove-Item $schema32
        Write-Host 'Removed schema 32 bit.'
    }

    if (Test-Path $schema64) {
        Remove-Item $schema64
        Write-Host 'Removed schema 64 bit.'
    }
}

function RemoveModuleFiles {
    $module32 = $program32 + $module
    $module64 = $program64 + $module
    if (Test-Path $module32) {
        Remove-Item $module32
        Write-Host 'Removed module 32 bit.'
    }

    if (Test-Path $module64) {
        Remove-Item $module64
        Write-Host 'Removed module 64 bit.'
    }
}

function UnpatchConfigFile([string]$source) {
    if (Test-Path $source) {
        [xml]$xmlDoc = Get-Content $source
        $existing = $xmlDoc.SelectSingleNode("/configuration/configSections/sectionGroup[@name=`"system.webServer`"]/section[@name=`"$sectionName`"]")
        if ($existing) {
            $parent = $xmlDoc.SelectSingleNode('/configuration/configSections/sectionGroup[@name="system.webServer"]')
            if ($parent) {
                $parent.RemoveChild($existing) | Out-Null
                $xmlDoc.Save($source)
                Write-Host "Removed section $sectionName."
            } else {
                Write-Host 'Invalid config file.'
            }
        } else {
            Write-Host "Section $sectionName not registered."
        }

        $global = $xmlDoc.SelectSingleNode("/configuration/system.webServer/globalModules/add[@name=`"$globalModuleName`"]")
        if ($global) {
            $parent = $xmlDoc.SelectSingleNode('/configuration/system.webServer/globalModules')
            if ($parent) {
                $parent.RemoveChild($global) | Out-Null
                $xmlDoc.Save($source)
                Write-Host 'Removed global module.'
            } else {
                Write-Warning 'Invalid config file.'
            }
        } else {
            Write-Host 'Global module not registered.'
        }

        $module = $xmlDoc.SelectSingleNode("/configuration/location[@path=`"`"]/system.webServer/modules/add[@name=`"$globalModuleName`"]")
        if ($module) {
            $parent = $xmlDoc.SelectSingleNode('/configuration/location[@path=""]/system.webServer/modules')
            if ($parent) {
                $parent.RemoveChild($module) | Out-Null
                $xmlDoc.Save($source)
                Write-Host 'Removed module.'
            } else {
                Write-Warning 'Invalid config file.'
            }
        } else {
            Write-Host 'Module not registered.'
        }
    } else {
        Write-Host 'Cannot find config file.'
    }
}

if ($fileName) {
    Write-Host "Configure $fileName."
    UnpatchConfigFile($fileName)
} else {
    Write-Host 'Configure all steps and default config file.'
    RemoveSchemaFiles
    RemoveModuleFiles
    UnpatchConfigFile([Environment]::GetFolderPath("MyDocuments") + "\IISExpress\config\applicationHost.config")
}

Write-Host 'All done.'
