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
        Write-Warning 'Cannot find the original schema.'
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

if ($msiFile) {
    $tempPath = [System.IO.Path]::GetTempPath()
    $tempDirName = 'IISCORS-{0:x}' -f (Get-Random)
    $tempDirPath = Join-Path $tempPath $tempDirName
    Start-Process msiexec "/a `"$msiFile`" /qn TARGETDIR=`"$tempDirPath`"" -Wait
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
