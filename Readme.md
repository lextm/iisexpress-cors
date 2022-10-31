# PowerShell Scripts to Install/Uninstall CORS Module for IIS Express

## Install

1. Install [IIS Express](https://docs.microsoft.com/en-us/iis/extensions/introduction-to-iis-express/iis-express-overview#installing-iis-express).
1. Download [Microsoft CORS Module for IIS](https://www.iis.net/downloads/microsoft/iis-cors-module) installer.
1. Run the proper command to install CORS module to IIS Express,

   > If the CORS module for IIS has been installed on this same machine in advance, then `-msiFile` switch can be omitted.

   * At command prompt
   ``` cmd
   powershell.exe -file install.ps1 -msiFile IISCORS_amd64.msi
   ```
   * In PowerShell console   
   ``` powershell
   install.ps1 -msiFile IISCORS_amd64.msi
   ```
The configuration is exactly the same as IIS, and just make sure you are editing the right files,

https://learn.microsoft.com/en-us/iis/extensions/cors-module/cors-module-configuration-reference

## Uninstall
1. Run `uninstall.ps1` to uninstall CORS module from IIS Express.

## Notes
The current release only supports Windows x64 machines.

The scripts must be run as administrator.

`install.ps1` and `uninstall.ps1` only manipulate the default IIS Express config file in current user's My Documents folder.

If a custom config file needs to be modified, run the PowerShell scripts with arguments (like `install.ps1 -fileName custom.config`).
