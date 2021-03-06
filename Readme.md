# PowerShell Scripts to Install/Uninstall CORS Module for IIS Express

## Install

1. [Install IIS](https://docs.microsoft.com/en-us/iis/install/installing-iis-7/installing-iis-on-windows-vista-and-windows-7).
1. Install [Microsoft CORS Module for IIS](https://www.iis.net/downloads/microsoft/iis-cors-module).
1. Install [IIS Express](https://docs.microsoft.com/en-us/iis/extensions/introduction-to-iis-express/iis-express-overview#installing-iis-express).
1. Run `install.ps1` to install CORS module to IIS Express.

## Uninstall
1. Run `uninstall.ps1` to uninstall CORS module from IIS Express.

## Notes
The current release only supports Windows x64 machines.

The scripts must be run as administrator.

`install.ps1` and `uninstall.ps1` only manipulate the default IIS Express config file in current user's My Documents folder.

If a custom config file needs to be modified, run the PowerShell scripts with arguments (like `install.ps1 -fileName custom.config`).
