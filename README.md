## TeslaTokens

This VB script helps you get your TESLA access_token and refresh_token in order to connect to third party applications (Teslalogger, Teslamate, TeslaFi, ABRP...) without providing them your TESLA credentials.
The script does not require your TESLA credentials. It will redirect you to TESLA site to authenticate as always (so only TESLA knowns your credentials).

## Prerequisites
 
- Microsoft Windows
- On Windows 64 bit: Chilkat 64-bit ActiveX MSI Installer: http://www.chilkatsoft.com/downloads_ActiveX.asp
- On Windows 32 bit: Chilkat 32-bit ActiveX MSI Installer: http://www.chilkatsoft.com/downloads_ActiveX.asp

## Execution

- Download and install Chilkat ActiveX control

- Download the vbs file for english and the localisation file for your language.

- Double click on the VB script

The script guides you through the process. First, it opens the TESLA login page where you have to logon to your TESLA account.

Then a "Page Not Found" error appears, but that's ok. You have to copy the URL from your browser to the input box of this script.

Then your tokens are written to the file TeslaTokens.txt in the same folder where the VB script is.
