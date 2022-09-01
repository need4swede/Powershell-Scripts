# Print Server Manager

PS script that lets you query all printers in a print server and output them in a simple GUI for easy installs.
Great for home users with advanced printer setups or IT Admins.

I was able to compile this into an executable and host it on a fileshare and distribute it to our org via 
GPO and our users have been using it to install and manage their printers. 

## Functions

* Fetches a list of all printers in a print server (excludes XPS Document Writer)
* Outputs that list into a simple GUI with the option of installing a selected printer
* Checks if the selected printer is already installed and asks if the user wants to remove and re-add it
* Asks users if they want to set the newly installed printer as their default
* Adjustable delay with prompt for install duration

## Download Executables
* <a href="https://github.com/need4swede/Powershell-Scripts/releases/download/v1.0/Print_Server_Manager.exe">Filename Manager (v1.0)</a>

## Change Log

* 1.0
    * Initial Release
* 2.0
    * Added install validation and support for multiple print servers
