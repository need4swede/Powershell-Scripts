# Print Server Manager

PS script that lets you query all printers in a print server and output them in a simple GUI for easy installs.
Great for home users with advanced printer setups or IT Admins.
I was able to compile this into an executable and host it on a fileshare and distribute it to our org via GPO and our users have been using it to install and manage their printers. 

## Functions

* Fetches a list of all printers in a print server (excludes XPS Document Writer)
* Outputs that list into a simple GUI with the option of installing a selected printer
* Checks if the selected printer is already installed and asks if the user wants to remove and re-add it
* Asks users if they want to set the newly installed printer as their default
* Adjustable delay with prompt for install duration

```
There's a footnote section that let's you include your own disclaimer.
Be sure to replace every mention of the phrase PRINTSERVER with the IP/Name of your print server.
If the aforementioned PRINTSERVER in the code doesn't include '' or "", then don't add them when updating my code.
If you see 'PRINTSERVER' or "PRINTSERVER", then keep the quotation marks and insert the IP/Name of your server.
```
## Download Executables
* None Available

## Change Log

* 1.0
    * Initial Release
