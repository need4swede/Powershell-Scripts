##### Powershell Printer Management Tool #####

<#
The MIT License (MIT)

Copyright (c) 2021, Need4Swede

Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in
the Software without restriction, including without limitation the rights to
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of
the Software, and to permit persons to whom the Software is furnished to do so,
subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR
COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>

<#
Replace all instances of 'PRINTSERVERADDRESS' with the IP of your print server!
#>

### Main Function ###
function PrintManager() {
while ( $true ) {

# Main Window
$window = New-Object System.Windows.Forms.Form
$window.Text = "Print Server Manager"
$window.Size = New-Object System.Drawing.Size(500, 400)
$window.StartPosition = 'CenterScreen'

# Install Button
$okButton = New-Object System.Windows.Forms.Button
$okButton.Location = New-Object System.Drawing.Point(320,100)
$okButton.Size = New-Object System.Drawing.Size(120,60)
$okButton.Text = 'INSTALL'
$okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
$window.AcceptButton = $okButton

# Exit / Quit Button
$cancelButton = New-Object System.Windows.Forms.Button
$cancelButton.Location = New-Object System.Drawing.Point(320,200)
$cancelButton.Size = New-Object System.Drawing.Size(120,60)
$cancelButton.Text = 'QUIT'
$cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
$window.CancelButton = $cancelButton

# Listbox Header
$label = New-Object System.Windows.Forms.Label
$label.Location = New-Object System.Drawing.Point(10,20)
$label.Size = New-Object System.Drawing.Size(320,20)
$label.Font = New-Object System.Drawing.Font("Franklin Gothic",12,[System.Drawing.FontStyle]::Regular)
$label.Text = 'Select printer from the list and click install'

# IT Disclaimer
$label2 = New-Object System.Windows.Forms.Label
$label2.Location = New-Object System.Drawing.Point(280,300)
$label2.Size = New-Object System.Drawing.Size(200,50)
$label2.Font = New-Object System.Drawing.Font("Franklin Gothic",7,[System.Drawing.FontStyle]::Regular)
$label2.Text = "This application was developed by the <blank> IT Department. For all inquiries, please contact us via <contact>."

# Print Server UI Output
$listBox = New-Object System.Windows.Forms.ListBox
$listBox.Location = New-Object System.Drawing.Point(10,60)
$listBox.Size = New-Object System.Drawing.Size(260,20)
$listBox.Height= 280
$micro = "Microsoft XPS Document Writer" # Excludes the default XPS Document Writer from the available list of printers
Get-Printer -ComputerName "\\PRINTSERVERADDRESS" | Sort-Object | Where-Object{$_.Name -ne $micro} |ForEach-Object { $listBox.Items.Add($_.Name)}

# Main Window Properties
$window.TopMost = $true
$window.Controls.Add($listBox)
$window.controls.Add($label)
$window.controls.Add($label2)
$window.Controls.Add($cancelButton)
$window.Controls.Add($okButton)
$result = $window.ShowDialog()

# Quit Function
if ($result -eq [System.Windows.Forms.DialogResult]::Cancel)
{
    exit
} else {

# Install Function
if ($result -eq [System.Windows.Forms.DialogResult]::OK){
    $x = $listBox.SelectedItem # Variable for selected printer
    
    # Checks if a printer was selected before install
    if (!$x){ 
        $b = new-object -comobject wscript.shell
        $b.popup("No printer was selected. Please select a printer and then click install.", `
        0,"No Printer Selected!",0)
        PrintManager}

    # Edits system registry to allow for local default printer management
    New-ItemProperty -LiteralPath 'Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows' -Name  'LegacyDefaultPrinterMode' -Value '00000001' -PropertyType 'DWORD' –Force
    Pop-Location # Get out of the Registry

    # Checks if the selected printer is already installed on the system
    $printers = Get-Printer
    if ($printers.Name -like "\\PRINTSERVERADDRESS\$x")
    {
        $rmv = new-object -comobject wscript.shell
        $intAnswer = $rmv.popup("The $x is already an installed printer on your system. Would you like to remove and re-add this printer?", `
        0,"Printer Found!",4)
        If ($intAnswer -eq 6) {    
            # Removes the printer
            Remove-Printer -Name  "\\PRINTSERVERADDRESS\$x"
            $rmv.popup("The $x has now been removed and will be re-installed on your system.", `
            0,"Printer Removed",0)
            } else {
            $rmv.popup("Nothing removed. Please run the tool again to add additional printers.", `
            0,"Print Manager",0)
            exit   }
    }

    ## Install selected printer w/ UI timer
    Add-Printer -ConnectionName \\PRINTSERVERADDRESS\$x

    #Adjust timer delay here
    $delay = 20

    $Counter_Form = New-Object System.Windows.Forms.Form
    $Counter_Form.Text = "Installing Printer"

    #Form size options
    $Counter_Form.Width = 350
    $Counter_Form.Height = 150

    #Centers form on screen
    $Counter_Form.StartPosition = "CenterScreen"

    #Places form on top of everything else
    $Counter_Form.TopMost = $true

    $Counter_Label = New-Object System.Windows.Forms.Label
    $Counter_Label2 = New-Object System.Windows.Forms.Label

    #Label2's text
    $Counter_Label2.Text = "INSTALLING. PLEASE WAIT."

    #Labels size and position
    $Counter_Label.AutoSize = $true
    $Counter_Label.Location = New-Object System.Drawing.Point(50,60)
    $Counter_Label2.AutoSize = $true
    $Counter_Label2.Location = New-Object System.Drawing.Point(90,30)


    $Counter_Form.Controls.Add($Counter_Label)
    $Counter_Form.Controls.Add($Counter_Label2)

    while ($delay -ge 0)
    {
      $Counter_Form.Show()


    #Timer label's text
      $Counter_Label.Text = "Printer installation will conclude in $($delay) seconds"


      start-sleep 1
      $delay -= 1
    }
    $Counter_Form.Close() 

    # Checks if the printer installed successfully
    $printers = Get-Printer
    if ($printers.Name -like "\\PRINTSERVERADDRESS\$x")
    {   
        # SUCCESSFUL INSTALL

        # Asks user if they want the printer to be their new default
        $a = new-object -comobject wscript.shell
        $intAnswer = $a.popup("Set the $x as your default printer?", `
        0,"Set as Default",4)
        If ($intAnswer -eq 6) { 
        Start-Sleep -s 3
        RUNDLL32 PRINTUI.DLL,PrintUIEntry /y /n \\PRINTSERVERADDRESS\$x
        $a.popup("The $x is now your default printer! You can change this setting at any time.")
        } else {}
    } 
    else {

        # FAILED INSTALL

        Add-Type -AssemblyName PresentationCore,PresentationFramework
        $ButtonType = [System.Windows.MessageBoxButton]::Ok
        $MessageIcon = [System.Windows.MessageBoxImage]::Error
        $MessageBody = "Printer Manager was unable to install $x. Please contact IT for support."
        $MessageTitle = "Install Failed"
        $Result = [System.Windows.MessageBox]::Show($MessageBody,$MessageTitle,$ButtonType,$MessageIcon)
        exit

} 
} 
    # Restart program
    $re = new-object -comobject wscript.shell
    $intAnswer = $re.popup("Would you like to add any additional printers?", `
    0,"Printer Manager",4)
    If ($intAnswer -eq 6) {    

    } else {
        break
      }
        }
                      }}

PrintManager
