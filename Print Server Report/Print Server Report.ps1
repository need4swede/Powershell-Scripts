##### Powershell Print Server Report Tool #####

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

Param (
    [string]$Printservers = "\\PRINTSERVERADDRESS"
)

# Create new Excel workbook
cls
$Excel = New-Object -ComObject Excel.Application
$Excel.Visible = $True
$Excel = $Excel.Workbooks.Add()
$Sheet = $Excel.Worksheets.Item(1)
$Sheet.Name = "Printer Inventory"
#======================================================
$Sheet.Cells.Item(1,1) = "Print Server"
$Sheet.Cells.Item(1,2) = "Printer Name"
$Sheet.Cells.Item(1,3) = "Location"
$Sheet.Cells.Item(1,4) = "Comment"
$Sheet.Cells.Item(1,5) = "IP Address"
$Sheet.Cells.Item(1,6) = "Driver Name"
$Sheet.Cells.Item(1,7) = "Driver Version"
$Sheet.Cells.Item(1,8) = "Driver"
$Sheet.Cells.Item(1,9) = "Shared"
$Sheet.Cells.Item(1,10) = "Share Name"
#=======================================================
$intRow = 2
$WorkBook = $Sheet.UsedRange
$WorkBook.Interior.ColorIndex = 40
$WorkBook.Font.ColorIndex = 11
$WorkBook.Font.Bold = $True
#=======================================================

# Get printer information
ForEach ($Printserver in $Printservers)
{   $Printers = Get-Printer -ComputerName "\\PRINTSERVERADDRESS" | Sort-Object | Where-Object{$_.Name -ne $micro}
    ForEach ($Printer in $Printers)
    {
        if ($Printer.Name -notlike "Microsoft XPS*")
        {
            $Sheet.Cells.Item($intRow, 1) = $Printer.PortName
            $Sheet.Cells.Item($intRow, 2) = $Printer.Name
            $Sheet.Cells.Item($intRow, 3) = $Printer.Location
            $Sheet.Cells.Item($intRow, 4) = $Printer.Comment
            
            If ($Printer.PortName -notlike "*\*")
            {   $Ports = Get-WmiObject Win32_TcpIpPrinterPort -Filter "name = '$($Printer.Portname)'" -ComputerName $Printserver
                ForEach ($Port in $Ports)
                {
                    $Sheet.Cells.Item($intRow, 5) = $Port.HostAddress
                }
            }
       
            ####################       
            $Drivers = Get-WmiObject Win32_PrinterDriver -Filter "__path like '%$($Printer.DriverName)%'" -ComputerName $Printserver
            ForEach ($Driver in $Drivers)
            {
                $Drive = $Driver.DriverPath.Substring(0,1)
                $Sheet.Cells.Item($intRow, 7) = (Get-ItemProperty ($Driver.DriverPath.Replace("$Drive`:","\\$PrintServer\$Drive`$"))).VersionInfo.ProductVersion
                $Sheet.Cells.Item($intRow,8) = Split-Path $Driver.DriverPath -Leaf
            }
            ####################      
            $Sheet.Cells.Item($intRow, 6) = $Printer.DriverName
            $Sheet.Cells.Item($intRow, 9) = $Printer.Shared
            $Sheet.Cells.Item($intRow, 10) = $Printer.ShareName
            $intRow ++
        }
    }
    $WorkBook.EntireColumn.AutoFit() | Out-Null
}
 

$intRow ++ 
$Sheet.Cells.Item($intRow,1) = "Printer inventory completed"
$Sheet.Cells.Item($intRow,1).Font.Bold = $True
$Sheet.Cells.Item($intRow,1).Interior.ColorIndex = 40
$Sheet.Cells.Item($intRow,2).Interior.ColorIndex = 40 
