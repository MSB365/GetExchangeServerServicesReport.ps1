<#
.SYNOPSIS
GetExchangeServerServicesReport.ps1 

.DESCRIPTION 
This PowerShell script will List you all Windows Services for Exchange and export it to an HTML.


.PARAMETER SourcePath
Specifies the source path for the HTML output.


.EXAMPLE
.\GetExchangeServerServicesReport.ps1 


.NOTES
Written by: Drago Petrovic

.TROUBLENOTES
if you get the following output: "Get-WmiObject : The RPC server is unavailable. (Exception from HRESULT: 0x800706BA)" check that the "Windows Management Instrumentation (WMI-In)" rule is enabled
in the firewall for each remote machine.
'netsh advfirewall firewall set rule group="Windows Management Instrumentation (WMI)" new enable=yes'


Find me on:

* LinkedIn:	https://www.linkedin.com/in/drago-petrovic-86075730/
* Xing:     https://www.xing.com/profile/Drago_Petrovic
* Website:  https://blog.abstergo.ch


Change Log
V1.00, 20/03/2017 - Initial version


--- keep it simple, but significant ---

.COPYRIGHT
Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, 
and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, 
WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
#>



$Style = "
<style>
    BODY{background-color:#b0c4de;}
    TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}
    TH{border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color:#778899}
    TD{border-width: 1px;padding: 3px;border-style: solid;border-color: black;}
    tr:nth-child(odd) { background-color:#d3d3d3;} 
    tr:nth-child(even) { background-color:white;}    
</style>
"
#Get all Exchange Servers
$servers = Get-ExchangeServer

#Get Date
$myDate = Get-Date -Format d

#Get Script Location
$settingspath = ($PSScriptRoot + "\ExchangeServicesReport_$mydate.html")

#Cell Color - Logic
$StatusColor = @{Disabled = ' bgcolor="Red">Disabled<';Auto = ' bgcolor="cyan">Auto<';Manual = ' bgcolor="Yellow">Manual<';UNKNOWN = ' bgcolor="#FF8000">UNKNOWN<';OK = ' bgcolor="#58FA58">OK<';}

#Exchange Service Status
$GService = Get-WmiObject win32_service -ComputerName $servers | select @{N='Exchange';E={$Servers}}, Name,Status,Startmode,StartName,Description | ConvertTo-HTML -AS Table -Fragment -PreContent '<h2>Exchange Services</h2>'|Out-String 

#Cell Color - Find\Replace
$StatusColor.Keys | foreach { $GService = $GService -replace ">$_<",($StatusColor.$_) }

#Last 5 Reboots
$Reboots = Get-WinEvent -FilterHashtable @{logname='system';id=6005} -MaxEvents 5| Select Message, TimeCreated  |ConvertTo-HTML -AS Table -Fragment -PreContent '<h2>Last 5 Reboots</h2>'|Out-String

#Save the HTML Web Page
ConvertTo-HTML -head $Style -PostContent $GService, $Reboots -PreContent '<h1>ExchangeServices</h1>'|Out-File $settingspath #Open TableHTML.html Invoke-Item $PSScriptRoot\ExchangeServices.html
#Open TableHTML.html
Invoke-Item $settingspath