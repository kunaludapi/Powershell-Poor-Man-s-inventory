#---------------------------------------------------------------------------------------------------------------------------
# Poor Mans server Inventory
# 
# Date: 31th October 2015, By: Kunal Udapi
# For more info check http://kunaludapi.blogspot.com\
#
# Initial Draft
#
# Code Name: Palash
#
# Verson 1.1 
# I have tested this in my lab, running windows server 2012 r2, windows 7, windows 8, Working fine, Powershell version 4
# 
#---------------------------------------------------------------------------------------------------------------------------

$ComputerList = .\ServersList.txt

$HtmlFilesFolder = "C:\Temp\WebSiteData"

if (-not(Test-Path -Path "$HtmlFilesFolder\Images")) {
   $null = New-Item -Name Images -Path $HtmlFilesFolder -ItemType Directory -Force
}

Set-Location  C:\Temp
Copy-Item  .\Images -Destination $HtmlFilesFolder -Recurse -Force

#region Head Style
$head = @"
<style>
#header {
    background-image: url('../Images/Header_Background.png');
    color:Black;
    text-align:center;
    padding:5px;
    font-family: Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;
}
#nav {
    line-height:30px;
    background-color:#E6E6E6;
    height:500px;
    width:Auto;
    float:left;
    padding:5px;
    font-family: Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;	      
}
#section {
    width:350px;
    float:left;
    padding:10px;
    font-family: Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;	 	 
}
#footer {
    background-color:black;
    color:white;
    clear:both;
    text-align:center;
    padding:5px;
    font-family: Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;	 	 
}

TH{
   width:100%; 
   font-size:0.9em;
   color:White;
   border-width:1px;
   padding: 2px;
   border-style: solid;
   border-color: #ADADAD;
   background-color:#666666};
   font-family: Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;

TD {
    width:100%;
    border-width: 1px;
    padding: 2px;
    border-style: solid;
    border-color: black;
    background-color: White};
    font-family: Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;
</style>

<!-- Table Style even and odd -->
<style type="text/css">
	.TFtable{
		border-collapse:collapse;
        width: 100%;
        table-layout: auto;
        white-space: nowrap;
	}
	.TFtable td{ 
		padding:5px; border:#ADADAD 1px solid;
	}
	/* provide some minimal visual accomodation for IE8 and below */
	.TFtable tr{
		background: #CFCFCF;
	}
	/*  Define the background color for all the ODD background rows  */
	.TFtable tr:nth-child(odd){ 
		background: #CFCFCF;
	}
	/*  Define the background color for all the EVEN background rows  */
	.TFtable tr:nth-child(even){
		background: #FFFFFF;
	}
</style>


<!-- Navigation bar -->
<style>
/* @import url(http://fonts.googleapis.com/css?family=Lato:300,400,700); */
/* Starter CSS for Flyout Menu */
#cssmenu,
#cssmenu ul,
#cssmenu ul li,
#cssmenu ul ul {
  float:left;
  list-style: none;
  margin: 0;
  padding: 0;
  border: 0;
  font-family: Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;
}
#cssmenu ul {
  position: relative;
  z-index: 597;
  float: left;
}
#cssmenu ul li {
  float: left;
  min-height: 1px;
  line-height: 1em;
  vertical-align: middle;
}
#cssmenu ul li.hover,
#cssmenu ul li:hover {
  position: relative;
  z-index: 599;
  cursor: default;
}
#cssmenu ul ul {
  margin-top: 1px;
  visibility: hidden;
  position: absolute;
  top: 1px;
  left: 99%;
  z-index: 598;
  width: 100%;
}
#cssmenu ul ul li {
  float: none;
}
#cssmenu ul ul ul {
  top: 1px;
  left: 99%;
}
#cssmenu ul li:hover > ul {
  visibility: visible;
}
#cssmenu ul li {
  float: none;
}
#cssmenu ul ul li {
  font-weight: normal;
}
/* Custom CSS Styles */
#cssmenu {
  font-family: font-family: Segoe UI,Frutiger,Frutiger Linotype,Dejavu Sans,Helvetica Neue,Arial,sans-serif;
  font-size: 18px;
  width: 200px;
}
#cssmenu ul a,
#cssmenu ul a:link,
#cssmenu ul a:visited {
  display: block;
  color: #848889;
  text-decoration: none;
  font-weight: 300;
}
#cssmenu > ul {
  float: none;
}
#cssmenu ul {
  background: #fff;
}
#cssmenu > ul > li {
  border-left: 3px solid #d7d8da;
}
#cssmenu > ul > li > a {
  padding: 10px 20px;
}
#cssmenu > ul > li:hover {
  border-left: 3px solid #3dbd99;
}
#cssmenu ul li:hover > a {
  color: #3dbd99;
}
#cssmenu > ul > li:hover {
  background: #f6f6f6;
}
/* Sub Menu */
#cssmenu ul ul a:link,
#cssmenu ul ul a:visited {
  font-weight: 400;
  font-size: 14px;
}
#cssmenu ul ul {
  width: 180px;
  background: none;
  border-left: 20px solid transparent;
}
#cssmenu ul ul a {
  padding: 8px 0;
  border-bottom: 1px solid #eeeeee;
}
#cssmenu ul ul li {
  padding: 0 20px;
  background: #fff;
}
#cssmenu ul ul li:last-child {
  border-bottom: 3px solid #d7d8da;
  padding-bottom: 10px;
}
#cssmenu ul ul li:first-child {
  padding-top: 10px;
}
#cssmenu ul ul li:last-child > a {
  border-bottom: none;
}
#cssmenu ul ul li:first-child:after {
  content: '';
  display: block;
  width: 0;
  height: 0;
  position: absolute;
  left: -20px;
  top: 13px;
  border-left: 10px solid transparent;
  border-right: 10px solid #fff;
  border-bottom: 10px solid transparent;
  border-top: 10px solid transparent;
}
} 
</style>
"@
#endregion Head Style
$workingcomputers = @()
$nonresponding = @()

foreach ($Computer in $ComputerList) {

write-host $computer
$workingcomputers += $Computer

#region Main Folder Path
$ComputerFolderPath = "{0}\{1}" -f $HtmlFilesFolder, $Computer
#endregion Main Folder Path
$title = "Status of $Computer"

#region Individual html file Path
if (-not(Test-Path -Path "$HtmlFilesFolder\$Computer")) {
   $null = New-Item -Name $Computer -Path $HtmlFilesFolder -ItemType Directory -Force
}

$HarWareFile =  $Computer + '-Hardware.html'
$HarWareNamePath = Join-Path -Path $ComputerFolderPath -ChildPath $HarWareFile

$OSFile =  $Computer + '-OS.html'
$OSFileNamePath = Join-Path -Path $ComputerFolderPath -ChildPath $OSFile

$DiskFile =  $Computer + '-Disk.html'
$DiskFileNamePath = Join-Path -Path $ComputerFolderPath -ChildPath $DiskFile

$ServicesFile =  $Computer + '-Services.html'
$ServicesFileNamePath = Join-Path -Path $ComputerFolderPath -ChildPath $ServicesFile

#endregion Individual html file Path

#region Indexing for Nextbutton and previousbutton
$Index = [array]::IndexOf($ComputerList, $computer)

#$previous = $ComputerList[($Index - 1)]
$previousbutton = "..\{0}\{1}-Hardware.html" -f $Computerlist[$Index - 1], $Computerlist[$Index - 1]

#$Next = $ComputerList[($Index + 1)]
if (($ComputerList.count - 1) -eq $Index) {
    $Nextbutton = "..\{0}\{1}-Hardware.html" -f $Computerlist[0], $Computerlist[0]
}
else {
    $Nextbutton = "..\{0}\{1}-Hardware.html" -f $Computerlist[$Index + 1], $Computerlist[$Index + 1]
}
#endregion Indexing for Nextbutton and previousbutton

#region navigation and footer
$Nav = @"
<div id='cssmenu'>
<ul>
   <li><a href="./$HarWareFile"><span>HARDWARE</span></a></li>
   <li class='active has-sub'><a href="./$OSFile"><span>Operating System</span></a></li>
   <li><a href=./$DiskFile><span>Storage</span></a></li>
   <li><a href=./$ServicesFile><span>Services</span></a></li>
   
   <li class='last'><a href='http://kunaludapi.blogspot.com'><span>Contact</span></a></li>
</ul>
</div>
"@

$Footer = @"
<div id="footer">
For more info checkout <a href='http://kunaludapi.blogspot.com'><span>http://kunaludapi.blogspot.com</span></a>
</div>
"@  
#endregion navigation and footer

#region Header button
$Header = @"
<div id="header">
<p>
<a href="..\Index.html">
  <img src="..\Images\HomeButton.png" alt="Home" style="width:70px;height:70px;border:0;align="middle";;margin: 0px 0px 30px 30px;">
</a>
<a href="$previousbutton">
  <img src="..\Images\Button-Previous-icon.png" alt="Previous" style="width:70px;height:70px;border:0;align="middle";margin: 0px 0px 30px 30px;">
</a> 
<a href="..\InventoryList.html">
  <img src="..\Images\Inventory-Button.png" alt="Inventory" style="width:70px;height:70px;border:0;align="middle";;margin: 0px 0px 30px 30px;">
</a>
<a href=$Nextbutton>
  <img src="..\Images\Button-Next-icon.png" alt="Next" style="width:70px;height:70px;border:0;align="middle";;margin: 0px 0px 30px 30px;">
</a>
<a href='http://kunaludapi.blogspot.in/p/about-me.html'>
  <img src="..\Images\Aboutme.png" alt="Next" style="width:62px;height:70px;border:0;align="middle";;margin: 0px 0px 30px 30px;">
</a>
</p>
<h1>$Computer Information</h1>
</div> 
"@
#endregion Header button
   
#region HardwareInfo   
$BiosInfo = Get-WmiObject win32_bios -ComputerName $Computer

$HardwareInfo = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer
$ProcessorInfo = Get-WmiObject -Class Win32_Processor -ComputerName $Computer

$DIMMSlots = Get-WmiObject -Class win32_PhysicalMemoryArray -ComputerName $Computer
$DIMMRAMs = Get-WmiObject -Class win32_PhysicalMemory -ComputerName $Computer

$NicInfo = Get-WmiObject -Class win32_networkadapter -ComputerName $Computer -Filter "PhysicalAdapter='True'"

#Motherboard
$MotherBoxObj = New-Object PSObject
$MotherBoxObj | Add-Member -MemberType NoteProperty -Name Manufacturer -Value $HardwareInfo.Name
$MotherBoxObj | Add-Member -MemberType NoteProperty -Name Model -Value $HardwareInfo.Model
$MotherBoxObj | Add-Member -MemberType NoteProperty -Name "Serial Number" -Value $BiosInfo.SerialNumber
$MotherBoxObjHTML = ($MotherBoxObj | ConvertTo-Html -Fragment -As List) -replace '<table>', '<table class="TFtable">'

#Processor
$ProcObj =  New-Object PSObject
$ProcObj | Add-Member -MemberType NoteProperty -Name Processor -Value $ProcessorInfo.Name
$ProcObj | Add-Member -MemberType NoteProperty -Name Manufacturer -Value $ProcessorInfo.Manufacturer
$ProcObj | Add-Member -MemberType NoteProperty -Name "Socket Designation" -Value $ProcessorInfo.SocketDesignation
$ProcObj | Add-Member -MemberType NoteProperty -Name "L2 Catche" -Value $ProcessorInfo.L2CacheSize
$ProcObj | Add-Member -MemberType NoteProperty -Name "L3 Catche" -Value $ProcessorInfo.L3CacheSize
$ProcObj | Add-Member -MemberType NoteProperty -Name "Max Clock Speed" -Value $ProcessorInfo.MaxClockSpeed
$ProcObj | Add-Member -MemberType NoteProperty -Name "Current Clock Speed" -Value $ProcessorInfo.CurrentClockSpeed
$ProcObj | Add-Member -MemberType NoteProperty -Name "Cores" -Value $ProcessorInfo.NumberOfCores
$ProcObj | Add-Member -MemberType NoteProperty -Name "Logical Processors" -Value $ProcessorInfo.NumberOfLogicalProcessors
$ProcObj | Add-Member -MemberType NoteProperty -Name "Virtulization Enabled" -Value $ProcessorInfo.VirtualizationFirmwareEnabled
$ProcObjHTML = ($ProcObj | ConvertTo-Html -Fragment -As List) -replace '<table>', '<table class="TFtable">'

#Physical Memory slots
$MemoryObj = New-Object PSObject
$MemoryObj | Add-Member -MemberType NoteProperty -Name "Installed Memory" -Value ("{0:n2} GB" -f ($HardwareInfo.TotalPhysicalMemory / 1GB))
$MemoryObj | Add-Member -MemberType NoteProperty -Name "Total Supported Max Memory" -Value ("{0:n2} GB" -f ($DIMMSlots.MaxCapacity / 1MB))
$MemoryObj | Add-Member -MemberType NoteProperty -Name "DIMM Slots" -Value $DIMMSlots.MemoryDevices
$MemoryObj | Add-Member -MemberType NoteProperty -Name "Filled Slots" -Value $DIMMRAMs.Count
$MemoryObjHTML = ($MemoryObj | ConvertTo-Html -Fragment -As List) -replace '<table>', '<table class="TFtable">'

#Physical Memory -as table
$DIMMReport = @()
foreach ($Dimm in $DIMMRAMs) {
    $RamObj = New-Object PSObject
    $RamObj | Add-Member -MemberType NoteProperty -Name "Memory Tag" -Value $DIMM.tag
    $RamObj | Add-Member -MemberType NoteProperty -Name "Memory Serial" -Value $DIMM.SerialNumber
    $RamObj | Add-Member -MemberType NoteProperty -Name "Memory Speed" -Value $DIMM.Speed
    $RamObj | Add-Member -MemberType NoteProperty -Name "Memory GB" -Value ("{0:n2} GB" -f ($DIMM.Capacity / 1GB))
    $RamObj | Add-Member -MemberType NoteProperty -Name "DIMM Label" -Value $DIMM.BankLabel
    $RamObj | Add-Member -MemberType NoteProperty -Name "DIMM Location" -Value $DIMM.DeviceLocator
    $RamObj | Add-Member -MemberType NoteProperty -Name "Memory PartNumber" -Value $DIMM.PartNumber
    $DIMMReport += $RamObj
}
$DIMMReportHTML = ($DIMMReport | ConvertTo-Html -Fragment) -replace '<table>', '<table class="TFtable">'

$NICreport = @()
#Physical Nic Info as table
foreach ($nic in $NicInfo) {
    $NICObj = New-Object PSObject
    $NICObj | Add-Member -MemberType NoteProperty -Name "Adapter Name" -Value $nic.Name
    $NICObj | Add-Member -MemberType NoteProperty -Name "Drivers" -Value $nic.ServiceName
    $NICObj | Add-Member -MemberType NoteProperty -Name "Device ID" -Value $nic.DeviceID
    $NICObj | Add-Member -MemberType NoteProperty -Name "Speed MB" -Value ("{0:n2} MB" -f ($nic.Speed / 1MB))
    $NICObj | Add-Member -MemberType NoteProperty -Name "MAC Address" -Value $nic.MACAddress
    $NICObj | Add-Member -MemberType NoteProperty -Name "Manufacturer" -Value $nic.Manufacturer
    $NICObj | Add-Member -MemberType NoteProperty -Name "Product Name" -Value $nic.ProductName 
    $NICReport +=$NICObj
}
$NicObjHTML = ($NICReport | ConvertTo-Html -Fragment) -replace '<table>', '<table class="TFtable">'

#HDD Info
$PhysicalDiskInfo = Get-WmiObject Win32_DiskDrive -ComputerName $Computer
$PhysicalDiskobjReport = @()
foreach ($PhysicalDisk in $PhysicalDiskInfo) {
    $PhysicalDiskobj =  New-Object PSObject
    $PhysicalDiskobj | Add-Member -MemberType NoteProperty -Name Name -Value $PhysicalDisk.Name
    $PhysicalDiskobj | Add-Member -MemberType NoteProperty -Name Model -Value $PhysicalDisk.Model
    $PhysicalDiskobj | Add-Member -MemberType NoteProperty -Name Size -Value ("{0:N2} GB" -f ($PhysicalDisk.Size / 1GB))
    $PhysicalDiskobj | Add-Member -MemberType NoteProperty -Name "SerialNumber" -Value $PhysicalDisk.SerialNumber
    $PhysicalDiskobjReport += $PhysicalDiskobj
}

$PhysicalDiskHTML = ($PhysicalDiskobjReport | ConvertTo-Html -Fragment) -replace '<table>', '<table class="TFtable">'

$HardwareSection = @"
<div id="section">
<h3>MotherBoard</h3>
<p>
$MotherBoxObjHTML
</p>
<h3>Processor</h3>
<p>
$ProcObjHTML
</p>
<h3>DIMM Slots</h3>
<p>
$MemoryObjHTML
</p>
<h3>Installed Memory</h3>
<p>
$DIMMReportHTML
</p>
<h3>Network Adapters</h3>
<p>
$NicObjHTML
</p>
<h3>Physical Disks</h3>
<p>
$PhysicalDiskHTML
</p>
</div>
"@

$HWbody = $Header + $Nav + $HardwareSection + $Footer  
ConvertTo-Html -Body $HWbody -Head $head | Out-File $HarWareNamePath
#endregion Hardware Info

#region OS info

$OSInfo = Get-WmiObject Win32_OperatingSystem -ComputerName $Computer
$NicIPinfo = Get-WmiObject win32_NetworkAdapterConfiguration -ComputerName $computer -Filter "IPEnabled='True'"

#Operating System -as list
$OSobj =  New-Object PSObject
$OSObj | Add-Member -MemberType NoteProperty -Name "Operating System" -Value $OSInfo.Caption
$OSObj | Add-Member -MemberType NoteProperty -Name "Build Number" -Value $OSInfo.BuildNumber
$OSObj | Add-Member -MemberType NoteProperty -Name "Computer Name" -Value $OSInfo.CSName
$OSObj | Add-Member -MemberType NoteProperty -Name "OS Install Date" -Value $OSInfo.ConverttoDateTime($OSInfo.InstallDate)
$OSObj | Add-Member -MemberType NoteProperty -Name "Last Boot Time" -Value $OSInfo.ConverttoDateTime($OSInfo.LastBootUpTime)
$OSObj | Add-Member -MemberType NoteProperty -Name "32 / 64 Bit" -Value $OSInfo.OSArchitecture 
$OSObj | Add-Member -MemberType NoteProperty -Name "Service Pack" -Value $OSInfo.ServicePackMajorVersion
$OSObj | Add-Member -MemberType NoteProperty -Name "Windows Directory" -Value $OSInfo.WindowsDirectory
$OSObj | Add-Member -MemberType NoteProperty -Name "DNS HostName" -Value $HardwareInfo.DNSHostName 
$OSObj | Add-Member -MemberType NoteProperty -Name "Domain" -Value $HardwareInfo.Domain
$OSObjHTML = ($OSObj | ConvertTo-Html -Fragment -As List) -replace '<table>', '<table class="TFtable">'

$NicIPReport = @()
#Network IP info
foreach ($nicip in $NicIPinfo) {
    $NICIPobj =  New-Object PSObject
    $NICIPobj | Add-Member -MemberType NoteProperty -Name "NIC Adapter" -Value $NICIP.Description
    $NICIPobj | Add-Member -MemberType NoteProperty -Name "Index" -Value $NICIP.Index
    $NICIPobj | Add-Member -MemberType NoteProperty -Name "DHCP" -Value $NICIP.DHCPEnabled
    $NICIPobj | Add-Member -MemberType NoteProperty -Name "IPAddress" -Value ($NICIP.IPAddress | Out-String)
    $NICIPobj | Add-Member -MemberType NoteProperty -Name "Gateway" -Value ($NICIP.DefaultIPGateway | Out-String)  
    $NICIPobj | Add-Member -MemberType NoteProperty -Name "DNS IP" -Value ($NICIP.DNSServerSearchOrder | Out-String)
    $NICIPobj | Add-Member -MemberType NoteProperty -Name "MacAddress" -Value $NICIP.MACAddress
    $NICIPobj | Add-Member -MemberType NoteProperty -Name "Nic Name" -Value ($NicInfo | where {$_.Index -match $nicip.index} | select -ExpandProperty NetConnectionID)
    $NicIPReport += $NICIPobj
}
$NICIPReportHTML = ($NicIPReport | ConvertTo-Html -Fragment) -replace '<table>', '<table class="TFtable">'


$OSSection = @"
<div id="section">
<h2>Operating System</h2>
<p>
$OSObjHTML
</p>
<h2>IP Address</h2>
<p>
$NICIPReportHTML
</p>
</div>
"@

$OSbody = $Header + $Nav + $OSSection + $Footer  
ConvertTo-Html -Body $OSbody -Head $head | Out-File $OSFileNamePath
#region OS info

#region Disk Info
$Logicaldiskreport = @()
$LogicalDiskInfo = Get-WmiObject Win32_LogicalDisk -ComputerName $computer
foreach ($LogicalDisk in $LogicalDiskInfo) {
    $LogicalDiskObj = New-Object PSObject
    $LogicalDiskObj | Add-Member -MemberType NoteProperty -Name Name -Value $LogicalDisk.Name
    $LogicalDiskObj | Add-Member -MemberType NoteProperty -Name Description -Value $LogicalDisk.Description
    $LogicalDiskObj | Add-Member -MemberType NoteProperty -Name "Drive Type" -Value $LogicalDisk.DriveType
    $LogicalDiskObj | Add-Member -MemberType NoteProperty -Name "File System" -Value $LogicalDisk.FileSystem
    $LogicalDiskObj | Add-Member -MemberType NoteProperty -Name "Total Space" -Value ("{0:n2} MB" -f ($LogicalDisk.Size / 1GB))
    $LogicalDiskObj | Add-Member -MemberType NoteProperty -Name "Free Space" -Value ("{0:n2} MB" -f ($LogicalDisk.FreeSpace / 1GB))
    $LogicalDiskObj | Add-Member -MemberType NoteProperty -Name "Volume Name" -Value $LogicalDisk.VolumeName
    $Logicaldiskreport += $LogicalDiskObj
}
$LogicaldiskHTML = ($Logicaldiskreport | ConvertTo-Html -Fragment) -replace '<table>', '<table class="TFtable">'
  
$StorageSection = @"
<div id="section">
<h2>Storage Drives</h2>
<p>
$LogicaldiskHTML
</p>
</div>
"@ 

$Storagebody = $Header + $Nav + $StorageSection + $Footer  
ConvertTo-Html -Body $Storagebody -Head $head | Out-File $DiskFileNamePath
#endregion Disk Info

#region services Info
$ServiceReport =@()
$ServiceInfo =  Get-WmiObject win32_Service -ComputerName $computer | Sort-Object State, StartMode
Foreach ($Service in $ServiceInfo) {
    $ServiceObj = New-Object PSObject
    $serviceObj | Add-Member -Name DisplayName -MemberType NoteProperty -Value $Service.DisplayName
    $serviceObj | Add-Member -Name Name -MemberType NoteProperty -Value $Service.Name
    $serviceObj | Add-Member -Name StartMode -MemberType NoteProperty -Value $Service.StartMode
    $serviceObj | Add-Member -Name Started -MemberType NoteProperty -Value $Service.Started
    $serviceObj | Add-Member -Name State -MemberType NoteProperty -Value $Service.State
    $ServiceReport += $serviceObj
}
$ServicesHTML = ($ServiceReport  | ConvertTo-Html -Fragment) -replace '<table>', '<table class="TFtable">'
$ServiceSection = @"
<div id="section">
<h2>Services</h2>
<p>
$ServicesHTML
</p>
</div>
"@ 

$Servicesbody = $Header + $Nav + $ServiceSection + $Footer  
ConvertTo-Html -Body $Servicesbody -Head $head | Out-File $ServicesFileNamePath
}
#endregion services Info

$workingReport = @()
ForEach ($working in $workingcomputers) {
$workingObj = "</p><a href=./$working/$working-Hardware.html>$working</a></p>"
$workingReport += $workingObj
}

$invhead = @"
<style>
body {
    background-image: url("./Images/Inventory_Page.png");
    background-repeat: no-repeat;
    background-position: right top;
    margin-right: 200px;
    background-attachment: fixed;
}
</style>
"@

$InventoryFile = "$HtmlFilesFolder\InventoryList.html"
$Inventorybody = $workingReport + $Footer  
ConvertTo-Html -Body $Inventorybody -Head $invhead | Out-File $InventoryFile

ii $InventoryFile

