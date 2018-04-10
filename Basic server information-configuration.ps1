<#
.SYNOPSIS
Get core information for server
.DESCRIPTION
Script will generate html reports with core information about server resources and configuration
.NOTES  
The script will execute the commands on multiple machines sequentially using non-concurrent sessions. This will process all servers from Serverlist.txt in the listed order.
The info will be exported to a csv format.
Requires: Serverlist.txt must be created in the same folder where the script is.
File Name  : Redirection-to-Onedrive.ps1
Author: Sasa Zelic 
sasa.zelic@outlook.com
#>

#region html style
$style = @"
<style>

h1 { text-align: center;background: #ffffff ;color:#404040;font-weight:normal}
h2,h3,h4, h5{ text-align: center;font-size:19px; color:#4d4d4d;font-weight:normal }
th { text-align: center;font-size:16px; color:#262626;font-weight:normal }
h10 { text-align: center;background: #ffffff ;color:#383838;font-size:13px;}



table { 
    margin: auto; 
    font-family: Segoe UI; 
    #box-shadow: 10px 10px 5px #888; 
    border-collapse: collapse;
}

th {
    background: #f3f3f3;
    border: 1px solid #CCCCCC;
    font-size: 16px;
    font-weight:normal;
    padding:7px;
    text-align:center;
    vertical-align:middle;
}
td {
    padding:10px 8px 10px 8px;
    font-size: 14px;
    border: 1px solid #CCCCCC;
    text-align:center;
    vertical-align:middle;
} 


.VMs{
    width:100%;
    /*height:200px;*/
    /*border: 1px solid;*/
    float:left;
    margin-bottom:22px;
    line-height:1.5;
}


tr{
    font-size: 12px;
}
tr:nth-child(even) { background: #ffffff; }
tr:nth-child(odd) { background: #f9f7f8; }




hr.style-six {
padding: 0;
border: none;
height: 1px;
background-image: -webkit-linear-gradient(left, rgba(0,0,0,0), rgba(0,0,0,0.75), rgba(0,0,0,0));
background-image: -moz-linear-gradient(left, rgba(0,0,0,0), rgba(0,0,0,0.75), rgba(0,0,0,0));
background-image: -ms-linear-gradient(left, rgba(0,0,0,0), rgba(0,0,0,0.75), rgba(0,0,0,0));
background-image: -o-linear-gradient(left, rgba(0,0,0,0), rgba(0,0,0,0.75), rgba(0,0,0,0));
color: #333;
text-align: center;
}

hr.style-six:after {
content:" ";
display: inline-block;
position: relative;
top: -22.1em;
font-size: 1.5em;
padding: 19px 1.75em;
background-size: 90px 90px;
height: 50px;
}

mark {
    background-color: #ff8080;
    
}

</style>
"@
#endregion
#region working directory variable
if (test-path 'c:\yw-data') {
    
    $wdir = 'c:\yw-data'
}
elseif (test-path 'd:\yw-data\') {
    $wdir = 'd:\yw-data'
}
elseif (test-path 'f:\yw-data\') {
    $wdir = 'f:\yw-data'
}
elseif (test-path 'g:\yw-data\') {
    $wdir = 'g:\yw-data'
}
elseif (test-path 'h:\yw-data\') {
    $wdir = 'h:\yw-data'
}
else {
    New-Item -ItemType directory  -Path 'C:\yw-data'
    $wdir = 'C:\yw-data'
} 
#endregion
#region check machine type
Function Get-MachineType {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    Param
    (
        # ComputerName
        [Parameter(Mandatory = $false,
            ValueFromPipeline = $true,
            ValueFromPipelineByPropertyName = $true,
            Position = 0)]
        [string[]]$ComputerName = $env:COMPUTERNAME,
        $Credential = [System.Management.Automation.PSCredential]::Empty
    )

    Begin {
    }
    Process {
        foreach ($Computer in $ComputerName) {
            Write-Verbose "Checking $Computer"
            try {
                # Check to see if $Computer resolves DNS lookup successfuly.
                $null = [System.Net.DNS]::GetHostEntry($Computer)
                
                $ComputerSystemInfo = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $Computer -ErrorAction Stop -Credential $Credential
                
                switch ($ComputerSystemInfo.Model) {
                    
                    # Check for Hyper-V Machine Type
                    "Virtual Machine" {
                        $MachineType = "VM"
                    }

                    # Check for VMware Machine Type
                    "VMware Virtual Platform" {
                        $MachineType = "VM"
                    }

                    # Check for Oracle VM Machine Type
                    "VirtualBox" {
                        $MachineType = "VM"
                    }

                    # Check for Xen
                    "HVM domU" {
                        $MachineType = "VM"
                    }

                    # Check for KVM
                    # I need the values for the Model for which to check.

                    # Otherwise it is a physical Box
                    default {
                        $MachineType = "Physical"
                    }
                }
                
                # Building MachineTypeInfo Object
                $MachineTypeInfo = New-Object -TypeName PSObject -Property ([ordered]@{
                        ComputerName = $ComputerSystemInfo.PSComputername
                        Type         = $MachineType
                        Manufacturer = $ComputerSystemInfo.Manufacturer
                        Model        = $ComputerSystemInfo.Model
                    })
                $MachineTypeInfo
            }
            catch [Exception] {
                Write-Output "$Computer`: $($_.Exception.Message)"
            }
        }
    }
    End {

    }
}
$machine = Get-MachineType
$machinemodel = Get-MachineType
$machinemodel = $machine.model
$machinetype = $machine.type
$machinemanufacturer = $machine.Manufacturer

#find hyperv host of VM
$hostofVM = get-item "HKLM:\SOFTWARE\Microsoft\Virtual Machine\Guest\Parameters" -ErrorAction SilentlyContinue
if ($hostofVM) {
    $hostofVM = (get-item "HKLM:\SOFTWARE\Microsoft\Virtual Machine\Guest\Parameters").GetValue("hostname")
    write-host "This is virtual machine" -ForegroundColor Yellow
}
else {
    $hostofVM = "This is physical machine"
    write-host "This is physical machine" -ForegroundColor Yellow
}

#generate machine type hash table whether it is physical or virtual machine
$serverCoreinfo=""
if ($machinetype -like "*physical*") {

    $hash = [ordered]@{
        #'ComputerName' = $env:COMPUTERNAME
        'Machine<br>Type'  = '!gray!'+$machinetype+'!spanend!' + '!asterisk!'
        'Manufacturer' = '!graysmall!'+ $machinemanufacturer + '!spanend!'
        'Model'        = '!graysmall!'+  $machinemodel + '!spanend!'
    }
    $hash1=New-Object -TypeName psobject -Property $hash
    $serverCoreinfo = $hash1 | ConvertTo-Html -Fragment -PreContent "<h2>&diams;Machine type:</h2>" | Out-string
} else {
    $hash = [ordered]@{
        #'ComputerName' = $env:COMPUTERNAME
        'Machine<br>Type'  = '!gray!'+$machinetype+'!spanend!' + '!asterisk!'
        'VmHost'      = '!graysmall!'+  $hostofVM + '!spanend!'

       
    }
    $hash1=New-Object -TypeName psobject -Property $hash
    $serverCoreinfo = $hash1 | ConvertTo-Html -Fragment -PreContent "<h2>&diams;Machine type:</h2>" | Out-string

}
#endregion

#region domain-worstation status

<#
.SYNOPSIS
get domain info

.DESCRIPTION
Long description

.EXAMPLE
An example

.NOTES
Thisi s how to get pc status 

$dr = gwmi -Class win32_computersystem | select -ExpandProperty domainrole
switch ($dr) {
        0 {"Standalone Workstation"} 
        1 {"Member Workstation"} 
        2 {"Standalone Server"} 
        3 {"Member Server"} 
        4 {"Backup Domain Controller"} 
        5 {"Primary Domain Controller"} 
        default {"Unknown"} 
        } # end switch
#> 
#check to see if pc is domain joined or not

$dr = gwmi -Class win32_computersystem | select -ExpandProperty domainrole
$serverstate = gwmi -Class win32_computersystem | select -ExpandProperty domainrole
if ($dr -eq 0) {
    $dr = "Standalone workstation"    
}
elseif ($dr -eq 1) {
    $dr = "Member workstation"
}
elseif ($dr -eq 2) {
    $dr = "Non domain joined server"
}
elseif ($dr -eq 3) {
    $dr = "Domain joined server"
}
elseif ($dr -eq 4) {
    $dr = "Secondary domain controller"
}
elseif ($dr -eq 5) {
    $dr = "Primary domain controller"
}
else {
    $dr = "Status unknown"
}
$domain = Get-WmiObject Win32_computersystem  #Get OS Information
$ifdomainjoined = $domain.PartOfDomain
if ($ifdomainjoined -eq $true) {
    $ifdomainjoined = $domain.Domain
}
else {
    $ifdomainjoined = "Not part of a domain"
}

#choose hash table if not part of domain or yes
if ($serverstate -eq 0 -or $serverstate -eq 2) {

    $hash = [ordered]@{
        'ServerName' = $env:COMPUTERNAME
        'ServerRole' = $dr
            
    }
    $domain1 = New-Object -TypeName PSObject -Property $hash

}
elseif ( $serverstate -eq 1 -or $serverstate -eq 3 ) {
    $hash = [ordered]@{
        'Server<br>Name' = $env:COMPUTERNAME
        'Server<br>Role' = $dr
        'Domain<br>Name' = $ifdomainjoined
            
    }
    $domain1 = New-Object -TypeName PSObject -Property $hash

}
else {
    $dinfo = Get-ADDomainController $env:COMPUTERNAME
    $hash = [ordered]@{
        #'ServerName'      = $env:COMPUTERNAME
        'Server<br>Role'      = if ($dr -like "*domain controller*"){'!cellgreen!'+$dr + '!spanend!'}else {$dr}
        'Domain<br>Name'      = $ifdomainjoined
        'AD-Site'         = $dinfo.site
        'Global<br>catalog' = $dinfo.isglobalcatalog





    }
    $domain1 = New-Object -TypeName PSObject -Property $hash
    
}

$domain2 = $domain1  | ConvertTo-Html -Fragment -PreContent '<h2>&diams; Domain joined/standalone?</h2>' | Out-String


#endregion

#region forest
$forest = " "
if ($serverstate -eq 4 -or $serverstate -eq 5) {

    $forestinfo = get-adforest
    

    $hash = [ordered]@{
        'Forest<br>Name'      = $forestinfo.name
        'Domains<br>In Forest' = ($forestinfo.Domains -join "!br!")
        'Global Catalog<br>Servers'       = ($forestinfo.Globalcatalogs -join "!br!")
        'Root<br>Domain'      = $forestinfo.Rootdomain 
        'AD<br>Sites'         = ($forestinfo.Sites -join "!br!")
        'Forest<br>Mode'      = $forestinfo.forestmode 
    }
    $forestinfo = New-Object -TypeName PSObject -Property $hash
    $forest = $forestinfo | ConvertTo-Html -Fragment -PreContent '<h2>&diams;Forest wide info</h2>' | Out-String
}
else {
    $forest = "." 
}



#get fsmo roles holders
if ($serverstate -eq 4 -or $serverstate -eq 5) {

    $fsmoforest = get-adforest
    $fsmodomain = Get-ADDomain

    $hash = [ordered]@{
        'Schema<br>Master'        = $fsmoforest.SchemaMaster
        'Domain<br>Naming<br>Master'  = $fsmoforest.DomainNamingMaster
        'PDC<br>Emulator'         = $fsmodomain.PDCEmulator
        'RID<br>Master'           = $fsmodomain.RIDMaster 
        'Infrastructure<br>Master' = $fsmodomain.InfrastructureMaster
            
    }
    $forestinfo = New-Object -TypeName PSObject -Property $hash
    $forest1 = $forestinfo | ConvertTo-Html -Fragment -PreContent '<h2>&diams;FSMO roles</h2>' | Out-String

}
else {
    $forest1 = "."
}


#endregion

#region network
function network {
    [cmdletbinding()]
param (
 [parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
    [string[]]$ComputerName = $env:computername
)

begin {}
process {
 foreach ($Computer in $ComputerName) {
  if(Test-Connection -ComputerName $Computer -Count 1 -ea 0) {
   try {
    $Networks =  
    Get-WmiObject Win32_NetworkAdapterConfiguration -ComputerName $Computer -EA Stop | ? {$_.IPEnabled }
   
   } catch {
        Write-Warning "Error occurred while querying $computer."
        Continue
   }
   foreach ($Network in $Networks) {
    $IPAddress  = $Network.IpAddress[0]
    $SubnetMask  = $Network.IPSubnet[0]
    $DefaultGateway = $Network.DefaultIPGateway
    $DNSServers  = $Network.DNSServerSearchOrder
    $WINS1 = $Network.WINSPrimaryServer
    $WINS2 = $Network.WINSSecondaryServer   
    $WINS = @($WINS1,$WINS2)         
    $IsDHCPEnabled = $false
    If($network.DHCPEnabled) {
     $IsDHCPEnabled = $true
    }
    $MACAddress  = $Network.MACAddress
    $OutputObj  = New-Object -Type PSObject
   # $OutputObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Computer.ToUpper()
    $OutputObj | Add-Member -MemberType NoteProperty -Name IPAddress -Value ($IPAddress+'!asterisk!')
    $OutputObj | Add-Member -MemberType NoteProperty -Name Subnet<br>Mask -Value ('!graysmall!'+  $SubnetMask + '!spanend!')
    $OutputObj | Add-Member -MemberType NoteProperty -Name Gateway -Value  ('!graysmall!'+ ( $DefaultGateway  -join "!br!" ) + '!spanend!')     
    $OutputObj | Add-Member -MemberType NoteProperty -Name '<span style="font-size:12px;color:#b3b3b3;"> IsDHCP<br>Enabled </spanend>' -Value ('!graysmall!'+ $IsDHCPEnabled + '!spanend!')
    $OutputObj | Add-Member -MemberType NoteProperty -Name DNS<br>Servers -Value ('!graysmall!'+ ( $DNSServers  -join "!br!" ) + '!spanend!')    
    $OutputObj | Add-Member -MemberType NoteProperty -Name '<span style="font-size:12px;color:#b3b3b3;"> WINS<br>Servers </spanend>' -Value ( '!graysmall!'+ $WINS -join "!br!")        
    $OutputObj | Add-Member -MemberType NoteProperty -Name '<span style="font-size:12px;color:#b3b3b3;"> MAC<br>Address </spanend>' -Value ('!graysmall!'+ $MACAddress + '!spanend!')
    $OutputObj
   }
  }
 }
}
 
end {}
        
}
   
$network = network | ConvertTo-Html -Fragment -PreContent '<h2>&diams;IP configuration</h2>' | Out-String
#endregion

#region osinfo
Function OSinfo {
    $server = 'localhost'

    $CPUInfo = Get-WmiObject Win32_Processor -ComputerName $server  #Get CPU Information 
    $OSInfo = Get-WmiObject Win32_OperatingSystem -ComputerName $server #Get OS Information
    #Get Memory Information. The data will be shown in a table as MB, rounded to the nearest second decimal.
    $OSTotalVirtualMemory = [math]::round($OSInfo.TotalVirtualMemorySize / 1MB, 2)
    $OSTotalVisibleMemory = [math]::round(($OSInfo.TotalVisibleMemorySize / 1MB), 2)
    $PhysicalMemory = Get-WmiObject CIM_PhysicalMemory -ComputerName $server | Measure-Object -Property capacity -Sum | ForEach-Object { [Math]::Round(($_.sum / 1GB), 2) }
    Foreach ($CPU in $CPUInfo) {
        $hash = [ordered]@{
            'OS Name'= '!arrow!'+$OSInfo.Caption + '!asterisk!'
           '<span style="font-size:12px;color:#b3b3b3;"> Version </spanend>'= '!graysmall!'+ $OSInfo.Version + '!spanend!'
            '<span style="font-size:12px;color:#b3b3b3;"> ServicePack </spanend>' = '!graysmall!'+ $OSInfo.ServicePackMajorVersion + '!spanend!'
          '<span style="font-size:12px;color:#b3b3b3;"> OS<br>architecture </spanend>' = '!graysmall!'+ $OSInfo.OSArchitecture + '!spanend!'
          '<span style="font-size:12px;color:#b3b3b3;"> Memory<br>GB </spanend>'  = '!graysmall!'+ $physicalmemory + '!spanend!'
        }
        new-object -TypeName psobject -Property $hash


	
    }
}
$serverinfo = OSinfo | ConvertTo-Html -Fragment -PreContent "<h2>&diams; OS</h2>" | Out-string



#endregion

#region cpu info
Function CPUinfo {
    $server = 'localhost'
    
    $CPUInfo = Get-WmiObject Win32_Processor -ComputerName $server #Get CPU Information
    $OSInfo = Get-WmiObject Win32_OperatingSystem -ComputerName $server #Get OS Information
    #Get Memory Information. The data will be shown in a table as MB, rounded to the nearest second decimal.
    $OSTotalVirtualMemory = [math]::round($OSInfo.TotalVirtualMemorySize / 1MB, 2)
    $OSTotalVisibleMemory = [math]::round(($OSInfo.TotalVisibleMemorySize / 1MB), 2)
    $PhysicalMemory = Get-WmiObject CIM_PhysicalMemory -ComputerName $server | Measure-Object -Property capacity -Sum | ForEach-Object { [Math]::Round(($_.sum / 1GB), 2) }

    Foreach ($CPU in $CPUInfo) {
    
        $cpushort  = ($cpu.name).split(' ')
        [string] $cpushort= $cpushort[0] +' ' + $cpushort[1]
        $cpumodelshort= ($cpu.description).split(' ')
        [string] $cpumodelshort = $cpumodelshort[0]
       

        $hash = [ordered]@{
            'Manufacturer'         = $CPU.Manufacturer;
            'Processor'            =  '!p1!'+ "$($cpushort)" + '!br!' + '!graysmall!' +  "$($CPU.name)" + '!spanend!' + '!asterisk!' + '!pend!';
            'Model'                = '!p1!'+ ("$($cpumodelshort)" + '!br!' + '!graysmallabbr!' + "$($cpu.description)" + '!quotes!' + 'More info' + '!abbrend!' + '!spanend!'+ '!pend!');
            'Physical<br>Cores'   = $CPU.NumberOfCores;
            'CPU<br>L2CacheSize'   = ('!graysmall!' +$CPU.L2CacheSize+ '!spanend!');
            'CPU<br>L3CacheSize'   = '!graysmall!' +$CPU.L3CacheSize+ '!spanend!';
            'Sockets'              = '!graysmall!' +$CPU.SocketDesignation + '!spanend!';
            'Logical<br>Cores'     = '!graysmall!' +$CPU.NumberOfLogicalProcessors + '!spanend!'
                
        }
        new-object -TypeName psobject -Property $hash
        
    }
}

$cpuinfo = cpuinfo | ConvertTo-Html -Fragment -PreContent "<h2>&diams; CPU</h2>" | Out-string

#endregion


#region memory info
Function memoryinfo {
    $server = 'localhost'
    
    $CPUInfo = Get-WmiObject Win32_Processor -ComputerName $server #Get CPU Information
    $OSInfo = Get-WmiObject Win32_OperatingSystem -ComputerName $server #Get OS Information
    #Get Memory Information. The data will be shown in a table as MB, rounded to the nearest second decimal.
    $OSTotalVirtualMemory = [math]::round($OSInfo.TotalVirtualMemorySize / 1MB, 2)
    $OSTotalVisibleMemory = [math]::round(($OSInfo.TotalVisibleMemorySize / 1MB), 2)
    $PhysicalMemory = Get-WmiObject CIM_PhysicalMemory -ComputerName $server | Measure-Object -Property capacity -Sum | % { [Math]::Round(($_.sum / 1GB), 2) }
    Foreach ($CPU in $CPUInfo) {
        $hash = [ordered]@{
                
            'Total<br>Physical Memory' = '!p1!' + "$PhysicalMemory" + '!br!' + '!graysmall!' + 'GB' + '!spanend!' + '!pend!'
            'Total<br>Virtual Memory'  = '!p1!' +'!graysmall!' +$OSTotalVirtualMemory+  '!br!''GB' +'!spanend!'+ '!pend!'
            'Total<br>Visible Memory'  = '!p1!' +'!graysmall!' +$OSTotalVisibleMemory+ '!br!''GB' +'!spanend!'+ '!pend!'
        }
        new-object -TypeName psobject -Property $hash
        
    }
}
$memoryinfo = memoryinfo | ConvertTo-Html -Fragment -PreContent "<h2>&diams; Memory </h2>" | Out-string

#endregion

#region disk info
function Get-InfoDisk {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)][string]$ComputerName = 'localhost'
    ) $drives = Get-WmiObject -class Win32_LogicalDisk -ComputerName $ComputerName `
        -Filter "DriveType=3"
    foreach ($drive in $drives) {
     [int]$drivefreespace = "{0:N1}" -f ( $drive.freespace / $drive.size * 100 -as [int])
         
        $props = [ordered]@{
           
            'Drive'     = '!arrow!' + $drive.DeviceID;
            'Size - GB' = $drive.size / 1GB -as [int] ;
            'Free GB' = '!graysmall!' + ("{0:N2}" -f ($drive.freespace / 1GB)) + '!spanend!' ;
            'Free<br>%'    = if ($drivefreespace -lt 14) {'!cellred!'+ $drivefreespace +'!spanend!' }else {$drivefreespace} ;
        }
        New-Object -TypeName PSObject -Property $props
    }
}
$diskafter = Get-InfoDisk | ConvertTo-Html -Fragment -PreContent '<h2>&diams; Disk </h2>' | Out-String
    
#endregion

#region shared folders#
#region shared folders#
function sharedfolders {

    $share = Get-SmbShare | where {$_.name -notlike "*$*"}
    foreach ($sh in $share) {
        $hash = [ordered]@{
            'Name' = $sh.Name;
            'Path' = '!graysmall!' + "$($sh.Path)" + '!spanend!'
         
 
        }
        New-Object -TypeName PSObject -Property $hash
    }
}
$sharedfolders = sharedfolders | ConvertTo-Html -Fragment -PreContent '<h2>&diams; Shared folders </h2>' | Out-String
#endregion
#endregion

#region features installed
function serverfeatures {
    $features = Get-WindowsFeature | ? { $_.Installed -AND $_.SubFeatures.Count -eq 0 }

    foreach ($fea in $features) {

        $hash = [ordered]@{
            'Name'         = $fea.name
            'Display name' = '!graysmall!' + $fea.Displayname + '!spanend!'
            'Type'         = if ($fea.featuretype -eq 'role'){'!amber!'+$fea.featuretype+ '!spanend!' + '!asterisk!'}else{$fea.featuretype}

        }
        New-Object -TypeName PSObject -Property $hash

    }


}
$serverfeatures = serverfeatures | ConvertTo-Html -Fragment -PreContent '<h2>&diams; Servers roles & Features </h2>' | Out-String
#endregion

function convertsize {

[cmdletbinding()]
param(
 
[parameter(Mandatory=$False,Position=0)][int64]$Size

)
 #Decide what is the type of size
Switch ($Size)
{
{$Size -gt 1TB}
{
Write-Verbose “Convert to TB”
$NewSize = “$([math]::Round(($Size / 1TB),2))TB”
Break
}
{$Size -gt 1GB}
{
Write-Verbose “Convert to GB”
$NewSize = “$([math]::Round(($Size / 1GB),2))GB”
Break
}
{$Size -gt 1MB}
{
Write-Verbose “Convert to MB”
$NewSize = “$([math]::Round(($Size / 1MB),2))MB”
Break
}
{$Size -gt 1KB}
{
Write-Verbose “Convert to MB”
$NewSize = “$([math]::Round(($Size / 1MB),2))MB”
Break
}
Default
{
Write-Verbose “Convert to Bytes”
$NewSize = “$([math]::Round($Size,2))Bytes”
Break
}
}
Return $NewSize

}




#region hyperv

#get OS
$OSname =Get-WmiObject -class Win32_OperatingSystem  
$script:osname = $osname.caption 

#FQDN
$script:FQDN=(Get-WmiObject win32_computersystem).DNSHostName+"."+(Get-WmiObject win32_computersystem).Domain

#HyperV node

#uptime region
   $oscapt= Get-WmiObject win32_operatingsystem
   $script:uptime= (Get-Date) - ($oscapt.ConvertToDateTime($oscapt.lastbootuptime)) 
   $script:updays = $uptime.Days
   $script:uphours=$uptime.Hours 
   $script:upminutes=$uptime.Minutes

  
function gethyperv {



$numbofVM= get-vm 
[string]$numbofVM=$numbofvm.count
$numbofvmrunning=get-vm | where {$_.state -eq 'running'}
[string]$numbofvmrunning=$numbofvmrunning.count

$vmhostinfo = Get-VMHost
$lcpu = $vmhostinfo.logicalprocessorcount
#get cpu model
$pCPUInfo = Get-WmiObject Win32_Processor  #Get CPU Information
$pCPUInfo=$pCPUInfo.name

$cpufullinfo= "" | Select @{n='TotalPhysicalProcessors';e={(,( gwmi Win32_Processor)).count}}, @{n='TotalPhysicalProcessorCores'; e={ (gwmi Win32_Processor | measure -Property NumberOfLogicalProcessors -sum).sum}}, @{n='TotalVirtualCPUs'; e={ (Get-VM | Get-VMProcessor | measure -Property Count -sum).sum }}, @{n='TotalVirtualCPUsInUse'; e={ (Get-VM | Where { $_.State -eq 'Running'} | Get-VMProcessor | measure -Property Count -sum).sum }}, @{n='TotalMSVMProcessors'; e={ (gwmi -ns root\virtualization\v2 MSVM_Processor).count }}, @{n='TotalMSVMProcessorsForVMs'; e={ (gwmi -ns root\virtualization\v2 MSVM_Processor -Filter "Description='Microsoft Virtual Processor'").count }}

#memory 
$os = Get-Ciminstance Win32_OperatingSystem
$pctFree = [math]::Round(($os.FreePhysicalMemory/$os.TotalVisibleMemorySize)*100,2)
$freeGB = [math]::Round($os.FreePhysicalMemory/1mb,1)
$totalGB = [int]($os.TotalVisibleMemorySize/1mb)
$usedGB = $totalGB - $freeGB
$hash = [ordered]@{


'Name'='!p1!'+ $fqdn + '!asterisk!' + '!br!' + '!graysmall!'+ $osname+'!spanend!' +'!pend!' ;
'UPtime' = '!p1!'+"$updays"  + '!graysmallr!'+ ' Days' + '!spanend!' + '!br!' + '!graysmall!'+ "$uphours"  +  ' Hours' + '!spanend!'+'!pend!';
'Logical<br>Processor'='!p1!'+ "$($lcpu)" + '!br!' + '!graysmallabbr!' + "$($pcpuinfo)" + '!quotes!' + 'Cpu info' + '!abbrend!' + '!spanend!'+'!pend!';
'Total<br>VM'='!p1!'+$numbofVM + '!br!' + '!ambersmall!' + $numbofvmrunning + ' Running' + '!spanend!' + '!asterisk!'+'!pend!'
'Used<br>Memory' = '!p1!'+"$($usedGB)" +  '!graysmall!' + ' GB' + '!spanend!' + '!pend!';
'Free<br>Memory' = if ($pctfree -lt  5){'!cellred!'+ '!p1!'+"$($freeGB)" +   ' GB' + '!spanend!' + '!pend!'   }`
elseif($pctfree -lt  10){'!cellyellow!'+ '!p1!'+"$($freeGB)" +   ' GB' + '!spanend!' + '!pend!'  }`
else{'!p1!'+"$($freeGB)" +  '!graysmall!' + ' GB' + '!spanend!' + '!pend!'};
'Total<br>Memory' = if ($pctfree -lt 5){'!p1!'+"$($totalGB)" +  '!graysmall!'+' GB' + '!spanend!' + '!br!'+ '!smallwhitebred!'+  "~% $($pctfree) free" + '!spanend!' +'!pend!'}`
elseif ($pctfree -lt 10){'!p1!'+"$($totalGB)" +  '!graysmall!'+' GB' + '!spanend!' + '!br!'+ '!smallwhitebyellow!'+  "~% $($pctfree) free" + '!spanend!' +'!pend!'}`
else {'!p1!'+"$($totalGB)" +  '!graysmall!'+' GB' + '!spanend!' + '!br!'+ '!graysmall!'+  "~% $($pctfree) free" + '!spanend!' +'!pend!'}`


}
New-Object -TypeName psobject -Property $hash


}


$hypervGeneral = Gethyperv | ConvertTo-Html -Fragment -PreContent '<h2>&diams; Node info </h2>' | Out-String



#endregion
#region virtual machines

function virtualmachines {

$vms = get-vm

    foreach ($vm in $vms) {
    
    
    $snapshot = Get-VMSnapshot -VMName $vm.name

    #region check if disk is missing and put results into $testdiskpath variable that I will use later in final table
    $testdiskpath = ""
    $checkdisk = Get-VMHardDiskDrive $vm
    $testdiskpath = foreach ($path in $checkdisk){
    
    if ((test-path $path.path) -eq $false) {
    $path.path 
    
    }else{
    $testdiskpath = ''
    }       
    }
    #endregion

    #region disk info
    $disks= Get-Vhd $vm.id 
    $diskn = 0
    $data = foreach ($disk in $disks) {
    
  
                            if ($disk.filesize -gt 1TB)
                    {$CurrentFileSize = "{0:N1}" -f ($disk.filesize / 1TB)  +  '!spanend!' + ' TB'                   
                    }
                        elseif ($disk.filesize -gt 1GB) {
                    $CurrentFileSize = "{0:N1}" -f ($disk.filesize / 1GB)  +  '!spanend!' + ' GB'                   
                    }
                    elseif ($disk.filesize -gt 1MB) {
                    $CurrentFileSize = "{0:N1}" -f ($disk.filesize / 1MB)  +  '!spanend!' + ' MB'
                    }
                    elseif ($Size.filesize -gt 1KB) {
                    $CurrentFileSize = "{0:N1}" -f ($disk.filesize / 1KB)  +  '!spanend!' + ' KB'
                    }else{
                        $CurrentFileSize = "{0:N1}" -f ($disk.filesize / 1GB)  +  '!spanend!' + ' GB'                       
                        }

                         if ($disk.size -gt 1TB)
                    {$MaxDisksize = "{0:N1}" -f ($disk.size / 1TB)  +   ' TB)'                    
                    }
                        elseif ($disk.size -gt 1GB) {
                    $MaxDisksize = "{0:N1}" -f ($disk.size / 1GB)   + ' GB)'                   
                    }
                    elseif ($disk.size -gt 1MB) {
                    $MaxDisksize = "{0:N1}" -f ($disk.size / 1MB)   + ' MB)'
                    }
                    elseif ($Size.size -gt 1KB) {
                    $MaxDisksize = "{0:N1}" -f ($disk.size / 1KB)   + ' KB)'
                    }else{
                        $MaxDisksize = "{0:N1}" -f ($disk.size / 1GB)   + ' GB)'                       
                        }
                    
                    #test if disk is differentcing, if yes, add parent path 
                    if ($disk.ParentPath){
                    $parent = 'Parent: ' + $disk.ParentPath 
                    }else{
                    $parent = ''
                    }

                   #count disk number
                   $diskn++               
                   
                  '!divleft!' +  '!Graymedium!' + '!abbr!'+ $disk.path + '!quotes!'+ 'Disk  ' + $diskn  + '!abbrend!' + '!spanend!' + '!asterisk!'                            
                  '!br!'
                   '!graymedium!' + '!arrow1!' + 'CurrentFileSize ' + '!gray!' + $CurrentFileSize +  '  (MaxDiskSize '+ $MaxDisksize +   '!spanend!'
                  '!br!'
                  '!graymedium!' + '!arrow1!' + '!abbr!'+ $parent + '!quotes!'+ $disk.vhdtype + '!abbrend!' +' ' + $disk.vhdformat  +  '!spanend!'
                  '!br!' + '!pend!' + '!divend!'                      
      
                         
                    }
    
    # change $data variable if any disk is unavailable for that VM
    [string]$vmnamecheck = $vm.vmName
    if ($testdiskpath -ne $null) {
    
    $data = $data  + '!divleft!' + '!redsmall!' + '!abbr!'+ $testdiskpath + '!quotes!'+ 'Failed disk(s)' + '!abbrend!' + '!spanend!' + '!asterisk!' + '!divend!'
    
    }else {
        
    $data = $data 
    }

  
 #endregion
    
        $hash = [ordered]@{
            'VM<br>Name'         = '!p1!' +'!abbr!' + $vm.path + '!quotes!'+  $vm.vmname + '!abbrend!'   + '!asterisk!' + '!br!'+ `
             '!graysmall!' + 'Gen'+$vm.generation + ' (' + $vm.version + ')' + '!spanend!' + '!br!'+`
             '!graysmall!' + '!abbr!' + $vm.notes + '!quotes!'+ 'Notes' +'!abbrend!' + '!spanend!'   + '!asterisk!' + '!pend!' ;
             'State'           = if ($vm.state -eq 'running'){'!cellgreen!' + $vm.state + '!spanend!' }elseif($vm.state -eq 'saved')`
                                {'!cellyellow!' + $vm.state + '!spanend!' }elseif($vm.state -eq 'off'){'!cellgray!' + $vm.state + '!spanend!' }`
                                elseif($vm.state -eq 'failed'){'!cellred!' + $vm.state + '!spanend!' }else{$vm.state};
            
            'Uptime'            = if ($vm.uptime -eq "00:00:00"){'!graysmall!'+'Stopped'+'!spanend!'}else{(($vm.uptime).days).tostring() +  '!graysmall!' +' day(s)' + '!spanend!' + '!br!' + '!graysmall!' +  (($vm.uptime).hours).tostring() +' H' + '!br!' +(($vm.uptime).minutes).tostring()+' Min' + '!spanend!'};
            'Check<br>point'= if ($snapshot.count -le 0){'!graysmall!'+'!abbr!' + $snapshot.count + ' Checkpoint(s)' + '!quotes!'+  'No' + '!abbrend!' }`
                                else{'!byellow!'+'!abbr!' + "$($snapshot.count) Checkpoint(s)" + '!quotes!'+  'Yes' + '!abbrend!' + '!spanend!'};
            'Integration<br>services' = if ($vm.integrationservicesstate -like '*update required*'){'!graysmall!'+'!abbr!' + "$($vm.integrationservicesversion)" + '!quotes!' + "$($vm.integrationservicesstate)" + '!abbrend!' + '!spanend!'}else{'!graysmall!' + 'N/A' + '!spanend!'};
            'vCPU'= '!graysmall!'+ $vm.processorCount;
            'Replica<br>health'=if($vm.replicationhealth -eq 'notapplicable'){'N/A'}else{$vm.replicationhealth};
            'Disk'=  "$data"
            
        
         
        }
        New-Object -TypeName PSObject -Property $hash

    }


}

#if there is not VM , $vms is empty



$vms = virtualmachines | ConvertTo-Html -Fragment -PreContent '<h2>&diams; Virtual machines </h2>' | Out-String


#endregion

$server = $env:COMPUTERNAME

$premain = "<h1>&diams; $($server) </h1>" | Out-String

#region Final report
if ($hypervGeneral){
}else{
$hypervGeneral = " "
}

if ($vms){
}else{
$vms = " "
}

$finalreport = convertto-html -as table -Body $style -PreContent $premain -PostContent $serverCoreinfo,   $serverinfo,$network,$diskafter, "<hr class='style-six'>",$hypervgeneral,$vms,"<hr class='style-six'>", $domain2, $forest, $forest1, "<hr class='style-six'>", `
     $sharedfolders, "<hr class='style-six'>", $cpuinfo, $memoryinfo, "<hr class='style-six'>", $serverfeatures, "<hr class='style-six'>"
$finalreport | out-file  c:\yw-data\serverinfo.html
(Get-Content C:\yw-data\serverinfo.html) -replace "!br!", "<br>" | Set-Content C:\yw-data\serverinfo.html

#CELL - RED 
(Get-Content $wdir\serverinfo.html) -replace '<td>!cellred!', '<td bgcolor=#ff9999><SPAN STYLE="font-size:12px;color:#ffffff">' | Set-Content $wdir\serverinfo.html

#CELL - GREEN 
(Get-Content $wdir\serverinfo.html) -replace '<td>!cellgreen!', '<td bgcolor=#8EF38B><SPAN STYLE="font-size:12px;color:#555F55">' | Set-Content $wdir\serverinfo.html

#CELL - yellow 
(Get-Content $wdir\serverinfo.html) -replace '<td>!cellyellow!', '<td bgcolor=#ffcc80><SPAN STYLE="font-size:12px;color:#ffffff">' | Set-Content $wdir\serverinfo.html
#CELL - gray 
(Get-Content $wdir\serverinfo.html) -replace '<td>!cellgray!', '<td bgcolor=#b3b3b3><SPAN STYLE="font-size:12px;color:#ffffff">' | Set-Content $wdir\serverinfo.html


# GRAY color - Big font
(Get-Content $wdir\serverinfo.html) -replace '!gray!', '<span style="font-size:19px;color:#b3b3b3;font-weight:bold;">' | Set-Content $wdir\serverinfo.html

# GRAY color - small font
(Get-Content $wdir\serverinfo.html) -replace '!graysmall!', '<span style="font-size:10px;color:#b3b3b3;">' | Set-Content $wdir\serverinfo.html
(Get-Content $wdir\serverinfo.html) -replace '!graymedium!', '<span style="font-size:12px;color:#b3b3b3;">' | Set-Content $wdir\serverinfo.html

# RED background and  white font - small font
(Get-Content $wdir\serverinfo.html) -replace '!smallwhitebred!', '<span style="background-color:#ff8080;font-size:10px;color:#ffffff;">' | Set-Content $wdir\serverinfo.html
# yellow background and  white font - small font
(Get-Content $wdir\serverinfo.html) -replace '!smallwhitebyellow!', '<span style="background-color:#ffb84d;font-size:10px;color:#ffffff;">' | Set-Content $wdir\serverinfo.html



# GRAY color - small font with right text indent
(Get-Content $wdir\serverinfo.html) -replace '!graysmallr!', '<span style="font-size:10px;color:#b3b3b3;text-indent:200px;">' | Set-Content $wdir\serverinfo.html

# GRAY color - small font with hyperlink
(Get-Content $wdir\serverinfo.html) -replace '!graysmallabbr!', '<span style="font-size:10px;color:#b3b3b3;"><abbr title="' | Set-Content $wdir\serverinfo.html
(Get-Content $wdir\serverinfo.html) -replace '!quotes!', '">' | Set-Content $wdir\serverinfo.html
(Get-Content $wdir\serverinfo.html) -replace '!abbrend!', '</abbr>' | Set-Content $wdir\serverinfo.html
#hyperlink with no change in font and color
(Get-Content $wdir\serverinfo.html) -replace '!abbr!', '<abbr title="' | Set-Content $wdir\serverinfo.html



# right arrow
(Get-Content $wdir\serverinfo.html) -replace '!arrow!', '&#8594 ' | Set-Content $wdir\serverinfo.html

(Get-Content $wdir\serverinfo.html) -replace '!arrow1!', '&#10148 ' | Set-Content $wdir\serverinfo.html


#abbreviation



# AMBER color
(Get-Content $wdir\serverinfo.html) -replace '!amber!', '<span style="color:#E5B053;font-size:14px">' | Set-Content $wdir\serverinfo.html
# AMBER with small font as gray
(Get-Content $wdir\serverinfo.html) -replace '!ambersmall!', '<span style="color:#E5B053;font-size:10px">' | Set-Content $wdir\serverinfo.html

# RED with small font as gray
(Get-Content $wdir\serverinfo.html) -replace '!redsmall!', '<span style="color:#ff4d4d;font-size:10px">' | Set-Content $wdir\serverinfo.html

#SPAN END
(Get-Content $wdir\serverinfo.html) -replace '!spanend!', '</span>' | Set-Content $wdir\serverinfo.html

#formating for *
(Get-Content $wdir\serverinfo.html) -replace '!asterisk!', '<span style="color:#ffbf00;font-size:20px"><sub> *</sub></span>' | Set-Content $wdir\serverinfo.html

# paragraph space 1, where there is no td formating
(Get-Content $wdir\serverinfo.html) -replace '!p1!', '<p style="line-height:0.8">' | Set-Content $wdir\serverinfo.html
(Get-Content $wdir\serverinfo.html) -replace '!p2!', '<p style="line-height:1.3;">' | Set-Content $wdir\serverinfo.html

#paragraph end
(Get-Content $wdir\serverinfo.html) -replace '!pend!', '</p>' | Set-Content $wdir\serverinfo.html

#backgrounds

#background red
(Get-Content $wdir\serverinfo.html) -replace '!bred!', '<span style="background-color:#ff8080;>' | Set-Content $wdir\serverinfo.html
(Get-Content $wdir\serverinfo.html) -replace '!byellow!', '<span style="background-color:#ffcc80";>' | Set-Content $wdir\serverinfo.html

#DIV arrange left
(Get-Content $wdir\serverinfo.html) -replace '!divleft!', '<div style="text-align:left;">' | Set-Content $wdir\serverinfo.html
(Get-Content $wdir\serverinfo.html) -replace '!divend!', '</div>' | Set-Content $wdir\serverinfo.html

#Invoke-item "$wdir\serverinfo.html"

`
# this is right arrow &#10148 more arrows https://websitebuilders.com/tools/html-codes/arrows/ https://www.w3schools.com/html/html_symbols.asp https://www.w3schools.com/charsets/ref_utf_arrows.asp
#http://html-css-js.com/html/character-codes/arrows/

# this is for hover       <abbr title=""$($vmDiskPath)
#ovako se koristi abreviation odnosno kada predjemo misem preko link-a   <abbr title=""$($vmDiskPath)"">$($vmDiskName)<span style=""font-size:10px;color:orange""> *</span></abbr>

# http://wiki.webperfect.ch/index.php?title=Hyper-V:_Capacity_Report



#region variables for email

$compinfo = Get-Content "$wdir\as.txt"
$subject= $compinfo
$to = "support@yanceyworks.com"
$from = "donotreply@yanceyworks.com"
#endregion

#region Credentials
$username= "donotreply@yanceyworks.com"
$password = "Dolisnotoko476*#" | ConvertTo-SecureString -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential($username,$password)
#endregion

#regionsend email
Send-MailMessage -to $to -From $from -Subject $subject `
-Bodyashtml "<h2>Server report</h2>"  -SmtpServer smtp.office365.com `
-Credential $cred -UseSsl -Attachments "$wdir\serverinfo.html"
#endregion