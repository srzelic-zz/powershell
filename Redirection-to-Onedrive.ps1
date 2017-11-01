<#
.SYNOPSIS
Redirect user's known data to OneDrive or business
.DESCRIPTION
Script will create folders in Onedrive if not exist (documents, music, pictures, desktop, favorites,videos) and redirect to onedrive and move data to a new location
.NOTES  
When you run a script, you will first get html overview report what you have in source destination, where your files are stored now, how much data you have before you start migration Script predict many situation like luck of write rights in a destination etc. 
When script is finished , you will get full information about how much data are in destionation, how much data was stored in source so you can compare that.
You will get information how many files/directories are failed during migration if any, how much data are left in source (for example if you already have newer files in destination) etc.

File Name  : Redirection-to-Onedrive.ps1
Author: Sasa Zelic 
https://www.linkedin.com/in/sasa-zelic-14a1533b/

#>

#region html style
$style = @"
<style>

h1 { text-align: center;background: #7aad7a ;color:#383838}
h2,h3,h4, h5, th{ text-align: center;font-size:23px; color:#383838; }
h10 { text-align: center;background: #7aad7a ;color:#383838;font-size:16px;}
table { margin: auto; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey; }
th { background: #7aad7a; color: #383838; max-width: 900px; padding: 5px 10px;background-color: #7aad7a;font-size:15px; }
td {text-align: left;font-size: 11px; padding: 5px 20px; color: black ;}
tr { background: white }
tr:nth-child(even) { background: #d9e4d9; }
tr:nth-child(odd) { background: #f4f7f4; }

</style>
"@
#endregion


#region functions
function getsize {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Path
    )
        
    "{0:N3} MB" -f ((Get-ChildItem -Path $path -Recurse  -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum / 1MB)
}

function Get-FolderSize {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        $Path,
        [ValidateSet("KB", "MB", "GB")]
        $Units = "MB"
    )
    if ( (Test-Path $Path) -and (Get-Item $Path).PSIsContainer ) {
        $Measure = Get-ChildItem $Path -Recurse -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum
        $Sum = $Measure.Sum / "1$Units"
        $sum = "{0:N3}" -f $sum
        [PSCustomObject]@{
            "Size($Units)" = $Sum
            "Path"         = $Path
            
        }
    }
}
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
elseif (test-path 'e:\yw-data\') {
    $wdir = 'h:\yw-data'
}
else {
    New-Item -ItemType directory  -Path 'C:\yw-data'
    $wdir = 'C:\yw-data'
} 
#endregion


#region Source User's data location before migration

$music = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -name 'my music').'my music'
$docs = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -name 'personal').'personal'
$desktop = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -name 'desktop').'desktop'
$videos = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -name 'my video').'my video'
$pictures = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -name 'my pictures').'my pictures'
$favorites = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -name 'favorites').'favorites'

#endregion
#region Info about location before migration
function dataloc {
    $array = @($docs, $desktop, $music, $videos, $pictures,$Favorites) 

    foreach ($ar in $array) {
        $size = Get-FolderSize -Path $ar -Units GB
        $hash = [ordered]@{
            'Size(GB)'  = $size.'Size(GB)'
            'Name'      = $ar.split('\')[-1]
            'Locations' = $ar;
            'Note'      = if ($ar -like "*OneDrive*") {'Stored in OD!'} else {'Not stored in OD'}
        
        }
        New-Object -TypeName psobject -Property $hash

    }     
}

$pre = "SOURCE - before migration"
$locationsbefore = dataloc | ConvertTo-Html  -Fragment -PreContent "<h2>&diams; $pre</h2>" | Out-string

$preoverview = "This is overview before migration. Take a look and close this window"
$overview = dataloc | ConvertTo-Html -Body $style -PreContent "<h2>&diams; $preoverview</h2>"
$overview = $overview -replace "Not stored in OD", "<font color='#993333'>Not stored in OD</font>"
$overview | Out-File "$wdir\overview.html"
invoke-item "$wdir\overview.html"
Write-Host "Press enter to continue migration" -ForegroundColor Green
Write-Host " "
read-host 'Continue?'
#endregion

#region determine Onedrive as a destination migration
$folderspath = Get-ChildItem -Directory -Path $env:userprofile -Depth 1 | Where-Object {$_.name -like "*onedrive -*"} | Select-Object FullName
#check if onedrive for business exist or there are more than one
if (!($folderspath)) {
    start-sleep 1
    Write-Host "Onedrive folder is not configured. Script will now exit" -ForegroundColor Yellow
    exit
}

elseif ($folderspath.Count -gt 1 ) {
    write-host "There are multiple onedrive for business folders!!!. " -ForegroundColor yellow
    write-host " "
    start-sleep 1
    write-host "Choose which one do you want to use for migration (Enter 1 for first, 2 for second etc.)"
    write-host " "
    $path1 = ($folderspath[0])
    $path2 = ($folderspath[1])
    $path3 = ($folderspath[2])
    write-host "$($path1.FullName)" -ForegroundColor Green
    write-host "$($path2.fullname)" -ForegroundColor Green
    if ($path3) {
        write-host "$($path3.fullname)" -ForegroundColor Green
    }
    write-host " "
    $confirmation = read-host 
    if ($confirmation -eq 1) {
        $folderspath = $path1.FullName
        Write-Host "Onedrive location is " -ForegroundColor Yellow -NoNewline
        write-host " $($folderspath)" -BackgroundColor Yellow -ForegroundColor Black 
        $answer = read-host "Do you want to migrate to $folderspath (yes/no)"
        if ($answer -eq 'yes'){
            write-host "Migration started"
            }else {
            write-host "Migration will now exit"

            }
    }
    elseif ($confirmation -eq 2) {
        $folderspath = $path2.FullName
        
        Write-Host "Onedrive location is " -ForegroundColor Yellow -NoNewline
        write-host " $($folderspath)" -BackgroundColor Yellow -ForegroundColor Black
        write-host " " 
        $answer = read-host "Do you want to migrate to $folderspath (yes/no)"
        if ($answer -eq 'yes'){
            write-host "Migration started"
            }else {
            write-host "Migration will now exit"

            }
    }
    else { 
        start-sleep 1
        write-host "You have chosen non existing option, script will exit now !!!" -ForegroundColor Yellow
        exit
    }
    #read-host "Script will exist now. Please pre Enter."
    
}
else {
    $folderspath = $folderspath.FullName
    start-sleep 1
    Write-Host "Onedrive path is  " -ForegroundColor Yellow -NoNewline
    write-host " $folderspath" -BackgroundColor Yellow -ForegroundColor Black 
    write-host " "
    $answer = read-host "Do you want to migrate to $folderspath (yes/no)"
    if ($answer -eq 'yes'){
        write-host "Migration started"
        }else {
        write-host "Migration will now exit"

        }
}

#endregion

#region user's data locations
$folders = @('Documents', 'Desktop', 'Pictures', 'Videos', 'Favorites', 'Music')

[string]$docstring = $folders[0]
[string]$desktopstring = $folders[1]
[string]$picturesstring = $folders[2]
[string]$videosstring = $folders[3]
[string]$favoritesstring = $folders[4]
[string]$musicstring = $folders[5]


$old = Get-ItemProperty  'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' 
$olddesktop = $old.Desktop
$oldmydocs = $old.Personal
$oldpictures = $old.'My Pictures'
$oldmusic = $old.'My Music'
$oldfavorites = $old.Favorites
$oldvideos = $old.'My Video'

#new destination locations
$newdesktop = $folderspath + '\' + $folders[1]
$newdocs = $folderspath + '\' + $folders[0]
$newpictures = $folderspath + '\' + $folders[2]
$newmusic = $folderspath + '\' + $folders[5]
$newfavorites = $folderspath + '\' + $folders[4]
$newvideos = $folderspath + '\' + $folders[3]
#endregion


#region Creating folders in onedrive and migrating
$folders = @($docstring, $musicstring, $picturesstring, $videosstring, $desktopstring, $favoritesstring)
$er = @() #collect number of successful migration .With this I determine whether to reboot pc or not 
$alreadyinOD = @() #With this array, I format Migration column in html report
foreach ($fol in $folders) {
    
    if (!(Test-Path "$folderspath\$fol" )) {
     
        try {        
            New-Item -Path $folderspath -ItemType Directory -Name $fol -ErrorAction Stop | Out-Null
            Write-Host " "
            start-sleep 1
            write-host "$fol" -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " folder is created" -ForegroundColor Yellow
            Start-Sleep 1
            
            
            #region documents migration                              
            if (($oldmydocs -notlike $folderspath + '\' + $docstring) -and ($fol -eq $docstring)) {
                
                try {
                                  
                    #with $m I build column in html to determine if migration is done or not.
                    $d = $true
                    #test write permissions on a migration destination folder before making registry changes
                    try {"" | out-file   $newdocs\write-testfileyw12347322322234330983.txt -ErrorAction Stop} catch {
                        write-host "It is not possible to write to destination folder $newdocs" -ForegroundColor Red
                    }
                    "" | out-file   $newdocs\write-testfileyw12347322322234330983.txt -ErrorAction Stop
                                            
                
                    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'personal' $newdocs -ErrorAction Stop
                    Set-ItemProperty -path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'personal' $newdocs -ErrorAction Stop
                    $USF = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop
                    $USFpath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' 
                                            
                    if ($USF.'{F42EE2D3-909F-4907-8871-4C22FC0BF756}') {
                        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name `
                            '{F42EE2D3-909F-4907-8871-4C22FC0BF756}' $newdocs -ErrorAction Stop
                    } 
                
                    else {
                        New-ItemProperty -Path $USFpath -Name '{F42EE2D3-909F-4907-8871-4C22FC0BF756}' -Value $newdocs -PropertyType expandstring -Force -ErrorAction Stop | Out-Null
                                                   
                    }
                    Start-Sleep 1
                    Write-Host $docstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " started migrating" -ForegroundColor Yellow
                                        
                    Robocopy $oldmydocs $newdocs /move /e /xo /xj /r:0 /ns /nc /np /njh  /log:"$wdir\robocopydocs.log" | out-null #use /fft if you copy over the network because of latency
                               
                    Write-Host $docstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " folder finished migrating" -ForegroundColor Yellow
                    Write-Host $docstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " migration result:" -ForegroundColor Yellow
                    #region robocopy migration result
                    $DOCSrobocopy = ''
                    $DOCSrobocopy = Get-Content "$wdir\robocopydocs.log" 
                    $DOCSrobocopy = $DOCSrobocopy  -match '^(?= *?\b(Total|Dirs|Files)\b)((?!    Files).)*$'
                    #convert to array
                    $DOCSrobocopyresult = @()
                    foreach ($line in $DOCSrobocopy ) {
                        $DOCSrobocopyresult += $line
                    }
                    $docsfailedfile = [int](($DOCSrobocopyresult[2] -split "\s+")[7])
                    $docsfaileddirs = [int](($DOCSrobocopyresult[1] -split "\s+")[7])
                    if (($docsfailedfile -eq 0) -and ($docsfaileddirs -eq 0) ){

                        write-host "There is no failed files or directories during the migration"
                    }
                    else{
                        write-host "There are some failed files or directories. Please see robocopy logs located in $wdir\robocopydocs.log" -ForegroundColor Red
                    }
                    #endregion
                    start-sleep 1
                    $er += $fol 
                        
                } 
                               
                catch {
                    write-host "$fol" -ForegroundColor White -BackgroundColor Red -NoNewline
                    write-host " migration will be skipped" -ForegroundColor Red
                    $d = $false
                    
                } 
                               
            }
            #endregion
            #region music migration
            elseif (($oldmusic -notlike $folderspath + '\' + $musicstring) -and ($fol -eq $musicstring)) {

                try {
                  
                    #with $m I build column in html to determine if migration is done or not.
                    $m = $true
                    #test write permissions on a migration destination folder before making registry changes
                    try {"" | out-file   $newmusic\write-testfileyw12347322322234330983.txt -ErrorAction Stop} catch {
                        write-host "It is not possible to write to destination folder $newmusic" -ForegroundColor Red
                    }
                    "" | out-file   $newmusic\write-testfileyw12347322322234330983.txt -ErrorAction Stop

                    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'my music' $newmusic -ErrorAction Stop
                    Set-ItemProperty -path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'my music' $newmusic -ErrorAction Stop
                    $USF = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop
                    $USFpath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' 
                            
                    if ($USF.'{A0C69A99-21C8-4671-8703-7934162FCF1D}') {
                        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name `
                            '{A0C69A99-21C8-4671-8703-7934162FCF1D}' $newmusic -ErrorAction Stop
                    } 

                    else {
                        New-ItemProperty -Path $USFpath -Name '{A0C69A99-21C8-4671-8703-7934162FCF1D}' -Value $newmusic -PropertyType expandstring -Force -ErrorAction Stop | Out-Null
                                   
                    }
                    Start-Sleep 1
                    Write-Host $musicstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " started migrating" -ForegroundColor Yellow
                        
                    Robocopy $oldmusic $newmusic /move /e /xo /xj /r:0 /ns /nc /np /njh /log:"$wdir\robocopymusic.log" | out-null #use /fft if you copy over the network because of latency
                    Write-Host $musicstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " folder finished migrating" -ForegroundColor Yellow
                    Write-Host $musicstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " migration result:" -ForegroundColor Yellow
                    #region robocopy migration result
                    $musicrobocopy = ''
                    $musicrobocopy = Get-Content "$wdir\robocopymusic.log" 
                    $musicrobocopy = $musicrobocopy  -match '^(?= *?\b(Total|Dirs|Files)\b)((?!    Files).)*$'
                    #convert to array
                    $musicrobocopyresult = @()
                    foreach ($line in $musicrobocopy ) {
                        $musicrobocopyresult += $line
                    }
                    $musicfailedfile = [int](($musicrobocopyresult[2] -split "\s+")[7])
                    $musicfaileddirs = [int](($musicrobocopyresult[1] -split "\s+")[7])
                    if (($musicfailedfile -eq 0) -and ($musicfaileddirs -eq 0) ){

                        write-host "There is no failed files or directories during the migration"
                    }
                    else{
                        write-host "There are some failed files or directories. Please see robocopy logs located in $wdir\robocopymusic.log" -ForegroundColor Red
                    }
                    #endregion
                    start-sleep 4
                    $er += $fol
           
                } 
               
                catch {
                    write-host "$fol" -ForegroundColor White -BackgroundColor Red -NoNewline
                    write-host " migration will be skipped" -ForegroundColor Red
                    $m = $false
                    
                } 
               
            }

            #endregion
            #region pictures migration          
            elseif (($oldpictures -notlike $folderspath + '\' + $picturesstring) -and ($fol -eq $picturesstring)) {
                
                try {
                                  
                    #with $p I build column in html to determine if migration is done or not.
                    $p = $true
                    #test write permissions on a migration destination folder before making registry changes
                    try {"" | out-file   $newpictures\write-testfileyw12347322322234330983.txt -ErrorAction Stop} catch {
                        write-host "It is not possible to write to destination folder $newpictures" -ForegroundColor Red
                    }
                    "" | out-file   $newpictures\write-testfileyw12347322322234330983.txt -ErrorAction Stop
                                            
                
                    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'my pictures' $newpictures -ErrorAction Stop
                    Set-ItemProperty -path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'my pictures' $newpictures -ErrorAction Stop
                    $USF = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop
                    $USFpath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' 
                                            
                    if ($USF.'{0DDD015D-B06C-45D5-8C4C-F59713854639}') {
                        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name `
                            '{0DDD015D-B06C-45D5-8C4C-F59713854639}' $newpictures -ErrorAction Stop
                    } 
                
                    else {
                        New-ItemProperty -Path $USFpath -Name '{0DDD015D-B06C-45D5-8C4C-F59713854639}' -Value $newpictures -PropertyType expandstring -Force -ErrorAction Stop | Out-Null
                                                   
                    }
                    Start-Sleep 1
                    Write-Host $picturesstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " started migrating" -ForegroundColor Yellow
                                        
                    Robocopy $oldpictures $newpictures /move /e /xo /xj /r:0 /ns /nc /np /njh /log:"$wdir\robocopypictures.log" | out-null #use /fft if you copy over the network because of latency
                                 
                    Write-Host $picturesstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " folder finished migrating" -ForegroundColor Yellow
                    Write-Host $picturesstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " migration result:" -ForegroundColor Yellow
                    
                    $picturesrobocopy = ''
                    $picturesrobocopy = Get-Content "$wdir\robocopypictures.log"
                    $picturesrobocopy = $picturesrobocopy  -match '^(?= *?\b(Total|Dirs|Files)\b)((?!    Files).)*$'
                    #convert to array
                    $picturesrobocopyresult = @()
                    foreach ($line in $picturesrobocopy ) {
                        $picturesrobocopyresult += $line
                    }
                    $picturesfailedfile = [int](($picturesrobocopyresult[2] -split "\s+")[7])
                    $picturesfaileddirs = [int](($picturesrobocopyresult[1] -split "\s+")[7])
                    if (($picturesfailedfile -eq 0) -and ($picturesfaileddirs -eq 0) ){

                        write-host "There is no failed files or directories during the migration"
                    }
                    else{
                        write-host "There are some failed files or directories. Please see robocopy logs located in $wdir\robocopypictures.log" -ForegroundColor Red
                    }
                    start-sleep 4
                    $er += $fol 
                           
                } 
                               
                catch {
                    write-host "$fol" -ForegroundColor White -BackgroundColor Red -NoNewline
                    write-host " migration will be skipped" -ForegroundColor Red
                    $p = $false
                    $musicreport = "Error in migration or already in OneDrive"
                } 
                               
            }
            #endregion
            #region desktop migration                                
            elseif (($olddesktop -notlike $folderspath + '\' + $desktopstring) -and ($fol -eq $desktopstring)) {
                    
                try {
                                      
                    #with $de I build column in html to determine if migration is done or not.
                    $de = $true
                    #test write permissions on a migration destination folder before making registry changes
                    try {"" | out-file   $newdesktop\write-testfileyw12347322322234330983.txt -ErrorAction Stop} catch {
                        write-host "It is not possible to write to destination folder $newdesktop" -ForegroundColor Red
                    }
                    "" | out-file   $newdesktop\write-testfileyw12347322322234330983.txt -ErrorAction Stop
                                                
                    
                    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'desktop' $newdesktop -ErrorAction Stop
                    Set-ItemProperty -path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'desktop' $newdesktop -ErrorAction Stop
                    $USF = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop
                    $USFpath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' 
                                                
                    if ($USF.'{754AC886-DF64-4CBA-86B5-F7FBF4FBCEF5}') {
                        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name `
                            '{754AC886-DF64-4CBA-86B5-F7FBF4FBCEF5}' $newdesktop -ErrorAction Stop
                    } 
                    
                    else {
                        New-ItemProperty -Path $USFpath -Name '{754AC886-DF64-4CBA-86B5-F7FBF4FBCEF5}' -Value $newdesktop -PropertyType expandstring -Force -ErrorAction Stop | Out-Null
                                                       
                    }
                    start-sleep 1
                    Write-Host $desktopstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " started migrating" -ForegroundColor Yellow
                                            
                    Robocopy $olddesktop $newdesktop /move /e /xo /xj /r:0 /ns /nc /np /njh /log:"$wdir\robocopydesktop.log" | out-null #use /fft if you copy over the network because of latency
                                       
                                        
                                       
                    Write-Host $desktopstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " folder finished migrating" -ForegroundColor Yellow
                    Write-Host $desktopstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " migration result:" -ForegroundColor Yellow
                    # robocopy migration result
                    $desktoprobocopy = ''
                    $desktoprobocopy = Get-Content "$wdir\robocopydesktop.log"
                    $desktoprobocopy = $desktoprobocopy  -match '^(?= *?\b(Total|Dirs|Files)\b)((?!    Files).)*$'
                    #convert to array
                    $desktoprobocopyresult = @()
                    foreach ($line in $desktoprobocopy ) {
                        $desktoprobocopyresult += $line
                    }
                    $desktopfailedfile = [int](($desktoprobocopyresult[2] -split "\s+")[7])
                    $desktopfaileddirs = [int](($desktoprobocopyresult[1] -split "\s+")[7])
                    if (($desktopfailedfile -eq 0) -and ($desktopfaileddirs -eq 0) ){

                        write-host "There is no failed files or directories during the migration"
                    }
                    else{
                        write-host "There are some failed files or directories. Please see robocopy logs located in $wdir\robocopydesktop.log" -ForegroundColor Red
                    }
                    start-sleep 2
                    $er += $fol  
                               
                } 
                                   
                catch {
                    write-host "$fol" -ForegroundColor White -BackgroundColor Red -NoNewline
                    write-host " migration will be skipped" -ForegroundColor Red
                    $de = $false
                    $musicreport = "Error in migration or already in OneDrive"
                } 
                                   
            }
            #endregion 
            #region videos migration                                                                
            elseif (($oldvideos -notlike $folderspath + '\' + $videosstring) -and ($fol -eq $videosstring)) {
                    
                try {
                                      
                    #with $v I build column in html to determine if migration is done or not.
                    $v = $true
                    #test write permissions on a migration destination folder before making registry changes
                    try {"" | out-file   $newvideos\write-testfileyw12347322322234330983.txt -ErrorAction Stop} catch {
                        write-host "It is not possible to write to destination folder $newvideos" -ForegroundColor Red
                    }
                    "" | out-file   $newvideos\write-testfileyw12347322322234330983.txt -ErrorAction Stop
                                                
                    
                    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'my video' $newvideos -ErrorAction Stop
                    Set-ItemProperty -path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'my video' $newvideos -ErrorAction Stop
                    $USF = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop
                    $USFpath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' 
                                                
                    if ($USF.'{35286A68-3C57-41A1-BBB1-0EAE73D76C95}') {
                        Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name `
                            '{35286A68-3C57-41A1-BBB1-0EAE73D76C95}' $newvideos -ErrorAction Stop
                    } 
                    
                    else {
                        New-ItemProperty -Path $USFpath -Name '{35286A68-3C57-41A1-BBB1-0EAE73D76C95}' -Value $newvideos -PropertyType expandstring -Force -ErrorAction Stop | Out-Null
                                                       
                    }
                    start-sleep 1
                    Write-Host $videosstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " started migrating" -ForegroundColor Yellow
                                            
                    Robocopy $oldvideos $newvideos /move /e /xo /xj /r:0 /ns /nc /np /njh /log:"$wdir\robocopyvideo.log" | out-null #use /fft if you copy over the network because of latency
                                    
                                      
                    Write-Host $videosstring  -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " folder finished migrating" -ForegroundColor Yellow
                    Write-Host $videosstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " migration result:" -ForegroundColor Yellow
                    # robocopy migration result
                    $videosrobocopy = ''
                    $videosrobocopy = Get-Content "$wdir\robocopyvideo.log" 
                    $videosrobocopy = $videosrobocopy  -match '^(?= *?\b(Total|Dirs|Files)\b)((?!    Files).)*$'
                    #convert to array
                    $videosrobocopyresult = @()
                    foreach ($line in $videosrobocopy ) {
                        $videosrobocopyresult += $line
                    }
                    $videosfailedfile = [int](($videosrobocopyresult[2] -split "\s+")[7])
                    $videosfaileddirs = [int](($videosrobocopyresult[1] -split "\s+")[7])
                    if (($videosfailedfile -eq 0) -and ($videosfaileddirs -eq 0) ){

                        write-host "There is no failed files or directories during the migration"
                    }
                    else{
                        write-host "There are some failed files or directories. Please see robocopy logs located in $wdir\robocopyvideo.log" -ForegroundColor Red
                    }
                    start-sleep 2
                    $er += $fol  
                               
                } 
                                   
                catch {
                    write-host "$fol" -ForegroundColor White -BackgroundColor Red -NoNewline
                    write-host " migration will be skipped" -ForegroundColor Red
                    $v = $false
                    $videoreport = "Error in migration"
                } 
                                   
            }

            #endregion                                                               
            #region favorites migration                                                                
            elseif (($oldfavorites -notlike $folderspath + '\' + $favoritesstring) -and ($fol -eq $favoritesstring)) {
                    
                try {
                                        
                    #with $f I build column in html to determine if migration is done or not.
                    $f = $true
                    #test write permissions on a migration destination folder before making registry changes
                    try {"" | out-file   $newfavorites\write-testfileyw12347322322234330983.txt -ErrorAction Stop} catch {
                        write-host "It is not possible to write to destination folder $newfavorites" -ForegroundColor Red
                    }
                    "" | out-file   $newfavorites\write-testfileyw12347322322234330983.txt -ErrorAction Stop
                                                
                    
                    Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'favorites' $newfavorites -ErrorAction Stop
                    Set-ItemProperty -path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'favorites' $newfavorites -ErrorAction Stop
                    $USF = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop
                    $USFpath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' 
                                                
                                        
                    start-sleep 1
                    Write-Host $favoritesstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " started migrating" -ForegroundColor Yellow
                                            
                    Robocopy $oldfavorites $newfavorites /move /e /xo /xj /r:0 /ns /nc /np /njh /log:"$wdir\robocopyfavorites.log" | out-null #use /fft if you copy over the network because of latency
                                      
                                      
                    Write-Host $favoritesstring  -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " folder finished migrating" -ForegroundColor Yellow
                    #
                    Write-Host $favoritesstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
                    write-host " migration result:" -ForegroundColor Yellow
                   # robocopy migration result
                   $favoritesrobocopy = ''
                   $favoritesrobocopy = Get-Content "$wdir\robocopyfavorites.log" 
                   $favoritesrobocopy = $favoritesrobocopy  -match '^(?= *?\b(Total|Dirs|Files)\b)((?!    Files).)*$'
                   #convert to array
                   $favoritesrobocopyresult = @()
                   foreach ($line in $favoritesrobocopy ) {
                       $favoritesrobocopyresult += $line
                   }
                   $favoritesfailedfile = [int](($favoritesrobocopyresult[2] -split "\s+")[7])
                   $favoritesfaileddirs = [int](($favoritesrobocopyresult[1] -split "\s+")[7])
                   if (($favoritesfailedfile -eq 0) -and ($favoritesfaileddirs -eq 0) ){

                       write-host "There is no failed files or directories during the migration"
                   }
                   else{
                       write-host "There are some failed files or directories. Please see robocopy logs located in $wdir\robocopyfavorites.log" -ForegroundColor Red
                   }
                    start-sleep 2
                                
                } 
                                    
                catch {
                    write-host "$fol" -ForegroundColor White -BackgroundColor Red -NoNewline
                    write-host " migration will be skipped" -ForegroundColor Red
                    $f = $false
                    $favoritesreport = "Error in migration"
                    #enable-NTFSAccessInheritance -Path 'C:\users\sasa\onedrive - sasa\Music'
                } 
                                    
            }    
            
            #endregion                                                               

                       
        }
        catch [System.UnauthorizedAccessException] {
            Write-Host "Access denied"
            start-sleep 1
        }
        catch {
            write-host "Unknown error"
        }
        
        
    }
    #endregion 


    #region if folder exists in destination
    
    #region elseif documents migration                                                                
    elseif (($oldmydocs -notlike $folderspath + '\' + $docstring) -and ($fol -eq $docstring)) {
        
        try {
                          
            #with $m I build column in html to determine if migration is done or not.
            $d = $true
            #test write permissions on a migration destination folder before making registry changes
            try {"" | out-file   $newdocs\write-testfileyw12347322322234330983.txt -ErrorAction Stop} catch {
                write-host "It is not possible to write to destination folder $newdocs" -ForegroundColor Red
            }
            "" | out-file   $newdocs\write-testfileyw12347322322234330983.txt -ErrorAction Stop
                                    
        
            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'personal' $newdocs -ErrorAction Stop
            Set-ItemProperty -path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'personal' $newdocs -ErrorAction Stop
            $USF = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop
            $USFpath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' 
                                    
            if ($USF.'{F42EE2D3-909F-4907-8871-4C22FC0BF756}') {
                Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name `
                    '{F42EE2D3-909F-4907-8871-4C22FC0BF756}' $newdocs -ErrorAction Stop
            } 
        
            else {
                New-ItemProperty -Path $USFpath -Name '{F42EE2D3-909F-4907-8871-4C22FC0BF756}' -Value $newdocs -PropertyType expandstring -Force -ErrorAction Stop | Out-Null
                                           
            }
            Start-Sleep 1
            Write-Host $docstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " started migrating" -ForegroundColor Yellow
                                
            Robocopy $oldmydocs $newdocs /move /e /xo /xj /r:0 /ns /nc /np /njh  /log:"$wdir\robocopydocs.log" | out-null #use /fft if you copy over the network because of latency
                       
            Write-Host $docstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " folder finished migrating" -ForegroundColor Yellow
            Write-Host $docstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " migration result:" -ForegroundColor Yellow
            #region robocopy migration result
            $DOCSrobocopy = ''
            $DOCSrobocopy = Get-Content "$wdir\robocopydocs.log" 
            $DOCSrobocopy = $DOCSrobocopy  -match '^(?= *?\b(Total|Dirs|Files)\b)((?!    Files).)*$'
            #convert to array
            $DOCSrobocopyresult = @()
            foreach ($line in $DOCSrobocopy ) {
                $DOCSrobocopyresult += $line
            }
            $docsfailedfile = [int](($DOCSrobocopyresult[2] -split "\s+")[7])
            $docsfaileddirs = [int](($DOCSrobocopyresult[1] -split "\s+")[7])
            if (($docsfailedfile -eq 0) -and ($docsfaileddirs -eq 0) ){

                write-host "There is no failed files or directories during the migration"
            }
            else{
                write-host "There are some failed files or directories. Please see robocopy logs located in $wdir\robocopydocs.log" -ForegroundColor Red
            }
            #endregion
            start-sleep 1
            $er += $fol 
                
        } 
                       
        catch {
            write-host "$fol" -ForegroundColor White -BackgroundColor Red -NoNewline
            write-host " migration will be skipped" -ForegroundColor Red
            $d = $false
            
        } 
                       
    }
    #endregion
    #region music migration
    elseif (($oldmusic -notlike $folderspath + '\' + $musicstring) -and ($fol -eq $musicstring)) {

        try {
          
            #with $m I build column in html to determine if migration is done or not.
            $m = $true
            #test write permissions on a migration destination folder before making registry changes
            try {"" | out-file   $newmusic\write-testfileyw12347322322234330983.txt -ErrorAction Stop} catch {
                write-host "It is not possible to write to destination folder $newmusic" -ForegroundColor Red
            }
            "" | out-file   $newmusic\write-testfileyw12347322322234330983.txt -ErrorAction Stop

            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'my music' $newmusic -ErrorAction Stop
            Set-ItemProperty -path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'my music' $newmusic -ErrorAction Stop
            $USF = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop
            $USFpath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' 
                    
            if ($USF.'{A0C69A99-21C8-4671-8703-7934162FCF1D}') {
                Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name `
                    '{A0C69A99-21C8-4671-8703-7934162FCF1D}' $newmusic -ErrorAction Stop
            } 

            else {
                New-ItemProperty -Path $USFpath -Name '{A0C69A99-21C8-4671-8703-7934162FCF1D}' -Value $newmusic -PropertyType expandstring -Force -ErrorAction Stop | Out-Null
                           
            }
            Start-Sleep 1
            Write-Host $musicstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " started migrating" -ForegroundColor Yellow
                
            Robocopy $oldmusic $newmusic /move /e /xo /xj /r:0 /ns /nc /np /njh /log:"$wdir\robocopymusic.log" | out-null #use /fft if you copy over the network because of latency
            Write-Host $musicstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " folder finished migrating" -ForegroundColor Yellow
            Write-Host $musicstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " migration result:" -ForegroundColor Yellow
            #region robocopy migration result
            $musicrobocopy = ''
            $musicrobocopy = Get-Content "$wdir\robocopymusic.log" 
            $musicrobocopy = $musicrobocopy  -match '^(?= *?\b(Total|Dirs|Files)\b)((?!    Files).)*$'
            #convert to array
            $musicrobocopyresult = @()
            foreach ($line in $musicrobocopy ) {
                $musicrobocopyresult += $line
            }
            $musicfailedfile = [int](($musicrobocopyresult[2] -split "\s+")[7])
            $musicfaileddirs = [int](($musicrobocopyresult[1] -split "\s+")[7])
            if (($musicfailedfile -eq 0) -and ($musicfaileddirs -eq 0) ){

                write-host "There is no failed files or directories during the migration"
            }
            else{
                write-host "There are some failed files or directories. Please see robocopy logs located in $wdir\robocopymusic.log" -ForegroundColor Red
            }
            #endregion
            start-sleep 4
            $er += $fol
   
        } 
       
        catch {
            write-host "$fol" -ForegroundColor White -BackgroundColor Red -NoNewline
            write-host " migration will be skipped" -ForegroundColor Red
            $m = $false
            
        } 
       
    }

    #endregion
    #region pictures migration          
    elseif (($oldpictures -notlike $folderspath + '\' + $picturesstring) -and ($fol -eq $picturesstring)) {
        
        try {
                          
            #with $p I build column in html to determine if migration is done or not.
            $p = $true
            #test write permissions on a migration destination folder before making registry changes
            try {"" | out-file   $newpictures\write-testfileyw12347322322234330983.txt -ErrorAction Stop} catch {
                write-host "It is not possible to write to destination folder $newpictures" -ForegroundColor Red
            }
            "" | out-file   $newpictures\write-testfileyw12347322322234330983.txt -ErrorAction Stop
                                    
        
            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'my pictures' $newpictures -ErrorAction Stop
            Set-ItemProperty -path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'my pictures' $newpictures -ErrorAction Stop
            $USF = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop
            $USFpath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' 
                                    
            if ($USF.'{0DDD015D-B06C-45D5-8C4C-F59713854639}') {
                Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name `
                    '{0DDD015D-B06C-45D5-8C4C-F59713854639}' $newpictures -ErrorAction Stop
            } 
        
            else {
                New-ItemProperty -Path $USFpath -Name '{0DDD015D-B06C-45D5-8C4C-F59713854639}' -Value $newpictures -PropertyType expandstring -Force -ErrorAction Stop | Out-Null
                                           
            }
            Start-Sleep 1
            Write-Host $picturesstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " started migrating" -ForegroundColor Yellow
                                
            Robocopy $oldpictures $newpictures /move /e /xo /xj /r:0 /ns /nc /np /njh /log:"$wdir\robocopypictures.log" | out-null #use /fft if you copy over the network because of latency
                         
            Write-Host $picturesstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " folder finished migrating" -ForegroundColor Yellow
            Write-Host $picturesstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " migration result:" -ForegroundColor Yellow
            
            $picturesrobocopy = ''
            $picturesrobocopy = Get-Content "$wdir\robocopypictures.log"
            $picturesrobocopy = $picturesrobocopy  -match '^(?= *?\b(Total|Dirs|Files)\b)((?!    Files).)*$'
            #convert to array
            $picturesrobocopyresult = @()
            foreach ($line in $picturesrobocopy ) {
                $picturesrobocopyresult += $line
            }
            $picturesfailedfile = [int](($picturesrobocopyresult[2] -split "\s+")[7])
            $picturesfaileddirs = [int](($picturesrobocopyresult[1] -split "\s+")[7])
            if (($picturesfailedfile -eq 0) -and ($picturesfaileddirs -eq 0) ){

                write-host "There is no failed files or directories during the migration"
            }
            else{
                write-host "There are some failed files or directories. Please see robocopy logs located in $wdir\robocopypictures.log" -ForegroundColor Red
            }
            start-sleep 4
            $er += $fol 
                   
        } 
                       
        catch {
            write-host "$fol" -ForegroundColor White -BackgroundColor Red -NoNewline
            write-host " migration will be skipped" -ForegroundColor Red
            $p = $false
            $musicreport = "Error in migration or already in OneDrive"
        } 
                       
    }
    #endregion
    #region desktop migration                                
    elseif (($olddesktop -notlike $folderspath + '\' + $desktopstring) -and ($fol -eq $desktopstring)) {
            
        try {
                              
            #with $de I build column in html to determine if migration is done or not.
            $de = $true
            #test write permissions on a migration destination folder before making registry changes
            try {"" | out-file   $newdesktop\write-testfileyw12347322322234330983.txt -ErrorAction Stop} catch {
                write-host "It is not possible to write to destination folder $newdesktop" -ForegroundColor Red
            }
            "" | out-file   $newdesktop\write-testfileyw12347322322234330983.txt -ErrorAction Stop
                                        
            
            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'desktop' $newdesktop -ErrorAction Stop
            Set-ItemProperty -path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'desktop' $newdesktop -ErrorAction Stop
            $USF = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop
            $USFpath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' 
                                        
            if ($USF.'{754AC886-DF64-4CBA-86B5-F7FBF4FBCEF5}') {
                Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name `
                    '{754AC886-DF64-4CBA-86B5-F7FBF4FBCEF5}' $newdesktop -ErrorAction Stop
            } 
            
            else {
                New-ItemProperty -Path $USFpath -Name '{754AC886-DF64-4CBA-86B5-F7FBF4FBCEF5}' -Value $newdesktop -PropertyType expandstring -Force -ErrorAction Stop | Out-Null
                                               
            }
            start-sleep 1
            Write-Host $desktopstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " started migrating" -ForegroundColor Yellow
                                    
            Robocopy $olddesktop $newdesktop /move /e /xo /xj /r:0 /ns /nc /np /njh /log:"$wdir\robocopydesktop.log" | out-null #use /fft if you copy over the network because of latency
                               
                                
                               
            Write-Host $desktopstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " folder finished migrating" -ForegroundColor Yellow
            Write-Host $desktopstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " migration result:" -ForegroundColor Yellow
            # robocopy migration result
            $desktoprobocopy = ''
            $desktoprobocopy = Get-Content "$wdir\robocopydesktop.log"
            $desktoprobocopy = $desktoprobocopy  -match '^(?= *?\b(Total|Dirs|Files)\b)((?!    Files).)*$'
            #convert to array
            $desktoprobocopyresult = @()
            foreach ($line in $desktoprobocopy ) {
                $desktoprobocopyresult += $line
            }
            $desktopfailedfile = [int](($desktoprobocopyresult[2] -split "\s+")[7])
            $desktopfaileddirs = [int](($desktoprobocopyresult[1] -split "\s+")[7])
            if (($desktopfailedfile -eq 0) -and ($desktopfaileddirs -eq 0) ){

                write-host "There is no failed files or directories during the migration"
            }
            else{
                write-host "There are some failed files or directories. Please see robocopy logs located in $wdir\robocopydesktop.log" -ForegroundColor Red
            }
            start-sleep 2
            $er += $fol  
                       
        } 
                           
        catch {
            write-host "$fol" -ForegroundColor White -BackgroundColor Red -NoNewline
            write-host " migration will be skipped" -ForegroundColor Red
            $de = $false
            $musicreport = "Error in migration or already in OneDrive"
        } 
                           
    }
    #endregion 
    #region videos migration                                                                
    elseif (($oldvideos -notlike $folderspath + '\' + $videosstring) -and ($fol -eq $videosstring)) {
            
        try {
                              
            #with $v I build column in html to determine if migration is done or not.
            $v = $true
            #test write permissions on a migration destination folder before making registry changes
            try {"" | out-file   $newvideos\write-testfileyw12347322322234330983.txt -ErrorAction Stop} catch {
                write-host "It is not possible to write to destination folder $newvideos" -ForegroundColor Red
            }
            "" | out-file   $newvideos\write-testfileyw12347322322234330983.txt -ErrorAction Stop
                                        
            
            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'my video' $newvideos -ErrorAction Stop
            Set-ItemProperty -path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'my video' $newvideos -ErrorAction Stop
            $USF = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop
            $USFpath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' 
                                        
            if ($USF.'{35286A68-3C57-41A1-BBB1-0EAE73D76C95}') {
                Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name `
                    '{35286A68-3C57-41A1-BBB1-0EAE73D76C95}' $newvideos -ErrorAction Stop
            } 
            
            else {
                New-ItemProperty -Path $USFpath -Name '{35286A68-3C57-41A1-BBB1-0EAE73D76C95}' -Value $newvideos -PropertyType expandstring -Force -ErrorAction Stop | Out-Null
                                               
            }
            start-sleep 1
            Write-Host $videosstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " started migrating" -ForegroundColor Yellow
                                    
            Robocopy $oldvideos $newvideos /move /e /xo /xj /r:0 /ns /nc /np /njh /log:"$wdir\robocopyvideo.log" | out-null #use /fft if you copy over the network because of latency
                            
                              
            Write-Host $videosstring  -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " folder finished migrating" -ForegroundColor Yellow
            Write-Host $videosstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " migration result:" -ForegroundColor Yellow
            # robocopy migration result
            $videosrobocopy = ''
            $videosrobocopy = Get-Content "$wdir\robocopyvideo.log" 
            $videosrobocopy = $videosrobocopy  -match '^(?= *?\b(Total|Dirs|Files)\b)((?!    Files).)*$'
            #convert to array
            $videosrobocopyresult = @()
            foreach ($line in $videosrobocopy ) {
                $videosrobocopyresult += $line
            }
            $videosfailedfile = [int](($videosrobocopyresult[2] -split "\s+")[7])
            $videosfaileddirs = [int](($videosrobocopyresult[1] -split "\s+")[7])
            if (($videosfailedfile -eq 0) -and ($videosfaileddirs -eq 0) ){

                write-host "There is no failed files or directories during the migration"
            }
            else{
                write-host "There are some failed files or directories. Please see robocopy logs located in $wdir\robocopyvideo.log" -ForegroundColor Red
            }
            start-sleep 2
            $er += $fol  
                       
        } 
                           
        catch {
            write-host "$fol" -ForegroundColor White -BackgroundColor Red -NoNewline
            write-host " migration will be skipped" -ForegroundColor Red
            $v = $false
            $videoreport = "Error in migration"
        } 
                           
    }

    #endregion                                                               
    #region favorites migration                                                                
    elseif (($oldfavorites -notlike $folderspath + '\' + $favoritesstring) -and ($fol -eq $favoritesstring)) {
            
        try {
                                
            #with $f I build column in html to determine if migration is done or not.
            $f = $true
            #test write permissions on a migration destination folder before making registry changes
            try {"" | out-file   $newfavorites\write-testfileyw12347322322234330983.txt -ErrorAction Stop} catch {
                write-host "It is not possible to write to destination folder $newfavorites" -ForegroundColor Red
            }
            "" | out-file   $newfavorites\write-testfileyw12347322322234330983.txt -ErrorAction Stop
                                        
            
            Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -Name 'favorites' $newfavorites -ErrorAction Stop
            Set-ItemProperty -path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders' -Name 'favorites' $newfavorites -ErrorAction Stop
            $USF = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop
            $USFpath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' 
                                        
                                
            start-sleep 1
            Write-Host $favoritesstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " started migrating" -ForegroundColor Yellow
                                    
            Robocopy $oldfavorites $newfavorites /move /e /xo /xj /r:0 /ns /nc /np /njh /log:"$wdir\robocopyfavorites.log" | out-null #use /fft if you copy over the network because of latency
                              
                              
            Write-Host $favoritesstring  -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " folder finished migrating" -ForegroundColor Yellow
            #
            Write-Host $favoritesstring -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            write-host " migration result:" -ForegroundColor Yellow
           # robocopy migration result
           $favoritesrobocopy = ''
           $favoritesrobocopy = Get-Content "$wdir\robocopyfavorites.log" 
           $favoritesrobocopy = $favoritesrobocopy  -match '^(?= *?\b(Total|Dirs|Files)\b)((?!    Files).)*$'
           #convert to array
           $favoritesrobocopyresult = @()
           foreach ($line in $favoritesrobocopy ) {
               $favoritesrobocopyresult += $line
           }
           $favoritesfailedfile = [int](($favoritesrobocopyresult[2] -split "\s+")[7])
           $favoritesfaileddirs = [int](($favoritesrobocopyresult[1] -split "\s+")[7])
           if (($favoritesfailedfile -eq 0) -and ($favoritesfaileddirs -eq 0) ){

               write-host "There is no failed files or directories during the migration"
           }
           else{
               write-host "There are some failed files or directories. Please see robocopy logs located in $wdir\robocopyfavorites.log" -ForegroundColor Red
           }
            start-sleep 2
                        
        } 
                            
        catch {
            write-host "$fol" -ForegroundColor White -BackgroundColor Red -NoNewline
            write-host " migration will be skipped" -ForegroundColor Red
            $f = $false
            $favoritesreport = "Error in migration"
            #enable-NTFSAccessInheritance -Path 'C:\users\sasa\onedrive - sasa\Music'
        } 
                            
    }    
    
    #endregion                                                               

  
    #endregion elseif favorites  migration                               
    else {
        
            write-host "$fol" -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
            Write-Host  "  already in Onedrive" -NoNewline -ForegroundColor Yellow
            Write-Host " "
            $alreadyinOD += $fol
            start-sleep 1
        }                       
}   



#endregion 

#region delete test files written to destination before migration
Remove-Item $newmusic\write-testfileyw12347322322234330983.txt -ErrorAction SilentlyContinue
Remove-Item $newdesktop\write-testfileyw12347322322234330983.txt -ErrorAction SilentlyContinue
Remove-Item $newdocs\write-testfileyw12347322322234330983.txt -ErrorAction SilentlyContinue
Remove-Item $newfavorites\write-testfileyw12347322322234330983.txt -ErrorAction SilentlyContinue
Remove-Item $newvideos\write-testfileyw12347322322234330983.txt -ErrorAction SilentlyContinue
Remove-Item $newpictures\write-testfileyw12347322322234330983.txt -ErrorAction SilentlyContinue
#endregion



#region Locations after migration 

# we can't use for example [Environment]::GetFolderPath("MyDocuments") because pc still not rebooted after registry changes and it will read old location
function dataloc {
    $musicafter = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -name 'my music').'my music'
    $docsafter = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -name 'personal').'personal'
    $desktopafter = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -name 'desktop').'desktop'
    $videosafter = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -name 'my video').'my video'
    $Picturesafter = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -name 'my pictures').'my pictures'
    $favoritesafter = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -name 'favorites').'favorites'
    $array = @($docsafter, $desktopafter, $musicafter, $videosafter, $picturesafter,$favoritesafter) 
    
   
    

    foreach ($ar in $array) {
        $size = Get-FolderSize -Path $ar -Units GB 
        $hash = [ordered]@{
            'Size(GB)'  = $size.'Size(GB)'
            'Name'      = $ar.split('\')[-1]
            'Locations' = $ar;
            'Migrated'  = if ((($ar -like "*music*") -and ($m -eq $false))) {'Error in migration'} elseif ((($ar -like "*music*") -and ($alreadyinod -contains "music"))) {'Already stored in OD'} elseif ((($ar -like "*music*") -and ($m -eq $true)) -and (($musicfailedfile -gt 0) -or ($musicfaileddirs -gt 0 ))) {'Partially migrated'} `
                elseif ((($ar -like "*documents*") -and ($d -eq $false))) {'Error in migration'} elseif ((($ar -like "*documents*") -and ($alreadyinod -contains "documents"))) {'Already stored in OD'} elseif ((($ar -like "*documents*") -and ($d -eq $true)) -and (($docsfailedfile -gt 0) -or ($docsfaileddirs -gt 0 ))) {'Partially migrated'} `
                elseif ((($ar -like "*desktop*") -and ($de -eq $false))) {'Error in migration'} elseif ((($ar -like "*desktop*") -and ($alreadyinod -contains "desktop"))) {'Already stored in OD'} elseif ((($ar -like "*desktop*") -and ($de -eq $true)) -and (($desktopfailedfile -gt 0) -or ($desktopfaileddirs -gt 0 ))) {'Partially migrated'} `
                elseif ((($ar -like "*video*") -and ($v -eq $false))) {'Error in migration'} elseif ((($ar -like "*video*") -and ($alreadyinod -contains "videos"))) {'Already stored in OD'} elseif ((($ar -like "*video*") -and ($v -eq $true)) -and (($videosfailedfile -gt 0) -or ($videosfaileddirs -gt 0 ))) {'Partially migrated'} `
                elseif ((($ar -like "*favorite*") -and ($f -eq $false))) {'Error in migration'} elseif ((($ar -like "*favorite*") -and ($alreadyinod -contains "favorites"))) {'Already stored in OD'} elseif ((($ar -like "*favorite*") -and ($f -eq $true)) -and (($favoritesfailedfile -gt 0) -or ($favoritesfaileddirs -gt 0 ))) {'Partially migrated'} `
                elseif ((($ar -like "*picture*") -and ($p -eq $false))) {'Error in migration'} elseif ((($ar -like "*picture*") -and ($alreadyinod -contains "pictures"))) {'Already stored in OD'} elseif ((($ar -like "*picture*") -and ($p -eq $true)) -and (($picturesfailedfile -gt 0) -or ($picturesfaileddirs -gt 0 ))) {'Partially migrated'} `
                else {'Migrated sucessfully'}
            
        
        }
        New-Object -TypeName psobject -Property $hash

    }     
}

$pre = "DESTINATIONS - after migration"
$locationsafter = dataloc | ConvertTo-Html  -Fragment -PreContent "<h2>&diams; $pre</h2>" | Out-string
#endregion

$date = Get-Date
$finalhtml = ConvertTo-Html -as Table -Body $style -PostContent $locationsbefore, $locationsafter, "<h2>$date</h2>" -PreContent "<h1>Report after onedrive migration: $env:COMPUTERNAME</h1>"  
$finalhtml = $finalhtml -replace "Already stored in OD", "<font color='#8080ff'>Already stored in OD</font>"
$finalhtml = $finalhtml -replace "Not stored in OD", "<font color='#993333'>Not stored in OD</font>"
$finalhtml = $finalhtml -replace "Error in migration", "<font color='red'>Error in migration</font>"
$finalhtml = $finalhtml -replace "Partially migrated", "<font color='blue'>Partially migrated</font>"
$finalhtml | Out-File C:\yw-data\Onedrivemigration.html
Write-Host " "
start-sleep 1
write-host "REPORT WILL BE AUTOMATICALLY OPENED IN  A FEW SECONDS." -ForegroundColor Yellow
start-sleep 1
Write-Host " "

[string]$docdataleft = Getsize $oldmydocs
[string]$musicdataleft = Getsize $oldmusic
[string]$videosdataleft = Getsize $oldvideos
[string]$picturesdataleft = Getsize $oldpictures
[string]$favoritesdataleft = Getsize $oldfavorites
[string]$desktopdataleft = Getsize $olddesktop
  

if ($er.count -gt 1) {
    Write-Host "To take full effect, please reboot your pc" -BackgroundColor Yellow -ForegroundColor Black  -NoNewline
}



#Robocopy results

#region filesleft #in case some newer files exist in destination, those sources files are not migrated. This part determines if we have that cases and if yes, display that information at the end of the report
if ((($docdataleft -gt 0.0001) -and ($oldmydocs -notlike $newdocs ) -and ($d -ne $false)) `
-or (($desktopdataleft -gt 0.0001) -and ($olddesktop -notlike $newdesktop ) -and ($de -ne $false)) `
-or (($picturesdataleft -gt 0.0001) -and ($oldpictures -notlike $newpictures ) -and ($p -ne $false)) `
-or (($musicdataleft -gt 0.0001) -and ($oldmusic -notlike $newmusic ) -and ($m -ne $false))`
-or (($favoritesdataleft -gt 0.0001) -and ($oldfavorites -notlike $newfavorites ) -and ($f -ne $false)) `
-or (($videosdataleft -gt 0.0001) -and ($oldvideos -notlike $newvideos ) -and ($v -ne $false))) `
{
'<h16 style="margin-left:28%;text-allign:center;font-size:12px">Note: Some files left in source because newer files exist in destination or access denied problem. </h16></br>' | out-file  C:\yw-data\Onedrivemigration.html -Append

}
if (($docdataleft -gt 0.0001) -and ($oldmydocs -notlike $newdocs) -and ($d -ne $false)) {
"<h21 style='margin-left:35%;font-size:10px'>Documents: $docdataleft  </h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append
}
if (($musicdataleft -gt 0.0001) -and ($oldmusic -notlike $newmusic) -and ($m -ne $false)) {
"<h21 style='margin-left:35%;font-size:10px'>Music: $musicdataleft  </h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append
}
if (($videosdataleft -gt 0.0001) -and ($oldvideos -notlike $newvideos) -and ($v -ne $false)) {
"<h21 style='margin-left:35%;font-size:10px'>Videos: $videosdataleft  </h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append
}
if (($picturesdataleft -gt 0.0001) -and ($oldpictures -notlike $newpictures) -and ($p -ne $false)) {
"<h21 style='margin-left:35%;font-size:10px'>Pictures: $picturesdataleft  </h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append
}
if (($favoritesdataleft -gt 0.0001) -and ($oldfavorites -notlike $newfavorites) -and ($f -ne $false)) {
"<h21 style='margin-left:35%;font-size:10px'>Favorites: $favoritesdataleft  </h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append
}
if (($desktopdataleft -gt 0.0001) -and ($olddesktop -notlike $newdesktop) -and ($de -ne $false)) {
"<h21 style='margin-left:35%;font-size:10px'>Desktop: $desktopdataleft  </h21></br></br></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append
}
#endregion


#region Robocopy results
"<h10> </h10></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append
'<h16 style="margin-left:35%;font-size:16px">Robocopy results:</h16></br>'  | out-file  C:\yw-data\Onedrivemigration.html -Append
#Documents report
if ($docsfailedfile -gt 0){

"<h21 style='color:#C43E0F;margin-left:35%;font-size:10px'>Documents: Failed files: $docsfailedfile  ||  check logs in $wdir\robocopydocs.log</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}else
{"<h21 style='margin-left:35%;font-size:10px'>Documents: Failed files: $docsfailedfile'</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}

if ($docsfaileddirs -gt 0){

"<h21 style='color:#C43E0F;margin-left:35%;font-size:10px'>Documents: Failed directories: $docsfaileddirs  ||  check logs in $wdir\robocopydocs.log</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}else
{"<h21 style='margin-left:35%;font-size:10px'>Documents: Failed directories: $docsfaileddirs</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}

#desktop report
if ($desktopfailedfile -gt 0){

"<h21 style='color:#C43E0F;margin-left:35%;font-size:10px'>Desktop: Failed files: $desktopfailedfile  ||  check logs in $wdir\robocopydesktop.log</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}else
{"<h21 style='margin-left:35%;font-size:10px'>Desktop: Failed files: $desktopfailedfile</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}

if ($desktopfaileddirs -gt 0){

"<h21 style='color:#C43E0F;margin-left:35%;font-size:10px'>Desktop: Failed directories: $desktopfaileddirs  ||  check logs in $wdir\robocopydesktop.log</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}else
{"<h21 style='margin-left:35%;font-size:10px'>Desktop: Failed directories: $desktopfaileddirs</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}

#music report
if ($musicfailedfile -gt 0){

"<h21 style='color:#C43E0F;margin-left:35%;font-size:10px'>Music: Failed files: $musicfailedfile  ||  check logs in $wdir\robocopymusic.log</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}else
{"<h21 style='margin-left:35%;font-size:10px'>Music: Failed files: $musicfailedfile</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}

if ($musicfaileddirs -gt 0){

"<h21 style='color:#C43E0F;margin-left:35%;font-size:10px'>Music: Failed directories: $musicfaileddirs  ||  check logs in $wdir\robocopymusic.log</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}else
{"<h21 style='margin-left:35%;font-size:10px'>Music: Failed directories: $musicfaileddirs</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}

#Videos report
if ($videosfailedfile -gt 0){

"<h21 style='color:#C43E0F;margin-left:35%;font-size:10px'>Videos: Failed files: $videosfailedfile  ||  check logs in $wdir\robocopyvideo.log</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}else
{"<h21 style='margin-left:35%;font-size:10px'>Videos: Failed files: $videosfailedfile</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}

if ($videosfaileddirs -gt 0){

"<h21 style='color:#C43E0F;margin-left:35%;font-size:10px'>Videos: Failed directories: $videosfaileddirs  ||  check logs in $wdir\robocopyvideo.log</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}else
{"<h21 style='margin-left:35%;font-size:10px'>Videos: Failed directories: $videosfaileddirs</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}

#Pictures report
if ($picturesfailedfile -gt 0){

"<h21 style='color:#C43E0F;margin-left:35%;font-size:10px'>Pictures: Failed files: $picturesfailedfile  ||  check logs in $wdir\robocopypictures.log</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}else
{"<h21 style='margin-left:35%;font-size:10px'>Pictures: Failed files: $picturesfailedfile</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}

if ($picturesfaileddirs -gt 0){

"<h21 style='color:#C43E0F;margin-left:35%;font-size:10px'>Pictures: Failed directories: $picturesfaileddirs  ||  check logs in $wdir\robocopypictures.log</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}else
{"<h21 style='margin-left:35%;font-size:10px'>Pictures: Failed directories: $picturesfaileddirs</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}


#FAvorites report
if ($favoritesfailedfile -gt 0){

"<h21 style='color:#C43E0F;margin-left:35%;font-size:10px'>Favorites: Failed files: $favoritesfailedfile  ||  check logs in $wdir\robocopyfavorites.log</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}else
{"<h21 style='margin-left:35%;font-size:10px'>Favorites: Failed files: $favoritesfailedfile</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}

if ($favoritesfaileddirs -gt 0){

"<h21 style='color:#C43E0F;margin-left:35%;font-size:10px'>Favorites: Failed directories: $favoritesfaileddirs  ||  check logs in $wdir\robocopyfavorites.log</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}else
{"<h21 style='margin-left:35%;font-size:10px'>Favorites: Failed directories: $favoritesfaileddirs</h21></br>"  | out-file  C:\yw-data\Onedrivemigration.html -Append}
#endregion
Start-Sleep 5
invoke-item C:\yw-data\Onedrivemigration.html
read-host





