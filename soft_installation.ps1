#Created by https://github.com/VladimirKosyuk

#Installs and updates desired soft list on local PC. 

#About:

 <#
Limitations:

1. Setup must be *.exe file
2. Setup installation keys must support key /S
3. PC OS must be windows 10 x64
4. OC release must be equal or higer than 10.0.19042
5. $Repo folder and all objects in it must have allow modify access to Domain computers, if script will be executed as task via system account
6. Inside $Repo must be exact same naming for each soft folder as in $programs array, inside each folder must be single file named *setup.exe, preferred as x64
7. $programs elements must be named as DisplayName in registry HKLM:\SOFTWARE\*\Uninstall\*
8. setup file must contain DisplayVersion populated as installed soft
9. SmtpServer port must be 25, well, it's hardcoded

Does:

1. Reads $programs array as desired soft list
2. If OS Windows x64 script will proceed, else - stops
3. Check registry, if soft from $programs is installed
4. If not installed, tries to get acces to install file, if fails - send email, 
5. Start installation, cli output, if fails - send email
6. If multiple versions found, like 32 and 64 simultaneously, deletes x32 version, cli output, if fails - send email
7. If installed version not match setup file, deletes installed and installes from setup file, cli output, if fails - send email

 #>

# Build date: 07.04.2021

#mail defined vars
$SmtpDomain = ""#need to be started as @ symbol, example - @contoso.com
$SmtpSrv = "" 
$To = ""
$From = $env:COMPUTERNAME+$SmtpDomain
$Subject = $env:COMPUTERNAME+" soft installation failed"
$Repo = "" #example repo - \\srv\repo, example setup filepath - \\srv\repo\7-zip\7z1900-x64_setup.exe
$programs = @(
#'7-zip'#example
#'WinDJView'#example
)
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
$OSData = Get-WMIObject win32_operatingsystem
if (($OSData.Caption -like "*Windows 10*") -and (($OSData.OSArchitecture) -like "64*")){
    Foreach ($program in $programs){
    #collect installed soft array
$IsInstalled = @(
Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | ?{($_.Publisher -notmatch "Microsoft") -and ($_.DisplayName -match $program)} | Select-Object DisplayName, DisplayVersion, InstallLocation, UninstallString, PSPath
Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |?{($_.Publisher -notmatch "Microsoft") -and ($_.DisplayName -match $program)} | Select-Object DisplayName, DisplayVersion, InstallLocation, UninstallString, PSPath
)
    #if not installed
    if(!($Setup = get-childitem $Repo\$program | ?{$_.Name -like "*setup.exe"})){Send-MailMessage -To $To -From $From -Subject $Subject -Body ($Repo+" is not reachable") -Port 25 -SmtpServer $SmtpSrv}
    if(!$IsInstalled){
    Write-Output ("DEBAG OUTPUT"+" "+$env:COMPUTERNAME+" "+(Get-Date)+" "+($program)+" "+"not found, start install")
    #install soft
        if (!(Start-Process -FilePath $Setup.FullName /S -NoNewWindow -Wait -PassThr)){Send-MailMessage -To $To -From $From -Subject $Subject -Body ("Cannot install "+$Setup.FullName) -Port 25 -SmtpServer $SmtpSrv}
        }
        #if installed multiple
        elseif($IsInstalled.count -notmatch "1"){
        $Delete32 = $IsInstalled | ?{$_.PSPath -notlike "*Wow6432Node*"}
        Write-Output ("DEBAG OUTPUT"+" "+$env:COMPUTERNAME+" "+(Get-Date)+" "+($program)+" "+"multiple versions found, start uninstall 32-bit")
            if (!(Start-Process $Delete32.UninstallString /S -NoNewWindow -Wait -PassThr)){Send-MailMessage -To $To -From $From -Subject $Subject -Body ("Cannot execute "+$Delete32.UninstallString) -Port 25 -SmtpServer $SmtpSrv}
        }
    #if outdated
$IsInstalled = @(
Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | ?{($_.Publisher -notmatch "Microsoft") -and ($_.DisplayName -match $program)} | Select-Object DisplayName, DisplayVersion, InstallLocation, UninstallString, PSPath
Get-ItemProperty HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* |?{($_.Publisher -notmatch "Microsoft") -and ($_.DisplayName -match $program)} | Select-Object DisplayName, DisplayVersion, InstallLocation, UninstallString, PSPath
)
    if($IsInstalled.DisplayVersion -notmatch $Setup.VersionInfo.ProductVersion){
    Write-Output ("DEBAG OUTPUT"+" "+$env:COMPUTERNAME+" "+(Get-Date)+" "+($program)+" "+"version is"+" "+($IsInstalled.DisplayVersion)+" "+"not match setup file version"+" "+($Setup.VersionInfo.ProductVersion))
        if (!(Start-Process $IsInstalled.UninstallString /S -NoNewWindow -Wait -PassThr)){Send-MailMessage -To $To -From $From -Subject $Subject -Body ("Cannot execute "+$IsInstalled.UninstallString) -Port 25 -SmtpServer $SmtpSrv}
        if (!(Start-Process -FilePath $Setup.FullName /S -NoNewWindow -Wait -PassThr)){Send-MailMessage -To $To -From $From -Subject $Subject -Body ("Cannot install "+$Setup.FullName) -Port 25 -SmtpServer $SmtpSrv}
        }
    }
}

Remove-Variable -Name * -Force -ErrorAction SilentlyContinue