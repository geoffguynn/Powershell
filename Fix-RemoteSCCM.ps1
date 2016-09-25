<# Begin ---- Global Variables - Script Limited Scope ---- #>
#Script Author Information
$script:ProgramName = "SCCM Reinstall Tool"
$script:ProgramDate = "19 Jun 2014"
$script:ProgramAuthor = "Geoffrey Guynn"
$script:ProgramAuthorEmail = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String("Z2VvZmZyZXlAZ3V5bm4ub3Jn"))

#Configuration Variables
# 1. A UNC path to the client installation files for SCCM client.
#    SCCM 2012 R2 default: \\<SCCM_Server>\Client or \\<SCCM_Server\SMS_<SiteCode>\Client
$Script:InstallerPath = ""
# 2. A computer to target for SCCM re-install
$Script:ComputerName = ""
# 3. Arguments that should be passed along to the SCCM installer program.
#    It is VITALLY IMPORTANT that you make sure this information matches the management subnet/scope for the computer(s) being fixed.
#    SCCM 2012 R2 default: /MP:<SCCM_Management_Point FQDN> SMSSITECODE=<SiteCode>
$Script:CCMSetupArguments = ""

#Script Information
$script:WorkingFileName = $MyInvocation.MyCommand.Definition
$script:WorkingDirectory = Split-Path $script:WorkingFileName -Parent
$script:ScriptVersion = "1.0.0.0"
<# End ---- Global Variables - Script Limited Scope ---- #>

<# Begin ---- SCCM Reinstall Tool Helper Functions ---- #>
Function Copy-SCCMFolder{
	param(
		[string]$ComputerName,
		[string]$Installer = $Script:InstallerPath
	)
	

	if ($Installer -and (Test-Path $Installer)){
		$SCCMUNCPath = Get-SCCMPath -ComputerName $ComputerName
		Robocopy $Installer $SCCMUNCPath /e /mir
	}
}
Function Fix-SCCMCache{
	param(
		[string]$ComputerName
	)

	$Local:ErrorActionPreference = "SilentlyContinue"
	$WMISCCMCache = Get-WmiObject -Class cacheinfoex -Namespace "ROOT\ccm\Softmgmtagent" -Computername $ComputerName
	$WMISCCMCache | % {
		$CurrentCacheItem = $_
		$CachePath = "\\$ComputerName\c`$$($CurrentCacheItem.location.substring(2))"

		if ((Test-Path $CachePath) -ne "true"){
			"Creating cache folder $CachePath"
			md $CachePath | Out-Null
		}
	}
}
Function Get-OSArchitecture{
	param(
		[string]$ComputerName
	)
	
	$Arch = (Get-WmiObject -class win32_operatingsystem -computername $ComputerName).OSArchitecture
	if ($Arch -like "*32*"){
		return "32"
	}
	else{
		return "64"
	}
}
Function Get-SCCMPath{
	param(
		[string]$ComputerName,
		[string]$Architecture,
		[switch]$Local
	)
	
	if (!$Architecture){
		$Architecture = Get-OSArchitecture -ComputerName $ComputerName 
	}
	if ($Local.IsPresent -eq $True){ #Return a local path
		if ($Architecture -eq "32"){
			return = "C:\windows\system32\ccmsetup"
		}
		else{
			return "C:\windows\ccmsetup"
		}
	}
	else{ #Return a unc path
		if($Architecture -eq "32"){
			return = "\\$ComputerName\C$\windows\system32\ccmsetup"
		}
		else{
			return "\\$ComputerName\C$\windows\ccmsetup"
		}
	}
}
Function New-Process{
	param(
		[string]$ComputerName,
		[string]$Command,
		[switch]$Wait
	)
		
	try{
		Write-Host "`nComputer: $ComputerName"
		Write-Host "Command: $Command"
		Write-Host "Status: " -NoNewLine
			
		$Process = ([WMIClass]"\\$ComputerName\ROOT\CIMv2:Win32_Process")
		$Return = $Process.create.Invoke($Command)
			
		if ($Return.ReturnValue -eq 0){
			Write-Host "Success" -ForegroundColor darkgreen -NoNewLine
			Write-Host ", Process ID - $($Return.ProcessID)`n`n"
			if ($wait.IsPresent){
				while ((Get-Process -computer $ComputerName -ID $Return.ProcessID -EA "SilentlyContinue") -ne $Null){
					$ProcName = (Get-Process -computer $ComputerName -ID $Return.ProcessID -EA "SilentlyContinue").Name
                    Write-Host "Waiting for " -nonewline
					Write-Host "$ProcName " -ForegroundColor darkblue -nonewline
					Write-Host "to complete..."
					Start-Sleep 10
				}
				Write-Host "`n$ProcName " -ForegroundColor darkgreen -nonewline
				Write-Host "has completed."
			}
		}
		else{
			Write-Host "Failure" -ForegroundColor red -NoNewLine
			Write-Host ", Return Code: $($Return.ReturnValue)" 
		}
	}
	catch{
		Write-Host "Failure" -ForegroundColor red -NoNewLine
		Write-Host ", Reason: $($_.exception.message)"
	}
	finally{
		$Return.psbase.Dispose()
	}
}
Function Test-RemoteAdmin{ #Verify system is online and remote access is available.
	param(
		[string]$ComputerName,
        [switch]$RemoteRegistry,
        [switch]$AdminShare,
        [switch]$Ping,
        [switch]$WMI,
        [switch]$AllTests = $True,
        [switch]$Verbose
	)
    
    if ($Verbose.IsPresent){$Local:VerbosePreference = "Continue"}


    if ($AdminShare.IsPresent){$AllTests = $False}
    if ($Ping.IsPresent){$AllTests = $False}
    if ($WMI.IsPresent){$AllTests = $False}
    
    if ($AllTests -or $Ping){
        if (Test-Connection $ComputerName -quiet -count 1){
            Write-Verbose "[$ComputerName] Ping Test: Passed"
        }
        else{
            Write-Verbose "[$ComputerName] Ping Test: Failed"
            return $False
        }
        
    }
    if ($AllTests -or $AdminShare){
        if (Test-Path "\\$ComputerName\c$"){
            try{
                [void](New-Item -path "\\$ComputerName\c$\Windows\temp\testfile" -type file -force)
                [void](rm -force "\\$ComputerName\c$\Windows\temp\testfile")
            }
            catch{
                Write-Verbose "[$ComputerName] AdminShare Test: Failed"
                return $False
            }
        }
        else{
            Write-Verbose "[$ComputerName] AdminShare Test: Failed"
            return $False
        }
        Write-Verbose "[$ComputerName] AdminShare Test: Passed"
    }
    if ($AllTests -or $WMI){
        try{
            Get-WmiObject -ComputerName $ComputerName -Class win32_Computersystem -ErrorAction "Stop"
        }
        catch{
            Write-Verbose "[$ComputerName] WMI Test: Failed"
            return $False
        }
        Write-Verbose "[$ComputerName] WMI Test: Passed"
    }
    if ($AllTests -or $RemoteRegistry){
        try{
            $RemRegService = Get-Service -ComputerName $ComputerName -Name "RemoteRegistry"
            if ($RemRegService.Status -ne "Running"){
                Write-Verbose "[$ComputerName] RemoteRegistry isn't running on $ComputerName, trying to start it."
                $Service = Set-Service -ComputerName $ComputerName -Name "RemoteRegistry" -Status "Running" -PassThru
                Write-Verbose "[$ComputerName] RemoteRegistry started successfully"
            } else {
                Write-Verbose "[$ComputerName] RemoteRegistry Test: Passed"
            }
        }
        catch{
            Write-Verbose "[$ComputerName] Couldn't start RemoteRegistry : $($_.Exception.Message)"
            return $false
        }
    }
    return $True
}
<# End ---- SCCM Reinstall Tool Helper Functions ---- #>

<# Begin ---- SCCM Reinstall Tool Main Execution Functions ---- #>
Function Install-SCCM{
	param(
		[string]$ComputerName,
		[string]$Installer = $Script:InstallerPath
	)

	if (Test-RemoteAdmin $ComputerName){
	    $Architecture = Get-OSArchitecture -ComputerName $ComputerName
	    $SCCMUNCPath = Get-SCCMPath -ComputerName $ComputerName -Architecture $Architecture
        $SCCMLocalPath = Get-SCCMPath -ComputerName $ComputerName -Architecture $Architecture -Local
	
		if (!$SCCMUNCPath){
			mkdir $SCCMUNCPath -force | Out-Null
		}
		Copy-SCCMFolder -ComputerName $ComputerName
		if (Test-Path "$SCCMUNCPath\ccmsetup.exe"){
			New-Process -ComputerName $ComputerName -Command "$SCCMLocalPath\ccmsetup.exe $Script:CCMSetupArguments" -Wait
		}
	}
    else{
        Write-Host "One of the admin tests has failed!"
        break
    }
}
Function Uninstall-SCCM{ #Uninstall lingering SCCM installation.
	param(
		[string]$ComputerName
	)
	if (Test-RemoteAdmin $ComputerName){	
	    $Architecture = Get-OSArchitecture -ComputerName $ComputerName
	    $SCCMUNCPath = Get-SCCMPath -ComputerName $ComputerName -Architecture $Architecture
        $SCCMLocalPath = Get-SCCMPath -ComputerName $ComputerName -Architecture $Architecture -Local
	
		if (Test-Path "$SCCMUNCPath\ccmsetup.exe"){ #Try to uninstall if the ccmsetup.exe is present.
			New-Process -ComputerName $ComputerName -Command "$SCCMLocalPath\ccmsetup.exe /uninstall" -Wait
		}
	}
    else{
        Write-Host "One of the admin tests has failed!"
        break
    }
}
<# End ---- SCCM Reinstall Tool Main Execution Functions ---- #>

<# Begin ---- SCCM Reinstall Tool Main Execution ---- #>
Uninstall-SCCM -ComputerName $ComputerName
Fix-SCCMCache -ComputerName $ComputerName
Install-SCCM -ComputerName $ComputerName -Installer $Script:InstallerPath
<# End ---- SCCM Reinstall Tool Main Execution ---- #>