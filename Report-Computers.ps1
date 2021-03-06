######################################################################################################################################
# Server reporting script
######################################################################################################################################
# Script Author Information
$script:ProgramName = "Computer Reporting Script"
$script:ProgramDate = "30 Jan 2014"
$script:ProgramAuthor = "Geoffrey Guynn"
$script:ProgramAuthorEmail = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String("Z2VvZmZyZXlAZ3V5bm4ub3Jn"))
$script:ProgramVersion = "1.0.0.1"

#Script Information
$script:WorkingFileName = $MyInvocation.MyCommand.Definition
$script:WorkingDirectory = Split-Path $script:WorkingFileName -Parent
######################################################################################################################################
# Purpose: This script retrieves metadata from remote systems and prints the data to a CSV. 
######################################################################################################################################
#                                    How to Use This Script     
# 1. Verify that powershell is running as an administrator. The title will say Administrator.
#-------------------------------------------------------------------------------------------------------------------------------------
# 2. If you haven't set execution policy, type the following command from an administrative session.
#    Set-ExecutionPolicy Unrestricted -force
#-------------------------------------------------------------------------------------------------------------------------------------
# 3. Create a text file that contains all servers (one per line) and put it in quotes on this line.
$script:ComputerList = "$script:WorkingDirectory\hosts.txt"
#-------------------------------------------------------------------------------------------------------------------------------------
# 4. Where should we save the report? (default is $env:userprofile\desktop, your desktop)
$script:SaveTo = "$script:WorkingDirectory"
# 5. Operating mode <console, report, both>
$script:OperatingMode = "report"
# 6. MaxJobs <30 - average CPU, 60 - fast CPU, 90 - VERY fast CPU> (USE CAUTION AS THIS WILL DIRECTLY IMPACT YOUR SYSTEM PERFORMANCE!)
$script:MaxJobs = 60
# 7. Wait Timeout (default is 30 seconds)
$script:Timeout = 30
######################################################################################################################################
#                         Do Not Modify Anything Below This Line!
######################################################################################################################################
<# Begin ---- Global Variables - Script Limited Scope ---- #>
$Results = @()
$i = 0
if ($ComputerList){
    if (Test-Path $ComputerList){
        $Computers = Get-Content $ComputerList | ? {$_ -notlike "*#*"}
    }
    else{
        Write-Warning "Computerlist $computerlist was not found, running against localcomputer."
        $computers = $env:computername
    }
}
else{
    Write-Warning "No computer list was provided, running against localcomputer."
    $computers = $env:computername
}
<# End ---- Global Variables - Script Limited Scope ---- #>

<# Begin ---- Helper Functions ---- #>
Function Report-Console{
	param(
		$Results
	)

	$Results | 
		Select HostName, 
               Model, 
               Manufacturer, 
               CPUName, 
               CPUClockSpeed, 
               CPUCores, 
               CPUCount, 
               OSName, 
               ServicePack, 
               TotalPhysicalMemory, 
               FreePhysicalMemory, 
               TotalVirtualMemory,
               FreeVirtualMemory, 
               HardDrives
}
Function Report-CSV{
	param(
		$Results
	)

    $SaveAs = Join-Path -Path $script:SaveTo -ChildPath "$($env:userdomain)_Client_Report_$(Get-Date -f "yyyy-MM-dd-hhmmsstt").csv"

	$Results |
		Select HostName, 
               Model, 
               Manufacturer, 
               CPUName, 
               CPUClockSpeed, 
               CPUCores, 
               CPUCount, 
               OSName, 
               ServicePack, 
               TotalPhysicalMemory, 
               FreePhysicalMemory, 
               TotalVirtualMemory, 
               FreeVirtualMemory, 
               HardDrives |
        Export-Csv -Path $SaveAs -NoTypeInformation
    
    Write-Host "Saved report data to $SaveAs"

}
<# End ---- Helper Functions ---- #>

<# Begin ---- Main Program Scriptblock ---- #>
$sb_ServerMetaData = {
	param(
		$CurrentComputer
	)

	<# Begin ---- Main Execution ---- #>
	try{
		$Local:ErrorActionPreference = "Stop"

		Write-Host "Processing: $CurrentComputer"
			
		$OSQuery = Get-WmiObject -ComputerName $CurrentComputer -Class win32_operatingsystem -Namespace "root/CIMv2"
		$HWQuery = Get-WmiObject -ComputerName $CurrentComputer -Class win32_computersystem -Namespace "root/CIMv2"
		$CPUQuery = @(Get-WmiObject -ComputerName $CurrentComputer -Class win32_processor -Namespace "root/CIMv2")
		try{
			$NumberofCores = @($CPUQuery  | Select -ExpandProperty NumberOfCores -unique)[0]
		}
		catch{
			$NumberofCores = 1
		}

		$Result = New-Object PSObject -Property @{
			HostName = $OSQuery.CSName
			Model = $HWQuery.Model
			Manufacturer = $HWQuery.Manufacturer
			CPUName = $CPUQuery | Select -ExpandProperty Name -unique
			CPUClockSpeed = "{0:N0} Mhz" -f @($CPUQuery | Select -ExpandProperty CurrentClockSpeed -unique)[0]
			CPUCores = $NumberofCores
			CPUCount = $CPUQuery.length
			OSName = $OSQuery.Caption
			ServicePack = $OSQuery.CSDVersion
			TotalPhysicalMemory = "{0:N0} MB" -f ($OSQuery.TotalVisibleMemorySize/1kb)
			FreePhysicalMemory = "{0:N0} MB" -f ($OSQuery.FreePhysicalMemory/1kb)
			TotalVirtualMemory = "{0:N0} MB" -f ($OSQuery.TotalVirtualMemorySize/1kb)
			FreeVirtualMemory = "{0:N0} MB" -f ($OSQuery.FreeVirtualMemory/1kb)
			HardDrives = Get-WmiObject `
				-Computer $CurrentComputer `
				-Class win32_LogicalDisk `
				-Namespace "root/CIMv2" |
					Where-Object {$_.Description -eq "Local Fixed Disk"} |
					Select-Object `
						Name, `
						VolumeName, `
						@{Label="FreeSpace"; Expression={ "{0:N0} GB" -f ($_.FreeSpace/1gb)}}, `
						@{Label="TotalSize"; Expression={ "{0:N0} GB" -f ($_.Size/1gb)}} | 
					ForEach-Object {"$($_.Name) $($_.VolumeName), $($_.FreeSpace) Free, $($_.TotalSize) Total;"}
		}
	$Result.HardDrives = $result.HardDrives -join " "
	}
	catch{
		Write-Warning $_.exception.message
		$Result = New-Object PSObject -Property @{
			HostName = $CurrentComputer
			Model = $_.exception.message
		}
	}
	finally{
		$Local:ErrorActionPreference = "Continue"
	}
	return $Result
}
<# End ---- Main Program Scriptblock ---- #>

<# Begin ---- Job Control Engine ---- #>
Get-Job | % { 
    Stop-Job $_ -ErrorAction "SilentlyContinue" 
    Remove-Job $_ -ErrorAction "SilentlyContinue"
}

$jobs = @()

$Computers | %{
	$currentComputer = $_

	$job_param = @{
		Name = $currentComputer
		ScriptBlock = $sb_ServerMetaData
		ArgumentList = $currentComputer	
	}
	
	while ((Get-Job -State "Running").count -ge [int]$script:MaxJobs){
		Write-Host "Job Throttle... waiting 3 seconds... ($((Get-Job -State "Running").count) jobs are running.)"
		Start-Sleep 3
	}
	
	$jobs += Start-Job @job_param
	Write-Host "Started job on: $CurrentComputer"
}

Write-Host "`r`nAll jobs have been started, please wait..."
Wait-Job -State "Running" | Out-Null

foreach ($job in (Get-Job -state "Completed")){
	$Results += Receive-Job -Id $job.Id -Keep
}
<# End ---- Job Control Engine ---- #>

switch ($script:OperatingMode){
    "both"{
        Report-CSV $Results
        Report-Console $Results
    }
    "report"{
        Report-CSV $Results
    }
    default{
        Report-Console $Results
    }
}
