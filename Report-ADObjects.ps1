####################################################################################################
# AD reporting script
####################################################################################################
# Script Author Information
$script:ProgramName = "AD Reporting Script"
$script:ProgramDate = "30 Jan 2014"
$script:ProgramAuthor = "Geoffrey Guynn"
$script:ProgramAuthorEmail = [System.Text.Encoding]::ASCII.GetString([System.Convert]::FromBase64String("Z2VvZmZyZXlAZ3V5bm4ub3Jn"))
$script:ProgramVersion = "1.1.0.0"

#Script Information
$script:WorkingFileName = $MyInvocation.MyCommand.Definition
$script:WorkingDirectory = Split-Path $script:WorkingFileName -Parent

#Constant Variables
$Const_Group_Security_Universal = -2147483640
$Const_Group_Security_Global = -2147483646
$Const_Group_Security_Local = -2147483644
$Const_Group_Distro_Universal = 8
$Const_Group_Distro_Global = 2
$Const_Group_Distro_Local = 4

#Create .Net Searcher
$DirectorySearcher = New-Object System.DirectoryServices.DirectorySearcher

#LDAP search filters
$LDAPFilter = @{
    ALLOUs = "objectCategory=organizationalUnit"
    ALLUsers = "samAccountType=805306368"
    AllComputers = "objectClass=computer"
    AllGroups = "objectClass=group"
    AllGPO = "objectClass=groupPolicyContainer"
}
####################################################################################################
# Purpose: This script retrieves metadata from Active Directory and prints the data into a CSV.
####################################################################################################
#                                    How to Use This Script     
# 1. Verify that powershell is running as an administrator. The title will say Administrator.
#---------------------------------------------------------------------------------------------------
# 2. If you haven't set execution policy, type the following command from an administrative session.
#    Set-ExecutionPolicy Unrestricted -force
#---------------------------------------------------------------------------------------------------
# 3. Where should we save the report?
#    Standard default location(s): 
#        $env:userprofile\desktop, (your desktop)
#        $script:WorkingDirectory, (the folder this script was saved to before execution)
$script:SaveTo = $script:WorkingDirectory
####################################################################################################
#                         Do Not Modify Anything Below This Line!
####################################################################################################

#Measure how long this search and sort takes to process.
$ExecutionTime = Measure-Command {

    #Search for all Users, Computers, Groups, GPOs, and OUs
    Write-Host "Performing Search in AD"
    $DirectorySearcher.Filter = "(|($($LDAPFilter.AllOUs))($($LDAPFilter.ALLUsers))($($LDAPFilter.AllComputers))($($LDAPFilter.AllGroups))($($LDAPFilter.AllGPO)))"
    $DirectorySearcher.PageSize = "10000"
    $Results = $DirectorySearcher.FindAll()

    #Filter search results into separate objects.
    $Computers = @()
    $OUs = @()
    $users = @()
    $groups = @()
    $gpo = @()
   
    Write-Host "Sorting results into separate arrays"
    $Results | % {
        switch ($_){
            {$_.properties.objectclass -contains "computer"}{
                $Computers += $_
            }
            {$_.properties.objectclass -contains "organizationalUnit"}{
                $OUs += $_
            }
            {$_.properties.samaccounttype -eq "805306368"}{
                $users += $_
            }
            {$_.properties.objectclass -contains "groupPolicyContainer"}{
                $gpo += $_
            }
            
            {$_.properties.objectclass -contains "group"}{
                switch ($_.properties.grouptype){
                    $Const_Group_Security_Universal{
                        $GroupType = "Security"
                        $GroupScope = "Universal"
                    }
                    $Const_Group_Security_Global{
                        $GroupType = "Security"
                        $GroupScope = "Global"
                    }
                    $Const_Group_Security_Local{
                        $GroupType = "Security"
                        $GroupScope = "Local"
                    }
                    $Const_Group_Distro_Universal{
                        $GroupType = "Distribution"
                        $GroupScope = "Universal"
                    }
                    $Const_Group_Distro_Global{
                        $GroupType = "Distribution"
                        $GroupScope = "Global"
                    }
                    $Const_Group_Distro_Local{
                        $GroupType = "Distribution"
                        $GroupScope = "Local"
                    }
                }
                $_ | Add-Member -MemberType "NoteProperty" -Name "GroupType" -Value $GroupType
                $_ | Add-Member -MemberType "NoteProperty" -Name "GroupScope" -Value $GroupScope
                $Groups += $_
            }
        }
    }

    #Parse data into useful Powershell objects.
    #(Because the search was done only once against AD, this will speed up the sorting of information dramatically.)
    Write-Host "Parsing GPO Data"
    $GPO_Sorted = $GPO `
        | Select `
            @{Name="DisplayName"; Expression={$_.properties.Item("displayName")}}, `
            @{Name="GUID"; Expression={$_.properties.Item("name")}}, `
            @{Name="Version"; Expression={$_.properties.Item("versionnumber")}}, `
            @{Name="FilePath"; Expression={$_.properties.Item("gpcfilesyspath")}}, `
            path `
        | Sort DisplayName
    Write-Host "Parsing Computer Data"
    $Computers_Sorted = $Computers `
        | Select `
            @{Name="Name"; Expression={$_.properties.Item("name")}}, `
            @{Name="LastLogon"; Expression={$_.properties.Item("lastlogontimestamp") | % {[DateTime]::FromFileTime($_)}}}, `
            @{Name="PwdLastSet"; Expression={$_.properties.Item("pwdlastset") | % {[DateTime]::FromFileTime($_)}}}, `
            @{Name="OperatingSystem"; Expression={$_.properties.Item("operatingsystem")}}, `
            @{Name="ServicePack"; Expression={$_.properties.Item("operatingsystemservicepack")}}, `
            @{Name="displayName"; Expression={$_.properties.Item("displayName")}}, `
            @{Name="sAMAccountName"; Expression={$_.properties.Item("sAMAccountName")}}, `
            @{Name="Description"; Expression={$_.properties.Item("description")}}, `
            @{Name="MemberOf"; Expression={($_.properties.Item("memberOf") | % {(($_ -split 'CN=')[1] -split ',OU=')[0]}) -join "; "}}, `
            @{Name="distinguishedName"; Expression={$_.properties.Item("distinguishedName")}}, `
            path `
        | Sort Name
    Write-Host "Parsing OU Data"
    $OUs_Sorted = $OUs `
        | Select `
            @{Name="Name"; Expression={$_.properties.Item("name")}}, `
            @{Name="Description"; Expression={$_.properties.Item("description")}}, `
            @{Name="Group Policy Links"; Expression={(($_.properties.Item("gpLink") -split 'LDAP://') | 
                ? {$_ -like "*{*}*"} | 
                % {(($_ -split '{')[1] -split '}')[0]}) -join "; "}}, `
            @{Name="distinguishedName"; Expression={$_.properties.Item("distinguishedName")}}, `
            path `
        | Sort Name
    Write-Host "Parsing User Data"
    $Users_Sorted = $Users `
        | Select `
            @{Name="Name"; Expression={$_.properties.Item("name")}}, `
            @{Name="sAMAccountName"; Expression={$_.properties.Item("sAMAccountName")}}, `
            @{Name="LastLogon"; Expression={$_.properties.Item("lastlogontimestamp") | 
                % {[DateTime]::FromFileTime($_)}}}, `
            @{Name="PwdLastSet"; Expression={$_.properties.Item("pwdlastset") | 
                % {[DateTime]::FromFileTime($_)}}}, `
            @{Name="userPrincipalName"; Expression={$_.properties.Item("userPrincipalName")}}, `
            @{Name="ExployeeType"; Expression={$_.properties.Item("EmployeeType")}}, `
            @{Name="Mail"; Expression={$_.properties.Item("mail")}}, `
            @{Name="TelephoneNumber"; Expression={$_.properties.Item("telephoneNumber")}}, `
            @{Name="Location"; Expression={$_.properties.Item("l")}}, `
            @{Name="Country"; Expression={$_.properties.Item("co")}}, `
            @{Name="Department"; Expression={$_.properties.Item("department")}}, `
            @{Name="Office"; Expression={$_.properties.Item("physicaldeliveryofficename")}}, `
            @{Name="ProfilePath"; Expression={$_.properties.Item("profilePath")}}, `
            @{Name="Mail Server"; Expression={$_.properties.Item("homeMDB")}}, `
            @{Name="MemberOf"; Expression={($_.properties.Item("memberOf") | 
                % {(($_ -split 'CN=')[1] -split ',OU=')[0]}) -join "; "}}, `
            path `
        | Sort Name
    Write-Host "Parsing Group Data"
    $Groups_Sorted = $Groups `
        | Select `
            @{Name="Name"; Expression={$_.properties.Item("name")}}, `
            GroupType, `
            GroupScope, `
            @{Name="sAMAccountName"; Expression={$_.properties.Item("sAMAccountName")}}, `
            @{Name="Mail"; Expression={$_.properties.Item("mail")}}, `
            @{Name="Member"; Expression={($_.properties.Item("member") | % {(($_ -split 'CN=')[1] -split ',OU=')[0]}) -join "; "}}, `
            @{Name="MemberOf"; Expression={($_.properties.Item("memberOf") | % {(($_ -split 'CN=')[1] -split ',OU=')[0]}) -join "; "}}, `
            @{Name="distinguishedName"; Expression={$_.properties.Item("distinguishedName")}}, `
            path `
        | Sort Name
   
    #Export the data into CSV files.
    Write-Host "Exporting Data to $SaveTo"

    Join-Path -Path $script:SaveTo -ChildPath "$($env:userdomain)_AD_Report_$(Get-Date -f "yyyy-MM-dd-hhmmsstt").csv"


    $Computers_Sorted | Select * | export-csv (Join-Path -Path $SaveTo -ChildPath "$($env:userdomain)_AD_Report_Computers_$(Get-Date -f "yyyy-MM-dd-hhmmsstt").csv") -NoTypeInformation
    $OUs_Sorted | Select * | export-csv (Join-Path -Path $SaveTo -ChildPath "$($env:userdomain)_AD_Report_OUs_$(Get-Date -f "yyyy-MM-dd-hhmmsstt").csv") -NoTypeInformation
    $Users_Sorted | Select * | export-csv (Join-Path -Path $SaveTo -ChildPath "$($env:userdomain)_AD_Report_Users_$(Get-Date -f "yyyy-MM-dd-hhmmsstt").csv") -NoTypeInformation
    $Groups_Sorted | Select * | export-csv (Join-Path -Path $SaveTo -ChildPath "$($env:userdomain)_AD_Report_Groups_$(Get-Date -f "yyyy-MM-dd-hhmmsstt").csv") -NoTypeInformation
    $GPO_Sorted | Select * | export-csv (Join-Path -Path $SaveTo -ChildPath "$($env:userdomain)_AD_Report_GPO_$(Get-Date -f "yyyy-MM-dd-hhmmsstt").csv") -NoTypeInformation
}

Write-Host "Completed in $ExecutionTime"

#************** Notes for administrator ******************
# $Computers_Sorted will contain all parsed computer objects.
# $OUs_Sorted will contain all parsed OU objects.
# $Users_Sorted will contain all parsed user objects.
# $groups_sorted will contain all parsed group objects.