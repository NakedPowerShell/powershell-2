#Requires -Version 3
#Requires -Modules ActiveDirectory

function Get-LastLoggedOnUser {

    <#

    .SYNOPSIS
        Finds the last logged on user of a computer and looks up their phone number.

    .DESCRIPTION
        Finds the last logged on user of a computer using WMI. Searches CUCM and AD for the user's phone number.

    .EXAMPLE
        Get-LastLoggedOnUser -ComputerName host1
        Finds the last logged on user on host1.

    .EXAMPLE
        'host1','host2' | Get-LastLoggedOnUser
        Finds the last logged on user on host1 and host2 via the pipeline.

    .PARAMETER ComputerName
        ComputerName or IP address to operate on. Defaults to localhost (actually... $env:COMPUTERNAME)
        Accepts values from the pipeline.

    .PARAMETER WmiFilter
        WQL filter used to refine which user profiles are returned.

    .PARAMETER CallManager
        Hostname of IP address of the Cisco Call Manager (CUCM) server.

    .PARAMETER Port
        TCP port that the XML directory web service is listening on.

    .PARAMETER Protocol
        HTTP or HTTPS.

    .PARAMETER AdSearchBase
        Distinguished Name used to constrain the Active Directory search.

    .PARAMETER UserNameFilter
        ScriptBlock to filter the UserName values not filtered by $WmiFilter.
    
    .INPUTS
        System.Object

    .OUTPUTS
        Selected.System.Management.Automation.PSCustomObject

    .NOTES
        #######################################################################################
        Author:     @oregon-national-guard/systems-administration
        Version:    1.0
        #######################################################################################
        License:    https://github.com/oregon-national-guard/powershell/blob/master/LICENCE
        #######################################################################################

    .LINK
        https://github.com/oregon-national-guard

    .LINK
        https://creativecommons.org/publicdomain/zero/1.0/

    #>

    [cmdletbinding()]

    param (
        
        [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
        [Alias('PSComputerName','DNSHostName','CN','Hostname')]
        [string[]]
        $ComputerName,

        [string]
        $WmiFilter = 'Special=False AND SID like "S-1-5-21-%" AND not SID like "%-500"',

        [string]
        $CallManager = 'cucm-hostname',

        [string]
        $Port = '8080',

        [parameter()]
        [ValidateSet('http', 'https')]
        [string]
        $Protocol = 'http',

        [string]
        $AdSearchBase,

        [scriptblock]
        $UserNameFilter

    ) #param

    begin {

        function ConvertFrom-Sid {
            [CmdletBinding()]
            param (
                [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
                [Alias('objectSid')]
                [string[]]
                $SID
            ) #param
            process {
                Add-Type -AssemblyName System.DirectoryServices.AccountManagement
                $SID | ForEach-Object {
                    $EachSid = $_
                    if ($EachSid -eq $null) {
                        $objSID = ([System.DirectoryServices.AccountManagement.UserPrincipal]::Current).Sid
                    } else {
                        $objSID = New-Object System.Security.Principal.SecurityIdentifier("$EachSid")
                    }
                    $objNtAccount = $objSID.Translate([System.Security.Principal.NTAccount])
                    $UserName = Split-Path -Path $($objNtAccount.Value) -Leaf
                    Write-Output -InputObject $UserName
                } #foreach
            } #process
        } #function ConvertFrom-Sid

    } #begin

    process {

        $ComputerName | ForEach-Object {

            $EachComputer = $_

            $SplatArgs = @{ Class = 'Win32_UserProfile'
                            Filter = $WmiFilter
                            ComputerName = $EachComputer }

            $WmiUserProfiles = Get-WmiObject @SplatArgs

            if ($WmiUserProfiles -ne $null) {

                $WmiUserProfiles |

                Sort-Object -Property LastUseTime -Descending | ForEach-Object {

                    $UserName = $_.SID | ConvertFrom-Sid

                    $LastUseTime = $_.ConvertToDateTime($_.LastUseTime)

                    New-Object -TypeName psobject -Property @{  ComputerName = $_.PSComputerName
                                                                UserName = $UserName
                                                                LastUseTime = $LastUseTime }

                } | Select-Object -Property ComputerName,UserName,LastUseTime |

                Where-Object $UserNameFilter | Select-Object -First 1 |

                Tee-Object -Variable LastLoggedOnUser | Out-Null

                $UserName = $($LastLoggedOnUser.UserName)

                $FirstName = $UserName.Split('.') | Select-Object -First 1
                $RawLastName = $UserName.Split('.') | Select-Object -Last 1
                $LastName = $RawLastName -replace '[0-9]+',''

                $DirectoryUri = "$Protocol`://$CallManager`:$Port/ccmcip/xmldirectorylist.jsp?l=$LastName&f=$FirstName"

                $CucmPhoneNumber = (Invoke-RestMethod -Uri $DirectoryUri).CiscoIPPhoneDirectory.DirectoryEntry |
                    Where-Object { $_.Telephone -ne $null } |
                    Select-Object -ExpandProperty Telephone

                $PhoneNumberList = @()

                if ($CucmPhoneNumber -ne $null) {

                    $CucmPhoneNumber | ForEach-Object {

                        $PhoneNumberList += $_

                    } #ForEach

                } #if

                if (-not $AdSearchBase) {

                    $AdSearchBase = Get-ADDomain | Select-Object -ExpandProperty DistinguishedName

                } #if

                $AdFilter = "SamAccountName -Like '*$UserName*'"
                $AdUser = Get-ADUser -SearchBase $AdSearchBase -Filter $AdFilter -Properties ipPhone
                $AdPhoneNumber = $AdUser | Select-Object -ExpandProperty ipPhone

                if ($AdPhoneNumber -ne $null) {

                    $PhoneNumberList += $AdPhoneNumber

                } #if

                if ($PhoneNumberList -ne $null) {

                    $PhoneNumberList = $PhoneNumberList | Sort-Object -Unique

                    [string] $PhoneNumber = ''

                    $PhoneNumberList | ForEach-Object {

                        $PhoneNumber = $PhoneNumber + $(($_ -as [string]).Trim()) + ',' + ' '

                    } #ForEach

                    $PhoneNumber = $PhoneNumber.Trim().Trim(',')

                } #if

                New-Object -TypeName psobject -Property @{  ComputerName = $($LastLoggedOnUser.ComputerName)
                                                            UserName = $($LastLoggedOnUser.UserName)
                                                            LastUseTime = $($LastLoggedOnUser.LastUseTime)
                                                            PhoneNumber = $PhoneNumber } |

                Select-Object -Property ComputerName,UserName,LastUseTime,PhoneNumber

            } else {

                New-Object -TypeName psobject -Property @{  ComputerName = $EachComputer
                                                            UserName = $null
                                                            LastUseTime = $null
                                                            PhoneNumber = $null } |

                Select-Object -Property ComputerName,UserName,LastUseTime,PhoneNumber

            } #if $WmiUserProfiles

        } #ForEach $ComputerName

    } #process

    end {} #end

} #function Get-LastLoggedOnUser