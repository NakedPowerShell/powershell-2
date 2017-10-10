function Get-TargetedEvents  {

    <#
    .SYNOPSIS
        Parses Windows logs for events related to a specific username or other metadata value
    .DESCRIPTION
        Parses Windows logs for events related to a specific username or other metadata value
    .PARAMETER ComputerName
        The computer name/names to operate on. Defaults to `$env:COMPUTERNAME
    .PARAMETER SearchTerm
        The term to search for
    .EXAMPLE
        Get-TargetedEvents -SearchTerm user.name
        Gets local events for user.name
    .EXAMPLE
        'host1','host2' | Get-TargetedEvents -SearchTerm user.name
        Gets events from both hosts input via the pipeline for user.name
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
        [string[]] $ComputerName = @($env:COMPUTERNAME),

        [string] $SearchTerm = $env:USERNAME

    ) #param

    begin {} #begin

    process  {

        $QueryFilterXpath = "*[EventData[Data and (Data='$SearchTerm')]]"

        $ComputerName | ForEach-Object {

            Get-WinEvent -LogName Security -ComputerName $_ -FilterXPath $QueryFilterXpath -ErrorAction SilentlyContinue

        }

    } #process

    end {} #end

} #function Get-TargetedEvents
