#Requires -Version 3.0

<#
.SYNOPSIS
    Parses saved evt log files and exports them to Splunk.
.DESCRIPTION
    Parses saved evt log files using advanced XML query filters.
    Extracts embedded EventData and appends it to the parent event.
    Exports the events to Splunk via the HTTP event collector API.
.EXAMPLE
    . C:\Scripts\Send-EvtLogsToSplunk.ps1 -LogFilePath '\\server2\path2' -DataProperty 'john.doe'
.INPUTS
    System.String
.OUTPUTS
    None
    System.Diagnostics.Eventing.Reader.EventLogRecord
.PARAMETER LogFilePath
    The path/paths where saved evt files are stored.
.PARAMETER DataProperty
    Event.EventData.Data property value to search for.
.PARAMETER SplunkServer
    FQDN of Splunk server.
.PARAMETER Port
    TCP port number for the HTTP connection.
.PARAMETER ApiKey
    API Key for authorization header.
.PARAMETER Protocol
    HTTP or HTTPS
.PARAMETER EndPoint
    Relative API endpoint for the Splunk HTTP event collector.
.PARAMETER SplunkHecUrl
    Splunk HTTP Event Collector endpoint to send events to.
.PARAMETER PassThru
    Switch parameter to output events to the pipeline
.PARAMETER Provider
    Provider for the XML EventLog filter.
.PARAMETER Headers
    Splunk HEC REST headers hashtable.
.PARAMETER Epoch
    Reference to epoch (1 JAN 1970) for calculating Splunk 'time' field.
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
.LINK
    http://dev.splunk.com/view/event-collector/SP-CAAAE6P#meta
.LINK
    https://github.com/RamblingCookieMonster/PowerShell/blob/master/Get-WinEventData.ps1
.LINK
    https://blogs.technet.microsoft.com/ashleymcglone/2013/08/28/
.LINK
    https://community.spiceworks.com/scripts/show/3239-select-winevent-make-custom-objects-from-get-winevent
#>

[cmdletbinding()]

param (

    [parameter(ValueFromPipeline=$True)]
    [string[]] $LogFilePath = '\\server1\path1',

    [string] $DataProperty = $env:USERNAME,

    [string] $SplunkServer = 'splunk.example.com',

    [string] $Port = '8088',

    [string] $ApiKey = 'XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX',

    [parameter()]
    [ValidateSet('http', 'https')]
    [string] $Protocol = 'http',

    [string] $EndPoint = '/services/collector',

    [string] $SplunkHecUrl = "$($Protocol + '://' + $SplunkServer + ':' + $Port + $EndPoint)",

    [switch] $PassThru,

    [string] $Provider = 'Microsoft-Windows-Security-Auditing',

    [System.Collections.Hashtable] $Headers = @{ Authorization = "Splunk $ApiKey" },

    [datetime] $Epoch = (Get-Date -Date '01/01/1970')

) #param

begin {} #begin

process {

    $LogFiles = Get-ChildItem -Path $LogFilePath -Filter *.evt | Sort-Object -Property LastWriteTime -Descending

    $LogFiles | ForEach-Object {

        # This filter will get events related to creating,
        # changing, or deleting Active Directory objects.
        $QueryString = @"
<QueryList>
  <Query Id='0' Path='file://$($_.FullName)'>
    <Select Path='file://$($_.FullName)'>
      *[System[Provider[@Name='$Provider'] and
        (
          EventID=4720 or
          EventID=4722 or
          (EventID &gt;= 4724 and EventID &lt;= 4729) or
          EventID=4732 or
          EventID=4733 or
          EventID=4740 or
          EventID=4741 or
          EventID=4743 or
          EventID=4756 or
          EventID=4757
        )
      ]]
      and
      *[EventData[Data='$DataProperty']]
    </Select>
  </Query>
</QueryList>
"@

        $QueryXml = [xml]$QueryString

        $ErrorActionPreferenceBak = $ErrorActionPreference

        # Set ErrorActionPreference to Stop for the trycatch
        $ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

        try {

            $Events = Get-WinEvent -FilterXml $QueryXml -Oldest

        } catch {

            $Events = $false

        } #trycatch

        $ErrorActionPreference = $ErrorActionPreferenceBak

        # If any events are returned, continiue.
        if (!($Events -eq $false)) {

            ForEach ($Event in $Events) {

                # Get a native XML version of each event
                $EventXml = [xml]$Event.ToXml()

                $XmlData = $null

                # Check to see if there's anything in the Event.EventData.Data nodes
                if ($XmlData = @($EventXml.Event.EventData.Data)) {

                    # Loop through the EventData fields.
                    # Use an integrated loop counter based on the total count of Data fields.
                    for ($i=0; $i -lt $XmlData.Count; $i++) {

                        # Append each one as a NoteProperty to the parent event.
                        $SplatArgs = @{ InputObject = $Event ;
                                        MemberType = "NoteProperty" ;
                                        Name = "$($XmlData[$i].name)" ;
                                        Value = "$($XmlData[$i].'#text')" ;
                                        Force = $true ;
                                        Passthru = $true }

                        $Event = Add-Member @SplatArgs

                    } #for

                } #if

            } #ForEach (`$Event in `$Events)

            $Events | ForEach-Object {

                $EachEvent = $_

                # Convert the events to JSON for easy Splunk input.
                $JsonEvent = $_ | ConvertTo-Json -Compress

                # Account for the fact that epoch needs to be in UTC
                [string] $EpochTime = $(
                    (
                        (
                            New-TimeSpan -Start $Epoch -End (
                                [system.timezoneinfo]::ConvertTime(
                                    ($_.TimeCreated),([system.timezoneinfo]::UTC)
                                )
                            )
                        ).TotalSeconds
                    ).ToString()
                )

                # Build the HTTP body. Be sure to follow proper JSON formatting.
                # You could also add a sourcetype to the top level of the JSON object.
                # See: 'http://dev.splunk.com/view/event-collector/SP-CAAAE6P#meta' for details.
                $Body = "{`"time`": `"$EpochTime`",`"host`": `"$($_.MachineName)`",`"event`": $JsonEvent}"

                $SplatArgs = @{    Uri = $SplunkHecUrl ;
                                Headers = $Headers ;
                                Method = 'Post' ;
                                Body = $Body }

                # Send the event to Splunk HTTP event collector.
                Invoke-RestMethod @SplatArgs | Out-Null

                if ($PassThru) {

                    $EachEvent

                }

            } #ForEach-Object (`$Events)

        } #if

    } #ForEach-Object (`$LogFiles)

} #process

end {} #end
