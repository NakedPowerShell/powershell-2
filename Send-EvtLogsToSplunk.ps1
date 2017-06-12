#Requires -Version 3.0

[cmdletbinding()]

Param (

	[parameter(ValueFromPipeline=$True)]
	[string[]] $LogFilePath = '\\server1\path1',

	[string] $DataProperty = $env:USERNAME,

	[string] $SplunkHecUri = 'http://splunk.example.com:8088/services/collector',

	[string] $SplunkHecApiKey = 'XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX',

	[switch] $TeeOutput,

	[System.Collections.Hashtable] $SplunkHecRestHeaders = @{ Authorization = "Splunk $SplunkHecApiKey" },

	[datetime] $Epoch = (Get-Date -Date '01/01/1970')

) #Param

Begin {} #Begin

Process {

	$LogFiles = Get-ChildItem -Path $LogFilePath -Filter *.evt | Sort-Object -Property LastWriteTime -Descending

	$LogFiles | ForEach-Object {

	# This filter will get events related to creating,
	# changing, or deleting Active Directory objects.
	$QueryString = @"
<QueryList>
  <Query Id='0' Path='file://$($_.FullName)'>
    <Select Path='file://$($_.FullName)'>
      *[System[Provider[@Name='Microsoft-Windows-Security-Auditing'] and
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

	$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop

	try {

		$Events = Get-WinEvent -FilterXml $QueryXml -Oldest

	} catch {

		$Events = $false

	} #trycatch

	$ErrorActionPreference = $ErrorActionPreferenceBak

	if (!($Events -eq $false)) {

		ForEach ($Event in $Events) {

			# Get a native XML version of each event
			$EventXml = [xml]$Event.ToXml()

			$XmlData = $null

			# Check to see if there's anything in the Event.EventData.Data nodes
			if ($XmlData = @($EventXml.Event.EventData.Data)) {

				# Iterate through the EventData fields.
				for ($i=0; $i -lt $XmlData.Count; $i++) {

					# Append each one as a NoteProperty to the parent event.
					$SplatArgs = @{	InputObject = $Event ;
									MemberType = "NoteProperty" ;
									Name = "$($XmlData[$i].name)" ;
									Value = "$($XmlData[$i].'#text')" ;
									Force = $true ;
									Passthru = $true }

					$Event = Add-Member @SplatArgs

				} #for

			} #if

		} #ForEach

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

			# Build the HTTP body being sure to follow JSON formatting.
			# You could also add a sourcetype to the top level of the JSON object.
			$Body = "{`"time`": `"$EpochTime`",`"host`": `"$($_.MachineName)`",`"event`": $JsonEvent}"

			$SplatArgs = @{	Uri = $SplunkHecUri ;
							Headers = $SplunkHecRestHeaders ;
							Method = 'Post' ;
							Body = $Body }

			# Send the event to Splunk HTTP event collector.
			Invoke-RestMethod @SplatArgs | Out-Null

			if ($TeeOutput) {

				$EachEvent

			}

		} #ForEach-Object

	} #if

	} #ForEach-Object

} #Process

End {} #End

<#
.SYNOPSIS
	Parses saved evt log files and exports them to Splunk.

.DESCRIPTION
	Parses saved evt log files using advanced XML query filters.
	Extracts embedded EventData and appends it to the parent event.
	Exports the events to Splunk via the HTTP event collector API.

.PARAMETER LogFilePath
	The path/paths where saved evt files are stored.

.PARAMETER DataProperty
	Event.EventData.Data property value to search for.

.PARAMETER SplunkHecUri
	Splunk HTTP Event Collector endpoint to send events to.

.PARAMETER SplunkHecApiKey
	API Key for authorization header.

.PARAMETER TeeOutput
	Switch parameter to output events to the pipeline

.PARAMETER SplunkHecRestHeaders
	Actual Splunk HEC REST headers hashtable

.PARAMETER Epoch
	Reference to epoch (1 JAN 1970) for calculating Splunk 'time' field.

.EXAMPLE
	. .\Send-EvtLogsToSplunk.ps1 -LogFilePath '\\server2\path2' -DataProperty 'john.doe'

.NOTES
    ###################################################################
    Author:     @oregon-national-guard/systems-administration
    Version:    1.0
    ###################################################################
    License
    -------
    This Work was prepared by a United States Government employee and,
    therefore, is excluded from copyright by Section 105 of the Copyright
    Act of 1976. Copyright and Related Rights in the Work worldwide are
    waived through the CC0 1.0 Universal license. Portions of specific
    scripts are licensed under Microsoft Limited Public License.
    ###################################################################
    Disclaimer of Warranty
    ----------------------
    This Work is provided "as is." Any express or implied warranties,
    including but not limited to, the implied warranties of merchantability
    and fitness for a particular purpose are disclaimed. In no event shall
    the United States Government be liable for any direct, indirect,
    incidental, special, exemplary or consequential damages (including,
    but not limited to, procurement of substitute goods or services,
    loss of use, data or profits, or business interruption) however caused
    and on any theory of liability, whether in contract, strict liability,
    or tort (including negligence or otherwise) arising in any way out of
    the use of this Guidance, even if advised of the possibility of such damage.
    The User of this Work agrees to hold harmless and indemnify the
    United States Government, its agents and employees from every claim or
    liability (whether in tort or in contract), including attorneys' fees,
    court costs, and expenses, arising in direct consequence of Recipient's
    use of the item, including, but not limited to, claims or liabilities
    made for injury to or death of personnel of User or third parties, damage
    to or destruction of property of User or third parties, and infringement or
    other violations of intellectual property or technical data rights.
    Nothing in this Work is intended to constitute an endorsement, explicit or implied,
    by the United States Government of any particular manufacturer's product or service.
    ###################################################################
    Disclaimer of Endorsement
    -------------------------
    Reference herein to any specific commercial product, process,
    or service by trade name, trademark, manufacturer, or otherwise,
    in this Work does not constitute an endorsement, recommendation,
    or favoring by the United States Government and shall not be used
    for advertising or product endorsement purposes.
    ###################################################################

.LINK
	https://github.com/oregon-national-guard

.LINK
	https://creativecommons.org/publicdomain/zero/1.0/

.LINK
	https://github.com/RamblingCookieMonster/PowerShell/blob/master/Get-WinEventData.ps1

.LINK
	https://blogs.technet.microsoft.com/ashleymcglone/2013/08/28/powershell-get-winevent-xml-madness-getting-details-from-event-logs/

.LINK
	https://community.spiceworks.com/scripts/show/3239-select-winevent-make-custom-objects-from-get-winevent
#>