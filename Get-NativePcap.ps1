#Requires -Modules PEF
Function Get-NativePcap {

[cmdletbinding()]

Param (

  [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
  [Alias('DNSHostName','PSComputerName','CN','Hostname')]
  [array] $ComputerName = @($env:COMPUTERNAME)

) # End of Param

Begin {

  $StartTime = (Get-Date)

  if (!(Get-WinEvent -ListProvider 'Packet-Capture-Script' -ErrorAction SilentlyContinue)) {

    New-EventLog -LogName Application -Source 'Packet-Capture-Script'

  } # if (Check for EventLog)

  ## This function will fail without the PEF module that comes with Microsoft Message Analyzer
  ## If Microsoft Message Analyzer isn't installed, uncomment the following,
  ## remove the '#Requires -Modules PEF' at the top, and run it again.

  <#

  $MessageAnalyzer = Get-CimInstance -ClassName CIM_Product -Filter "Name LIKE '%Microsoft Message Analyzer%'"

  if (!($MessageAnalyzer)) {

    (New-Object System.Net.WebClient).DownloadFile(
      'https://download.microsoft.com/download/2/8/3/283DE38A-5164-49DB-9883-9D1CC432174D/MessageAnalyzer64.msi',
      'C:\Windows\Temp\MessageAnalyzer64.msi'
    )

    Unblock-File -Path 'C:\Windows\Temp\MessageAnalyzer64.msi'

    [array] $InstallArgs = @(
      '/i',
      'C:\Windows\Temp\MessageAnalyzer64.msi',
      '/quiet',
      '/norestart'
    )

    Start-Process -FilePath 'C:\Windows\System32\msiexec.exe' -ArgumentList $InstallArgs -Verb runas -Wait | Out-Null

  } # if (!(`$MessageAnalyzer))

  #>

  Import-Module PEF

} # End of Begin ScriptBlock

Process {

  $ComputerName | ForEach-Object {

    Invoke-Command -ComputerName $_ -ScriptBlock {

      Write-Verbose 'setting arguments to start the trace'

      [array] $NetshStartArgs = @(
        'trace',
        'start',
        'capture=yes',
        'persistent=yes',
        'overwrite=yes',
        'tracefile=C:\Windows\Temp\trace.etl'
      )

      Write-Verbose 'starting trace with netsh'

      Start-Process -FilePath 'C:\Windows\System32\netsh.exe' -ArgumentList $NetshStartArgs -Verb runas -Wait | Out-Null

      Write-Verbose 'starting 60 second timer'

      Start-Sleep -Seconds 60

      Write-Verbose 'setting arguments to stop the trace'

      [array] $NetshStopArgs = @(
        'trace',
        'stop'
      )

      Write-Verbose 'stopping trace'

      Start-Process -FilePath 'C:\Windows\System32\netsh.exe' -ArgumentList $NetshStopArgs -Verb runas -Wait | Out-Null

    } # End of Invoke-Command ScriptBlock

    $RemoteEtlPath = '\\' + $_ + '\C$\Windows\Temp\trace.etl'

    $DestCapPath = $('C:\Windows\Temp\' + $(Get-Date -Format 'YYYY-M-D') + '-' + $_ + '-' + 'trace.cap')

    $Capture = New-PefTraceSession -Path $DestCapPath -SaveOnStop

    $Capture | Add-PefMessageProvider -Provider $RemoteEtlPath | Out-Null

    $Capture | Start-PefTraceSession | Out-Null

    Write-Host ''

    Write-Host "Packet capture for $_ saved to: $DestCapPath"

    Write-EventLog -LogName Application -EntryType Information -EventId 1 -Source 'Packet-Capture-Script' -Message (
      "$env:USERNAME just Captured Packets on $($($_.ToUpper())) from $($($env:COMPUTERNAME.ToUpper()))"
    )

  } # End of ForEach-Object (`$ComputerName)

} # Process

End {

  $EndTime = (Get-Date)

  Write-Verbose "Elapsed Time: $(($EndTime - $StartTime).TotalSeconds) seconds"

}

<#

  .SYNOPSIS

    Gets an native ETL trace from a computer and converts it to a CAP/PCAP file.


  .DESCRIPTION

    Uses PSRemoting to collect a native ETL trace from a remote machine, 
    then uses the PEF module to convert from ETL to CAP/PCAP for analysis. 


  .PARAMETER ComputerName

    The computer name/names to operate on. Defaults to `$env:COMPUTERNAME


  .EXAMPLE
    Get-NativePcap -ComputerName host1
    Gets a PCAP file from host1

  .EXAMPLE

    'host1','host2','host3' | Get-NativePcap
    Gets a PCAP file from the three host input via the pipeline


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

    https://download.microsoft.com/download/2/8/3/283DE38A-5164-49DB-9883-9D1CC432174D/MessageAnalyzer64.msi

#>

} # End of Get-NativePcap Function