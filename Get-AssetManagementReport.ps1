Function Get-AssetManagementReport {

[cmdletbinding()]

Param (

  [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
  [Alias('DNSHostName','PSComputerName','CN','Hostname')]
  [array] $ComputerName = @($env:COMPUTERNAME),

  [string] $IpAddressRegEx = (
    '^(?:(?:1\d\d|2[0-5][0-5]|2[0-4]\d|0?[1-9]\d|0?0?\d)\.){3}(?:1\d\d|2[0-5][0-5]|2[0-4]\d|0?[1-9]\d|0?0?\d)$'
  )

)

Begin {

  Write-Verbose "Setting ErrorActionPreference = 'Stop'"

  $ErrorActionPreference = 'Stop'

} # End of Begin ScriptBlock

Process {

$ComputerName | ForEach-Object {

  Write-Verbose "Testing resolution for $_"

  try {

    $DnsHostEntry = [System.Net.Dns]::GetHostEntry("$_")

    $ResolvedComputerName = (($DnsHostEntry.HostName) -ireplace '\..+','')

    $IpAddress = $DnsHostEntry.AddressList.IPAddressToString |
      Where-Object { $_ -match $IpAddressRegEx }

  } catch [System.Net.Sockets.SocketException] {

    Write-Warning "Failed to resolve $_ to a ComputerName!"

  } # End try catch for Name/Address resolutuion

  Write-Verbose "Testing reachability for $ResolvedComputerName"

  try {

    Write-Verbose "Testing to see if $ResolvedComputerName is reachable"

    $Reachable = Test-Connection -Count 1 -ComputerName $ResolvedComputerName -quiet

    Write-Verbose "$ResolvedComputerName is reachable"

  } catch {

    if ($ResolvedComputerName -ne $null ) {

      Write-Warning "$ResolvedComputerName is NOT reachable!"

    } # End of if

  } # End try catch for Reachability

  if ($Reachable) {

    try {

      $CimSession = New-CimSession -ComputerName $ResolvedComputerName

    } catch {

      try {

        Write-Warning "Failed to connect to $ResolvedComputerName with WSMAN, trying DCOM"

        $CimSessionDcomOption = New-CimSessionOption -Protocol Dcom

        $CimSession = New-CimSession `
          -ComputerName $ResolvedComputerName `
          -SessionOption $CimSessionDcomOption

      } catch {

        Write-Warning "Failed to connect to $ResolvedComputerName, skipping."

        $failure = $true

      } # End try catch (last resort)

    } # End try catch (Set up CIM session)

    if (!($failure -eq $true)) {

      $CimUserProfile = Get-CimInstance -CimSession $CimSession `
        -ClassName Win32_UserProfile `
        -Filter 'Special=False AND SID like "S-1-5-21-%" AND not SID like "%-500"' |
        Sort-Object -Property LastUseTime -Descending |
        Select-Object -First 1 |
        ForEach-Object {
          Add-Member -InputObject $_ -MemberType NoteProperty -Name UserName -Value (
            ($_.LocalPath).Replace('C:\Users\','')
          ) -Force -PassThru
        } | Select-Object -Property PSComputerName,
                                    SID,
                                    UserName,
                                    LastUseTime

      $CimComputerSystemProduct = Get-CimInstance -CimSession $CimSession `
        -ClassName Win32_ComputerSystemProduct

      $CimDefaultInterfaceIndex = Get-CimInstance -CimSession $CimSession `
        -ClassName Win32_IP4RouteTable `
        -Filter "Destination='0.0.0.0'" |
        Select-Object -ExpandProperty InterfaceIndex

      $CimMacAddress = Get-CimInstance -CimSession $CimSession `
        -ClassName CIM_NetworkAdapter |
        Where-Object {
          $_.InterfaceIndex -in $CimDefaultInterfaceIndex
        } |
        Select-Object -ExpandProperty MACAddress

      $CimDiskDrive = Get-CimInstance -CimSession $CimSession `
        -ClassName CIM_DiskDrive `
        -Filter "Name='\\\\.\\PHYSICALDRIVE0'" |
        Select-Object -Property Model,
                                FirmwareRevision,
                                SerialNumber

      ##
      ## The Monitor SerialNumber query is based on one of Jason Hofferle's blog posts.
      ## http://www.hofferle.com/retrieve-monitor-serial-numbers-with-powershell/
      ##

      $CimMonitors = Get-CimInstance -CimSession $CimSession `
        -ClassName wmiMonitorID -Namespace 'root\wmi'

      Remove-CimSession -CimSession $CimSession

      $monitorInfo = @()

      ForEach ($Monitor in $CimMonitors) {

        $mon = New-Object -TypeName psobject
        $name = $null
        $serial = $null

        $Monitor.SerialNumberID |
          ForEach-Object {
            $serial += [char] $_
          }

        $Monitor.UserFriendlyName |
          ForEach-Object {
            $name += [char] $_
          }

        $mon | Add-Member -MemberType NoteProperty -Name Name -Value $name
        $mon | Add-Member -MemberType NoteProperty -Name SerialNumber -Value $serial

        $MonitorInfo += $mon

      } # End of ForEach ($Monitor in $CimMonitors)

      $MonitorCount = ($MonitorInfo | Measure-Object).Count

      if ($MonitorCount -eq 1) {

        $Monitor1 = $MonitorInfo[0]
        $Monitor2 = New-Object -TypeName psobject -Property (
          @{ 'Name' = $null ;
             'SerialNumber' = $null
          }
        )

      }

      if ($MonitorCount -eq 2) {

        $Monitor1 = $MonitorInfo[0]
        $Monitor2 = $MonitorInfo[1]

      }

      $AssetManagementReport = New-Object -TypeName psobject -Property (
        @{ 'ComputerName' = $ResolvedComputerName ;
           'IpAddress' = $IpAddress ;
           'SerialNumber' = $CimComputerSystemProduct.IdentifyingNumber ;
           'LastLoggedOnUser' = $CimUserProfile.UserName ;
           'LastLogonTime' = $CimUserProfile.LastUseTime ;
           'Vendor' = $CimComputerSystemProduct.Vendor ;
           'Model' = $CimComputerSystemProduct.Name ;
           'MacAddress' = $CimMacAddress ;
           'DiskDriveModel' = $CimDiskDrive.Model ;
           'DiskDriveFirmwareRevision' = $CimDiskDrive.FirmwareRevision ;
           'DiskDriveSerialNumber' = $CimDiskDrive.SerialNumber ;
           'MonitorCount' = $MonitorCount ;
           'Monitor1Name' = $Monitor1.Name ;
           'Monitor1SerialNumber' = $Monitor1.SerialNumber ;
           'Monitor2Name' = $Monitor2.Name ;
           'Monitor2SerialNumber' = $Monitor2.SerialNumber
        }
      )

      $AssetManagementReport |

      Select-Object -Property ComputerName,
                              IpAddress,
                              MacAddress,
                              Vendor,
                              Model,
                              SerialNumber,
                              LastLoggedOnUser,
                              LastLogonTime,
                              DiskDriveModel,
                              DiskDriveFirmwareRevision,
                              DiskDriveSerialNumber,
                              MonitorCount,
                              Monitor1Name,
                              Monitor1SerialNumber,
                              Monitor2Name,
                              Monitor2SerialNumber

    } # End of if (and caring)

  } # End if (Reachability)

} # End of ForEach-Object `$ComputerName ScriptBlock

} # End of Process ScriptBlock

End {} # End of End ScriptBlock

<#

  .SYNOPSIS

    Gathers asset intelligence on endpoints.


  .DESCRIPTION

    Gathers asset intelligence on endpoints.


  .PARAMETER ComputerName

    The computer name to query. Defaults to localhost (actually... `$env:COMPUTERNAME)
    Accepts values from the pipeline.


  .EXAMPLE

    Get-AssetManagementReport -ComputerName host1
    Gets info from host1.


  .EXAMPLE

    'host1','host2','host3' | Get-AssetManagementReport
    Gets info from the three host input via the pipeline.


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

#>

} # End of Function Get-AssetManagementReport