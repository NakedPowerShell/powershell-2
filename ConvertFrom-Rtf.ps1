#Requires -Version 3.0
function ConvertFrom-Rtf {

[cmdletbinding()]

param (

    [parameter(ValueFromPipeline=$True,ValueFromPipelineByPropertyName=$True)]
    [Alias('PSPath','FullName')]
    [array] $PathToFile = @((Get-Item 'C:\Windows\Help\en-US\credits.rtf')),

    [switch] $AsObject,

    [switch] $Hash,

    [datetime] $epoch = '01/01/1970'

) # param

begin {

    $StartTime = (Get-Date)

    $NameOfFunction = 'ConvertFrom-Rtf'

    

    $CountFiles = ($PathToFile | Measure-Object).Count

    $ErrorActionPreference = 'Stop'

    Add-Type -AssemblyName System.Windows.Forms

    #
    # Thank @kmsigma for the rtf conversion part.
    # I just adapted it for the pipeline and put some
    # window dressing on it to fit my use case.
    #

} # begin

process {

    $PathToFile | ForEach-Object {

        $RichTextBox = New-Object -TypeName System.Windows.Forms.RichTextBox

        try {

            $RichTextBox.Rtf = [System.IO.File]::ReadAllText($_.FullName)

        } catch {

            Remove-Variable RichTextBox -ErrorAction SilentlyContinue

            $ExceptionName = $_.Exception.GetType().FullName

            $RichTextBox = New-Object -TypeName psobject -Property @{
                "Text" = "#!#!# Failed to convert with a $ExceptionName exception #!#!#"
            }

        } finally {

            if ($AsObject) {

                if ($Hash) {

                    New-Object -TypeName psobject -Property @{

                        "Name" = $_.Name
                        "FullName" = $_.FullName
                        "Text" = $RichTextBox.Text
                        "LastWriteTime" = $(($_.LastWriteTime) -as [datetime])
                        "EpochTime" = $(
                            (New-TimeSpan -Start $epoch -End (
                                [system.timezoneinfo]::ConvertTime(
                                    ($_.LastWriteTime),([system.timezoneinfo]::UTC)
                                ))
                            ).TotalSeconds
                        )
                        "TimeStamp" = $((($_.LastWriteTime).GetDateTimeFormats('s')) -as [string])
                        "Length" = $_.Length
                        "SHA256" = $((Get-FileHash -Path $_.FullName -Algorithm SHA256).Hash)

                    } |

                    Select-Object -Property LastWriteTime,
                                            TimeStamp,
                                            EpochTime,
                                            Name,
                                            FullName,
                                            Length,
                                            SHA256,
                                            Text

                    Remove-Variable RichTextBox -ErrorAction SilentlyContinue

                } elseif (!($Hash)) {

                    New-Object -TypeName psobject -Property @{

                        "Name" = $_.Name
                        "FullName" = $_.FullName
                        "Text" = $RichTextBox.Text
                        "LastWriteTime" = $(($_.LastWriteTime) -as [datetime])
                        "EpochTime" = $(
                            (New-TimeSpan -Start $epoch -End (
                                [system.timezoneinfo]::ConvertTime(
                                    ($_.LastWriteTime),([system.timezoneinfo]::UTC)
                                ))
                            ).TotalSeconds
                        )
                        "TimeStamp" = $((($_.LastWriteTime).GetDateTimeFormats('s')) -as [string])
                        "Length" = $_.Length

                    } |

                    Select-Object -Property LastWriteTime,
                                            TimeStamp,
                                            EpochTime,
                                            Name,
                                            FullName,
                                            Length,
                                            Text

                    Remove-Variable RichTextBox -ErrorAction SilentlyContinue

                } # if else (`$Hash)

            } elseif (!($AsObject)) {

                $RichTextBox.Text

                Remove-Variable RichTextBox -ErrorAction SilentlyContinue

            } # if else (`$AsObject)

        } # try catch finally

    } # ForEach-Object

} # process

end {

    $EndTime = (Get-Date)

    Write-Verbose "Finished running $NameOfFunction on $CountFiles files"

    Write-Verbose "Elapsed Time: $(($EndTime - $StartTime).TotalSeconds) seconds"

} # end

<#

    .SYNOPSIS

        Converts richtext (rtf) documents to plaintext


    .DESCRIPTION

        Converts richtext (rtf) documents to plaintext and outputs
        text strings or psobjects based on a switch.


    .EXAMPLE

        Get-ChildItem 'C:\Windows'  -Recurse `
                                    -File `
                                    -Filter *.rtf `
                                    -ErrorAction SilentlyContinue |
                                    Select-Object -First 5 |
                                    ConvertFrom-Rtf
        Converts first five rtf files in C:\Windows to plaintext


    .INPUTS

        System.IO.FileInfo


    .OUTPUTS

        System.Management.Automation.PSCustomObject


    .PARAMETER PathToFile

        System.IO.FileInfo objects to operate on


    .PARAMETER AsObject

        Switch to output psobjects instead of strings


    .PARAMETER Hash

        Switch to hash files or not


    .PARAMETER epoch

        DateTime of epoch (01/01/1970)


    .NOTES

        ###################################################################

        Author:     @oregon-national-guard/cyberspace-operations
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

        https://creativecommons.org/publicdomain/zero/1.0/


    .LINK

        https://github.com/oregon-national-guard/powershell


    .LINK

        http://blog.kmsigma.com/2014/10/01/converting-rtf-to-txt-via-powershell/

#>

} # function ConvertFrom-Rtf