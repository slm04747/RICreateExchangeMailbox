# RI-Create_Mailboxes_byFile

#region Variables

# Paths
$InputPath = 'C:\PSScripts\RapidIdentity\Input\'
$ProcessedPath = 'C:\PSScripts\RapidIdentity\Processed\'
$TranscriptPath = 'C:\PSScripts\RapidIdentity\Transcripts\'

# Files
$InputFileNames = $InputPath + 'Users_*.txt'
$TranscriptFile = $TranscriptPath + 'RI-CreateMBFromFile.txt'
$StatusFile = $TranscriptPath + 'RI_UsersProcessed.csv'
$DestinationName = ''

# To Be processed

[array]$FilesToProcess = @()
[array]$usersToProcess = @()

# Status Object

$Status = [ordered]@{
    Object     = '';
    ObjectType = '';
    Status     = ''
}

# Miscellaneous

$RemoteSuffix = '@mymailpomona.mail.onmicrosoft.com'
$FileCounter = 0

#endregion Variables


Start-Transcript -Path $TranscriptFile -Append

If (Test-Path $InputFileNames) {

    $FilesToProcess = Get-Childitem $InputFileNames
    Add-PSSnapin Microsoft.Exchange.Management.PowerShell.RecipientManagement

    Foreach ($File in ($FilesToProcess)) {

        $usersToProcess = Get-Content $File
        $UserCount = $UsersToProcess.Count

        $Status.Object = $file.Name
        $Status.ObjectType = 'File'
        $Status.Status = "Processing $UserCount entries."

        [pscustomobject]$Status | Export-csv -Path $StatusFile -NoTypeInformation -Append

        ForEach ($User in $usersToProcess) {


            # Remove leading and trailing spaces
            $User = $User.trim() 
            $Status.Object = $User
            $Status.ObjectType = 'User'

            If ($User.length -gt 0) {
                Try {
                    # Does User Exist

                    $UserType = Get-User $User -ErrorAction Stop | Select-Object recipienttype

                    # Enable the remote mailbox 
                    If ($UserType.recipienttype -eq 'User') {
                        Enable-Remotemailbox $User -RemoteRoutingAddress ($User + $RemoteSuffix) -ErrorAction Stop
                        $Status.Status = 'Success'
                    }
                    Else {
                        $Status.Status = 'User already processed - ' + $UserType.recipienttype
                    }
                }
                Catch [Microsoft.Exchange.Configuration.Tasks.ManagementObjectNotFoundException ] {
                    $Status.Status = 'User not found.'
                }
                Catch {
                    $Status.Status = 'Enable remote MBX failure.'
                }
                [pscustomobject]$Status | Export-CSV -Path $StatusFile -NoTypeInformation -Append
            }

        }

        $Status.Object = $file.Name
        $Status.ObjectType = 'File'

        Try {
            $DestinationName = $ProcessedPath + $File.name
            While (Test-Path $DestinationName) {
                $DestinationName = $ProcessedPath + $File.BaseName + '-' + [string]$FileCounter + $File.Extension
                $FileCounter += 1
            }
            Move-Item $File.fullname $DestinationName
            $Status.Status = "Moved to processed."
        }
        Catch {
            $Status.Status = "File move failed."
        }
        [pscustomobject]$status | Export-CSV -Path $StatusFile -NoTypeInformation -Append

    }


}

Stop-Transcript