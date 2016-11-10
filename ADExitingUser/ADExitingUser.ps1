<#
 AD actions script for exiting staff
 By Alex Castaneda (dewtturbo@gmail.com)

 This script does the following:

 * Removes AD Account from all groups except groups specified in $KeepGroups
 * Adds AD Account to group specified in $0DayVaultGroup
 * Adds expiration date 10 days from today
 * Moves AD Account to the OU specified in $DisabledUsersOU

 The script will create a backup before modifying account info that can be used to recover the groups and location previous to processing

 Comment out the $TestMode variable to make the script write changes
#>

$TestMode = $true


#region setup
$KeepGroups = 'CN=Group 1,OU=Groups,DC=Contoso,DC=lcl
CN=Group 1,OU=Groups,DC=Contoso,DC=lcl
CN=Group 2,OU=Groups,DC=Contoso,DC=lcl
CN=Group 3,OU=Groups,DC=Contoso,DC=lcl
CN=Group 4,OU=Groups,DC=Contoso,DC=lcl'
$0DayVaultGroup = 'CN=0 day vault group,OU=Groups,DC=contoso,DC=lcl'
[string]$filedate = "$((get-date -UFormat %H-%M-%S_%m-%d-%y).ToString())"
$filedate = "$filedate"
$GroupLoop = $true
$ADUserObj = $null
$ContinueAfterBackupCheck = "Y"
$DisabledUsersOU = 'OU=Disabled Accounts,OU=Users,DC=contoso,DC=lcl'

#endregion setup

#region main
If ($Testmode -eq $true) {
    Write-Host -ForegroundColor Green "*** SCRIPT RUNNING IN TEST MODE - NO CHANGES WILL BE MADE ***`n"
    $ExtraFlags = @{whatif = $true}
} else {
    Write-Host -ForegroundColor Red "*** SCRIPT RUNNING IN LIVE MODE - CHANGES WILL BE MADE ***`n"
    $ExtraFlags = @{whatif = $false}
}

While (-not $ADUserObj) {
    
    $ADUser = Read-Host "$(Write-Host -NoNewLine -ForegroundColor yellow 'Enter the ADUser that is leaving or needs to be recovered from backup') "

    Try {
        $ADUserObj = Get-ADUser $ADUser -Properties *
    }

    Catch {
        Write-Warning "User not found or no user entered. Please enter a valid AD account name to continue."
    }

    If ($ADUserObj -ne $null) {
        $ADUserGroups = Get-aduser $ADUserObj -properties memberof  -ErrorAction SilentlyContinue | % memberof -ErrorAction SilentlyContinue
        Write-Host -ForegroundColor yellow "User Info:`n"
        Write-Host -NoNewLine -ForegroundColor Yellow "Name: "
        Write-Host -ForegroundColor cyan "$($ADUserObj.Name)"
        Write-Host -NoNewLine -ForegroundColor yellow "ADUser: "
        Write-Host -ForegroundColor cyan "$($ADUserObj.SamAccountName)"
        Write-Host -NoNewLine -ForegroundColor yellow "DN: "
        Write-Host -ForegroundColor cyan "$($ADUserObj.DistinguishedName)"
        Write-Host -NoNewline -ForegroundColor Yellow "Home Server: "
        Write-Host -ForegroundColor Cyan "$((($ADUserObj.HomeDirectory).Split('\'))[2])"
        Write-Host -NoNewLine -ForegroundColor yellow "Office: "
        Write-Host -ForegroundColor cyan "$($ADUserObj.Office)"
        Write-Host -NoNewLine -ForegroundColor yellow "E-Mail: "
        Write-Host -ForegroundColor cyan "$($ADUserObj.emailaddress)"
        Write-Host -NoNewline -ForegroundColor Yellow "Manager: "
        Write-Host -ForegroundColor Cyan "$($ADUserObj.Manager | Get-ADUser | % name)`n`n"

        While ($GroupLoop) {
            $KeepGroups = $KeepGroups -split '[\r\n]' |? {$_}
            Try {
                $RemoveUserGroups = $ADUserGroups | ? {$KeepGroups -notcontains $_}
                }
            Catch {
                }
            Write-Host -ForegroundColor Yellow "All groups specified in `$KeepGroups will be kept. The following groups will be removed:`n"
            Try {
                $RemoveUserGroups | get-adgroup -EA SilentlyContinue | % {Write-Host -ForegroundColor Cyan $_.name}| sort -Unique
                }
            Catch {
            }
            $RemoveUserGroups = $RemoveUserGroups | sort -Unique
            Write-Host -ForegroundColor yellow "`n0 Day vault group will be added to account`n"
            Write-Host -ForegroundColor yellow "Account will be moved to:"
            Write-Host -ForegroundColor cyan "$($DisabledUsersOU)`n"
            Write-Host -NoNewLine -ForegroundColor yellow "Expiration will be set to "
            Write-Host -ForegroundColor cyan "$(((get-date).AddDays(10)).ToShortDateString())`n"
            Write-Host -ForegroundColor Yellow "Enter k to select additional existing groups to keep, r to recover a previously processed user or y to continue`n"
            $Continue = Read-Host "$(Write-Host -NoNewLine -ForegroundColor red 'Continue? (y/N/k/r)') "
            $ADObjectBackupPath = "$($pwd.path)\$(($ADUserObj.SamAccountName)).ADObject.$filedate.clixml"

            If ($Continue.ToUpper() -eq "Y") {
                $GroupLoop = $false
                Write-Host -NoNewLine -ForegroundColor Yellow "Attempting to backup user object to " 
                Write-Host -ForegroundColor Cyan $ADObjectBackupPath
                $ADUserBackupObj = $ADUserObj.DistinguishedName,$ADUserObj.memberof
                Try {
                    Export-Clixml -InputObject $ADUserObj -Path $ADObjectBackupPath
                    }
                Catch {
                    Write-Warning "Something went wrong and a backup was not saved to $ADObjectBackupPath."
                    $ContinueAfterBackupCheck = Read-Host "$(Write-Host -NoNewLine -ForegroundColor red 'Continue? (y/N)') "
                    }
                Switch ($ContinueAfterBackupCheck.ToUpper()) {
                    "Y" {
                        Start-Sleep 3
                        $RemoveUserGroups | % {
                            Write-Host -NoNewLine -ForegroundColor Yellow "Removing group "
                            Write-Host -ForegroundColor Cyan "$($_ | get-adgroup | % name)"
                            Remove-ADGroupMember -Confirm:$false $_ $ADUser @ExtraFlags
                        }
                    Write-Host -NoNewLine -ForegroundColor Yellow "Adding Group "
                    Write-Host -ForegroundColor Cyan "$(get-adgroup $0DayVaultGroup | % name)"
                    Add-ADGroupMember -Confirm:$false $0DayVaultGroup $ADUser @ExtraFlags
                    Write-Host -NoNewLine -ForegroundColor Yellow "Setting account expiration date to "
                    Write-Host -ForegroundColor Cyan "$(((get-date).AddDays(10)).ToShortDateString())"
                    Set-ADUser $ADUserObj -Confirm:$false -Replace @{info="Expires: $(((get-date).AddDays(10)).ToShortDateString())"} @ExtraFlags
                    Write-Host -ForegroundColor Yellow "Disabling Account"
                    Disable-ADAccount -Confirm:$false $ADUserObj @ExtraFlags
                    Write-Host -NoNewLine -ForegroundColor Yellow "Moving User to "
                    Write-Host -ForegroundColor Cyan "$DisabledUsersOU"
                    Move-ADObject $ADUserObj.DistinguishedName $DisabledUsersOU @ExtraFlags
                    Get-ADUser $ADUserObj.SamAccountName -Properties memberof,info
                    }
                    "default" {
                        Write-Host -ForegroundColor Yellow "Aborting..."
                    }
                }
            } elseif ($Continue.ToUpper() -eq "K") {
                If ($RemoveUserGroups) {
                    Write-Host -ForegroundColor Yellow "Select additional groups to keep from open window"
                    $KeepGroupsAdditional = Compare-Object  $KeepGroups $ADUserGroups -PassThru | get-adgroup | % name | Out-GridView -PassThru | Get-ADGroup |`
                    % DistinguishedName
                    $Keepgroups = $KeepGroups + $KeepGroupsAdditional
                } else {
                    Write-Warning "No current groups to select from"
                }
            } elseif ($Continue.ToUpper() -eq "R") {
                    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null
                    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
                    $OpenFileDialog.initialDirectory = "$($pwd.path)\"
                    $OpenFileDialog.filter = "CliXML Files | *ADObject*.clixml"
                    $OpenFileDialog.Title = "Select a backed up user to continue."
                    $OpenFileDialog.ShowDialog() | Out-Null
                    Write-Host -ForegroundColor Yellow "Loading selected ADObject backup..."
                    Try {
                        $BackupObject = Import-Clixml -Path $OpenFileDialog.filename
                    }
                    Catch {
                        Write-Error 'Unable to load ADObject from backup file...Exiting'
                    }
                    If ($BackupObject.DistinguishedName -eq $null) {
                        $GroupLoop = $false
                        return
                    } else {
                        $BackupDN = $BackupObject.DistinguishedName
                        $BackupDNSplit = ($BackupObject.DistinguishedName.Replace('OU=',';')).Split(';')
                        # The following line may need editing to make it work
                        $BackupDNOUOnly = "OU=" + ((-join $BackupDNSplit[1..($BackupDNSplit.GetUpperBound(0)-1)]).Replace(',',',OU=')).`
                            Replace('R3,OU=',"R3,OU=$($BackupDNSplit[($BackupDNSplit.GetUpperBound(0))])")
                        If ($BackupObject.MemberOf -ne $null) {
                            $BackupObject.MemberOf | % {
                                Write-Host -NoNewLine -ForegroundColor Yellow "Restoring group "
                                Write-Host -ForegroundColor Cyan "$($_ | get-adgroup | % name)"
                                Add-ADGroupMember -Confirm:$false $_ $ADUser @ExtraFlags
                            }
                        } else {
                            Write-Host -ForegroundColor Yellow "No groups found to re-add"
                        }
                        Write-Host -ForegroundColor Yellow "Removing expiration and re-enabling account"
                        Set-ADUser $ADUserObj -Remove @{info=$ADUserObj.info} @ExtraFlags
                        Enable-ADAccount $ADUserObj @ExtraFlags
                        Write-Host -NoNewLine -ForegroundColor Yellow "Attempting to move ADObject back to "
                        Write-Host -Foregroundcolor Cyan "$BackupDNOUOnly"
                        $ADUserObj| Move-ADObject -TargetPath $BackupDNOUOnly @ExtraFlags
                        Get-ADUser $ADUser -Properties memberof,info
                        Write-Host -ForegroundColor Yellow "Done."
                        $GroupLoop = $false
                    }
            } else {
                $GroupLoop = $false
                return
            }   
        }
    }
}

#endregion main