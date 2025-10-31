# SharedMailboxAutomation.ps1
# This script automates the creation of shared mailboxes based on ServiceNow tasks.
# Replace all placeholders (e.g., [YOUR_INSTANCE], [YOUR_PASSWORD]) with your actual values before running.

# Secure Credentials Setup
$snUsername = "[YOUR_SERVICENOW_USERNAME]"
$snPassword = ConvertTo-SecureString "[YOUR_SERVICENOW_PASSWORD]" -AsPlainText -Force
$snCreds = New-Object System.Management.Automation.PSCredential ($snUsername, $snPassword)

$name = Get-Date -Format "dd-MM-yyyy"
$nameApi = "$name-API"

$adUsername = Get-Content "[PATH_TO_AD_USERNAME_FILE]"  # e.g., path to file with AD username
$adPassword = Get-Content "[PATH_TO_AD_PASSWORD_FILE]" | ConvertTo-SecureString -AsPlainText -Force  # e.g., path to file with AD password
$adCreds = New-Object System.Management.Automation.PSCredential ($adUsername, $adPassword)

$exUsername = Get-Content "[PATH_TO_EXCHANGE_USERNAME_FILE]"  # e.g., path to file with Exchange username
$exCred = New-Object System.Management.Automation.PSCredential ($exUsername, $adPassword)

$currentDC = "[YOUR_DOMAIN_CONTROLLER]"  # e.g., your domain controller hostname

Import-Module ActiveDirectory

function Get-TaskDetails {
    <#
    .SYNOPSIS
    Retrieves task details from ServiceNow and exports to CSV.
    #>
    $uri = "https://[YOUR_INSTANCE].service-now.com/api/now/table/sc_task?sysparm_query=assignment_group.name=[YOUR_TEAM]^active=true^stateIN1,2^short_description=[YOUR_TASK_NAME]&sysparm_fields=number"
    $response = Invoke-RestMethod -Uri $uri -Method Get -Credential $snCreds -Headers @{ "Accept" = "application/json" }
    $tasks = $response.result

    foreach ($task in $tasks) {
        $scTask = $task.number
        $uri = "https://[YOUR_INSTANCE].service-now.com/api/v1/sctask_inbound?sysparm_query=$scTask"
        $response = Invoke-WebRequest -Uri $uri -Method Get -Credential $snCreds
        $data = $response.Content | ConvertFrom-Json
        $taskDetails = $data.result

        $userSysIds = $taskDetails.variables.owner_s_of_mailbox -split ","
        $userSams = @()
        $userEmails = @()

        foreach ($sysId in $userSysIds) {
            $userUri = "https://[YOUR_INSTANCE].service-now.com/api/now/table/sys_user?sysparm_query=sys_id=$sysId&sysparm_fields=email"
            $userResponse = Invoke-RestMethod -Uri $userUri -Method Get -Credential $snCreds -Headers @{ Accept = "application/json" }
            $userEmail = $userResponse.result[0].email

            $userSam = Get-ADUser -Server $currentDC -Credential $adCreds -Filter { EmailAddress -eq $userEmail } | Select-Object -ExpandProperty SamAccountName
            $userEmai = Get-ADUser -Server $currentDC -Credential $adCreds -Properties EmailAddress -Filter { EmailAddress -eq $userEmail } | Select-Object -ExpandProperty EmailAddress

            $userEmails += $userEmai
            $userSams += $userSam
        }

        $mailboxPassword = "[YOUR_DEFAULT_MAILBOX_PASSWORD]"  # e.g., default password for mailboxes
        $genGroupName = $taskDetails.variables.name_of_mailbox
        $cleanName = $genGroupName -replace '[^a-zA-Z0-9]', ''
        $primOwner = ($userSams -split ";")[0]

        $doma = switch ($taskDetails.variables.email_domain) {
            "[YOUR_DOMAIN_1]" { "[DOMAIN_LABEL_1]" }  # e.g., "example.co.uk" -> "Location1"
            "[YOUR_DOMAIN_2]" { "[DOMAIN_LABEL_2]" }  # e.g., "example.com" -> "Location2"
            "[YOUR_DOMAIN_3]" { "[DOMAIN_LABEL_3]" }  # e.g., "example.com.hk" -> "Location3"
            default { "Unknown" }
        }

        $premTemp = if ($taskDetails.variables.please_select_if_this_is_for_a_permanent_or_temporary_mailbox -eq "permanent") { "Perm" } else { "Temp" }

        [PSCustomObject]@{
            SCTASK_Number     = $taskDetails.number
            Short_Description = $taskDetails.short_description
            State             = $taskDetails.state
            Assigned_To       = $taskDetails.assigned_to_display_value
            BeforeDot         = $taskDetails.variables.name_of_mailbox
            Domain            = $doma
            Group             = "_GEN" + $cleanName.ToUpper() + "01G"
            Period            = $premTemp
            Expiry            = $taskDetails.variables.if_mailbox_is_temporary_provide_a_closure_date
            Owner             = $primOwner
            Password          = $mailboxPassword
            Department        = $taskDetails.variables.department_responsible_for_mailbox
            SOwners           = $userSams -join ";"
            Mail              = $userEmails -join ";"
        } | Export-Csv "[PATH_TO_EXPORT_CSV]\$nameApi.csv" -Append -NoTypeInformation  # e.g., path to output CSV
    }
}

function Validate-Requests {
    <#
    .SYNOPSIS
    Validates the exported CSV data for mailbox creation.
    #>
    $errorList = @()
    $successList = @()
    $csvPath = "[PATH_TO_EXPORT_CSV]\$nameApi.csv"

    if (Test-Path $csvPath) {
        $csv = Import-Csv $csvPath
        foreach ($m in $csv) {
            $rowErrors = @()
            $mailbox = $m.BeforeDot
            $accessGroup = $m.Group
            $location = $m.Domain
            $permTemp = $m.Period
            $expiry = $m.Expiry
            $owner = $m.Owner
            $password = $m.Password
            $department = $m.Department
            $sOwner = $m.SOwners
            $mail = $m.Mail
            $scTaskNumber = $m.SCTASK_Number

            if ([string]::IsNullOrWhiteSpace($mailbox) -or [string]::IsNullOrWhiteSpace($accessGroup)) {
                $rowErrors += "Missing mailbox or group information"
            } else {
                # Check if mailbox exists
                $adCheck = Get-ADUser -Server $currentDC -Credential $adCreds $mailbox -ErrorAction SilentlyContinue
                if ($adCheck) { $rowErrors += "$mailbox already exists in AD" }

                # Check mailbox length
                if ($mailbox.Length -gt 20) { $rowErrors += "$mailbox exceeds 20 characters" }

                # Check for invalid characters
                if ($mailbox -notmatch '^[a-zA-Z0-9._]+$') {
                    $invalidChars = ($mailbox -split '') | Where-Object { $_ -notmatch '[a-zA-Z0-9._]' } | Select-Object -Unique
                    $rowErrors += "Invalid characters found: $($invalidChars -join ',')"
                }

                # Check if group exists
                $groupCheck = Get-ADGroup -Server $currentDC -Credential $adCreds $accessGroup -ErrorAction SilentlyContinue
                if ($groupCheck) { $rowErrors += "Group $accessGroup already exists in AD" }
            }

            if ($rowErrors.Count -gt 0) {
                $errorList += [PSCustomObject]@{
                    Mailbox       = $mailbox
                    SCTASK_Number = $scTaskNumber
                    Mail          = $mail
                    ErrorDetail   = ($rowErrors -join "; ")
                }
            } else {
                $successList += [PSCustomObject]@{
                    BeforeDot     = $mailbox
                    Location      = $location
                    Group         = $accessGroup
                    'Perm/Temp'   = $permTemp
                    Expiry        = $expiry
                    Owner         = $owner
                    Password      = $password
                    Department    = $department
                    SOwner        = $sOwner
                    Mail          = $mail
                    SCTASK_Number = $scTaskNumber
                    Status        = "Validation Passed"
                }
            }
        }

        $errorList | Export-Csv -Path "[PATH_TO_ERROR_CSV]\$name-ValidationErrors.csv" -NoTypeInformation -Force
        $successList | Export-Csv -Path "[PATH_TO_SUCCESS_CSV]\$name-ValidEntries.csv" -NoTypeInformation -Force
    }
}

function Send-FailureEmails {
    <#
    .SYNOPSIS
    Sends emails for failed validations and updates ServiceNow.
    #>
    $errorCsvPath = "[PATH_TO_ERROR_CSV]\$name-ValidationErrors.csv"
    $errors = Import-Csv $errorCsvPath

    foreach ($error in $errors) {
        $from = "[YOUR_FROM_EMAIL]"  # e.g., your team's email address
        $to = $error.Mail -split ";"
        $scTaskNumber = $error.SCTASK_Number
        $server = "[YOUR_SMTP_SERVER]"

        $subject = "$scTaskNumber - Create a Shared Mailbox: Failed"
        $body = @"
Hello,

This is an automated process. The request for mailbox creation has failed. Please find the error(s) below:

Mailbox: $($error.Mailbox)
Error: $($error.ErrorDetail)

Kindly resubmit the form via the request link: [YOUR_FORM_LINK]

Thanks & Regards,
[YOUR_TEAM_NAME]
"@

        Send-MailMessage -From $from -To $to -Subject $subject -Body $body -SmtpServer $server -Bcc $from

        $headers = @{
            "Accept"       = "application/json"
            "Content-Type" = "application/json"
        }
        $wrkNotes = @{
            u_work_notes   = $body
            u_sctask_number = $scTaskNumber
            u_status       = "3"  # Closed status
        } | ConvertTo-Json

        $response = Invoke-RestMethod -Uri "[YOUR_SERVICENOW_TASK_URI]" -Method Get -Credential $snCreds -Headers $headers
        $sysId = $response.result[0].sys_id
        $updateUri = "https://[YOUR_INSTANCE].service-now.com/api/v2/sctask_inbound"
        Invoke-RestMethod -Uri $updateUri -Method Post -Credential $snCreds -Headers $headers -Body $wrkNotes
    }
}

function Connect-ToExchange {
    <#
    .SYNOPSIS
    Connects to Exchange server.
    #>
    $exSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "[YOUR_EXCHANGE_URI]" -Authentication Kerberos -Credential $adCreds

    if ($null -eq $exSession) {
        Write-Host "Failed to connect to Exchange" -ForegroundColor Red
        Remove-PSSession $exSession
        Start-Sleep 10
        exit
    }

    Import-PSSession $exSession -AllowClobber
}

function Create-Mailboxes {
    <#
    .SYNOPSIS
    Creates AD users and enables remote mailboxes.
    #>
    $import = Import-Csv -Path "[PATH_TO_SUCCESS_CSV]\$name-ValidEntries.csv"

    foreach ($m in $import) {
        $firstName = $m.BeforeDot -replace '\s', ''
        $fullDot = $firstName
        $fullName = $firstName
        $location = $m.Location
        $groupName = $m.Group -replace '\s', ''
        $title = "Generic Mailbox - $fullName"
        $department = $m.Department
        $company = "[YOUR_COMPANY_NAME]"  # e.g., your company name
        $expiry = $m.Expiry
        $password = ConvertTo-SecureString $m.Password -AsPlainText -Force -ErrorAction SilentlyContinue
        $owner = $m.Owner
        $permTemp = $m.'Perm/Temp'
        $genOu = $m.Location
        $smtpDomain = "[YOUR_SMTP_DOMAIN]"  # e.g., "@example.com"
        $expiryCheck = Get-Date $expiry -ErrorAction SilentlyContinue

        Write-Host "Creating mailbox: $fullDot" -ForegroundColor Cyan

        New-ADUser -Name $fullName `
                   -DisplayName $fullName `
                   -GivenName $firstName `
                   -Enabled $true `
                   -Department $department `
                   -Title $title `
                   -Company $company `
                   -SamAccountName $fullDot `
                   -Description "Generic Mailbox - $fullName" `
                   -Path "[YOUR_MAILBOX_OU]" `
                   -UserPrincipalName ($fullDot + $smtpDomain) `
                   -EmailAddress ($fullDot + $smtpDomain) `
                   -PasswordNotRequired $true `
                   -ChangePasswordAtLogon $false `
                   -Credential $adCreds `
                   -Server $currentDC

        if ($expiryCheck) {
            Set-ADUser $fullDot -AccountExpirationDate (Get-Date $expiry).AddDays(1) -Server $currentDC -Credential $adCreds
        }

        Set-ADUser -Identity $fullDot -Add @{ info = "Owner: $owner, Group: $groupName"; extensionAttribute2 = "Mailbox"; extensionAttribute14 = "GEN" } -Server $currentDC -Credential $adCreds
        Set-ADUser -Identity $fullDot -Add @{ MSExchVersion = "" } -Server $currentDC -Credential $adCreds

        Get-ADUser $fullDot | Move-ADObject -TargetPath "[YOUR_TARGET_OU]" -Server $currentDC -Credential $adCreds

        Write-Host "Enabling remote mailbox: $fullDot"
        Enable-RemoteMailbox -Identity $fullDot -RemoteRoutingAddress ($fullDot + "[YOUR_ONMICROSOFT_DOMAIN]") -PrimarySmtpAddress ($fullDot + $smtpDomain) -DomainController $currentDC | Out-Null
        Enable-RemoteMailbox -Identity ($fullDot + $smtpDomain) -Archive -DomainController $currentDC | Out-Null

        Set-ADUser -Identity $fullDot -Server $currentDC -Credential $adCreds -Replace @{
            "ExtensionAttribute3" = "FT"
            "ExtensionAttribute4" = $permTemp
            "ExtensionAttribute5" = $location
            "ExtensionAttribute6" = $fullDot
        }

        Set-ADUser -Identity $fullDot -Server $currentDC -Add @{ "proxyAddresses" = "smtp:" + $fullDot + $smtpDomain } -Credential $adCreds
        Set-ADUser -Identity $fullDot -Server $currentDC -Add @{ "proxyAddresses" = "smtp:" + $fullDot + "[YOUR_ONMICROSOFT_DOMAIN]" } -Credential $adCreds

        Write-Host "Mailbox $fullDot created" -ForegroundColor Green
    }
}

function Create-Groups {
    <#
    .SYNOPSIS
    Creates AD security groups for mailbox access.
    #>
    $import = Import-Csv -Path "[PATH_TO_SUCCESS_CSV]\$name-ValidEntries.csv"

    foreach ($m in $import) {
        $firstName = $m.BeforeDot -replace '\s', ''
        $fullDot = $firstName
        $location = $m.Location
        $groupName = $m.Group -replace '\s', ''
        $sOwner = $m.SOwner
        $department = $m.Department
        $company = "[YOUR_COMPANY_NAME]"
        $owner = $m.Owner

        New-ADGroup -Name $groupName `
                    -DisplayName $groupName `
                    -SamAccountName $groupName `
                    -GroupCategory Security `
                    -GroupScope Universal `
                    -Description "Owner: $sOwner" `
                    -ManagedBy $owner `
                    -Server $currentDC `
                    -Path "[YOUR_GROUP_OU]" `
                    -Credential $adCreds

        Set-ADGroup -Identity $groupName -Add @{ info = "Mailbox: $fullDot Dept: $department" } -Server $currentDC -Credential $adCreds
        Add-ADGroupMember -Identity $groupName -Members $owner -Credential $adCreds -Server $currentDC
        Enable-DistributionGroup -Identity $groupName -DomainController $currentDC

        Write-Host "Group $groupName created" -ForegroundColor Green
    }

    Start-Sleep 5  # Pause for synchronization
}

function Set-GroupSecurity {
    <#
    .SYNOPSIS
    Sets ACLs for group management.
    #>
    $csvPath = "[PATH_TO_SUCCESS_CSV]\$name-ValidEntries.csv"
    $logPath = "[PATH_TO_LOG]\log-$name.txt"

    $import = Import-Csv -Path $csvPath

    foreach ($m in $import) {
        $gName = $m.Group
        $sOwners = $m.SOwner -split ";"

        foreach ($u in $sOwners) {
            try {
                Add-ADGroupMember -Identity $gName -Members $u -Credential $adCreds -Server $currentDC

                $uName = Get-ADUser $u -Credential $adCreds -Server $currentDC
                $sid = New-Object System.Security.Principal.SecurityIdentifier($uName.SID)
                $group = Get-ADGroup -Identity $gName -Credential $adCreds -Server $currentDC
                $dn = $group.DistinguishedName

                $de = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$dn", $adCreds.UserName, $adCreds.GetNetworkCredential().Password)
                $acl = $de.ObjectSecurity
                $ace = New-Object System.DirectoryServices.ActiveDirectoryAccessRule($sid, "ListChildren, ReadProperty, GenericWrite", "Allow")
                $acl.AddAccessRule($ace)
                $de.ObjectSecurity = $acl
                $de.CommitChanges()

                Add-Content $logPath "ACL updated for $u on $gName at $(Get-Date)"
            } catch {
                Add-Content $logPath "ACL ERROR for $u on $gName : $_ at $(Get-Date)"
            }
        }
    }
}

function Grant-ExchangeAccess {
    <#
    .SYNOPSIS
    Grants permissions in Exchange Online.
    .NOTES
    Requires Exchange Online PowerShell module.
    #>
    $import = Import-Csv "[PATH_TO_SUCCESS_CSV]\$name-ValidEntries.csv"
    Connect-ExchangeOnline -Credential $exCred -ErrorAction Stop

    foreach ($m in $import) {
        $firstName = $m.BeforeDot
        $groupName = $m.Group

        Set-Mailbox $firstName -Type Shared
        Add-RecipientPermission $firstName -AccessRights SendAs -Trustee $groupName -Confirm:$false
        Add-MailboxPermission $firstName -User $groupName -AccessRights FullAccess -Confirm:$false

        Write-Host "Cloud Mailbox Access applied for $firstName & $groupName" -ForegroundColor Cyan
    }
}

function Send-CompletionEmails {
    <#
    .SYNOPSIS
    Sends completion emails and updates ServiceNow.
    #>
    $import = Import-Csv -Path "[PATH_TO_SUCCESS_CSV]\$name-ValidEntries.csv"

    foreach ($m in $import) {
        $firstName = $m.BeforeDot -replace '\s', ''
        $fullDot = $firstName
        $fullName = $firstName
        $location = $m.Location
        $sOwner = $m.SOwner
        $groupName = $m.Group -replace '\s', ''
        $title = "Generic Mailbox - $fullName"
        $department = $m.Department
        $company = "[YOUR_COMPANY_NAME]"
        $owner = $m.Owner
        $permTemp = $m.'Perm/Temp'
        $scTaskNumber = $m.SCTASK_Number
        $value = $scTaskNumber

        $subject = "$fullName New Mailbox - *$fullDot* - *$value*"
        $to = $m.Mail -split ";"
        $from = "[YOUR_FROM_EMAIL]"
        $bcc = "[YOUR_BCC_EMAIL]"
        $server = "[YOUR_SMTP_SERVER]"

        $body = @"
<font face="Calibri">Hello,

This is an automated process. Your request for a new generic mailbox has been completed. Details below:

Mailbox: $fullDot
Owner(s): $sOwner
Access Group: $groupName

Users should access the mailbox from their own domain account and Outlook client. Refer to the attached guides for adding the shared mailbox to Outlook and managing group membership. Changes may take up to 2 hours to propagate.

For support, contact [YOUR_SUPPORT_CONTACT].

Thanks & Regards,
[YOUR_TEAM_NAME]
$company
</font>
"@

        Send-MailMessage -From $from `
                         -To $to `
                         -Bcc $bcc `
                         -SmtpServer $server `
                         -Subject $subject `
                         -BodyAsHtml:$true `
                         -Body $body `
                         -Attachments "[PATH_TO_ATTACHMENT_1]", "[PATH_TO_ATTACHMENT_2]"  # e.g., paths to documentation

        $headers = @{
            "Accept"       = "application/json"
            "Content-Type" = "application/json"
        }
        $wrkNotes = @{
            u_work_notes   = "Automated Request Completed"
            u_sctask_number = $value
            u_status       = "3"  # Closed status
        } | ConvertTo-Json

        $updateUri = "https://[YOUR_INSTANCE].service-now.com/api/v2/sctask_inbound"
        Invoke-RestMethod -Uri $updateUri -Method Post -Credential $snCreds -Headers $headers -Body $wrkNotes
    }
}

# Main Execution
$ErrorActionPreference = "SilentlyContinue"
Get-TaskDetails
Validate-Requests
Send-FailureEmails
Connect-ToExchange
Create-Mailboxes
Create-Groups
Set-GroupSecurity

Write-Host "Waiting for directory synchronization..."
for ($time = 300; $time -gt 0; $time--) {
    Write-Progress -Activity "Waiting for DirSync" -SecondsRemaining $time -CurrentOperation "Synchronizing" -Status "Counting Down"
    Start-Sleep -Seconds 1
}

Grant-ExchangeAccess
Send-CompletionEmails

Get-PSSession | Remove-PSSession

Write-Host "Script Completed" -ForegroundColor Green
