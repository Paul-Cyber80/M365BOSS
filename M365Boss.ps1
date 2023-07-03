# Clear the screen
Clear-Host

# Prompt for case folder name
$caseFolderName = Read-Host "Create a case folder name"

# Create the case folder if it doesn't exist
$caseFolderPath = Join-Path -Path (Get-Location) -ChildPath $caseFolderName
if (-not (Test-Path $caseFolderPath)) {
    New-Item -ItemType Directory -Path $caseFolderPath | Out-Null
}

$continue = $true

while ($continue) {
    # Clear the screen
    Clear-Host

    # Print menu
	Write-Host
	Write-Host
    Write-Host "***********************************"
    Write-Host "*            M365 BOSS            *"
	Write-Host "*    presented by Cyber80.com     *"
    Write-Host "***********************************"
    Write-Host "* 1. Read Me                      *"
    Write-Host "* 2. Install Modules/Authenticate *"
    Write-Host "* 3. UAL                          *"
    Write-Host "* 4. Admin Audit Log Config       *"
    Write-Host "* 5. Forwarding Rules             *"
    Write-Host "* 6. Email Rules                  *"
    Write-Host "* 7. Exit                         *"
	Write-Host "***********************************"
    Write-Host

    # Prompt for menu choice
    $menuChoice = Read-Host "Enter your choice (1-7):"

    switch ($menuChoice) {
        "1" {
            # Option 1: Read Me
            Clear-Host
            Write-Host "Thank you for using M365 BOSS."  
			Write-Host
			Write-Host "This program was created and tested"
			Write-Host "with Windows 10 and PowerShell version 5.1.19041.3031."  
			Write-Host
			Write-Host "Please make sure to run step #2 and install all the required modules"
			Write-Host "and authenticate with a Global Admin before acquiring any data."
			Write-Host
			Write-Host "Please visit Cyber80.com for updates and other free software."
			Write-Host "- Paul Adie/Cyber80"
            Write-Host
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")
        }
        "2" {
            # Option 2: Install Modules
            Clear-Host
            Write-Host "Installing modules..."
            Write-Host

            # Install MSOnline module
            Write-Host "Installing MSOnline module..."
            Install-Module -Name MSOnline -Force -AllowClobber -Scope CurrentUser -Confirm:$false

            # Install ExchangeOnlineManagement module
            Write-Host "Installing ExchangeOnlineManagement module..."
            Install-Module -Name ExchangeOnlineManagement -Force -AllowClobber -Scope CurrentUser -Confirm:$false

            # Import ExchangeOnlineManagement module
            Write-Host "Importing ExchangeOnlineManagement module..."
            Write-Host
            Import-Module ExchangeOnlineManagement -Force
            Write-Host
            Write-Host

            # Prompt for Global Admin email address
            $globalAdminEmail = Read-Host "Enter Global Admin email address"

            do {
                try {
                    # Connect to Exchange Online
                    Write-Host "Connecting to Exchange Online..."
                    Connect-ExchangeOnline -UserPrincipalName $globalAdminEmail -ShowProgress $true -ErrorAction Stop

                    # Confirm successful authentication
                    Write-Host
                    Write-Host "Authentication successful."
                    Write-Host
                    Write-Host "Press any key to continue..."
                    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")

                    $retry = $false
                }
                catch {
                    Write-Host
                    Write-Host "Error occurred during authentication. Please try again."
                    Write-Host
                    $retry = $true
                }
            } while ($retry)
        }
        "3" {
            # Option 3: Start UAL Acquisition
            # Clear the screen
            Clear-Host

            # Print tool information
            Write-Host "UAL Acquisition"
            Write-Host

            # Prompt for the start date
            $startDate = Read-Host "Enter the start date (yyyy-MM-dd):"

            # Prompt for the end date
            $endDate = Read-Host "Enter the end date (yyyy-MM-dd):"

            # Prompt for the user ID
            $userId = Read-Host "Enter the user email address (e.g., user@example.com):"

            # Prompt for the output file
            $outputFile = Read-Host "Enter the output file name (e.g., Name_UAL.csv):"

            # Create UAL folder inside the case folder if it doesn't exist
            $ualFolder = Join-Path -Path $caseFolderName -ChildPath "UAL"
            $ualFolderPath = Join-Path -Path $caseFolderPath -ChildPath "UAL"
            if (-not (Test-Path $ualFolderPath)) {
                New-Item -ItemType Directory -Path $ualFolderPath | Out-Null
            }

            # Loop through each day within the specified period
            $currentDate = Get-Date $startDate
            $endDateTime = Get-Date $endDate
            while ($currentDate -le $endDateTime) {
                $formattedStartDate = Get-Date $currentDate -Format "yyyy-MM-dd"
                $formattedEndDate = Get-Date $currentDate -Format "yyyy-MM-ddT23:59:59"

                # Retrieve the audit log for the user and append it to the output file
                $auditLog = Search-UnifiedAuditLog -StartDate $formattedStartDate -EndDate $formattedEndDate -UserIds $userId -ResultSize 5000 |
                    Select-Object PSComputerName, RunspaceID, PSShowComputerName, RecordType, CreationDate, UserIds, Operations, AuditData, ResultIndex, ResultCount, Identity, IsValid, ObjectState
                $auditLog | Export-Csv -Path (Join-Path -Path $ualFolderPath -ChildPath $outputFile) -Append -NoTypeInformation

                # Move to the next day
                $currentDate = $currentDate.AddDays(1)
            }

            # Calculate and save the SHA-256 hash of the output file
            $hashOutputFile = Get-FileHash -Path (Join-Path -Path $ualFolderPath -ChildPath $outputFile) -Algorithm SHA256
            $hashOutputFile | Out-File -FilePath (Join-Path -Path $ualFolderPath -ChildPath "Hash_$outputFile.txt")

            # Prompt to press any key to continue
            Write-Host "`nPress any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")
        }
        "4" {
            # Option 4: Admin Audit Log Config
            $auditLogConfigFolder = Join-Path -Path $caseFolderPath -ChildPath "Admin Audit Log Config"
            if (-not (Test-Path $auditLogConfigFolder)) {
                New-Item -ItemType Directory -Path $auditLogConfigFolder | Out-Null
            }

            # Retrieve admin audit log config and save it to file
            $adminAuditLogConfig = Get-AdminAuditLogConfig
            $adminAuditLogConfig | Out-File -FilePath (Join-Path -Path $auditLogConfigFolder -ChildPath "AdminAuditLogConfig.txt")

            # Calculate and save the SHA-256 hash of the admin audit log config file
            $hashAdminAuditLogConfig = Get-FileHash -Path (Join-Path -Path $auditLogConfigFolder -ChildPath "AdminAuditLogConfig.txt") -Algorithm SHA256
            $hashAdminAuditLogConfig | Out-File -FilePath (Join-Path -Path $auditLogConfigFolder -ChildPath "Hash_AdminAuditLogConfig.txt")

            # Prompt to press any key to continue
            Write-Host "`nPress any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")
        }
        "5" {
            # Option 5: Forwarding Rules
            $forwardingRulesFolder = Join-Path -Path $caseFolderPath -ChildPath "Forwarding Rules"
            if (-not (Test-Path $forwardingRulesFolder)) {
                New-Item -ItemType Directory -Path $forwardingRulesFolder | Out-Null
            }

            # Retrieve forwarding rules and save them to file
            $forwardingRules = Get-Mailbox | Select-Object UserPrincipalName, ForwardingSmtpAddress, DeliverToMailboxAndForward
            $forwardingRules | Out-File -FilePath (Join-Path -Path $forwardingRulesFolder -ChildPath "Forwarding Rules.txt") 

            # Calculate and save the SHA-256 hash of the forwarding rules file
            $hashForwardingRules = Get-FileHash -Path (Join-Path -Path $forwardingRulesFolder -ChildPath "Forwarding Rules.txt") -Algorithm SHA256
            $hashForwardingRules | Out-File -FilePath (Join-Path -Path $forwardingRulesFolder -ChildPath "Hash_Forwarding Rules.txt")

            # Prompt to press any key to continue
            Write-Host "`nPress any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")
        }
        "6" {
            # Option 6: Email Rules
            # Prompt for the email address to check
            $clientEmailAddress = Read-Host "Enter the email address to check for email rules:"

            # Create the Email Rules folder if it doesn't exist
            $emailRulesFolder = Join-Path -Path $caseFolderPath -ChildPath "Email Rules"
            if (-not (Test-Path $emailRulesFolder)) {
                New-Item -ItemType Directory -Path $emailRulesFolder | Out-Null
            }

            # Run the command to retrieve email rules
            $emailRules = Get-InboxRule -Mailbox $clientEmailAddress | Format-List Name, Description, ForwardTo, MarkAsRead, MoveToFolder, StopProcessingRules

            # Save the email rules to a file
            $emailRulesFile = Join-Path -Path $emailRulesFolder -ChildPath "${clientEmailAddress}_EmailRules.txt"
            $emailRules | Out-File -FilePath $emailRulesFile

            # Calculate and save the SHA-256 hash of the email rules file
            $hashEmailRules = Get-FileHash -Path $emailRulesFile -Algorithm SHA256
            $hashEmailRules | Out-File -FilePath (Join-Path -Path $emailRulesFolder -ChildPath "Hash_${clientEmailAddress}_EmailRules.txt")

            # Prompt to press any key to continue
            Write-Host "`nPress any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")
        }
        "7" {
            # Option 7: Exit
            $continue = $false
        }
        default {
            Write-Host "Invalid choice. Please enter a valid option."
            Write-Host
            Write-Host "Press any key to continue..."
            $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyUp")
        }
    }
}

# Clear the screen
Clear-Host

# Print closing message
Write-Host "Thank you for using M365 BOSS"
Write-Host

