# Shared Mailbox Automation Script

This PowerShell script automates the creation of shared mailboxes based on ServiceNow tasks. It integrates with Active Directory, Exchange, and ServiceNow APIs to create mailboxes, security groups, and manage permissions, then sends notifications.

## Features
- Fetches task details from ServiceNow.
- Validates mailbox and group data.
- Creates AD users and groups.
- Configures Exchange Online mailbox permissions.
- Sends success/failure notifications.
- Updates ServiceNow task status.

## Prerequisites
- PowerShell 5.1 or later.
- Active Directory module (`Import-Module ActiveDirectory`).
- Exchange Online PowerShell module (`Connect-ExchangeOnline`).
- ServiceNow API access with valid credentials.
- Secure storage for credentials (e.g., encrypted files).

## Setup
1. Replace all placeholders (e.g., `[YOUR_INSTANCE]`, `[YOUR_SMTP_SERVER]`) with your environment-specific values.
2. Ensure file paths for CSVs and attachments are valid.
3. Store credentials securely in files or a credential manager.

## Usage
```powershell
.\SharedMailboxAutomation.ps1
