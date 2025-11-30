# Teams Feedback Policy Helper

This script provides an **interactive**, admin-friendly way to assign Microsoft Teams feedback policies to users or groups.

It’s designed for Global Admins / Teams Admins who want to control the post-meeting / call-feedback experience (surveys, prompts, etc.) without having to remember all the PowerShell details.

---

## What the script does

- Ensures required modules are installed:
  - `MicrosoftTeams`
  - `AzureAD`
- Connects to:
  - Microsoft Teams (`Connect-MicrosoftTeams`)
  - Azure AD (`Connect-AzureAD`)
- Lets you choose **who** will be affected:
  - A single user (by UPN)
  - An Azure AD group (by name; script resolves users)
- Lets you choose **how** they will be affected:
  - Which **Teams feedback policy** to apply
- Shows a **summary** before making any changes:
  - Scope (User vs Group)
  - Selected policy
  - **All affected users by UPN**
- Applies the changes via `Grant-CsTeamsFeedbackPolicy`
- Logs **every step** to:
  - The terminal
  - A timestamped log file (`TeamsFeedbackPolicy_YYYYMMDD_HHMMSS.log`)

---

## Prerequisites

- Windows PowerShell (recommended to run as Administrator)
- Sufficient privileges:
  - **Global Administrator** or **Teams Administrator**
- Internet access to:
  - Install modules from the PowerShell Gallery (if not already installed)
- Permissions to sign in to:
  - Microsoft Teams
  - Azure AD

The script will automatically install/import the following modules if they are missing:

- `MicrosoftTeams`
- `AzureAD`

If module installation fails (e.g., due to missing admin rights, blocked PSGallery, etc.), the script will stop and tell you what went wrong.

---

## How to run

1. Download the script (for example):

   ```powershell
   Set-TeamsFeedbackPolicyInteractive.ps1

2. Open **Windows PowerShell** as Administrator:

3. Navigate to the script folder:

    ```powershell
    cd C:\Path\To\Script

4. If it's your first time running scripts, you may need to loosen the execution policy (optional example):

    ```powershell
    Set-ExecutionPolicy RemoteSigned

5. Run the script:

    ```powershell
    .\Set-TeamsFeedbackPolicyInteractive.ps1

6. Follow the prompts in the terminal

---

## Interactive Flow

1. **Choose scope (who is affected)**

    You’ll be asked:
    
    Select WHO will be affected by this change:
    - **`Single user (enter a single UPN)`**
    - **`Azure AD group (enter group display name; script resolves users)`**



Single user: you’ll enter one UPN (e.g., user@contoso.com).

Azure AD group:

Enter (part of) the group’s display name.

If multiple groups match, you’ll select one from a numbered list.

The script will:

Resolve the group’s ObjectId.

Retrieve its user members.

List their UPNs.

2. Choose which policy to assign

You’ll be shown:

Please choose which Teams feedback policy to assign:

Global

Tag:Enabled

Tag:Disabled

Tag:UserChoice

Each option comes with a short description in the script. More details are below.

3. Review the summary

Before anything changes, you’ll see something like:

Scope (User or Group)

Policy selected

Number of users

All affected UPNs