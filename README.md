# Teams Feedback Policy Helper

This script provides an **interactive**, admin-friendly way to assign Microsoft Teams feedback policies to users or groups.

It’s designed for Global Admins / Teams Admins who want to control the post-meeting / call feedback experience (surveys, prompts, etc.) without having to remember all the PowerShell details.

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
  - An Azure AD group (by name; script resolves the group and its members)
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
