# Teams Feedback Policy Helper

This script provides an **interactive**, admin-friendly way to assign Microsoft Teams feedback policies to users or groups. In addition to assigning policies, the script can also check the *current* Teams feedback policy applied to a specific user or all members of an Azure AD group, without making any changes.

It’s designed for Global Admins / Teams Admins who want to control the post-meeting / call-feedback experience (surveys, prompts, etc.) without having to remember all the PowerShell details.

---

## What the script does

- Ensures required modules are installed:
  - `MicrosoftTeams`
  - `AzureAD`
- Connects to:
  - Microsoft Teams (`Connect-MicrosoftTeams`)
  - Azure AD (`Connect-AzureAD`)
- Lets you choose **what** you want to do:
  - **Assign/Update** a Teams feedback policy
  - **Check current** Teams feedback policy (no changes)
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
- In **check-only** mode:
  - Retrieves each user's current `TeamsFeedbackPolicy` via `Get-CsOnlineUser`
  - Displays the results in a simple table in the console
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

1. **Choose operation (what you want to do)**

    You’ll first be asked to choose:
    
    - **Assign/Update Teams feedback policy**
    - **Check current Teams feedback policy (no changes)**

    This determines whether the script will **modify** policies or only **report** on existing ones.

2. **Choose scope (who is affected/checked)**

    You'll then be asked:

    Select WHO will be affected/checked:

    - **Single user (enter a single UPN)**
    - **Azure AD group (enter group display name; script resolves users)**
    
  - **Single user**: you’ll enter one UPN (e.g., `user@contoso.com`).
    
  - **Azure AD group**:
    - Enter (part of) the group’s display name.
    - If multiple groups match, you’ll select one from a numbered list.
    - The script will:
    - Resolve the group’s ObjectId.
    - Retrieve its user members.
    - List their UPNs.

3. If you chose **Assign/Update** - choose which policy to assign

    You’ll be shown:

    Please choose which Teams feedback policy to assign:
     - **"Global"**
     - **"Tag:Enabled"**
     - **"Tag:Disabled"**
     - **"Tag:UserChoice"**

    Each option comes with a short description in the script. More details are below in **Policy options explained**.

4. **Review the summary & confirm**

    Before anything changes(Assign), or before the check runs (Check), you’ll see something like:
    - Operation (Assign or Check)
    - Scope (User or Group)
    - Policy selected (Assign only)
    - Number of users
    - All affected/checked UPNs

    Example(Assign mode):
    ```text
    Operation   : Assign
    Scope       : Group
    Policy      : Tag:Disabled
    User Count  : 12
    Users (UPN) :
        - alice@contoso.com
        - bob@contoso.com
        ...
    ```

    Example(Check mode):
    ```text
    Operation   : Check
    Scope       : User
    User Count  : 1
    Users (UPN) :
        - user@contoso.com
    ```

    - In **Assign** mode you'll see:

    ```text
    Do you want to apply this policy to ALL of the users listed above? (Y/N)
    ```
    - Enter **Y** to proceed.
    - Enter **N** to cancel with no changes made.

    - In **Check** mode you'll see:

    ```text
    Proceed to CHECK current policies for the users listed above? (Y/N)
    ```
    - Enter **Y** to fetch and display current policies.
    - Enter **N** to cancel with no changes made.

    ---

5. **Checking current policies (read-only)**

    When you choose the **Check current Teams feedback policy** operation, the script:

    - Resolves the target users (UPN or group members).
    - Calls `Get-CsOnlineUser` for each user.
    -  Builds a small table showing:
      - `DisplayName`
      - `UserPrincipalName`
      - `TeamsFeedbackPolicy`
    -  Outputs this table to the console

    Example Output:
    
    ```text
    DisplayName   UserPrincipalName      TeamsFeedbackPolicy
    -----------   -------------------    -------------------
    Alice Smith   alice@contoso.com      Tag:Disabled
    Bob Jones     bob@contoso.com        Global
    ```

    No changes are made in this mode; it is purely for inspection/reporting.


## Policy options explained

The script lets you choose from the four built-in Teams feedback policies:

1. `Global`
- Uses the tenant-wide **default** Teams feedback policy.
- Whatever your organization’s global policy is set to, users will follow that.
- Choosing Global is effectively “resetting” the user to whatever the default is, instead of a custom/tagged policy.

2. `Tag:Enabled`
- Feedback features are **enabled**.
- Users will continue to see feedback prompts and surveys after calls/meetings.
- Best if you want to ensure users can always send feedback and see surveys.

3. `Tag:Disabled`
- Feedback features are **fully disabled** for affected users.
- This typically means:
    - No post-meeting / call-quality surveys
    - No feedback prompts
    - Feedback UI options may be removed/hidden for those users
- Use this if you want to **stop Teams feedback surveys for a specific user or group**.

4. `Tag:UserChoice`
- Users get to **decide for themselves** in the Teams client whether to participate in feedback/surveys.
- The org provides the capability, but each user controls their own preference.

**Note**: The exact behavior of each policy is defined by Microsoft and may evolve over time, but in general:
   - `Tag:Disabled` = **no feedback/surveys**
   - `Tag:Enabled` = **feedback/surveys allowed**
   - `Tag:UserChoice` = **user decides in client**
   - `Global` = **follow your tenant default**

---

## Logging

The script uses PowerShell’s transcript functionality:
- A log is created in the same folder as the script:
    - `TeamsFeedbackPolicy_YYYYMMDD_HHMMSS.log`
- The log includes:
    - Module checks and installs
    - Connections to Teams/Azure AD
    - Operation choice (Assign vs Check)
    - Policy inspection results (check mode)
    - Scope choices
    - Policy choices
    - All resolved UPNs
    - Per-user grant attempts and any errors
    - Final summary
You can open the log in any text editor for auditing or troubleshooting.

---

## Safety and rollback

If you want to undo changes later:
- Re-run the script and:
    - Reassign `Global` to the same scope (single user or group), or
    - Assign a different policy (e.g., switch from `Tag:Disabled` to `Tag:UserChoice`).
Because the script uses standard `Grant-CsTeamsFeedbackPolicy` cmdlets, it is fully compatible with manual admin operations and other automation.

---

## Disclaimer

This script is provided as-is.
Always test in a lab or with a small pilot scope before running it against large or production-sensitive groups.