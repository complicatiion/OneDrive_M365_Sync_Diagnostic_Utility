# OneDrive / M365 / SharePoint Sync Diagnostic Utility

A Windows batch-based diagnostic utility for analyzing OneDrive client issues, Microsoft 365 Office sync problems, and SharePoint library sync issues on Windows 11 systems.

## Purpose

This script is intended for administrative and support troubleshooting scenarios where OneDrive, SharePoint, or Office file synchronization is not working as expected.

It focuses primarily on the local OneDrive sync client and checks common causes such as policy restrictions, Known Folder Move behavior, account configuration, client status, environment paths, and event logs.

## Scope

The script is designed to help investigate issues such as:

- OneDrive sync not starting
- SharePoint libraries not syncing
- Microsoft 365 Office files not saving or syncing correctly
- Known Folder Move (KFM) being blocked by policy
- Desktop, Documents, or Pictures not being redirected into OneDrive
- OneDrive client installed but not functioning correctly
- OneDrive account or tenant configuration issues
- Cached Office / OneDrive sync inconsistencies
- Policy-driven sync restrictions on managed Windows 11 devices

## Main Checks Included

### Policy and Group Policy Checks
- Generates a Group Policy Result report via:
  - `gpresult /h C:\Temp\gp.html`
- Reads OneDrive policy keys from:
  - `HKLM\SOFTWARE\Policies\Microsoft\OneDrive`
  - `HKCU\SOFTWARE\Policies\Microsoft\OneDrive`

### Relevant OneDrive Policy Values
The script checks for important policy entries including:

- `KFMBlockOptIn`
- `KFMBlockOptOut`
- `KFMOptInWithWizard`
- `KFMOptInNoWizard`
- `SilentAccountConfig`
- `FilesOnDemandEnabled`
- `DisablePersonalSync`
- `TenantAutoMount`

### KFM Policy Interpretation
Examples:

- `KFMBlockOptIn=1`  
  Blocks moving known folders into OneDrive.

- `KFMBlockOptOut=1`  
  Prevents moving known folders back to the local PC.

These settings are especially relevant when Desktop, Documents, or Pictures are not syncing as expected.

### OneDrive Client Checks
- Detects OneDrive executable path
- Checks OneDrive process status
- Reads installed OneDrive version
- Reviews local user-based client installation paths
- Checks service and sync-related runtime state where applicable

### Account and Sync Configuration Checks
- Enumerates OneDrive-related account registry entries
- Checks whether OneDrive appears configured for the current user
- Reviews relevant environment and profile locations
- Checks common OneDrive and Office cache folders

### Office / SharePoint / Sync Environment Checks
- Reviews local paths relevant to:
  - OneDrive
  - Office file cache
  - SharePoint sync cache
  - user profile sync context
- Helps identify broken or missing local sync structures

### Event Log Analysis
The script also checks relevant Windows event logs for indicators related to:

- OneDrive
- Office
- SharePoint
- sync-related application issues

## Output

The script provides interactive console-based diagnostics and can also generate a report on the desktop, depending on the version you use.

A temporary Group Policy HTML report is generated here:
C:\Temp\gp.html

## Requirements

* Windows 11
* PowerShell available on the system
* Administrative rights recommended
* OneDrive client installed for full client-side analysis
* Microsoft 365 Apps / Office installed for Office-related checks

## Recommended Use Cases

This utility is useful when troubleshooting cases such as:

* OneDrive known folders cannot be redirected
* SharePoint document libraries fail to sync locally
* Office documents remain in upload pending state
* users receive policy-related OneDrive or KFM errors
* sync works on one machine but not another
* managed devices show different OneDrive behavior due to policy application

## Notes

* The script is diagnostic-focused and does not make configuration changes by default.
* Results should always be interpreted together with tenant policies, Intune/GPO configuration, user licensing, and the actual OneDrive client state.
* On managed enterprise devices, policy settings are often the root cause of OneDrive or SharePoint sync behavior.

## Suggested Follow-Up Checks

After running the script, typical next analysis steps may include:

* opening the generated `gp.html` report
* reviewing OneDrive admin templates and applied policy objects
* comparing affected and unaffected devices
* checking whether the user is signed into the correct organizational account
* validating tenant restrictions and SharePoint library sync permissions
* checking Known Folder Move rollout configuration
* reviewing Office sign-in state and licensing

## Disclaimer

This script is intended for diagnostic and administrative troubleshooting purposes only.
It should be reviewed and tested in your own environment before broad operational use.

