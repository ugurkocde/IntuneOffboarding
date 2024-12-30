# Intune Offboarding Tool

<div align="center">
  <p>
    <a href="https://twitter.com/UgurKocDe">
      <img src="https://img.shields.io/badge/Follow-@UgurKocDe-1DA1F2?style=flat&logo=x&logoColor=white" alt="Twitter Follow"/>
    </a>
    <a href="https://www.linkedin.com/in/ugurkocde/">
      <img src="https://img.shields.io/badge/LinkedIn-Connect-0A66C2?style=flat&logo=linkedin" alt="LinkedIn"/>
    </a>
    <img src="https://img.shields.io/github/license/ugurkocde/IntuneAssignmentChecker?style=flat" alt="License"/>
  </p>
</div>

---

This PowerShell script provides a WPF GUI-based tool that facilitates the offboarding of devices from Microsoft's Intune, AutoPilot, and Entra ID services. The tool leverages Microsoft Graph APIs to authenticate, search, and remove devices.

# Quickstart:

`Install-Script -Name Get-IntuneOffboardingTool`

https://github.com/ugurkocde/IntuneOffboarding/assets/43906965/8fdc12ea-cc6c-4bd5-96f6-6956ca211317

## Features

- Connect to Microsoft Graph
- Search for devices
- Display device information
- Offboard devices from Intune, AutoPilot, and Entra ID

## Requirements

- Microsoft PowerShell 5.1 or later
- Necessary modules:
- Microsoft.Graph.Identity.DirectoryManagement
- Microsoft.Graph.DeviceManagement
- Microsoft.Graph.DeviceManagement.Enrollment

- Permissions:
- DeviceManagementManagedDevices.ReadWrite.All,
- DeviceManagementServiceConfig.ReadWrite.All

These modules will be installed automatically if not present, but you need to have administrative permissions.
Instructions

1. Clone this repository or download the script.
2. Open PowerShell.
3. Navigate to the directory containing the script.
4. Run the script by typing .\Get-IntuneOffboardingTool.ps1

## Using the Tool

1. Connect to Microsoft Graph: Click the "Connect to MS Graph" button to authenticate. The button will display "Successfully Connected" when the connection is established.
2. Install Modules: Click the "Install Modules" button to install the necessary PowerShell modules for the script to run. The button will display "Modules Installed" when the modules are successfully installed.
3. Search: Enter the device name in the provided text box and click the "Search" button. The device details will be displayed in the text blocks below, and the availability status of the device in Intune, Autopilot, and AzureAD will also be shown.
4. Offboard: Click the "Offboard" button to remove the device from Intune, AutoPilot, and Entra ID. A success message will be displayed after each successful removal.

## Notes

- Ensure you have administrative permissions to install necessary modules.
- Ensure you have the necessary privileges in your Microsoft 365 subscription to offboard devices from Intune, AutoPilot, and Entra ID.

## Troubleshooting

If you face any issues while running the script, please check the following:

    You have a valid Microsoft 365 subscription.
    You have installed the necessary modules.
    You have administrative privileges.
    The device name you're searching for exists in the respective services.

## Disclaimer

This tool is provided as-is, with no warranties. Always test scripts and tools in a safe and recoverable environment before deploying them in production. Please ensure you have proper backups before performing any offboarding.
Contributions

## Contributions, issues, and feature requests are welcome.

Feel free to check the issues page.

## License

This project is MIT licensed.
