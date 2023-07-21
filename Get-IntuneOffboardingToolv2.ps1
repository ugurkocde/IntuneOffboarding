Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Define WPF XAML
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Intune Offboarding Tool" Height="500" Width="570" ResizeMode="NoResize">



    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="5*"/>
            <RowDefinition Height="212*"/>
        </Grid.RowDefinitions>
        <Button x:Name="SearchButton" Content="Search" HorizontalAlignment="Left" Margin="393,100,0,0" VerticalAlignment="Top" Width="92" Grid.Row="1" Height="22"/>
        <Button x:Name="AuthenticateButton" Content="Connect to MS Graph" HorizontalAlignment="Left" Margin="76,42,0,0" VerticalAlignment="Top" Width="131" Grid.Row="1"/>
        <Button x:Name="InstallModulesButton" Content="Install/Update Modules" HorizontalAlignment="Left" Margin="318,42,0,0" VerticalAlignment="Top" Width="167" Height="20" Grid.Row="1"/>
        <TextBox x:Name="SearchInputText" VerticalScrollBarVisibility="Auto" TextWrapping="Wrap"  HorizontalAlignment="Left" Margin="176,73,0,0" VerticalAlignment="Top" Width="213" Height="22" Grid.Row="1" FontSize="12"/>
        <Button x:Name="OffboardButton" Content="Offboard device(s) from Intune, AutoPilot and Azure AD" HorizontalAlignment="Left" Margin="78,333,0,0" VerticalAlignment="Top" Width="308" Grid.Row="1"/>
        <Label Content="Intune Offboarding Tool" HorizontalAlignment="Left" VerticalAlignment="Top" Height="49" Width="264" FontSize="24" Grid.Row="1" Margin="160,0,0,0"/>
        <TextBlock x:Name="intune_status" HorizontalAlignment="Left" Margin="79,251,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Top" Width="154"><Run Language="de-de" Text="Intune Status"/></TextBlock>
        <TextBlock x:Name="autopilot_status" HorizontalAlignment="Left" Margin="79,272,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Top" Width="154"><Run Language="de-de" Text="Autopilot"/><Run Text=" Status"/></TextBlock>
        <TextBlock x:Name="aad_status" HorizontalAlignment="Left" Margin="79,293,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Top" Width="154"><Run Language="de-de" Text="Azure AD"/><Run Text=" Status"/></TextBlock>
        <ComboBox x:Name="dropdown" HorizontalAlignment="Left" Margin="76,73,0,0" Grid.Row="1" VerticalAlignment="Top" Width="95"/>
        <DataGrid x:Name="SearchResultsDataGrid" Grid.Row="1" Margin="78,127,71,210"/>
        <Button x:Name="logs_button" Content="Logs" HorizontalAlignment="Left" Margin="504,422,0,0" VerticalAlignment="Top" Grid.Row="1"/>
        <Button x:Name="export_button" Content="Export results" HorizontalAlignment="Left" Margin="409,252,0,0" Grid.Row="1" VerticalAlignment="Top"/>
        <Button x:Name="bulk_import_button" Content="Bulk Import" HorizontalAlignment="Left" Margin="392,73,0,0" VerticalAlignment="Top" Width="93" Grid.Row="1" Height="22"/>
        <ComboBox x:Name="dropdown_lastsync_days" HorizontalAlignment="Left" Margin="79,367,0,0" Grid.Row="1" VerticalAlignment="Top" Width="94"/>
        <Button x:Name="export_stale_devices_button" Content="Export Stale devices" HorizontalAlignment="Left" Margin="195,367,0,0" VerticalAlignment="Top" Width="118" Grid.Row="1" Height="23"/>
        <ComboBox x:Name="dropdown_lastsync_platform" HorizontalAlignment="Left" Margin="79,396,0,0" Grid.Row="1" VerticalAlignment="Top" Width="94"/>
        <Button x:Name="disconnect_button" Content="Disconnect" Width="80" HorizontalAlignment="Left" Margin="419,422,0,0" VerticalAlignment="Top" Grid.Row="1"/>
        <Button x:Name="check_permissions_button" Content="Check Permissions" HorizontalAlignment="Left" Margin="212,42,0,0" Grid.Row="1" VerticalAlignment="Top" />
    </Grid>


    </Window>
"@

# Parse XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml) 
$Window = [Windows.Markup.XamlReader]::Load( $reader )

$script:LogFilePath = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "IntuneOffboardingTool_Log.txt")

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string] $Message
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "$timestamp - $Message"

    Add-Content -Path $script:LogFilePath -Value $logMessage
}

# Connect to Controls
$SearchButton = $Window.FindName("SearchButton")
$OffboardButton = $Window.FindName("OffboardButton")
$AuthenticateButton = $Window.FindName("AuthenticateButton")
$SearchInputText = $Window.FindName("SearchInputText")
$InstallModulesButton = $Window.FindName("InstallModulesButton")
$bulk_import_button = $Window.FindName('bulk_import_button')
$Dropdown = $Window.FindName("dropdown")
$Disconnect = $Window.FindName('disconnect_button')
$Export_Button = $Window.FindName('export_button')
$logs_button = $Window.FindName('logs_button')
$CheckPermissionsButton = $Window.FindName('check_permissions_button')




$Dropdown_LastSync_Platform = $Window.FindName('dropdown_lastsync_platform')
$Dropdown_LastSync_Days = $Window.FindName('dropdown_lastsync_days')
$Export_Stale_Devices = $Window.FindName('export_stale_devices_button')

$SearchInputText.Add_GotFocus({
        $SearchInputText.Height = 50  
        $SearchInputText.Width = 213
    })

$SearchInputText.Add_LostFocus({
        $SearchInputText.Height = 22  
        $SearchInputText.Width = 213
    })
    

$Window.Add_Loaded({
        $Dropdown.Items.Add("Devicename")
        $Dropdown.Items.Add("Serialnumber")
        $Dropdown.SelectedIndex = 0

        $Dropdown_LastSync_Platform.Items.Add("Windows")
        $Dropdown_LastSync_Platform.Items.Add("Android")
        $Dropdown_LastSync_Platform.Items.Add("iOS")
        $Dropdown_LastSync_Platform.Items.Add("macOS")
        $Dropdown_LastSync_Platform.SelectedIndex = 0

        $Dropdown_LastSync_Days.Items.Add("30 days")
        $Dropdown_LastSync_Days.Items.Add("60 days")
        $Dropdown_LastSync_Days.Items.Add("90 days")
        $Dropdown_LastSync_Days.Items.Add("120 days")
        $Dropdown_LastSync_Days.Items.Add("150 days")
        $Dropdown_LastSync_Days.Items.Add("180 days")
        $Dropdown_LastSync_Days.Items.Add("210 days")
        $Dropdown_LastSync_Days.Items.Add("240 days")
        $Dropdown_LastSync_Days.Items.Add("270 days")
        $Dropdown_LastSync_Days.Items.Add("300 days")
        $Dropdown_LastSync_Days.Items.Add("330 days")
        $Dropdown_LastSync_Days.Items.Add("365 days")
        $Dropdown_LastSync_Days.SelectedIndex = 0    
    })


$Window.Add_Loaded({
        try {
            Write-Log "Window is loading..."
    
            $context = Get-MgContext
    
            if ($null -eq $context) {
                Write-Log "Not connected to MS Graph"
                $AuthenticateButton.Content = "Connect to MS Graph"
                $AuthenticateButton.IsEnabled = $true
                $Disconnect.IsEnabled = $false
            }
            else {
                Write-Log "Successfully connected to MS Graph"
                $AuthenticateButton.Content = "Successfully connected"
                $AuthenticateButton.IsEnabled = $false
                $Disconnect.IsEnabled = $true
            }
        }
        catch {
            Write-Log "Error occurred: $_"
            $AuthenticateButton.Content = "Not Connected to MS Graph"
            $AuthenticateButton.IsEnabled = $true
        }
    })
    
$Disconnect.Add_Click({
        try {
            Write-Log "Attempting to disconnect from MS Graph..."
            Disconnect-MgGraph
            $Disconnect.Content = "Disconnected"
            $Disconnect.IsEnabled = $false  # Disable Disconnect button after disconnect
            $AuthenticateButton.Content = "Connect to MS Graph"
            $AuthenticateButton.IsEnabled = $true
            $CheckPermissionsButton.IsEnabled = $false  # Disable CheckPermissionsButton after disconnect
        }
        catch {
            Write-Log "Error occurred while attempting to disconnect from MS Graph: $_"
            [System.Windows.MessageBox]::Show("Error in disconnect operation.")
        }
    })
    
$AuthenticateButton.Add_Click({
        try {
            Connect-MgGraph -Scopes "DeviceManagementManagedDevices.ReadWrite.All", "DeviceManagementServiceConfig.ReadWrite.All" -ErrorAction Stop
            $context = Get-MgContext
        
            if ($null -eq $context) {
                Write-Log "Authentication Failed"
                $AuthenticateButton.Content = "Authentication Failed"
                $AuthenticateButton.IsEnabled = $true
                $Disconnect.Content = "Disconnected"  
                $Disconnect.IsEnabled = $false  
                $CheckPermissionsButton.IsEnabled = $false  
            }
            else {
                Write-Log "Authentication Successful"
                $AuthenticateButton.Content = "Authentication Successful"
                $AuthenticateButton.IsEnabled = $false
                $Disconnect.Content = "Disconnect"  
                $Disconnect.IsEnabled = $true  
                $CheckPermissionsButton.IsEnabled = $true  
            }
        }
        catch {
            Write-Log "Error occurred during authentication. Exception: $_"
            $AuthenticateButton.Content = "Authentication Failed"
            $AuthenticateButton.IsEnabled = $true
            $Disconnect.Content = "Disconnected"  
            $Disconnect.IsEnabled = $false  
            $CheckPermissionsButton.IsEnabled = $false  
        }
    })
    
    
$CheckPermissionsButton.Add_Click({
        try {
            Write-Log "Check Permissions button clicked, attempting to retrieve user context..."

            $context = Get-MGContext

            if ($context) {
                $username = $context.Account

                # Check if the required scopes are present
                $requiredScopes = @("DeviceManagementManagedDevices.ReadWrite.All", "DeviceManagementServiceConfig.ReadWrite.All")
                $missingScopes = $requiredScopes | Where-Object { $_ -notin $context.Scopes }

                if ($missingScopes) {
                    $missingScopesList = $missingScopes -join ', '
                    $message = "WARNING: The following required permissions are missing for $username : $missingScopesList"
                }
                else {
                    $message = "All required permissions are granted for $username."
                }

                [System.Windows.MessageBox]::Show($message)
            }
            else {
                [System.Windows.MessageBox]::Show("No user context found. Please authenticate first.")
            }
        }
        catch {
            Write-Log "Error in retrieving user context: $_"
            [System.Windows.MessageBox]::Show("Error in retrieving user context.")
        }
    })

    
$Export_Button.Add_Click({
        try {
            Write-Log "Export button clicked, attempting to export data..."

            $data = $Window.FindName('SearchResultsDataGrid').ItemsSource
    
            if ($data -ne $null -and $data.Count -gt 0) {

                $outputPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "SearchResults.csv")
    

                $data | Export-Csv -Path $outputPath -NoTypeInformation
    
                Write-Log "Search results exported successfully to SearchResults.csv on your Desktop."
                [System.Windows.MessageBox]::Show("Search results exported successfully to SearchResults.csv on your Desktop.")
            }
            else {
                Write-Log "No search results to export."
                [System.Windows.MessageBox]::Show("No search results to export.")
            }
        }
        catch {
            Write-Log "Error occurred during the export operation: $_"
            [System.Windows.MessageBox]::Show("Error in export operation.")
        }
    })
    

$Export_Stale_Devices.Add_Click({
        try {
            Write-Log "Export stale devices button clicked, attempting to export data..."
            $daysRaw = $Window.FindName('dropdown_lastsync_days').SelectedItem
            $OS = $Window.FindName('dropdown_lastsync_platform').SelectedItem
    
            $days = [int]($daysRaw -replace " days", "")
    
            if (![string]::IsNullOrEmpty($days) -and ![string]::IsNullOrEmpty($OS)) {
                $pastDate = (Get-Date).AddDays(-$days).ToString('yyyy-MM-ddTHH:mm:ssZ')
    
                $StaleDevices = Get-MgDeviceManagementManagedDevice -Filter "lastSyncDateTime le $pastDate and OperatingSystem eq '$OS'"
    
                if ($StaleDevices) {

                    $deviceNames = $StaleDevices | ForEach-Object { $_.DeviceName } | Out-String
                
                    $outputPath = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "StaleDevices.txt")
                    Set-Content -Path $outputPath -Value $deviceNames
    
                    Write-Log "Stale devices exported successfully to StaleDevices.txt on your Desktop."
                    [System.Windows.MessageBox]::Show("Stale devices exported successfully to StaleDevices.txt on your Desktop.")
                }
                else {
                    Write-Log "No stale devices found."
                    [System.Windows.MessageBox]::Show("No stale devices found.")
                }
            }
            else {
                Write-Log "Number of days and platform not selected."
                [System.Windows.MessageBox]::Show("Please select the number of days and platform.")
            }
        }
        catch {
            Write-Log "Error in export operation. Please ensure you have selected the correct number of days and platform: $_"
            [System.Windows.MessageBox]::Show("Error in export operation. Please ensure you have selected the correct number of days and platform.")
        }
    })
    
$Window.Add_Loaded({
        try {
            Write-Log "Window is loading..."
      
            $modules = @(
                "Microsoft.Graph.Identity.DirectoryManagement",
                "Microsoft.Graph.DeviceManagement",
                "Microsoft.Graph.DeviceManagement.Enrollment"
            )
      
            if ($modules | ForEach-Object { Get-Module -ListAvailable -Name $_ }) {
                $InstallModulesButton.Content = "Modules Installed"
                $InstallModulesButton.IsEnabled = $false
                Write-Log "Modules already installed."
            }
        }
        catch {
            Write-Log "Error occurred while checking if modules are installed: $_"
            $AuthenticateButton.Content = "Not Connected to MS Graph"
            $AuthenticateButton.IsEnabled = $true
        }
    })
    
$InstallModulesButton.Add_Click({
        try {
            Write-Log "Install Modules button clicked, attempting to install modules..."
    
            $modules = @(
                "Microsoft.Graph.Identity.DirectoryManagement",
                "Microsoft.Graph.DeviceManagement",
                "Microsoft.Graph.DeviceManagement.Enrollment"
            )
    
            $InstallModulesButton.Content = "Installing..."
            $InstallModulesButton.IsEnabled = $false
    
            $job = Start-Job -ScriptBlock {
                param($modules)
    
                foreach ($module in $modules) {
                    if (!(Get-Module -ListAvailable -Name $module)) {
                        Write-Host "Installing module: $module..."
                        Install-Module $module -Scope CurrentUser -Force -ErrorAction Stop
                    }
                }
            } -ArgumentList $modules
    
            # Register event to capture installation completion and update GUI accordingly
            Register-ObjectEvent -InputObject $job -EventName StateChanged -Action {
                if ($event.SourceEventArgs.JobStateInfo.State -eq 'Completed') {
                    if ($modules | ForEach-Object { Get-Module -ListAvailable -Name $_ }) {
                        $InstallModulesButton.Dispatcher.Invoke({
                                $InstallModulesButton.Content = "Modules Installed"
                            })
                        Write-Log "All modules installed successfully."
                    }
                }
                elseif ($event.SourceEventArgs.JobStateInfo.State -eq 'Failed') {
                    $InstallModulesButton.Dispatcher.Invoke({
                            $InstallModulesButton.Content = "Install Modules"
                            $InstallModulesButton.IsEnabled = $true
                        })
                    Write-Log "Error in installing modules. Please ensure you have administrative permissions: $_"
                    [System.Windows.MessageBox]::Show("Error in installing modules. Please ensure you have administrative permissions.")
                }
            }
        }
        catch {
            Write-Log "Exception: $_"
            [System.Windows.MessageBox]::Show("Error in installing modules. Please ensure you have administrative permissions.")
        }
    })
    

$SearchButton.Add_Click({

        if ($AuthenticateButton.IsEnabled) {
            
            Write-Log "User is not connected to MS Graph. Attempted search operation."
            [System.Windows.MessageBox]::Show("You are not connected to MS Graph. Please connect first.")
            return
        }

        try {
            $SearchTexts = $SearchInputText.Text -split ', '
            Write-Log "$SearchTexts"
            $searchOption = $Dropdown.SelectedItem
    
            $searchResults = New-Object 'System.Collections.Generic.List[System.Object]'
            $AADCount = 0
            $IntuneCount = 0
            $AutopilotCount = 0
    
            foreach ($SearchText in $SearchTexts) {
                if (![string]::IsNullOrEmpty($SearchText)) {
                    if ($searchOption -eq "Devicename") {
                        $AADDevices = Get-MgDevice -Search "displayName:$SearchText" -CountVariable CountVar -ConsistencyLevel eventual -ErrorAction Stop | Select-Object -Property displayName, ApproximateLastSignInDateTime, IsManaged, OperatingSystem
                        $IntuneDevices = Get-MgDeviceManagementManagedDevice -Filter "deviceName eq '$SearchText'" -ErrorAction Stop | Select-Object -Property UserDisplayName, OperatingSystem, SerialNumber, DeviceName, LastSyncDateTime
    
                        foreach ($AADDevice in $AADDevices) {
                            foreach ($IntuneDevice in $IntuneDevices) {
                                if ($IntuneDevice.DeviceName -eq $AADDevice.displayName) {
                                    $AutopilotDevice = Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -Filter "contains(serialNumber,'$($IntuneDevice.SerialNumber)')" -ErrorAction Stop | Select-Object -Property GroupTag, SerialNumber, LastContactedDateTime
    
                                    $CombinedDevice = New-Object PSObject -Property @{
                                        "DeviceName"             = $IntuneDevice.DeviceName
                                        "SerialNumber"           = $IntuneDevice.SerialNumber
                                        "OperatingSystem"        = $AADDevice.OperatingSystem
                                        "Primary User"           = $IntuneDevice.UserDisplayName
                                        "AzureAD Last Contact"   = $AADDevice.ApproximateLastSignInDateTime
                                        "Intune Last Contact"    = $IntuneDevice.LastSyncDateTime
                                        "Autopilot Last Contact" = $AutopilotDevice.LastContactedDateTime
                                    }
                                    
    
                                    $searchResults.Add($CombinedDevice) | Select-Object "DeviceName", "SerialNumber", "OperatingSystem", "Primary User", "AzureAD Last Contact", "Intune Last Contact", "Autopilot Last Contact"
                                    if ($AADDevice) { $AADCount++ }
                                    if ($IntuneDevice) { $IntuneCount++ }
                                    if ($AutopilotDevice) { $AutopilotCount++ }
                                }
                            }
                        }
                    }
                    elseif ($searchOption -eq "Serialnumber") {
                        $IntuneDevices = Get-MgDeviceManagementManagedDevice -Filter "SerialNumber eq '$SearchText'" -ErrorAction Stop | Select-Object -Property UserDisplayName, OperatingSystem, SerialNumber, DeviceName, LastSyncDateTime
                        $displayName = $IntuneDevices.DeviceName
                        $AADDevices = Get-MgDevice -Search "displayName:$displayName" -CountVariable CountVar -ConsistencyLevel eventual -ErrorAction Stop | Select-Object -Property displayName, ApproximateLastSignInDateTime, IsManaged, OperatingSystem
    
                        foreach ($AADDevice in $AADDevices) {
                            foreach ($IntuneDevice in $IntuneDevices) {
                                if ($IntuneDevice.DeviceName -eq $AADDevice.displayName) {
                                    $AutopilotDevice = Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -Filter "contains(serialNumber,'$SearchText')" -ErrorAction Stop | Select-Object -Property GroupTag, SerialNumber, LastContactedDateTime
                        
                                    $CombinedDevice = New-Object PSObject -Property @{
                                        "DeviceName"             = $IntuneDevice.DeviceName
                                        "SerialNumber"           = $IntuneDevice.SerialNumber
                                        "OperatingSystem"        = $AADDevice.OperatingSystem
                                        "Primary User"           = $IntuneDevice.UserDisplayName
                                        "AzureAD Last Contact"   = $AADDevice.ApproximateLastSignInDateTime
                                        "Intune Last Contact"    = $IntuneDevice.LastSyncDateTime
                                        "Autopilot Last Contact" = $AutopilotDevice.LastContactedDateTime
                                    }
                        
                                    $searchResults.Add($CombinedDevice) | Select-Object "DeviceName", "SerialNumber", "OperatingSystem", "Primary User", "AzureAD Last Contact", "Intune Last Contact", "Autopilot Last Contact"
                                    if ($AADDevice) { $AADCount++ }
                                    if ($IntuneDevice) { $IntuneCount++ }
                                    if ($AutopilotDevice) { $AutopilotCount++ }
                                }
                            }
                        }
                        
                    }
                }
            }
    
            $Window.FindName('intune_status').Text = "Intune: $IntuneCount Available"
            $Window.FindName('intune_status').Foreground = if ($IntuneCount -gt 0) { 'Green' } else { 'Red' }
            $Window.FindName('autopilot_status').Text = "Autopilot: $AutopilotCount Available"
            $Window.FindName('autopilot_status').Foreground = if ($AutopilotCount -gt 0) { 'Green' } else { 'Red' }
            $Window.FindName('aad_status').Text = "AzureAD: $AADCount Available"
            $Window.FindName('aad_status').Foreground = if ($AADCount -gt 0) { 'Green' } else { 'Red' }
    
            $Window.FindName('SearchResultsDataGrid').ItemsSource = $searchResults
        }
        catch {
            Write-Log "Error occurred during search operation. Exception: $_"
            [System.Windows.MessageBox]::Show("Error in search operation. Please ensure the Serialnumber or Devicename is valid.")
        }
    })
    
        
$bulk_import_button.Add_Click({
        try {

            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.filter = "CSV files (*.csv)|*.csv|TXT files (*.txt)|*.txt"
            $OpenFileDialog.ShowDialog() | Out-Null

            $filePath = $OpenFileDialog.FileName

            if (Test-Path $filePath) {

                $devices = Get-Content -Path $filePath
                $deviceNames = $devices

                $deviceNamesString = $deviceNames -join ", "

                $SearchInputText.Text = $deviceNamesString

            }
        }
        catch {
            Write-Log "Exception: $_"
            [System.Windows.MessageBox]::Show("Error in bulk import operation. Please ensure the file is valid and try again.")
        }
    })

$OffboardButton.Add_Click({

        if ($AuthenticateButton.IsEnabled) {
            Write-Log "User is not connected to MS Graph. Attempted offboarding operation."
            [System.Windows.MessageBox]::Show("You are not connected to MS Graph. Please connect first.")
            return
        }

        $confirmationResult = [System.Windows.MessageBox]::Show("Are you sure you want to proceed with offboarding? This action cannot be undone.", "Confirm Offboarding", [System.Windows.MessageBoxButton]::YesNo)
        if ($confirmationResult -eq 'No') {
            Write-Log "User canceled offboarding operation."
            return
        }

        try {
            $SearchTexts = $SearchInputText.Text -split ', '

            foreach ($SearchText in $SearchTexts) {
                if (![string]::IsNullOrEmpty($SearchText)) {
                    $AADDevice = Get-MgDevice -Search "displayName:$SearchText" -CountVariable CountVar -ConsistencyLevel eventual -ErrorAction Stop
                    $IntuneDevice = Get-MgDeviceManagementManagedDevice -Filter "deviceName eq '$SearchText'" -ErrorAction Stop
                    $AutopilotDevice = Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -Filter "contains(serialNumber,'$($IntuneDevice.SerialNumber)')" -ErrorAction Stop

                    if ($AADDevice) {
                        Remove-MgDevice -DeviceId $AADDevice.Id -ErrorAction Stop
                        [System.Windows.MessageBox]::Show("Successfully removed device $SearchText from AzureAD.")
                        $Window.FindName('aad_status').Text = "AzureAD: Unavailable"
                        Write-Log "Successfully removed device $SearchText from Azure AD."
                    }
                    else {
                        [System.Windows.MessageBox]::Show("Device $SearchText not found in AzureAD.")
                    }

                    if ($IntuneDevice) {
                        Remove-MgDeviceManagementManagedDevice -ManagedDeviceId $IntuneDevice.Id -PassThru -ErrorAction Stop
                        [System.Windows.MessageBox]::Show("Successfully removed device $SearchText from Intune.")
                        $Window.FindName('intune_status').Text = "Intune: Unavailable"
                        Write-Log "Successfully removed device $SearchText from Intune."
                    }
                    else {
                        [System.Windows.MessageBox]::Show("Device $SearchText not found in Intune.")
                    }

                    if ($AutopilotDevice) {
                        Remove-MgDeviceManagementWindowsAutopilotDeviceIdentity -WindowsAutopilotDeviceIdentityId $AutopilotDevice.Id -PassThru -ErrorAction Stop
                        [System.Windows.MessageBox]::Show("Successfully removed device $SearchText from Autopilot.")
                        $Window.FindName('autopilot_status').Text = "Autopilot: Unavailable"
                        Write-Log "Successfully removed device $SearchText from Autopilot."
                    }
                    else {
                        [System.Windows.MessageBox]::Show("Device $SearchText not found in Autopilot.")
                    }
                }
                else {
                    [System.Windows.MessageBox]::Show("Please provide a valid device name.")
                }
            }
        }
        catch {
            Write-Log "Error in offboarding operation. Exception: $_"
            [System.Windows.MessageBox]::Show("Error in offboarding operation. Please ensure device names are valid.")
        }
    })
    

$logs_button.Add_Click({
        $logFilePath = [System.IO.Path]::Combine([Environment]::GetFolderPath("Desktop"), "IntuneOffboardingTool_Log.txt")
        if (Test-Path $logFilePath) {
            Invoke-Item $logFilePath
        }
        else {
            Write-Host "Log file not found."
        }
    })
        
# Show Window
$Window.ShowDialog() | Out-Null