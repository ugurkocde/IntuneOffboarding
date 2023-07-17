Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Define WPF XAML
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="MainWindow" Height="270" Width="570" ResizeMode="NoResize">
     <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="5*"/>
            <RowDefinition Height="212*"/>
        </Grid.RowDefinitions>
        <Button Name="SearchButton" Content="Search" HorizontalAlignment="Left" Margin="401,69,0,0" VerticalAlignment="Top" Width="76" Grid.Row="1"/>
        <Button Name="AuthenticateButton" Content="Connect to MS Graph" HorizontalAlignment="Left" Margin="167,42,0,0" VerticalAlignment="Top" Width="121" Grid.Row="1"/>
        <Button Name="InstallModulesButton" Content="Install Modules" HorizontalAlignment="Left" Margin="293,42,0,0" VerticalAlignment="Top" Width="126" Height="20" Grid.Row="1"/>
        <Label Content="Devicename:" HorizontalAlignment="Left" Margin="83,62,0,0" VerticalAlignment="Top" Grid.Row="1" FontSize="18" RenderTransformOrigin="0.673,0.602"/>
        <TextBox Name="DeviceNameTextBox" HorizontalAlignment="Left" Margin="194,65,0,0" VerticalAlignment="Top" Width="195" Height="28" Grid.Row="1" FontSize="18"/>
        <Button Name="OffboardButton" Content="Offboard from Intune, AutoPilot and Azure AD" HorizontalAlignment="Center" Margin="0,198,0,0" VerticalAlignment="Top" Width="268" Grid.Row="1"/>
        <Label Content="Intune Offboarding Tool" HorizontalAlignment="Left" VerticalAlignment="Top" Height="49" Width="264" FontSize="24" Grid.Row="1" Margin="160,0,0,0"/>
        <TextBlock x:Name="intune_status" HorizontalAlignment="Left" Margin="55,108,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Top" Width="154"><Run Language="de-de" Text="Intune Status"/></TextBlock>
        <TextBlock x:Name="autopilot_status" HorizontalAlignment="Left" Margin="55,129,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Top" Width="154"><Run Language="de-de" Text="Autopilot"/><Run Text=" Status"/></TextBlock>
        <TextBlock x:Name="aad_status" HorizontalAlignment="Left" Margin="55,150,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Top" Width="154"><Run Language="de-de" Text="Azure AD"/><Run Text=" Status"/></TextBlock>
        <TextBlock HorizontalAlignment="Left" Margin="247,129,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Top" Text="Serialnumber:" Width="80"/>
        <TextBlock x:Name="serialnumber" HorizontalAlignment="Left" Margin="332,129,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Top" Width="199"/>
        <TextBlock HorizontalAlignment="Left" Margin="246,108,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Top" Text="Devicename:" Width="74"/>
        <TextBlock HorizontalAlignment="Left" Margin="246,150,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Top" Text="OS:" Width="20"/>
        <TextBlock x:Name="devicename" HorizontalAlignment="Left" Margin="332,108,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Top" Width="199" Text=""/>
        <TextBlock x:Name="os" HorizontalAlignment="Left" Margin="332,150,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Top" Width="199" Text=""/>
        <TextBlock HorizontalAlignment="Left" Margin="246,171,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Top" Width="28"><Run Text="User"/><Run Language="de-de" Text=":"/></TextBlock>
        <TextBlock x:Name="user" HorizontalAlignment="Left" Margin="332,171,0,0" Grid.Row="1" TextWrapping="Wrap" VerticalAlignment="Top" Width="199" Text=""/>
    </Grid>
</Window>
"@

# Parse XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml) 
$Window = [Windows.Markup.XamlReader]::Load( $reader )

# Connect to Controls
$SearchButton = $Window.FindName("SearchButton")
$OffboardButton = $Window.FindName("OffboardButton")
$AuthenticateButton = $Window.FindName("AuthenticateButton")
$DeviceNameTextBox = $Window.FindName("DeviceNameTextBox")
$InstallModulesButton = $Window.FindName("InstallModulesButton")

$Window.Add_Loaded({
        try {

            $context = Get-MgContext

            if ($null -eq $context) {

                $AuthenticateButton.Content = "Connect to MS Graph"
                $AuthenticateButton.IsEnabled = $true
            }
            else {

                $AuthenticateButton.Content = "Successfully connected"
                $AuthenticateButton.IsEnabled = $false
            }
        }
        catch {
            $AuthenticateButton.Content = "Not Connected to MS Graph"
            $AuthenticateButton.IsEnabled = $true
        }
    })

$Window.Add_Loaded({
        try {
  
            # Check if the modules are installed
            $modules = @(
                "Microsoft.Graph.Identity.DirectoryManagement",
                "Microsoft.Graph.DeviceManagement",
                "Microsoft.Graph.DeviceManagement.Enrollment"
            )
    
            # If all the modules are already installed, change the button's text and disable it
            if ($modules | ForEach-Object { Get-Module -ListAvailable -Name $_ }) {
                $InstallModulesButton.Content = "Modules Installed"
                $InstallModulesButton.IsEnabled = $false
            }
        }
        catch {
            $AuthenticateButton.Content = "Not Connected to MS Graph"
            $AuthenticateButton.IsEnabled = $true
        }
    })
    

$InstallModulesButton.Add_Click({
        try {
            # Define the modules to be installed
            $modules = @(
                "Microsoft.Graph.Identity.DirectoryManagement",
                "Microsoft.Graph.DeviceManagement",
                "Microsoft.Graph.DeviceManagement.Enrollment"
            )

            # Loop through the modules and install if not already installed
            foreach ($module in $modules) {
                if (!(Get-Module -ListAvailable -Name $module)) {
                    Install-Module $module –Scope CurrentUser –Force –ErrorAction Stop
                }
            }

            # If all the modules have been installed, change the button's text and disable it
            if ($modules | ForEach-Object { Get-Module -ListAvailable -Name $_ }) {
                $InstallModulesButton.Content = "Modules Installed"
                $InstallModulesButton.IsEnabled = $false
            }
        }
        catch {
            [System.Windows.MessageBox]::Show("Error in installing modules. Please ensure you have administrative permissions.")
        }
    })


$AuthenticateButton.Add_Click({
        try {
            Connect-MgGraph –Scopes "DeviceManagementManagedDevices.ReadWrite.All", "DeviceManagementServiceConfig.ReadWrite.All" –ErrorAction Stop
            $context = Get-MgContext

            if ($null -eq $context) {
                $AuthenticateButton.Content = "Authentication Failed"
                $AuthenticateButton.IsEnabled = $true
            }
            else {
                $AuthenticateButton.Content = "Authentication Successful"
                $AuthenticateButton.IsEnabled = $false
            }
        }
        catch {
            $AuthenticateButton.Content = "Authentication Failed"
            $AuthenticateButton.IsEnabled = $true
        }
    })


$SearchButton.Add_Click({
        try {
            $DeviceName = $DeviceNameTextBox.Text
    
            if (![string]::IsNullOrEmpty($DeviceName)) {
                $AADDevice = Get-MgDevice -Search "displayName:$DeviceName" -Property "displayName, ApproximateLastSignInDateTime, IsManaged, OperatingSystem" -CountVariable CountVar -ConsistencyLevel eventual -ErrorAction Stop
                $IntuneDevice = Get-MgDeviceManagementManagedDevice -Filter "deviceName eq '$DeviceName'" -ErrorAction Stop
                $AutopilotDevice = Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -Filter "contains(serialNumber,'$($IntuneDevice.SerialNumber)')" -ErrorAction Stop
    
                $Window.FindName('intune_status').Text = if ($IntuneDevice) { "Intune: Available" } else { "Intune: Unavailable" }
                $Window.FindName('intune_status').Foreground = if ($IntuneDevice) { 'Green' } else { 'Red' }
                $Window.FindName('autopilot_status').Text = if ($AutopilotDevice) { "Autopilot: Available" } else { "Autopilot: Unavailable" }
                $Window.FindName('autopilot_status').Foreground = if ($AutopilotDevice) { 'Green' } else { 'Red' }
                $Window.FindName('aad_status').Text = if ($AADDevice) { "AzureAD: Available" } else { "AzureAD: Unavailable" }
                $Window.FindName('aad_status').Foreground = if ($AADDevice) { 'Green' } else { 'Red' }
    
                $Window.FindName('serialnumber').Text = $IntuneDevice.SerialNumber
                $Window.FindName('devicename').Text = $IntuneDevice.DeviceName
                $Window.FindName('os').Text = $IntuneDevice.OperatingSystem
                $Window.FindName('user').Text = $IntuneDevice.UserDisplayName
            }
        }
        catch {
            [System.Windows.MessageBox]::Show("Error in search operation. Please ensure device name is valid.")
        }
    })
    
$OffboardButton.Add_Click({
        try {
            $DeviceName = $DeviceNameTextBox.Text
    
            if (![string]::IsNullOrEmpty($DeviceName)) {
                $AADDevice = Get-MgDevice -Search "displayName:$DeviceName" -CountVariable CountVar -ConsistencyLevel eventual -ErrorAction Stop
                $IntuneDevice = Get-MgDeviceManagementManagedDevice -Filter "deviceName eq '$DeviceName'" -ErrorAction Stop
                $AutopilotDevice = Get-MgDeviceManagementWindowsAutopilotDeviceIdentity -Filter "contains(serialNumber,'$($IntuneDevice.SerialNumber)')" -ErrorAction Stop
    
                if ($AADDevice) {
                    Remove-MgDevice -DeviceId $AADDevice.Id -ErrorAction Stop
                    [System.Windows.MessageBox]::Show("Successfully removed device from AzureAD.")
                    $Window.FindName('aad_status').Text = "AzureAD: Unavailable"
                }
                else {
                    [System.Windows.MessageBox]::Show("Device not found in AzureAD.")
                }
    
                if ($IntuneDevice) {
                    Remove-MgDeviceManagementManagedDevice -ManagedDeviceId $IntuneDevice.Id -PassThru -ErrorAction Stop
                    [System.Windows.MessageBox]::Show("Successfully removed device from Intune.")
                    $Window.FindName('intune_status').Text = "Intune: Unavailable"
                }
                else {
                    [System.Windows.MessageBox]::Show("Device not found in Intune.")
                }
    
                if ($AutopilotDevice) {
                    Remove-MgDeviceManagementWindowsAutopilotDeviceIdentity -WindowsAutopilotDeviceIdentityId $AutopilotDevice.Id -PassThru -ErrorAction Stop
                    [System.Windows.MessageBox]::Show("Successfully removed device from Autopilot.")
                    $Window.FindName('autopilot_status').Text = "Autopilot: Unavailable"
                }
                else {
                    [System.Windows.MessageBox]::Show("Device not found in Autopilot.")
                }
            }
            else {
                [System.Windows.MessageBox]::Show("Please provide a valid device name.")
            }
        }
        catch {
            [System.Windows.MessageBox]::Show("Error in offboarding operation. Please ensure device name is valid.")
        }
    })
        
# Show Window
$Window.ShowDialog() | Out-Null