Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Add the DeviceObject class definition
if (-not ([System.Management.Automation.PSTypeName]'DeviceObject').Type) {
    Add-Type -TypeDefinition @"
    using System;
    using System.ComponentModel;

    public class DeviceObject : INotifyPropertyChanged
    {
        private bool isSelected;
        public bool IsSelected
        {
            get { return isSelected; }
            set 
            { 
                isSelected = value;
                OnPropertyChanged("IsSelected");
            }
        }
        
        public string DeviceName { get; set; }
        public string SerialNumber { get; set; }
        public string OperatingSystem { get; set; }
        public string PrimaryUser { get; set; }
        public DateTime? AzureADLastContact { get; set; }
        public DateTime? IntuneLastContact { get; set; }
        public DateTime? AutopilotLastContact { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;

        protected void OnPropertyChanged(string name)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
"@
}

# Define a helper function for paginated Graph API calls
function Get-GraphPagedResults {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Uri
    )
    
    $results = @()
    $nextLink = $Uri
    
    do {
        try {
            $response = Invoke-MgGraphRequest -Uri $nextLink -Method GET
            if ($response.value) {
                $results += $response.value
            }
            $nextLink = $response.'@odata.nextLink'
        }
        catch {
            Write-Log "Error in pagination: $_"
            break
        }
    } while ($nextLink)
    
    return $results
}

# Define WPF XAML
[xml]$xaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Device Offboarding Manager" Height="700" Width="1200" 
    Background="#F0F0F0"
    WindowStartupLocation="CenterScreen" 
    ResizeMode="NoResize">
    
    <Window.Resources>
        <!-- Base Button Style -->
        <Style TargetType="Button">
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="Padding" Value="12,5"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Height" Value="28"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                CornerRadius="2" 
                                BorderThickness="{TemplateBinding BorderThickness}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#106EBE"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#CCCCCC"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Menu Button Style -->
        <Style x:Key="MenuButtonStyle" TargetType="RadioButton">
            <Setter Property="Foreground" Value="#808080"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Height" Value="40"/>
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RadioButton">
                        <Border x:Name="border" 
                                Background="{TemplateBinding Background}"
                                BorderThickness="0">
                            <Grid>
                                <Border x:Name="indicator" 
                                        Width="3" 
                                        Background="Transparent"
                                        HorizontalAlignment="Left"/>
                                <ContentPresenter Margin="20,0,0,0" 
                                                VerticalAlignment="Center"/>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#404040"/>
                                <Setter Property="Foreground" Value="White"/>
                                <Setter TargetName="indicator" Property="Background" Value="#0078D4"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter Property="Background" Value="#404040"/>
                                <Setter Property="Foreground" Value="White"/>
                                <Setter TargetName="indicator" Property="Background" Value="#0078D4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Sidebar Connection Button Style -->
        <Style x:Key="SidebarButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="#404040"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="Height" Value="32"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="2"
                                BorderThickness="0">
                            <ContentPresenter HorizontalAlignment="Center" 
                                            VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#505050"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Background" Value="#333333"/>
                                <Setter Property="Foreground" Value="#808080"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Playbook Button Style -->
        <Style x:Key="PlaybookButtonStyle" TargetType="Button">
            <Setter Property="Background" Value="#28A745"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Padding" Value="20,15"/>
            <Setter Property="Margin" Value="0,0,0,10"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="4"
                                Padding="{TemplateBinding Padding}">
                            <Grid>
                                <ContentPresenter HorizontalAlignment="Left" 
                                                VerticalAlignment="Center"/>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#218838"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- TextBox Style -->
        <Style TargetType="TextBox">
            <Setter Property="Height" Value="28"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#CCCCCC"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <!-- ComboBox Style -->
        <Style TargetType="ComboBox">
            <Setter Property="Height" Value="28"/>
            <Setter Property="Padding" Value="8,5"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="#CCCCCC"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <!-- DataGrid Style -->
        <Style TargetType="DataGrid">
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#CCCCCC"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="RowHeight" Value="35"/>
            <Setter Property="RowBackground" Value="White"/>
            <Setter Property="AlternatingRowBackground" Value="#F8F8F8"/>
            <Setter Property="HorizontalGridLinesBrush" Value="#E0E0E0"/>
            <Setter Property="VerticalGridLinesBrush" Value="#E0E0E0"/>
            <Setter Property="ColumnHeaderHeight" Value="32"/>
        </Style>

        <!-- DataGridColumnHeader Style -->
        <Style TargetType="DataGridColumnHeader">
            <Setter Property="Background" Value="#F5F5F5"/>
            <Setter Property="Foreground" Value="#323130"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Padding" Value="8,0"/>
            <Setter Property="BorderBrush" Value="#E0E0E0"/>
            <Setter Property="BorderThickness" Value="0,0,1,1"/>
        </Style>

        <!-- Authentication Radio Button Style -->
        <Style x:Key="AuthRadioButtonStyle" TargetType="RadioButton">
            <Setter Property="Margin" Value="0,8,8,8"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Foreground" Value="#2D3748"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RadioButton">
                        <Border x:Name="border" 
                                Background="{TemplateBinding Background}"
                                BorderBrush="#E2E8F0"
                                BorderThickness="1"
                                CornerRadius="6"
                                Padding="12">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="24"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Ellipse x:Name="radioOuter"
                                         Width="18" Height="18"
                                         Stroke="#CBD5E0"
                                         StrokeThickness="2"
                                         Fill="Transparent"/>
                                <Ellipse x:Name="radioInner"
                                         Width="10" Height="10"
                                         Fill="#0078D4"
                                         Opacity="0"/>
                                <ContentPresenter Grid.Column="1"
                                                Margin="12,0,0,0"
                                                VerticalAlignment="Center"/>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#F7FAFC"/>
                                <Setter TargetName="radioOuter" Property="Stroke" Value="#0078D4"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="radioInner" Property="Opacity" Value="1"/>
                                <Setter TargetName="radioOuter" Property="Stroke" Value="#0078D4"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#0078D4"/>
                                <Setter TargetName="border" Property="Background" Value="#F0F9FF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- TextBox Style -->
        <Style x:Key="AuthTextBoxStyle" TargetType="TextBox">
            <Setter Property="Height" Value="36"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="Margin" Value="0,4"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#E2E8F0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="BorderBrush" Value="#0078D4"/>
                </Trigger>
                <Trigger Property="IsFocused" Value="True">
                    <Setter Property="BorderBrush" Value="#0078D4"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Password Box Style -->
        <Style x:Key="AuthPasswordBoxStyle" TargetType="PasswordBox">
            <Setter Property="Height" Value="36"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="Margin" Value="0,4"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#E2E8F0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <!-- Button Style -->
        <Style x:Key="AuthButtonStyle" TargetType="Button">
            <Setter Property="Height" Value="40"/>
            <Setter Property="Padding" Value="24,0"/>
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" 
                                            VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#106EBE"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#005A9E"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Secondary Button Style -->
        <Style x:Key="SecondaryButtonStyle" TargetType="Button">
            <Setter Property="Height" Value="40"/>
            <Setter Property="Padding" Value="24,0"/>
            <Setter Property="Background" Value="#F0F0F0"/>
            <Setter Property="Foreground" Value="#2D3748"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" 
                                            VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#E2E2E2"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#D4D4D4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="200"/>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>

        <!-- Sidebar -->
        <Border Grid.Column="0" Background="#2D2D2D">
            <DockPanel>
                <!-- Menu Items -->
                <StackPanel DockPanel.Dock="Bottom" Margin="0,0,0,0">
                    <!-- Prominent Connect Button -->
                    <Border Margin="15,5,15,10" 
                            Background="#0078D4" 
                            CornerRadius="4">
                        <Button x:Name="AuthenticateButton" 
                                Content="Connect to MS Graph" 
                                Style="{StaticResource SidebarButtonStyle}"
                                Background="Transparent"
                                Foreground="White"
                                Height="40"
                                Margin="0"/>
                    </Border>

                    <Button x:Name="CheckPermissionsButton" 
                            Content="Check Permissions" 
                            Style="{StaticResource SidebarButtonStyle}"
                            IsEnabled="False"
                            Margin="15,5"/>
                    <Button x:Name="InstallModulesButton" 
                            Content="Install/Update Modules"
                            Style="{StaticResource SidebarButtonStyle}"
                            Margin="15,5"/>
                    <Button x:Name="logs_button" 
                            Content="Logs"
                            Style="{StaticResource SidebarButtonStyle}"
                            Margin="15,5"/>
                    <Button x:Name="disconnect_button" 
                            Content="Disconnect"
                            Style="{StaticResource SidebarButtonStyle}"
                            IsEnabled="False"
                            Margin="15,5"/>
                </StackPanel>
                
                <!-- Navigation Menu -->
                <StackPanel Margin="0,10,0,0">
                    <RadioButton x:Name="MenuHome"
                                Content="Home"
                                Style="{StaticResource MenuButtonStyle}"
                                IsChecked="True"/>
                    <RadioButton x:Name="MenuDashboard"
                                Content="Dashboard"
                                Style="{StaticResource MenuButtonStyle}"
                                IsEnabled="False"/>
                    <RadioButton x:Name="MenuDeviceManagement"
                                Content="Device Management"
                                Style="{StaticResource MenuButtonStyle}"
                                IsEnabled="False"/>
                    <RadioButton x:Name="MenuPlaybooks"
                                Content="Playbooks"
                                Style="{StaticResource MenuButtonStyle}"
                                IsEnabled="False"/>
                </StackPanel>
            </DockPanel>
        </Border>

        <!-- Main Content Area -->
        <Grid x:Name="MainContent" Grid.Column="1" Margin="20">
            <!-- Home Page -->
            <Grid x:Name="HomePage" Visibility="Visible">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <!-- Header -->
                <TextBlock Grid.Row="0"
                          Text="Welcome to the Device Offboarding Manager"
                          FontSize="24"
                          FontWeight="SemiBold"
                          Margin="0,0,0,20"/>

                <!-- Content -->
                <StackPanel Grid.Row="1" Margin="0,20,0,0">
                    <TextBlock Text="This tool helps you manage and offboard devices from:"
                             FontSize="16"
                             Margin="0,0,0,15"/>
                    <TextBlock Text="• Microsoft Intune"
                             FontSize="14"
                             Margin="20,0,0,10"/>
                    <TextBlock Text="• Windows Autopilot"
                             FontSize="14"
                             Margin="20,0,0,10"/>
                    <TextBlock Text="• Microsoft Entra ID (formerly Azure AD)"
                             FontSize="14"
                             Margin="20,0,0,10"/>
                    
                    <TextBlock Text="Getting Started:"
                             FontSize="16"
                             FontWeight="SemiBold"
                             Margin="0,30,0,15"/>
                    <TextBlock Text="1. Connect to MS Graph using the button in the sidebar"
                             FontSize="14"
                             Margin="20,0,0,10"/>
                    <TextBlock Text="2. Check permissions to ensure you have the required access"
                             FontSize="14"
                             Margin="20,0,0,10"/>
                    <TextBlock Text="3. Navigate to Device Management to search and manage devices"
                             FontSize="14"
                             Margin="20,0,0,10"/>
                    <TextBlock Text="4. Use the Dashboard to view device statistics"
                             FontSize="14"
                             Margin="20,0,0,10"/>
                    <TextBlock Text="5. Explore Playbooks for common management tasks"
                             FontSize="14"
                             Margin="20,0,0,10"/>
                </StackPanel>
            </Grid>

            <!-- Dashboard Page -->
            <Grid x:Name="DashboardPage">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <!-- Header -->
                <TextBlock Grid.Row="0"
                          Text="Dashboard" 
                          FontSize="24"
                          FontWeight="SemiBold"
                          Margin="0,0,0,20"/>

                <!-- Top Row Statistics -->
                <UniformGrid Grid.Row="1" Rows="1" Margin="0,0,0,20">
                    <Border Background="#2D3A4F" Margin="0,0,10,0" CornerRadius="8">
                        <StackPanel Margin="20">
                            <TextBlock Text="Number of Devices in Intune" 
                                     Foreground="#8B95A5" 
                                     FontSize="14"/>
                            <TextBlock x:Name="IntuneDevicesCount"
                                     Text="0"
                                     Foreground="White"
                                     FontSize="32"
                                     FontWeight="Bold"
                                     Margin="0,10,0,5"/>
                            <TextBlock Text="Total Devices"
                                     Foreground="#8B95A5"/>
                        </StackPanel>
                    </Border>

                    <Border Background="#2D3A4F" Margin="10,0" CornerRadius="8">
                        <StackPanel Margin="20">
                            <TextBlock Text="Number of Devices in Autopilot"
                                     Foreground="#8B95A5"
                                     FontSize="14"/>
                            <TextBlock x:Name="AutopilotDevicesCount"
                                     Text="0"
                                     Foreground="White"
                                     FontSize="32"
                                     FontWeight="Bold"
                                     Margin="0,10,0,5"/>
                            <TextBlock Text="Total Devices"
                                     Foreground="#8B95A5"/>
                        </StackPanel>
                    </Border>

                    <Border Background="#2D3A4F" Margin="10,0,0,0" CornerRadius="8">
                        <StackPanel Margin="20">
                            <TextBlock Text="Number of Devices in EntraID"
                                     Foreground="#8B95A5"
                                     FontSize="14"/>
                            <TextBlock x:Name="EntraIDDevicesCount"
                                     Text="0"
                                     Foreground="White"
                                     FontSize="32"
                                     FontWeight="Bold"
                                     Margin="0,10,0,5"/>
                            <TextBlock Text="Total Devices"
                                     Foreground="#8B95A5"/>
                        </StackPanel>
                    </Border>
                </UniformGrid>

                <!-- Middle Row - Stale Devices -->
                <UniformGrid Grid.Row="2" Rows="1" Margin="0,0,0,20">
                    <Border Background="#2D3A4F" Margin="0,0,10,0" CornerRadius="8">
                        <StackPanel Margin="20">
                            <TextBlock Text="Stale Devices (30 days)"
                                     Foreground="#8B95A5"
                                     FontSize="14"/>
                            <TextBlock x:Name="StaleDevices30Count"
                                     Text="0"
                                     Foreground="White"
                                     FontSize="32"
                                     FontWeight="Bold"
                                     Margin="0,10,0,5"/>
                            <TextBlock Text="Did not Sync with Intune in the last 30 days"
                                     Foreground="#8B95A5"
                                     TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>

                    <Border Background="#2D3A4F" Margin="10,0" CornerRadius="8">
                        <StackPanel Margin="20">
                            <TextBlock Text="Stale Devices (90 days)"
                                     Foreground="#8B95A5"
                                     FontSize="14"/>
                            <TextBlock x:Name="StaleDevices90Count"
                                     Text="0"
                                     Foreground="White"
                                     FontSize="32"
                                     FontWeight="Bold"
                                     Margin="0,10,0,5"/>
                            <TextBlock Text="Did not Sync with Intune in the last 90 days"
                                     Foreground="#8B95A5"
                                     TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>

                    <Border Background="#2D3A4F" Margin="10,0,0,0" CornerRadius="8">
                        <StackPanel Margin="20">
                            <TextBlock Text="Stale Devices (180 days)"
                                     Foreground="#8B95A5"
                                     FontSize="14"/>
                            <TextBlock x:Name="StaleDevices180Count"
                                     Text="0"
                                     Foreground="White"
                                     FontSize="32"
                                     FontWeight="Bold"
                                     Margin="0,10,0,5"/>
                            <TextBlock Text="Did not Sync with Intune in the last 180 days"
                                     Foreground="#8B95A5"
                                     TextWrapping="Wrap"/>
                        </StackPanel>
                    </Border>
                </UniformGrid>

                <!-- Bottom Row - Personal/Corporate and Charts -->
                <Grid Grid.Row="3">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="*"/>
                    </Grid.ColumnDefinitions>

                    <!-- Personal/Corporate Devices -->
                    <UniformGrid Grid.Column="0" Rows="1" Margin="0,0,10,0">
                        <Border Background="#2D3A4F" Margin="0,0,10,0" CornerRadius="8">
                            <StackPanel Margin="20">
                                <TextBlock Text="Number of Personal Devices in Intune"
                                         Foreground="#8B95A5"
                                         FontSize="14"/>
                                <TextBlock x:Name="PersonalDevicesCount"
                                         Text="0"
                                         Foreground="White"
                                         FontSize="32"
                                         FontWeight="Bold"
                                         Margin="0,10,0,5"/>
                                <TextBlock Text="Total personal devices"
                                         Foreground="#8B95A5"/>
                            </StackPanel>
                        </Border>

                        <Border Background="#2D3A4F" Margin="10,0,0,0" CornerRadius="8">
                            <StackPanel Margin="20">
                                <TextBlock Text="Number of Corporate Devices in Intune"
                                         Foreground="#8B95A5"
                                         FontSize="14"/>
                                <TextBlock x:Name="CorporateDevicesCount"
                                         Text="0"
                                         Foreground="White"
                                         FontSize="32"
                                         FontWeight="Bold"
                                         Margin="0,10,0,5"/>
                                <TextBlock Text="Total corporate devices"
                                         Foreground="#8B95A5"/>
                            </StackPanel>
                        </Border>
                    </UniformGrid>

                    <!-- Charts -->
                    <Grid Grid.Column="1" Margin="10,0,0,0">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="*"/>
                            <RowDefinition Height="*"/>
                        </Grid.RowDefinitions>

                        <!-- Platform Distribution Chart -->
                        <Border Grid.Row="0" Background="#2D3A4F" CornerRadius="8" Margin="0,0,0,10">
                            <StackPanel Margin="20">
                                <TextBlock Text="Number of Devices per Platform (Intune)"
                                         Foreground="#8B95A5"
                                         FontSize="14"
                                         Margin="0,0,0,10"/>
                                <!-- Add chart here -->
                            </StackPanel>
                        </Border>

                        <!-- OS Versions Chart -->
                        <Border Grid.Row="1" Background="#2D3A4F" CornerRadius="8" Margin="0,10,0,0">
                            <StackPanel Margin="20">
                                <TextBlock Text="OS Versions (Intune)"
                                         Foreground="#8B95A5"
                                         FontSize="14"
                                         Margin="0,0,0,10"/>
                                <!-- Add chart here -->
                            </StackPanel>
                        </Border>
                    </Grid>
                </Grid>
            </Grid>

            <!-- Device Management Page -->
            <Grid x:Name="DeviceManagementPage">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Search Controls -->
                <Grid Grid.Row="1" Margin="0,0,0,10">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="150"/>
                        <ColumnDefinition Width="*"/>
                        <ColumnDefinition Width="Auto"/>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <ComboBox x:Name="dropdown" 
                              Margin="0,0,8,0"/>
                    <TextBox x:Name="SearchInputText" 
                             Grid.Column="1" 
                             Margin="0,0,8,0"
                             TextWrapping="Wrap"
                             AcceptsReturn="True"
                             VerticalScrollBarVisibility="Auto"/>
                    <Button x:Name="bulk_import_button" 
                            Grid.Column="2" 
                            Content="Bulk Import" 
                            Margin="0,0,8,0"/>
                    <Button x:Name="SearchButton" 
                            Grid.Column="3" 
                            Content="Search"/>
                </Grid>

                <!-- Results Grid -->
                <DataGrid x:Name="SearchResultsDataGrid" 
                          Grid.Row="3"
                          Margin="0,0,0,15"
                          AutoGenerateColumns="False"
                          IsReadOnly="False"
                          HeadersVisibility="Column"
                          GridLinesVisibility="All"
                          CanUserResizeRows="False"
                          CanUserReorderColumns="False"
                          SelectionMode="Extended"
                          SelectionUnit="FullRow"
                          CanUserAddRows="False">
                    <DataGrid.Columns>
                        <DataGridCheckBoxColumn Binding="{Binding IsSelected, UpdateSourceTrigger=PropertyChanged, Mode=TwoWay}" 
                                              Header="Select" 
                                              Width="50"
                                              IsReadOnly="False"/>
                        <DataGridTextColumn Binding="{Binding DeviceName}" 
                                                  Header="Device Name" 
                                                  Width="*"
                                                  IsReadOnly="True"/>
                        <DataGridTextColumn Binding="{Binding SerialNumber}" 
                                                  Header="Serial Number" 
                                                  Width="*"
                                                  IsReadOnly="True"/>
                        <DataGridTextColumn Binding="{Binding OperatingSystem}" 
                                                  Header="OS" 
                                                  Width="*"
                                                  IsReadOnly="True"/>
                        <DataGridTextColumn Binding="{Binding PrimaryUser}" 
                                                  Header="Primary User" 
                                                  Width="*"
                                                  IsReadOnly="True"/>
                        <DataGridTextColumn Binding="{Binding AzureADLastContact}" 
                                                  Header="Entra ID Last Contact" 
                                                  Width="*"
                                                  IsReadOnly="True"/>
                        <DataGridTextColumn Binding="{Binding IntuneLastContact}" 
                                                  Header="Intune Last Contact" 
                                                  Width="*"
                                                  IsReadOnly="True"/>
                        <DataGridTextColumn Binding="{Binding AutopilotLastContact}" 
                                                  Header="Autopilot Last Contact" 
                                                  Width="*"
                                                  IsReadOnly="True"/>
                    </DataGrid.Columns>
                </DataGrid>

                <!-- Status Section -->
                <StackPanel Grid.Row="4" 
                            Margin="0,0,0,15">
                    <TextBlock x:Name="intune_status" 
                              Margin="0,0,0,4" 
                              FontSize="12"/>
                    <TextBlock x:Name="autopilot_status" 
                              Margin="0,0,0,4" 
                              FontSize="12"/>
                    <TextBlock x:Name="aad_status" 
                              Margin="0,0,0,4" 
                              FontSize="12"/>
                </StackPanel>

                <!-- Bottom Section -->
                <Grid Grid.Row="5">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition Width="Auto"/>
                    </Grid.ColumnDefinitions>

                    <!-- Left Side -->
                    <Button x:Name="OffboardButton" 
                            Content="Offboard device(s)" 
                            Background="#D83B01"
                            Grid.Column="0"
                            Margin="0,0,8,0"/>
                </Grid>
            </Grid>

            <!-- Playbooks Page -->
            <Grid x:Name="PlaybooksPage" Visibility="Collapsed">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>

                <StackPanel Grid.Row="0" Margin="20">
                    <TextBlock Text="Playbooks" 
                              FontSize="24"
                              FontWeight="SemiBold"
                              Foreground="#323130"
                              Margin="0,0,0,10"/>
                    <TextBlock Text="Select a playbook to execute and view device information."
                              Opacity="0.7"
                              Margin="0,0,0,20"/>
                </StackPanel>

                <ScrollViewer Grid.Row="1" 
                             Margin="20,0,20,20" 
                             VerticalScrollBarVisibility="Auto">
                    <StackPanel>
                        <Button x:Name="PlaybookAutopilotNotIntune"
                                Style="{StaticResource PlaybookButtonStyle}"
                                Content="List all devices that are in Autopilot but not in Intune"/>
                        <Button x:Name="PlaybookIntuneNotAutopilot"
                                Style="{StaticResource PlaybookButtonStyle}"
                                Content="List all devices that are in Intune but not in Autopilot"/>
                        <Button x:Name="PlaybookCorporateDevices"
                                Style="{StaticResource PlaybookButtonStyle}"
                                Content="List all Corporate devices in Intune"/>
                        <Button x:Name="PlaybookPersonalDevices"
                                Style="{StaticResource PlaybookButtonStyle}"
                                Content="List all Personal devices in Intune"/>
                        <Button x:Name="PlaybookStaleDevices"
                                Style="{StaticResource PlaybookButtonStyle}"
                                Content="List all Stale Devices in Intune"/>
                        <Button x:Name="PlaybookSpecificOS"
                                Style="{StaticResource PlaybookButtonStyle}"
                                Content="List all devices with a specific OS in Intune"/>
                        <Button x:Name="PlaybookNotLatestOS"
                                Style="{StaticResource PlaybookButtonStyle}"
                                Content="List all devices that are not on the latest OS"/>
                        <Button x:Name="PlaybookEOLOS"
                                Style="{StaticResource PlaybookButtonStyle}"
                                Content="List all devices in Intune with a End-of-life OS Version"/>
                        <Button x:Name="PlaybookBitLocker"
                                Style="{StaticResource PlaybookButtonStyle}"
                                Content="List all Windows Devices and BitLocker Keys"/>
                        <Button x:Name="PlaybookFileVault"
                                Style="{StaticResource PlaybookButtonStyle}"
                                Content="List all macOS Devices and FileVault Keys"/>
                    </StackPanel>
                </ScrollViewer>

                <!-- Playbook Results -->
                <Grid x:Name="PlaybookResultsGrid" 
                      Visibility="Collapsed"
                      Grid.Row="1">
                    <DataGrid x:Name="PlaybookResultsDataGrid"
                             Margin="20"
                             Style="{StaticResource {x:Type DataGrid}}"/>
                </Grid>
            </Grid>
        </Grid>
    </Grid>
</Window>
"@

# Define Authentication Modal XAML
[xml]$authModalXaml = @"
<Window 
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Authentication" Height="500" Width="650"
    WindowStartupLocation="CenterScreen"
    Background="#F8F9FA">
    
    <Window.Resources>
        <!-- Radio Button Style -->
        <Style x:Key="AuthRadioButtonStyle" TargetType="RadioButton">
            <Setter Property="Margin" Value="0,8,8,8"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Foreground" Value="#2D3748"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RadioButton">
                        <Border x:Name="border" 
                                Background="{TemplateBinding Background}"
                                BorderBrush="#E2E8F0"
                                BorderThickness="1"
                                CornerRadius="6"
                                Padding="12">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="24"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Ellipse x:Name="radioOuter"
                                         Width="18" Height="18"
                                         Stroke="#CBD5E0"
                                         StrokeThickness="2"
                                         Fill="Transparent"/>
                                <Ellipse x:Name="radioInner"
                                         Width="10" Height="10"
                                         Fill="#0078D4"
                                         Opacity="0"/>
                                <ContentPresenter Grid.Column="1"
                                                Margin="12,0,0,0"
                                                VerticalAlignment="Center"/>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="border" Property="Background" Value="#F7FAFC"/>
                                <Setter TargetName="radioOuter" Property="Stroke" Value="#0078D4"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="radioInner" Property="Opacity" Value="1"/>
                                <Setter TargetName="radioOuter" Property="Stroke" Value="#0078D4"/>
                                <Setter TargetName="border" Property="BorderBrush" Value="#0078D4"/>
                                <Setter TargetName="border" Property="Background" Value="#F0F9FF"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- TextBox Style -->
        <Style x:Key="AuthTextBoxStyle" TargetType="TextBox">
            <Setter Property="Height" Value="36"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="Margin" Value="0,4"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#E2E8F0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="BorderBrush" Value="#0078D4"/>
                </Trigger>
                <Trigger Property="IsFocused" Value="True">
                    <Setter Property="BorderBrush" Value="#0078D4"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- Password Box Style -->
        <Style x:Key="AuthPasswordBoxStyle" TargetType="PasswordBox">
            <Setter Property="Height" Value="36"/>
            <Setter Property="Padding" Value="12,8"/>
            <Setter Property="Margin" Value="0,4"/>
            <Setter Property="Background" Value="White"/>
            <Setter Property="BorderBrush" Value="#E2E8F0"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
        </Style>

        <!-- Button Style -->
        <Style x:Key="AuthButtonStyle" TargetType="Button">
            <Setter Property="Height" Value="40"/>
            <Setter Property="Padding" Value="24,0"/>
            <Setter Property="Background" Value="#0078D4"/>
            <Setter Property="Foreground" Value="White"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" 
                                            VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#106EBE"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#005A9E"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- Secondary Button Style -->
        <Style x:Key="SecondaryButtonStyle" TargetType="Button">
            <Setter Property="Height" Value="40"/>
            <Setter Property="Padding" Value="24,0"/>
            <Setter Property="Background" Value="#F0F0F0"/>
            <Setter Property="Foreground" Value="#2D3748"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}"
                                CornerRadius="6"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" 
                                            VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter Property="Background" Value="#E2E2E2"/>
                            </Trigger>
                            <Trigger Property="IsPressed" Value="True">
                                <Setter Property="Background" Value="#D4D4D4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>

    <Border Background="White" 
            CornerRadius="8" 
            Margin="16">
        <DockPanel Margin="24">
            <!-- Header -->
            <StackPanel DockPanel.Dock="Top" 
                       Margin="0,0,0,24">
                <TextBlock Text="Connect to Microsoft Graph" 
                          FontSize="24" 
                          FontWeight="SemiBold" 
                          Foreground="#1A202C"/>
                <TextBlock Text="Choose your preferred authentication method to connect to Microsoft Graph API"
                          Foreground="#4A5568"
                          FontSize="14"
                          Margin="0,8,0,0"/>
            </StackPanel>

            <!-- Action Buttons -->
            <StackPanel DockPanel.Dock="Bottom" 
                       Orientation="Horizontal" 
                       HorizontalAlignment="Right"
                       Margin="0,24,0,0">
                <Button x:Name="CancelAuthButton" 
                        Content="Cancel" 
                        Style="{StaticResource SecondaryButtonStyle}"
                        Width="120" 
                        Margin="0,0,12,0"/>
                <Button x:Name="ConnectButton" 
                        Content="Connect" 
                        Style="{StaticResource AuthButtonStyle}"
                        Width="120"/>
            </StackPanel>

            <!-- Scrollable Content -->
            <ScrollViewer VerticalScrollBarVisibility="Auto"
                         HorizontalScrollBarVisibility="Disabled"
                         Padding="0,0,16,0">
                <StackPanel Margin="0,0,8,0">
                    <RadioButton x:Name="InteractiveAuth" 
                                Style="{StaticResource AuthRadioButtonStyle}"
                                Content="Interactive Login (Admin User)" 
                                IsChecked="True"/>
                    
                    <RadioButton x:Name="CertificateAuth" 
                                Style="{StaticResource AuthRadioButtonStyle}"
                                Content="App Registration with Certificate"/>
                    
                    <Grid x:Name="CertificateInputs" 
                          Margin="44,8,0,16" 
                          Visibility="Collapsed">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="100"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>

                        <TextBlock Text="App ID" 
                                  Grid.Row="0" 
                                  VerticalAlignment="Center"
                                  Foreground="#4A5568"/>
                        <TextBox x:Name="CertAppId" 
                                 Grid.Row="0" 
                                 Grid.Column="1"
                                 Style="{StaticResource AuthTextBoxStyle}"/>

                        <TextBlock Text="Tenant ID" 
                                  Grid.Row="1" 
                                  VerticalAlignment="Center"
                                  Foreground="#4A5568"/>
                        <TextBox x:Name="CertTenantId" 
                                 Grid.Row="1" 
                                 Grid.Column="1"
                                 Style="{StaticResource AuthTextBoxStyle}"/>

                        <TextBlock Text="Thumbprint" 
                                  Grid.Row="2" 
                                  VerticalAlignment="Center"
                                  Foreground="#4A5568"/>
                        <TextBox x:Name="CertThumbprint" 
                                 Grid.Row="2" 
                                 Grid.Column="1"
                                 Style="{StaticResource AuthTextBoxStyle}"/>

                        <!-- Import Button -->
                        <Button x:Name="ImportCertButton"
                                Grid.Row="3"
                                Grid.Column="1"
                                Content="Import"
                                HorizontalAlignment="Right"
                                Style="{StaticResource SecondaryButtonStyle}"
                                Height="32"
                                Width="120"
                                Margin="0,12,0,0"/>

                        <!-- Help Text -->
                        <TextBlock Grid.Row="4" 
                                  Grid.Column="0"
                                  Grid.ColumnSpan="2"
                                  Text="Import format: JSON file (.json) containing AppId, TenantId, and Thumbprint"
                                  Foreground="#718096"
                                  HorizontalAlignment="Right"
                                  FontSize="12"
                                  Margin="0,8,0,0"
                                  TextWrapping="Wrap"/>
                    </Grid>

                    <RadioButton x:Name="SecretAuth" 
                                Style="{StaticResource AuthRadioButtonStyle}"
                                Content="App Registration with Secret"/>
                    
                    <Grid x:Name="SecretInputs" 
                          Margin="44,8,0,16" 
                          Visibility="Collapsed">
                        <Grid.RowDefinitions>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                            <RowDefinition Height="Auto"/>
                        </Grid.RowDefinitions>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="100"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>

                        <TextBlock Text="App ID" 
                                  Grid.Row="0" 
                                  VerticalAlignment="Center"
                                  Foreground="#4A5568"/>
                        <TextBox x:Name="SecretAppId" 
                                 Grid.Row="0" 
                                 Grid.Column="1"
                                 Style="{StaticResource AuthTextBoxStyle}"/>

                        <TextBlock Text="Tenant ID" 
                                  Grid.Row="1" 
                                  VerticalAlignment="Center"
                                  Foreground="#4A5568"/>
                        <TextBox x:Name="SecretTenantId" 
                                 Grid.Row="1" 
                                 Grid.Column="1"
                                 Style="{StaticResource AuthTextBoxStyle}"/>

                        <TextBlock Text="Client Secret" 
                                  Grid.Row="2" 
                                  VerticalAlignment="Center"
                                  Foreground="#4A5568"/>
                        <PasswordBox x:Name="ClientSecret" 
                                    Grid.Row="2" 
                                    Grid.Column="1"
                                    Style="{StaticResource AuthPasswordBoxStyle}"/>

                        <!-- Import Button -->
                        <Button x:Name="ImportSecretButton"
                                Grid.Row="3"
                                Grid.Column="1"
                                Content="Import"
                                HorizontalAlignment="Right"
                                Style="{StaticResource SecondaryButtonStyle}"
                                Height="32"
                                Width="120"
                                Margin="0,12,0,0"/>

                        <!-- Help Text -->
                        <TextBlock Grid.Row="4" 
                                  Grid.Column="0"
                                  Grid.ColumnSpan="2"
                                  Text="Import format: JSON file (.json) containing AppId, TenantId, and ClientSecret"
                                  Foreground="#718096"
                                  HorizontalAlignment="Right"
                                  FontSize="12"
                                  Margin="0,8,0,0"
                                  TextWrapping="Wrap"/>
                    </Grid>
                </StackPanel>
            </ScrollViewer>
        </DockPanel>
    </Border>
</Window>
"@

# Define required permissions with reasons
$script:requiredPermissions = @(
    @{
        Permission = "User.Read.All"
        Reason     = "Required to read user profile information and check group memberships"
    },
    @{
        Permission = "Group.Read.All"
        Reason     = "Needed to read group information and memberships"
    },
    @{
        Permission = "DeviceManagementConfiguration.Read.All"
        Reason     = "Allows reading Intune device configuration policies and their assignments"
    },
    @{
        Permission = "DeviceManagementApps.Read.All"
        Reason     = "Necessary to read mobile app management policies and app configurations"
    },
    @{
        Permission = "DeviceManagementManagedDevices.ReadWrite.All"
        Reason     = "Required to read and modify managed device information and compliance policies"
    },
    @{
        Permission = "Device.Read.All"
        Reason     = "Needed to read device information from Entra ID"
    },
    @{
        Permission = "DeviceManagementServiceConfig.ReadWrite.All"
        Reason     = "Required for Autopilot configuration and management"
    }
)

function Show-AuthenticationDialog {
    $reader = (New-Object System.Xml.XmlNodeReader $authModalXaml)
    $authWindow = [Windows.Markup.XamlReader]::Load($reader)

    # Get controls
    $interactiveAuth = $authWindow.FindName('InteractiveAuth')
    $certificateAuth = $authWindow.FindName('CertificateAuth')
    $secretAuth = $authWindow.FindName('SecretAuth')
    $certificateInputs = $authWindow.FindName('CertificateInputs')
    $secretInputs = $authWindow.FindName('SecretInputs')
    $connectButton = $authWindow.FindName('ConnectButton')
    $cancelAuthButton = $authWindow.FindName('CancelAuthButton')
    $importCertButton = $authWindow.FindName('ImportCertButton')
    $importSecretButton = $authWindow.FindName('ImportSecretButton')

    # Add event handlers for radio buttons
    $certificateAuth.Add_Checked({
            $certificateInputs.Visibility = 'Visible'
            $secretInputs.Visibility = 'Collapsed'
        })

    $secretAuth.Add_Checked({
            $secretInputs.Visibility = 'Visible'
            $certificateInputs.Visibility = 'Collapsed'
        })

    $interactiveAuth.Add_Checked({
            $certificateInputs.Visibility = 'Collapsed'
            $secretInputs.Visibility = 'Collapsed'
        })

    # Add import button handlers
    $importCertButton.Add_Click({
            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.Filter = "JSON files (*.json)|*.json"
            $OpenFileDialog.Title = "Import Certificate Configuration"
        
            if ($OpenFileDialog.ShowDialog() -eq 'OK') {
                try {
                    $config = Get-Content $OpenFileDialog.FileName | ConvertFrom-Json
                
                    if ($config.AppId -and $config.TenantId -and $config.Thumbprint) {
                        $authWindow.FindName('CertAppId').Text = $config.AppId
                        $authWindow.FindName('CertTenantId').Text = $config.TenantId
                        $authWindow.FindName('CertThumbprint').Text = $config.Thumbprint
                    }
                    else {
                        [System.Windows.MessageBox]::Show(
                            "Invalid configuration file. Please ensure it contains AppId, TenantId, and Thumbprint.",
                            "Invalid Configuration",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                    }
                }
                catch {
                    [System.Windows.MessageBox]::Show(
                        "Error reading configuration file: $_",
                        "Error",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Error
                    )
                }
            }
        })

    $importSecretButton.Add_Click({
            $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $OpenFileDialog.Filter = "JSON files (*.json)|*.json"
            $OpenFileDialog.Title = "Import Secret Configuration"
        
            if ($OpenFileDialog.ShowDialog() -eq 'OK') {
                try {
                    $config = Get-Content $OpenFileDialog.FileName | ConvertFrom-Json
                
                    if ($config.AppId -and $config.TenantId -and $config.ClientSecret) {
                        $authWindow.FindName('SecretAppId').Text = $config.AppId
                        $authWindow.FindName('SecretTenantId').Text = $config.TenantId
                        $authWindow.FindName('ClientSecret').Password = $config.ClientSecret
                    }
                    else {
                        [System.Windows.MessageBox]::Show(
                            "Invalid configuration file. Please ensure it contains AppId, TenantId, and ClientSecret.",
                            "Invalid Configuration",
                            [System.Windows.MessageBoxButton]::OK,
                            [System.Windows.MessageBoxImage]::Warning
                        )
                    }
                }
                catch {
                    [System.Windows.MessageBox]::Show(
                        "Error reading configuration file: $_",
                        "Error",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Error
                    )
                }
            }
        })

    # Add event handlers for buttons
    $cancelAuthButton.Add_Click({
            $script:authCancelled = $true
            $authWindow.DialogResult = $false
            $authWindow.Close()
        })

    $connectButton.Add_Click({
            $script:authCancelled = $false
            $authWindow.DialogResult = $true
            $authWindow.Close()
        })

    # Show dialog and return result
    $result = $authWindow.ShowDialog()
    
    if ($result) {
        # Return authentication details based on selected method
        if ($interactiveAuth.IsChecked) {
            return @{
                Method = 'Interactive'
            }
        }
        elseif ($certificateAuth.IsChecked) {
            return @{
                Method     = 'Certificate'
                AppId      = $authWindow.FindName('CertAppId').Text
                TenantId   = $authWindow.FindName('CertTenantId').Text
                Thumbprint = $authWindow.FindName('CertThumbprint').Text
            }
        }
        else {
            return @{
                Method   = 'Secret'
                AppId    = $authWindow.FindName('SecretAppId').Text
                TenantId = $authWindow.FindName('SecretTenantId').Text
                Secret   = $authWindow.FindName('ClientSecret').Password
            }
        }
    }
    return $null
}

function Connect-ToGraph {
    param (
        [Parameter(Mandatory = $true)]
        [hashtable]$AuthDetails
    )

    try {
        Write-Log "Attempting to connect to Microsoft Graph using $($AuthDetails.Method) authentication..."
        
        # Get required permissions
        $permissionsList = ($script:requiredPermissions | ForEach-Object { $_.Permission })

        # Connect based on authentication method
        switch ($AuthDetails.Method) {
            'Interactive' {
                $connectionResult = Connect-MgGraph -Scopes $permissionsList -NoWelcome -ErrorAction Stop
            }
            'Certificate' {
                if ([string]::IsNullOrWhiteSpace($AuthDetails.AppId) -or
                    [string]::IsNullOrWhiteSpace($AuthDetails.TenantId) -or
                    [string]::IsNullOrWhiteSpace($AuthDetails.Thumbprint)) {
                    throw "Certificate authentication requires App ID, Tenant ID, and Certificate Thumbprint"
                }
                
                # Disconnect any existing connections first
                Disconnect-MgGraph -ErrorAction SilentlyContinue
                
                $connectionResult = Connect-MgGraph -ClientId $AuthDetails.AppId -TenantId $AuthDetails.TenantId -CertificateThumbprint $AuthDetails.Thumbprint -NoWelcome -ErrorAction Stop
                
                # Verify connection was successful
                $context = Get-MgContext
                if (-not $context) {
                    throw "Failed to establish connection with certificate authentication"
                }
            }
            'Secret' {
                if ([string]::IsNullOrWhiteSpace($AuthDetails.AppId) -or
                    [string]::IsNullOrWhiteSpace($AuthDetails.TenantId) -or
                    [string]::IsNullOrWhiteSpace($AuthDetails.Secret)) {
                    throw "Secret authentication requires App ID, Tenant ID, and Client Secret"
                }
                
                # Create client secret credential
                $SecuredPasswordPassword = ConvertTo-SecureString -String $AuthDetails.Secret -AsPlainText -Force
                
                $ClientSecretCredential = New-Object -TypeName System.Management.Automation.PSCredential `
                    -ArgumentList $AuthDetails.AppId, $SecuredPasswordPassword
                
                # Get required permissions
                $permissionsList = ($script:requiredPermissions | ForEach-Object { $_.Permission })
                
                $connectionResult = Connect-MgGraph -TenantId $AuthDetails.TenantId `
                    -ClientSecretCredential $ClientSecretCredential `
                    -NoWelcome -ErrorAction Stop
            }
            default {
                throw "Invalid authentication method specified"
            }
        }

        # Check permissions
        $context = Get-MgContext
        if (-not $context) {
            throw "Failed to get Microsoft Graph context after connection"
        }

        $currentPermissions = $context.Scopes
        $missingPermissions = @()

        foreach ($permissionInfo in $script:requiredPermissions) {
            $permission = $permissionInfo.Permission
            if (-not ($currentPermissions -contains $permission -or
                    $currentPermissions -contains $permission.Replace(".Read", ".ReadWrite"))) {
                $missingPermissions += $permission
            }
        }

        if ($missingPermissions.Count -gt 0) {
            $missingList = $missingPermissions -join ", "
            Write-Log "Warning: Missing permissions: $missingList"
            [System.Windows.MessageBox]::Show(
                "The following permissions are missing: `n$missingList`n`nThe application may not function correctly.",
                "Missing Permissions",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
        }

        Write-Log "Successfully connected to Microsoft Graph"
        
        # Close any open authentication dialogs
        $authWindow = [System.Windows.Application]::Current.Windows | Where-Object { $_.Title -eq "Authentication" }
        if ($authWindow) {
            $authWindow.Close()
        }
        
        return $true
    }
    catch {
        Write-Log "Failed to connect to Microsoft Graph: $_"
        [System.Windows.MessageBox]::Show(
            "Failed to connect to Microsoft Graph: $_",
            "Connection Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
        return $false
    }
}

# Parse XAML
$reader = (New-Object System.Xml.XmlNodeReader $xaml)
$Window = [Windows.Markup.XamlReader]::Load($reader)

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
$logs_button = $Window.FindName('logs_button')
$CheckPermissionsButton = $Window.FindName('CheckPermissionsButton')

$SearchInputText.Add_GotFocus({
        # Empty - no resizing needed
    })

$SearchInputText.Add_LostFocus({
        # Empty - no resizing needed
    })
    
$Window.Add_Loaded({
        $Dropdown.Items.Add("Devicename")
        $Dropdown.Items.Add("Serialnumber")
        $Dropdown.SelectedIndex = 0
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
                $CheckPermissionsButton.IsEnabled = $false
                
                # Disable navigation menus
                $MenuDashboard.IsEnabled = $false
                $MenuDeviceManagement.IsEnabled = $false
                $MenuPlaybooks.IsEnabled = $false
                
                # Force Home menu selection
                $MenuHome.IsChecked = $true
            }
            else {
                Write-Log "Successfully connected to MS Graph"
                $AuthenticateButton.Content = "Successfully connected"
                $AuthenticateButton.IsEnabled = $false
                $Disconnect.IsEnabled = $true
                $CheckPermissionsButton.IsEnabled = $true
                
                # Enable navigation menus
                $MenuDashboard.IsEnabled = $true
                $MenuDeviceManagement.IsEnabled = $true
                $MenuPlaybooks.IsEnabled = $true
                
                # Update dashboard statistics for existing connection
                Update-DashboardStatistics
                
                # Verify permissions for existing connection
                $currentPermissions = $context.Scopes
                $missingPermissions = @()
                
                foreach ($permissionInfo in $script:requiredPermissions) {
                    $permission = $permissionInfo.Permission
                    if (-not ($currentPermissions -contains $permission -or
                            $currentPermissions -contains $permission.Replace(".Read", ".ReadWrite"))) {
                        $missingPermissions += $permission
                    }
                }
                
                if ($missingPermissions.Count -gt 0) {
                    $missingList = $missingPermissions -join ", "
                    Write-Log "Warning: Missing permissions for existing connection: $missingList"
                    [System.Windows.MessageBox]::Show(
                        "The following permissions are missing: `n$missingList`n`nThe application may not function correctly.",
                        "Missing Permissions",
                        [System.Windows.MessageBoxButton]::OK,
                        [System.Windows.MessageBoxImage]::Warning
                    )
                }
            }
        }
        catch {
            Write-Log "Error occurred during window load: $_"
            $AuthenticateButton.Content = "Not Connected to MS Graph"
            $AuthenticateButton.IsEnabled = $true
            $Disconnect.IsEnabled = $false
            $CheckPermissionsButton.IsEnabled = $false
            
            # Disable navigation menus
            $MenuDashboard.IsEnabled = $false
            $MenuDeviceManagement.IsEnabled = $false
            $MenuPlaybooks.IsEnabled = $false
        }
    })
    
$Disconnect.Add_Click({
        try {
            Write-Log "Attempting to disconnect from MS Graph..."
            
            # Disconnect from Graph
            Disconnect-MgGraph -ErrorAction Stop
            
            # Reset UI state
            $Disconnect.Content = "Disconnected"
            $Disconnect.IsEnabled = $false
            $AuthenticateButton.Content = "Connect to MS Graph"
            $AuthenticateButton.IsEnabled = $true
            $CheckPermissionsButton.IsEnabled = $false
            
            # Disable navigation menus and force Home selection
            $MenuDashboard.IsEnabled = $false
            $MenuDeviceManagement.IsEnabled = $false
            $MenuPlaybooks.IsEnabled = $false
            $MenuHome.IsChecked = $true
            
            # Clear any sensitive data from the dashboard
            $Window.FindName('IntuneDevicesCount').Text = "0"
            $Window.FindName('AutopilotDevicesCount').Text = "0"
            $Window.FindName('EntraIDDevicesCount').Text = "0"
            $Window.FindName('StaleDevices30Count').Text = "0"
            $Window.FindName('StaleDevices90Count').Text = "0"
            $Window.FindName('StaleDevices180Count').Text = "0"
            $Window.FindName('PersonalDevicesCount').Text = "0"
            $Window.FindName('CorporateDevicesCount').Text = "0"
            
            Write-Log "Successfully disconnected from MS Graph"
        }
        catch {
            Write-Log "Error occurred while attempting to disconnect from MS Graph: $_"
            [System.Windows.MessageBox]::Show(
                "Error disconnecting from Microsoft Graph: $_",
                "Disconnect Error",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Error
            )
        }
    })
    
$AuthenticateButton.Add_Click({
        try {
            # Check if already connected
            $context = Get-MgContext
            if ($context) {
                Write-Log "Already connected to MS Graph, skipping authentication dialog"
                return
            }
            
            Write-Log "Authentication button clicked, showing authentication dialog..."
        
            # Show authentication dialog
            $authDetails = Show-AuthenticationDialog
            if (-not $authDetails) {
                Write-Log "Authentication cancelled by user"
                return
            }

            # Attempt to connect
            $connected = Connect-ToGraph -AuthDetails $authDetails
            if ($connected) {
                Write-Log "Authentication Successful"
                $AuthenticateButton.Content = "Authentication Successful"
                $AuthenticateButton.IsEnabled = $false
                $Disconnect.Content = "Disconnect"
                $Disconnect.IsEnabled = $true
                $CheckPermissionsButton.IsEnabled = $true

                # Enable navigation menus
                $MenuDashboard.IsEnabled = $true
                $MenuDeviceManagement.IsEnabled = $true
                $MenuPlaybooks.IsEnabled = $true

                # Update dashboard statistics after successful authentication
                Update-DashboardStatistics
            }
            else {
                Write-Log "Authentication Failed"
                $AuthenticateButton.Content = "Authentication Failed"
                $AuthenticateButton.IsEnabled = $true
                $Disconnect.Content = "Disconnected"
                $Disconnect.IsEnabled = $false
                $CheckPermissionsButton.IsEnabled = $false
                
                # Disable navigation menus
                $MenuDashboard.IsEnabled = $false
                $MenuDeviceManagement.IsEnabled = $false
                $MenuPlaybooks.IsEnabled = $false
            }
        }
        catch {
            Write-Log "Error occurred during authentication. Exception: $_"
            $AuthenticateButton.Content = "Authentication Failed"
            $AuthenticateButton.IsEnabled = $true
            $Disconnect.Content = "Disconnected"
            $Disconnect.IsEnabled = $false
            $CheckPermissionsButton.IsEnabled = $false
            
            # Disable navigation menus
            $MenuDashboard.IsEnabled = $false
            $MenuDeviceManagement.IsEnabled = $false
            $MenuPlaybooks.IsEnabled = $false
        }
    })
    
$CheckPermissionsButton.Add_Click({
        try {
            Write-Log "Check Permissions button clicked, attempting to retrieve user context..."

            $context = Get-MGContext

            if ($context) {
                $username = $context.Account

                # Check if the required scopes are present
                $requiredScopes = @("Device.Read.All", "DeviceManagementManagedDevices.ReadWrite.All", "DeviceManagementServiceConfig.ReadWrite.All")
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

$Window.Add_Loaded({
        try {
            Write-Log "Window is loading..."
      
            $modules = @(
                "Microsoft.Graph.Authentication"
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
            $SearchTexts = $SearchInputText.Text -split ', ' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            Write-Log "Searching for devices: $SearchTexts"
            $searchOption = $Dropdown.SelectedItem
    
            $searchResults = New-Object 'System.Collections.Generic.List[DeviceObject]'
            $AADCount = 0
            $IntuneCount = 0
            $AutopilotCount = 0
    
            foreach ($SearchText in $SearchTexts) {
                if ($searchOption -eq "Devicename") {
                    # Get Entra ID Device
                    $uri = "https://graph.microsoft.com/v1.0/devices?`$filter=displayName eq '$SearchText'"
                    $AADDevices = Get-GraphPagedResults -Uri $uri
                    
                    # Get Intune Device
                    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=deviceName eq '$SearchText'"
                    $IntuneDevices = Get-GraphPagedResults -Uri $uri

                    if ($AADDevices -and $IntuneDevices) {
                        foreach ($AADDevice in $AADDevices) {
                            $matchingIntuneDevice = $IntuneDevices | Where-Object { $_.deviceName -eq $AADDevice.displayName } | Select-Object -First 1
                            if ($matchingIntuneDevice) {
                                # Get Autopilot Device
                                $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,'$($matchingIntuneDevice.serialNumber)')"
                                $AutopilotDevice = (Get-GraphPagedResults -Uri $uri) | Select-Object -First 1

                                $CombinedDevice = New-Object DeviceObject
                                $CombinedDevice.IsSelected = $false
                                $CombinedDevice.DeviceName = $matchingIntuneDevice.deviceName
                                $CombinedDevice.SerialNumber = $matchingIntuneDevice.serialNumber
                                $CombinedDevice.OperatingSystem = $AADDevice.operatingSystem
                                $CombinedDevice.PrimaryUser = $matchingIntuneDevice.userDisplayName
                                $CombinedDevice.AzureADLastContact = $AADDevice.approximateLastSignInDateTime
                                $CombinedDevice.IntuneLastContact = $matchingIntuneDevice.lastSyncDateTime
                                $CombinedDevice.AutopilotLastContact = $AutopilotDevice.lastContactedDateTime
                                
                                $searchResults.Add($CombinedDevice)
                                $AADCount++
                                $IntuneCount++
                                if ($AutopilotDevice) { $AutopilotCount++ }
                            }
                        }
                    }
                }
                elseif ($searchOption -eq "Serialnumber") {
                    # Get Intune Device
                    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=serialNumber eq '$SearchText'"
                    $IntuneDevices = Get-GraphPagedResults -Uri $uri
                    
                    if ($IntuneDevices) {
                        foreach ($IntuneDevice in $IntuneDevices) {
                            $displayName = $IntuneDevice.deviceName
                            # Get Entra ID Device
                            $uri = "https://graph.microsoft.com/v1.0/devices?`$filter=displayName eq '$displayName'"
                            $AADDevice = (Get-GraphPagedResults -Uri $uri) | Select-Object -First 1

                            if ($AADDevice) {
                                # Get Autopilot Device
                                $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,'$SearchText')"
                                $AutopilotDevice = (Get-GraphPagedResults -Uri $uri) | Select-Object -First 1
                    
                                $CombinedDevice = New-Object DeviceObject
                                $CombinedDevice.IsSelected = $false
                                $CombinedDevice.DeviceName = $IntuneDevice.deviceName
                                $CombinedDevice.SerialNumber = $IntuneDevice.serialNumber
                                $CombinedDevice.OperatingSystem = $AADDevice.operatingSystem
                                $CombinedDevice.PrimaryUser = $IntuneDevice.userDisplayName
                                $CombinedDevice.AzureADLastContact = $AADDevice.approximateLastSignInDateTime
                                $CombinedDevice.IntuneLastContact = $IntuneDevice.lastSyncDateTime
                                $CombinedDevice.AutopilotLastContact = $AutopilotDevice.lastContactedDateTime
                    
                                $searchResults.Add($CombinedDevice)
                                $AADCount++
                                $IntuneCount++
                                if ($AutopilotDevice) { $AutopilotCount++ }
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
    
            if ($searchResults.Count -gt 0) {
                $SearchResultsDataGrid.ItemsSource = $searchResults
            }
            else {
                $SearchResultsDataGrid.ItemsSource = $null
                [System.Windows.MessageBox]::Show("No devices found matching the search criteria.")
            }
            
            # Ensure Offboard button is disabled until selection
            $OffboardButton.IsEnabled = $false
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

        $selectedDevices = $SearchResultsDataGrid.ItemsSource | Where-Object { $_.IsSelected }
        
        if (-not $selectedDevices) {
            [System.Windows.MessageBox]::Show("Please select at least one device to offboard.")
            return
        }

        $confirmationResult = [System.Windows.MessageBox]::Show("Are you sure you want to proceed with offboarding the selected device(s)? This action cannot be undone.", "Confirm Offboarding", [System.Windows.MessageBoxButton]::YesNo)
        if ($confirmationResult -eq 'No') {
            Write-Log "User canceled offboarding operation."
            return
        }

        try {
            foreach ($device in $selectedDevices) {
                $deviceName = $device.DeviceName
                # Get Entra ID Device
                $uri = "https://graph.microsoft.com/v1.0/devices?`$filter=displayName eq '$deviceName'"
                $AADDevice = (Invoke-MgGraphRequest -Uri $uri -Method GET).value | Select-Object -First 1

                # Get Intune Device
                $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices?`$filter=deviceName eq '$deviceName'"
                $IntuneDevice = (Invoke-MgGraphRequest -Uri $uri -Method GET).value | Select-Object -First 1

                # Get Autopilot Device
                if ($IntuneDevice) {
                    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities?`$filter=contains(serialNumber,'$($IntuneDevice.serialNumber)')"
                    $AutopilotDevice = (Invoke-MgGraphRequest -Uri $uri -Method GET).value | Select-Object -First 1
                }

                if ($AADDevice) {
                    $uri = "https://graph.microsoft.com/v1.0/devices/$($AADDevice.id)"
                    Invoke-MgGraphRequest -Uri $uri -Method DELETE
                    [System.Windows.MessageBox]::Show("Successfully removed device $deviceName from AzureAD.")
                    $Window.FindName('aad_status').Text = "AzureAD: Unavailable"
                    Write-Log "Successfully removed device $deviceName from Entra ID."
                }
                else {
                    [System.Windows.MessageBox]::Show("Device $deviceName not found in AzureAD.")
                }

                if ($IntuneDevice) {
                    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices/$($IntuneDevice.id)"
                    Invoke-MgGraphRequest -Uri $uri -Method DELETE
                    [System.Windows.MessageBox]::Show("Successfully removed device $deviceName from Intune.")
                    $Window.FindName('intune_status').Text = "Intune: Unavailable"
                    Write-Log "Successfully removed device $deviceName from Intune."
                }
                else {
                    [System.Windows.MessageBox]::Show("Device $deviceName not found in Intune.")
                }

                if ($AutopilotDevice) {
                    $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities/$($AutopilotDevice.id)"
                    Invoke-MgGraphRequest -Uri $uri -Method DELETE
                    [System.Windows.MessageBox]::Show("Successfully removed device $deviceName from Autopilot.")
                    $Window.FindName('autopilot_status').Text = "Autopilot: Unavailable"
                    Write-Log "Successfully removed device $deviceName from Autopilot."
                }
                else {
                    [System.Windows.MessageBox]::Show("Device $deviceName not found in Autopilot.")
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
        
# Add new control connections
$MenuHome = $Window.FindName('MenuHome')
$MenuDashboard = $Window.FindName('MenuDashboard')
$MenuDeviceManagement = $Window.FindName('MenuDeviceManagement')
$MenuPlaybooks = $Window.FindName('MenuPlaybooks')
$HomePage = $Window.FindName('HomePage')
$DashboardPage = $Window.FindName('DashboardPage')
$DeviceManagementPage = $Window.FindName('DeviceManagementPage')
$PlaybooksPage = $Window.FindName('PlaybooksPage')
$PlaybookResultsGrid = $Window.FindName('PlaybookResultsGrid')
$PlaybookResultsDataGrid = $Window.FindName('PlaybookResultsDataGrid')

# Set initial page visibility
$Window.Add_Loaded({
        # Set initial page visibility
        $HomePage.Visibility = 'Visible'
        $DashboardPage.Visibility = 'Collapsed'
        $DeviceManagementPage.Visibility = 'Collapsed'
        $PlaybooksPage.Visibility = 'Collapsed'
        $PlaybookResultsGrid.Visibility = 'Collapsed'

        # Update dashboard statistics if connected
        if (-not $AuthenticateButton.IsEnabled) {
            Update-DashboardStatistics
        }
    })

# Add menu switching functionality
$MenuHome.Add_Checked({
        $HomePage.Visibility = 'Visible'
        $DashboardPage.Visibility = 'Collapsed'
        $DeviceManagementPage.Visibility = 'Collapsed'
        $PlaybooksPage.Visibility = 'Collapsed'
        $PlaybookResultsGrid.Visibility = 'Collapsed'
    })

$MenuDashboard.Add_Checked({
        $HomePage.Visibility = 'Collapsed'
        $DashboardPage.Visibility = 'Visible'
        $DeviceManagementPage.Visibility = 'Collapsed'
        $PlaybooksPage.Visibility = 'Collapsed'
        $PlaybookResultsGrid.Visibility = 'Collapsed'
        
        # Update dashboard statistics if connected
        if (-not $AuthenticateButton.IsEnabled) {
            Update-DashboardStatistics
        }
    })

$MenuDeviceManagement.Add_Checked({
        $HomePage.Visibility = 'Collapsed'
        $DashboardPage.Visibility = 'Collapsed'
        $DeviceManagementPage.Visibility = 'Visible'
        $PlaybooksPage.Visibility = 'Collapsed'
        $PlaybookResultsGrid.Visibility = 'Collapsed'
    })

$MenuPlaybooks.Add_Checked({
        $HomePage.Visibility = 'Collapsed'
        $DashboardPage.Visibility = 'Collapsed'
        $DeviceManagementPage.Visibility = 'Collapsed'
        $PlaybooksPage.Visibility = 'Visible'
        $PlaybookResultsGrid.Visibility = 'Collapsed'
    })

function Update-DashboardStatistics {
    try {
        Write-Log "Updating dashboard statistics..."

        # Get all managed devices
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/managedDevices"
        $intuneDevices = Get-GraphPagedResults -Uri $uri

        # Get all Autopilot devices
        $uri = "https://graph.microsoft.com/v1.0/deviceManagement/windowsAutopilotDeviceIdentities"
        $autopilotDevices = Get-GraphPagedResults -Uri $uri

        # Get all EntraID devices
        $uri = "https://graph.microsoft.com/v1.0/devices"
        $entraDevices = Get-GraphPagedResults -Uri $uri

        # Update top row counts
        $Window.FindName('IntuneDevicesCount').Text = $intuneDevices.Count
        $Window.FindName('AutopilotDevicesCount').Text = $autopilotDevices.Count
        $Window.FindName('EntraIDDevicesCount').Text = $entraDevices.Count

        # Calculate stale devices
        $thirtyDaysAgo = (Get-Date).AddDays(-30).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $ninetyDaysAgo = (Get-Date).AddDays(-90).ToString('yyyy-MM-ddTHH:mm:ssZ')
        $oneEightyDaysAgo = (Get-Date).AddDays(-180).ToString('yyyy-MM-ddTHH:mm:ssZ')

        $stale30 = ($intuneDevices | Where-Object { $_.lastSyncDateTime -lt $thirtyDaysAgo }).Count
        $stale90 = ($intuneDevices | Where-Object { $_.lastSyncDateTime -lt $ninetyDaysAgo }).Count
        $stale180 = ($intuneDevices | Where-Object { $_.lastSyncDateTime -lt $oneEightyDaysAgo }).Count

        $Window.FindName('StaleDevices30Count').Text = $stale30
        $Window.FindName('StaleDevices90Count').Text = $stale90
        $Window.FindName('StaleDevices180Count').Text = $stale180

        # Update personal/corporate counts
        $personalDevices = ($intuneDevices | Where-Object { $_.managementState -eq 'userEnrollment' }).Count
        $corporateDevices = ($intuneDevices | Where-Object { $_.managementState -eq 'managed' }).Count

        $Window.FindName('PersonalDevicesCount').Text = $personalDevices
        $Window.FindName('CorporateDevicesCount').Text = $corporateDevices

        Write-Log "Dashboard statistics updated successfully."
    }
    catch {
        Write-Log "Error updating dashboard statistics: $_"
        [System.Windows.MessageBox]::Show("Error updating dashboard statistics. Please ensure you are connected to MS Graph.")
    }
}

# Connect playbook buttons
$PlaybookButtons = @(
    $Window.FindName('PlaybookAutopilotNotIntune'),
    $Window.FindName('PlaybookIntuneNotAutopilot'),
    $Window.FindName('PlaybookCorporateDevices'),
    $Window.FindName('PlaybookPersonalDevices'),
    $Window.FindName('PlaybookStaleDevices'),
    $Window.FindName('PlaybookSpecificOS'),
    $Window.FindName('PlaybookNotLatestOS'),
    $Window.FindName('PlaybookEOLOS'),
    $Window.FindName('PlaybookBitLocker'),
    $Window.FindName('PlaybookFileVault')
)

# Add click handlers for playbook buttons
foreach ($button in $PlaybookButtons) {
    $button.Add_Click({
            # We'll implement the playbook functionality in the next step
            Write-Log "Playbook clicked: $($this.Content)"
        })
}

# Results Grid
$SearchResultsDataGrid = $Window.FindName('SearchResultsDataGrid')
$OffboardButton = $Window.FindName('OffboardButton')

# Initially disable the Offboard button
$OffboardButton.IsEnabled = $false

# Add selection changed event handler for the DataGrid
$SearchResultsDataGrid.Add_SelectionChanged({
        # Update the Offboard button state based on selected devices
        $selectedDevices = $SearchResultsDataGrid.ItemsSource | Where-Object { $_.IsSelected }
        $OffboardButton.IsEnabled = ($null -ne $selectedDevices -and $selectedDevices.Count -gt 0)
    })

# Add handler for checkbox selection changes
$SearchResultsDataGrid.Add_LoadingRow({
        param($sender, $e)
        $row = $e.Row
        $dataContext = $row.DataContext
        if ($dataContext -and $dataContext.GetType().Name -eq 'DeviceObject') {
            $dataContext.add_PropertyChanged({
                    param($s, $ev)
                    if ($ev.PropertyName -eq 'IsSelected') {
                        # Update Offboard button state
                        $selectedDevices = $SearchResultsDataGrid.ItemsSource | Where-Object { $_.IsSelected }
                        $OffboardButton.IsEnabled = ($null -ne $selectedDevices -and $selectedDevices.Count -gt 0)
                    }
                })
        }
    })

# Update dashboard when switching to Dashboard tab
$MenuDashboard.Add_Checked({
        $HomePage.Visibility = 'Collapsed'
        $DashboardPage.Visibility = 'Visible'
        $DeviceManagementPage.Visibility = 'Collapsed'
        $PlaybooksPage.Visibility = 'Collapsed'
        $PlaybookResultsGrid.Visibility = 'Collapsed'
        
        # Update dashboard statistics if connected
        if (-not $AuthenticateButton.IsEnabled) {
            Update-DashboardStatistics
        }
    })

# Show Window
$Window.ShowDialog() | Out-Null