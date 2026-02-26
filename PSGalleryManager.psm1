function Start-PSGalleryManager {
    <#
    .SYNOPSIS
        Launch PSGalleryManager.
    .DESCRIPTION
        Opens a WPF-based graphical tool for managing PowerShell modules.
        Features: Search and install from PSGallery, per-module scope detection,
        update checking, module details, bulk updates, batch operations, 
        uninstall, and JSON cache for fast startup.

        Run as Administrator to manage AllUsers (system-wide) modules.
    .EXAMPLE
        Start-PSGalleryManager
        Launches the GUI module manager.
    .NOTES
        Author  : Nick Geoffroy
        Cache stored in: %APPDATA%\PSGalleryManager\
    #>
    [CmdletBinding()]
    param()

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms

# Ensure PowerShellGet is loaded before using its commands
Import-Module PowerShellGet -ErrorAction SilentlyContinue

# Theme colours
$accent       = "#0078D4"
$accentHover  = "#106EBE"
$accentLight  = "#E8F1FB"
$bgDark       = "#1E1E1E"
$bgPanel      = "#252526"
$bgCard       = "#2D2D2D"
$bgCardHover  = "#333333"
$fgPrimary    = "#CCCCCC"
$fgSecondary  = "#999999"
$fgBright     = "#FFFFFF"
$green        = "#16C60C"
$orange       = "#F9A825"
$red          = "#E74856"
$border       = "#3E3E3E"
$purple       = "#B388FF"

# Cache configuration – separate cache per PS major version (5 vs 7 have different module paths)
$script:cacheDir  = Join-Path $env:APPDATA "PSGalleryManager\PS$($PSVersionTable.PSVersion.Major)"
$script:cacheInstalled = Join-Path $script:cacheDir "installed.json"
$script:cacheUpdates   = Join-Path $script:cacheDir "updates.json"

if (-not (Test-Path $script:cacheDir)) {
    New-Item -Path $script:cacheDir -ItemType Directory -Force | Out-Null
}

[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="PSGalleryManager"
    Height="780" Width="1100" MinHeight="600" MinWidth="850"
    WindowStartupLocation="CenterScreen"
    Background="$bgDark" Foreground="$fgPrimary"
    FontFamily="Segoe UI" FontSize="13">

    <Window.Resources>
        <Style TargetType="ScrollViewer">
            <Setter Property="VerticalScrollBarVisibility" Value="Auto"/>
        </Style>
        <Style x:Key="BtnBase" TargetType="Button">
            <Setter Property="Padding" Value="14,7"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="FontSize" Value="12"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="bd" Background="{TemplateBinding Background}"
                                CornerRadius="4" Padding="{TemplateBinding Padding}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                BorderBrush="{TemplateBinding BorderBrush}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="bd" Property="Opacity" Value="0.85"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter Property="Opacity" Value="0.4"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="BtnAccent" TargetType="Button" BasedOn="{StaticResource BtnBase}">
            <Setter Property="Background" Value="$accent"/>
            <Setter Property="Foreground" Value="White"/>
        </Style>
        <Style x:Key="BtnDanger" TargetType="Button" BasedOn="{StaticResource BtnBase}">
            <Setter Property="Background" Value="$red"/>
            <Setter Property="Foreground" Value="White"/>
        </Style>
        <Style x:Key="BtnGhost" TargetType="Button" BasedOn="{StaticResource BtnBase}">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="$fgPrimary"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="BorderBrush" Value="$border"/>
        </Style>
        <Style x:Key="BtnSuccess" TargetType="Button" BasedOn="{StaticResource BtnBase}">
            <Setter Property="Background" Value="$green"/>
            <Setter Property="Foreground" Value="White"/>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="$bgCard"/>
            <Setter Property="Foreground" Value="$fgBright"/>
            <Setter Property="BorderBrush" Value="$border"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="8,6"/>
            <Setter Property="CaretBrush" Value="$fgBright"/>
            <Setter Property="SelectionBrush" Value="$accent"/>
        </Style>
    </Window.Resources>

    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <!-- HEADER -->
        <Border Grid.Row="0" Background="$bgPanel" BorderBrush="$border" BorderThickness="0,0,0,1" Padding="20,14">
            <Grid>
                <Grid.ColumnDefinitions>
                    <ColumnDefinition Width="Auto"/>
                    <ColumnDefinition Width="*"/>
                </Grid.ColumnDefinitions>
                <StackPanel Grid.Column="0" Orientation="Horizontal" VerticalAlignment="Center">
                    <TextBlock Text="PS" FontSize="18" FontWeight="Bold" Foreground="$accent" Margin="0,0,0,0" VerticalAlignment="Center"/>
                    <TextBlock Text="GalleryManager" FontSize="18" FontWeight="Bold" Foreground="$fgBright" VerticalAlignment="Center"/>
                    <TextBlock Text="v1.0.5" FontSize="11" Foreground="$fgSecondary" Margin="10,4,0,0" VerticalAlignment="Center"/>
                </StackPanel>
                <Border Grid.Column="1" Margin="30,0" MaxWidth="550" CornerRadius="4"
                        Background="$bgCard" BorderBrush="$border" BorderThickness="1">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="Auto"/>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>
                        <TextBlock Grid.Column="0" Text=" &gt; " FontSize="14" Margin="8,0,2,0"
                                   VerticalAlignment="Center" Foreground="$fgSecondary"/>
                        <TextBox x:Name="txtSearch" Grid.Column="1" BorderThickness="0" Background="Transparent"
                                 VerticalAlignment="Center" FontSize="13"/>
                        <Button x:Name="btnSearch" Grid.Column="2" Content="Search Gallery"
                                Style="{StaticResource BtnAccent}" Margin="4" Padding="12,5"/>
                    </Grid>
                </Border>
            </Grid>
        </Border>

        <!-- MAIN CONTENT -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="200"/>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition x:Name="detailColumn" Width="0"/>
            </Grid.ColumnDefinitions>

            <!-- LEFT NAV -->
            <Border Grid.Column="0" Background="$bgPanel" BorderBrush="$border" BorderThickness="0,0,1,0">
                <StackPanel Margin="0,8">
                    <Button x:Name="navInstalled" Content="  Installed Modules" Style="{StaticResource BtnBase}"
                            Background="$accent" Foreground="White" HorizontalContentAlignment="Left"
                            Padding="16,10" Margin="6,2" FontWeight="SemiBold"/>
                    <Button x:Name="navUpdates" Style="{StaticResource BtnBase}"
                            Background="Transparent" Foreground="$fgPrimary" HorizontalContentAlignment="Left"
                            Padding="16,10" Margin="6,2">
                        <StackPanel Orientation="Horizontal">
                            <TextBlock Text="  Updates Available"/>
                            <Border x:Name="badgeUpdates" Background="$orange" CornerRadius="8"
                                    Padding="6,1" Margin="8,0,0,0" Visibility="Collapsed">
                                <TextBlock x:Name="txtBadge" Text="0" FontSize="11" Foreground="$bgDark" FontWeight="Bold"/>
                            </Border>
                        </StackPanel>
                    </Button>
                    <Button x:Name="navSearch" Content="  Gallery Search" Style="{StaticResource BtnBase}"
                            Background="Transparent" Foreground="$fgPrimary" HorizontalContentAlignment="Left"
                            Padding="16,10" Margin="6,2"/>
                    <Separator Background="$border" Margin="16,10"/>
                    <Button x:Name="btnRefresh" Content="  Refresh (live scan)" Style="{StaticResource BtnGhost}"
                            Margin="12,4" Padding="12,8"/>
                    <Button x:Name="btnClearCache" Content="  Clear Cache" Style="{StaticResource BtnGhost}"
                            Margin="12,4" Padding="12,8"/>
                    <Button x:Name="btnUpdateAll" Content="  Update All" Style="{StaticResource BtnSuccess}"
                            Margin="12,4" Padding="12,8" Visibility="Collapsed"/>
                    <Separator Background="$border" Margin="16,10"/>
                    <TextBlock x:Name="txtCacheAge" Text="" Foreground="$fgSecondary" FontSize="10.5"
                               Margin="16,2" TextWrapping="Wrap"/>
                </StackPanel>
            </Border>

            <!-- CENTER: MODULE LIST -->
            <Grid Grid.Column="1">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                </Grid.RowDefinitions>
                <Border Grid.Row="0" Padding="16,10" Background="$bgDark" BorderBrush="$border" BorderThickness="0,0,0,1">
                    <Grid>
                        <TextBlock x:Name="txtViewTitle" Text="Installed Modules" FontSize="15"
                                   FontWeight="SemiBold" Foreground="$fgBright" VerticalAlignment="Center"/>
                        <TextBlock x:Name="txtCount" HorizontalAlignment="Right" Foreground="$fgSecondary"
                                   VerticalAlignment="Center" FontSize="12"/>
                    </Grid>
                </Border>
                <!-- BATCH ACTION BAR -->
                <Border x:Name="pnlBatchBar" Grid.Row="1" Padding="12,8" Background="#1A0078D4"
                        BorderBrush="$accent" BorderThickness="0,0,0,1" Visibility="Collapsed">
                    <Grid>
                        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                            <TextBlock x:Name="txtSelected" Text="0 selected" Foreground="$fgBright"
                                       FontWeight="SemiBold" FontSize="12" VerticalAlignment="Center" Margin="4,0,12,0"/>
                            <Button x:Name="btnBatchUpdate" Content="Update Selected" Style="{StaticResource BtnSuccess}"
                                    Padding="10,5" Margin="0,0,8,0" Visibility="Collapsed"/>
                            <Button x:Name="btnBatchUninstall" Content="Uninstall Selected" Style="{StaticResource BtnDanger}"
                                    Padding="10,5" Margin="0,0,8,0" Visibility="Collapsed"/>
                        </StackPanel>
                        <Button x:Name="btnClearSelection" Content="Clear Selection" Style="{StaticResource BtnGhost}"
                                HorizontalAlignment="Right" Padding="10,5"/>
                    </Grid>
                </Border>
                <ListBox x:Name="lstModules" Grid.Row="2" Background="Transparent" BorderThickness="0"
                         ScrollViewer.HorizontalScrollBarVisibility="Disabled"
                         SelectionMode="Extended"
                         HorizontalContentAlignment="Stretch" Padding="8">
                    <ListBox.ItemContainerStyle>
                        <Style TargetType="ListBoxItem">
                            <Setter Property="Padding" Value="0"/>
                            <Setter Property="Margin" Value="0,2"/>
                            <Setter Property="HorizontalContentAlignment" Value="Stretch"/>
                            <Setter Property="Background" Value="Transparent"/>
                            <Setter Property="Template">
                                <Setter.Value>
                                    <ControlTemplate TargetType="ListBoxItem">
                                        <Border x:Name="Bd" Background="$bgCard" CornerRadius="6"
                                                Padding="14,10" BorderBrush="$border" BorderThickness="1"
                                                Margin="4,2">
                                            <ContentPresenter/>
                                        </Border>
                                        <ControlTemplate.Triggers>
                                            <Trigger Property="IsMouseOver" Value="True">
                                                <Setter TargetName="Bd" Property="Background" Value="$bgCardHover"/>
                                                <Setter TargetName="Bd" Property="BorderBrush" Value="$accent"/>
                                            </Trigger>
                                            <Trigger Property="IsSelected" Value="True">
                                                <Setter TargetName="Bd" Property="Background" Value="#1A0078D4"/>
                                                <Setter TargetName="Bd" Property="BorderBrush" Value="$accent"/>
                                            </Trigger>
                                        </ControlTemplate.Triggers>
                                    </ControlTemplate>
                                </Setter.Value>
                            </Setter>
                        </Style>
                    </ListBox.ItemContainerStyle>
                </ListBox>
                <Border x:Name="pnlLoading" Grid.Row="2" Background="#CC1E1E1E" Visibility="Collapsed">
                    <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                        <TextBlock x:Name="txtLoading" Text="Loading..." FontSize="16" Foreground="$fgBright"
                                   HorizontalAlignment="Center"/>
                        <ProgressBar x:Name="progBar" IsIndeterminate="True" Width="260" Height="4"
                                     Margin="0,12,0,0" Foreground="$accent" Background="$bgCard"/>
                    </StackPanel>
                </Border>
                <Border x:Name="pnlEmpty" Grid.Row="2" Visibility="Collapsed">
                    <StackPanel VerticalAlignment="Center" HorizontalAlignment="Center">
                        <TextBlock Text="(empty)" FontSize="28" HorizontalAlignment="Center" Foreground="$fgSecondary"/>
                        <TextBlock x:Name="txtEmpty" Text="No modules found" FontSize="15"
                                   Foreground="$fgSecondary" Margin="0,8,0,0" HorizontalAlignment="Center"/>
                    </StackPanel>
                </Border>
            </Grid>

            <!-- RIGHT: DETAIL PANEL -->
            <Border Grid.Column="2" Background="$bgPanel" BorderBrush="$border" BorderThickness="1,0,0,0">
                <ScrollViewer Padding="18,14">
                    <StackPanel x:Name="pnlDetail">
                        <Grid>
                            <TextBlock x:Name="txtDetailName" FontSize="18" FontWeight="Bold" Foreground="$fgBright"
                                       TextWrapping="Wrap"/>
                            <Button x:Name="btnCloseDetail" Content="X" HorizontalAlignment="Right"
                                    Style="{StaticResource BtnBase}" Background="Transparent"
                                    Foreground="$fgSecondary" Padding="6,2" FontSize="14"/>
                        </Grid>
                        <TextBlock x:Name="txtDetailVersion" Foreground="$fgSecondary" Margin="0,4,0,0"/>
                        <TextBlock x:Name="txtDetailAuthor" Foreground="$fgSecondary" Margin="0,2,0,0"/>
                        <TextBlock x:Name="txtDetailScope" Foreground="$purple" Margin="0,2,0,0" FontWeight="SemiBold"/>
                        <Separator Background="$border" Margin="0,12"/>
                        <TextBlock Text="Description" FontWeight="SemiBold" Foreground="$fgBright"/>
                        <TextBlock x:Name="txtDetailDesc" TextWrapping="Wrap" Foreground="$fgPrimary"
                                   Margin="0,6,0,0" LineHeight="20"/>
                        <Separator Background="$border" Margin="0,12"/>
                        <Grid x:Name="gridInfo" Margin="0,0,0,8">
                            <Grid.ColumnDefinitions>
                                <ColumnDefinition Width="110"/>
                                <ColumnDefinition Width="*"/>
                            </Grid.ColumnDefinitions>
                            <Grid.RowDefinitions>
                                <RowDefinition/><RowDefinition/><RowDefinition/>
                                <RowDefinition/><RowDefinition/><RowDefinition/>
                            </Grid.RowDefinitions>
                            <TextBlock Grid.Row="0" Grid.Column="0" Text="Published" Foreground="$fgSecondary" Margin="0,3"/>
                            <TextBlock x:Name="txtDetailDate" Grid.Row="0" Grid.Column="1" Foreground="$fgPrimary" Margin="0,3"/>
                            <TextBlock Grid.Row="1" Grid.Column="0" Text="Downloads" Foreground="$fgSecondary" Margin="0,3"/>
                            <TextBlock x:Name="txtDetailDownloads" Grid.Row="1" Grid.Column="1" Foreground="$fgPrimary" Margin="0,3"/>
                            <TextBlock Grid.Row="2" Grid.Column="0" Text="License" Foreground="$fgSecondary" Margin="0,3"/>
                            <TextBlock x:Name="txtDetailLicense" Grid.Row="2" Grid.Column="1" Foreground="$fgPrimary" Margin="0,3"/>
                            <TextBlock Grid.Row="3" Grid.Column="0" Text="Project URL" Foreground="$fgSecondary" Margin="0,3"/>
                            <TextBlock x:Name="txtDetailUrl" Grid.Row="3" Grid.Column="1" Foreground="$accent" Margin="0,3"
                                       Cursor="Hand" TextWrapping="Wrap" TextDecorations="Underline"/>
                            <TextBlock Grid.Row="4" Grid.Column="0" Text="Tags" Foreground="$fgSecondary" Margin="0,3"/>
                            <TextBlock x:Name="txtDetailTags" Grid.Row="4" Grid.Column="1" Foreground="$fgPrimary"
                                       TextWrapping="Wrap" Margin="0,3"/>
                            <TextBlock Grid.Row="5" Grid.Column="0" Text="Dependencies" Foreground="$fgSecondary" Margin="0,3"/>
                            <TextBlock x:Name="txtDetailDeps" Grid.Row="5" Grid.Column="1" Foreground="$fgPrimary"
                                       TextWrapping="Wrap" Margin="0,3"/>
                        </Grid>
                        <Separator Background="$border" Margin="0,6"/>
                        <WrapPanel x:Name="pnlActions" Margin="0,10,0,0">
                            <Button x:Name="btnDetailInstall" Content="Install" Style="{StaticResource BtnAccent}"
                                    Margin="0,0,8,8" Visibility="Collapsed"/>
                            <Button x:Name="btnDetailUpdate" Content="Update" Style="{StaticResource BtnSuccess}"
                                    Margin="0,0,8,8" Visibility="Collapsed"/>
                            <Button x:Name="btnDetailUninstall" Content="Uninstall" Style="{StaticResource BtnDanger}"
                                    Margin="0,0,8,8" Visibility="Collapsed"/>
                            <Button x:Name="btnOpenGallery" Content="PS Gallery" Style="{StaticResource BtnGhost}"
                                    Margin="0,0,8,8"/>
                            <Button x:Name="btnOpenLocation" Content="Open Location" Style="{StaticResource BtnGhost}"
                                    Margin="0,0,8,8" Visibility="Collapsed"/>
                        </WrapPanel>
                    </StackPanel>
                </ScrollViewer>
            </Border>
        </Grid>

        <!-- STATUS BAR -->
        <Border Grid.Row="2" Background="$bgPanel" BorderBrush="$border" BorderThickness="0,1,0,0" Padding="14,6">
            <Grid>
                <TextBlock x:Name="txtStatus" Text="Ready" Foreground="$fgSecondary" FontSize="11.5"/>
                <StackPanel Orientation="Horizontal" HorizontalAlignment="Right">
                    <TextBlock x:Name="txtCacheStatus" Text="" Foreground="$fgSecondary" FontSize="11.5" Margin="0,0,16,0"/>
                    <TextBlock x:Name="txtIsAdmin" Foreground="$fgSecondary" FontSize="11.5"/>
                </StackPanel>
            </Grid>
        </Border>
    </Grid>
</Window>
"@

# Build Window
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

$ui = @{}
$xaml.SelectNodes('//*[@*[contains(translate(name(),"X","x"),"x:Name")]]') | ForEach-Object {
    $ui[$_.Name] = $window.FindName($_.Name)
}

# State
$script:currentView     = 'installed'
$script:installedCache  = @()
$script:updatesCache    = @()
$script:searchResults   = @()
$script:selectedModule  = $null

# Admin detection
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if ($isAdmin) {
    $ui.txtIsAdmin.Text = "Administrator"
}
else {
    $ui.txtIsAdmin.Text = "Standard User"
    $ui.txtIsAdmin.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($orange)
}

# Install scope: admin = AllUsers, standard = CurrentUser
$script:installScope = if ($isAdmin) { 'AllUsers' } else { 'CurrentUser' }

# Check if -AcceptLicense is supported (PowerShellGet 1.6+)
$script:supportsAcceptLicense = $false
try {
    $installCmd = Get-Command Install-Module -ErrorAction SilentlyContinue
    if ($installCmd.Parameters.ContainsKey('AcceptLicense')) {
        $script:supportsAcceptLicense = $true
    }
}
catch {}

# -- Helpers --

function Set-Status {
    param([string]$msg)
    $ui.txtStatus.Text = $msg
    [System.Windows.Forms.Application]::DoEvents()
}

function Show-Loading {
    param([string]$msg = "Loading...")
    $ui.txtLoading.Text = $msg
    $ui.pnlLoading.Visibility = 'Visible'
    $ui.pnlEmpty.Visibility   = 'Collapsed'
    [System.Windows.Forms.Application]::DoEvents()
}

function Hide-Loading {
    $ui.pnlLoading.Visibility = 'Collapsed'
    [System.Windows.Forms.Application]::DoEvents()
}

function Format-Downloads {
    param([long]$n)
    if ($n -ge 1000000) { return "{0:N1}M" -f ($n / 1000000) }
    elseif ($n -ge 1000) { return "{0:N1}K" -f ($n / 1000) }
    else { return $n.ToString() }
}

function Format-CacheAge {
    param([datetime]$cachedAt)
    $span = (Get-Date) - $cachedAt
    if ($span.TotalMinutes -lt 1) { return "just now" }
    if ($span.TotalMinutes -lt 60) { return [math]::Floor($span.TotalMinutes).ToString() + "m ago" }
    if ($span.TotalHours -lt 24) { return [math]::Floor($span.TotalHours).ToString() + "h ago" }
    return [math]::Floor($span.TotalDays).ToString() + "d ago"
}

function Update-CacheAgeDisplay {
    $lines = @()
    if (Test-Path $script:cacheInstalled) {
        try {
            $data = Get-Content $script:cacheInstalled -Raw | ConvertFrom-Json
            $age = Format-CacheAge -cachedAt ([datetime]$data.CachedAt)
            $lines += "Installed: " + $age
        }
        catch {}
    }
    if (Test-Path $script:cacheUpdates) {
        try {
            $data = Get-Content $script:cacheUpdates -Raw | ConvertFrom-Json
            $age = Format-CacheAge -cachedAt ([datetime]$data.CachedAt)
            $lines += "Updates: " + $age
        }
        catch {}
    }
    if ($lines.Count -gt 0) {
        $ui.txtCacheAge.Text = "Cache:`n" + ($lines -join "`n")
        $ui.txtCacheStatus.Text = "Cached"
    }
    else {
        $ui.txtCacheAge.Text = "Cache: none"
        $ui.txtCacheStatus.Text = "No cache"
    }
}

# Detect module scope from install path
function Get-ModuleScope {
    param([string]$installedLocation)
    if ([string]::IsNullOrWhiteSpace($installedLocation)) { return "Unknown" }
    $userProfile = $env:USERPROFILE
    if ($installedLocation.StartsWith($userProfile, [System.StringComparison]::OrdinalIgnoreCase)) {
        return "CurrentUser"
    }
    return "AllUsers"
}

function Set-NavActive {
    param([string]$name)
    $brushTransparent = [System.Windows.Media.BrushConverter]::new().ConvertFromString('Transparent')
    $brushFg          = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgPrimary)
    $brushAccent      = [System.Windows.Media.BrushConverter]::new().ConvertFromString($accent)
    $brushBright      = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgBright)

    foreach ($btn in @('navInstalled','navUpdates','navSearch')) {
        $ui[$btn].Background = $brushTransparent
        $ui[$btn].Foreground = $brushFg
        $ui[$btn].FontWeight = 'Normal'
    }

    $map = @{ 'installed' = 'navInstalled'; 'updates' = 'navUpdates'; 'search' = 'navSearch' }
    $active = $map[$name]
    if ($active -and $ui.ContainsKey($active)) {
        $ui[$active].Background = $brushAccent
        $ui[$active].Foreground = $brushBright
        $ui[$active].FontWeight = 'SemiBold'
    }
}

# -- JSON Cache Functions --

function Save-CacheInstalled {
    param([array]$modules)
    try {
        $export = @()
        foreach ($m in $modules) {
            $export += @{
                Name           = [string]$m.Name
                Version        = [string]$m.Version
                Author         = [string]$m.Author
                Description    = [string]$m._Description
                PublishedDate  = [string]$m._PublishedDate
                License        = [string]$m._License
                ProjectUri     = [string]$m._ProjectUri
                Tags           = [string]$m._Tags
                Dependencies   = [string]$m._Dependencies
                Downloads      = [long]$m._Downloads
                Status         = [string]$m._Status
                LatestVersion  = [string]$m._LatestVersion
                Scope          = [string]$m._Scope
                InstalledLocation = [string]$m._InstalledLocation
            }
        }
        $wrapper = @{
            CachedAt = (Get-Date).ToString("o")
            Modules  = $export
        }
        $wrapper | ConvertTo-Json -Depth 4 | Set-Content -Path $script:cacheInstalled -Encoding UTF8 -Force
    }
    catch {}
}

function Save-CacheUpdates {
    param([array]$modules)
    try {
        $export = @()
        foreach ($m in $modules) {
            $export += @{
                Name              = [string]$m.Name
                Version           = [string]$m.Version
                Author            = [string]$m.Author
                Description       = [string]$m._Description
                PublishedDate     = [string]$m._PublishedDate
                License           = [string]$m._License
                ProjectUri        = [string]$m._ProjectUri
                Tags              = [string]$m._Tags
                Dependencies      = [string]$m._Dependencies
                Downloads         = [long]$m._Downloads
                Status            = [string]$m._Status
                LatestVersion     = [string]$m._LatestVersion
                Scope             = [string]$m._Scope
                InstalledLocation = [string]$m._InstalledLocation
            }
        }
        $wrapper = @{
            CachedAt = (Get-Date).ToString("o")
            Modules  = $export
        }
        $wrapper | ConvertTo-Json -Depth 4 | Set-Content -Path $script:cacheUpdates -Encoding UTF8 -Force
    }
    catch {}
}

function Load-CacheFile {
    param([string]$path)
    if (-not (Test-Path $path)) { return $null }
    try {
        $raw = Get-Content $path -Raw | ConvertFrom-Json
        $modules = @()
        foreach ($entry in $raw.Modules) {
            $obj = New-Object PSObject
            $obj | Add-Member -NotePropertyName 'Name'           -NotePropertyValue $entry.Name
            $obj | Add-Member -NotePropertyName 'Version'        -NotePropertyValue $entry.Version
            $obj | Add-Member -NotePropertyName 'Author'         -NotePropertyValue $entry.Author
            $obj | Add-Member -NotePropertyName 'Description'    -NotePropertyValue $entry.Description
            $obj | Add-Member -NotePropertyName '_Status'        -NotePropertyValue $entry.Status
            $obj | Add-Member -NotePropertyName '_Description'   -NotePropertyValue $entry.Description
            $obj | Add-Member -NotePropertyName '_Downloads'     -NotePropertyValue ([long]$entry.Downloads)
            $obj | Add-Member -NotePropertyName '_PublishedDate' -NotePropertyValue $entry.PublishedDate
            $obj | Add-Member -NotePropertyName '_License'       -NotePropertyValue $entry.License
            $obj | Add-Member -NotePropertyName '_ProjectUri'    -NotePropertyValue $entry.ProjectUri
            $obj | Add-Member -NotePropertyName '_Tags'          -NotePropertyValue $entry.Tags
            $obj | Add-Member -NotePropertyName '_Dependencies'  -NotePropertyValue $entry.Dependencies
            $obj | Add-Member -NotePropertyName '_LatestVersion' -NotePropertyValue $entry.LatestVersion
            $obj | Add-Member -NotePropertyName '_Scope'             -NotePropertyValue $entry.Scope
            $obj | Add-Member -NotePropertyName '_InstalledLocation' -NotePropertyValue $entry.InstalledLocation
            $modules += $obj
        }
        return @{
            CachedAt = [datetime]$raw.CachedAt
            Modules  = $modules
        }
    }
    catch {
        return $null
    }
}

function Clear-AllCache {
    if (Test-Path $script:cacheInstalled) { Remove-Item $script:cacheInstalled -Force }
    if (Test-Path $script:cacheUpdates)   { Remove-Item $script:cacheUpdates -Force }
    Update-CacheAgeDisplay
}

# -- Build module card --

function New-ModuleCard {
    param($mod)

    $brushBright  = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgBright)
    $brushSec     = [System.Windows.Media.BrushConverter]::new().ConvertFromString($fgSecondary)
    $brushAccent  = [System.Windows.Media.BrushConverter]::new().ConvertFromString($accent)
    $brushGreen   = [System.Windows.Media.BrushConverter]::new().ConvertFromString($green)
    $brushOrange  = [System.Windows.Media.BrushConverter]::new().ConvertFromString($orange)
    $brushTransp  = [System.Windows.Media.BrushConverter]::new().ConvertFromString('Transparent')
    $brushPurple  = [System.Windows.Media.BrushConverter]::new().ConvertFromString($purple)

    $grid = New-Object System.Windows.Controls.Grid
    $grid.Tag = $mod

    $col1 = New-Object System.Windows.Controls.ColumnDefinition
    $col1.Width = [System.Windows.GridLength]::new(1, 'Star')
    $col2 = New-Object System.Windows.Controls.ColumnDefinition
    $col2.Width = [System.Windows.GridLength]::Auto
    $grid.ColumnDefinitions.Add($col1)
    $grid.ColumnDefinitions.Add($col2)

    # Left side
    $left = New-Object System.Windows.Controls.StackPanel
    [System.Windows.Controls.Grid]::SetColumn($left, 0)

    $nameRow = New-Object System.Windows.Controls.StackPanel
    $nameRow.Orientation = 'Horizontal'

    $txtName = New-Object System.Windows.Controls.TextBlock
    $txtName.Text = $mod.Name
    $txtName.FontWeight = 'SemiBold'
    $txtName.FontSize = 14
    $txtName.Foreground = $brushBright
    $nameRow.Children.Add($txtName) | Out-Null

    $txtVer = New-Object System.Windows.Controls.TextBlock
    $txtVer.Text = "  v" + $mod.Version
    $txtVer.Foreground = $brushSec
    $txtVer.FontSize = 12
    $txtVer.VerticalAlignment = 'Center'
    $nameRow.Children.Add($txtVer) | Out-Null

    # Scope badge (for installed modules)
    if ($mod._Scope -and $mod._Scope -ne 'NotInstalled') {
        $scopeBadge = New-Object System.Windows.Controls.Border
        $scopeBadge.CornerRadius = [System.Windows.CornerRadius]::new(4)
        $scopeBadge.Padding = [System.Windows.Thickness]::new(5, 1, 5, 1)
        $scopeBadge.Margin  = [System.Windows.Thickness]::new(6, 0, 0, 0)
        $scopeBadge.VerticalAlignment = 'Center'
        $scopeBadge.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#1AB388FF')
        $scopeTxt = New-Object System.Windows.Controls.TextBlock
        $scopeTxt.FontSize = 10
        $scopeTxt.Foreground = $brushPurple
        $scopeTxt.FontWeight = 'SemiBold'
        if ($mod._Scope -eq 'AllUsers') {
            $scopeTxt.Text = 'AllUsers'
        }
        else {
            $scopeTxt.Text = 'User'
        }
        $scopeBadge.Child = $scopeTxt
        $nameRow.Children.Add($scopeBadge) | Out-Null
    }

    # Status badge
    if ($mod._Status) {
        $badge = New-Object System.Windows.Controls.Border
        $badge.CornerRadius = [System.Windows.CornerRadius]::new(4)
        $badge.Padding = [System.Windows.Thickness]::new(6, 1, 6, 1)
        $badge.Margin  = [System.Windows.Thickness]::new(6, 0, 0, 0)
        $badge.VerticalAlignment = 'Center'
        $badgeTxt = New-Object System.Windows.Controls.TextBlock
        $badgeTxt.FontSize = 10.5
        $badgeTxt.FontWeight = 'SemiBold'

        if ($mod._Status -eq 'Installed') {
            $badge.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#1A16C60C')
            $badgeTxt.Foreground = $brushGreen
            $badgeTxt.Text = 'Installed'
        }
        elseif ($mod._Status -eq 'Update') {
            $badge.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#1AF9A825')
            $badgeTxt.Foreground = $brushOrange
            $badgeTxt.Text = "Update -> " + $mod._LatestVersion
        }
        elseif ($mod._Status -eq 'NotInstalled') {
            $badge.Background = [System.Windows.Media.BrushConverter]::new().ConvertFromString('#1A0078D4')
            $badgeTxt.Foreground = $brushAccent
            $badgeTxt.Text = 'Available'
        }
        $badge.Child = $badgeTxt
        $nameRow.Children.Add($badge) | Out-Null
    }

    $left.Children.Add($nameRow) | Out-Null

    # Description
    $txtDesc = New-Object System.Windows.Controls.TextBlock
    $descText = ""
    if ($mod.Description) { $descText = $mod.Description }
    elseif ($mod._Description) { $descText = $mod._Description }
    if ($descText.Length -gt 140) { $descText = $descText.Substring(0,137) + "..." }
    $txtDesc.Text = $descText
    $txtDesc.Foreground = $brushSec
    $txtDesc.FontSize = 12
    $txtDesc.Margin = [System.Windows.Thickness]::new(0, 4, 20, 0)
    $txtDesc.TextWrapping = 'Wrap'
    $left.Children.Add($txtDesc) | Out-Null

    # Author + downloads
    $metaRow = New-Object System.Windows.Controls.StackPanel
    $metaRow.Orientation = 'Horizontal'
    $metaRow.Margin = [System.Windows.Thickness]::new(0, 4, 0, 0)
    if ($mod.Author) {
        $txtAuth = New-Object System.Windows.Controls.TextBlock
        $txtAuth.Text = "by " + $mod.Author
        $txtAuth.Foreground = $brushSec
        $txtAuth.FontSize = 11
        $txtAuth.Margin = [System.Windows.Thickness]::new(0, 0, 12, 0)
        $metaRow.Children.Add($txtAuth) | Out-Null
    }
    if ($mod._Downloads -gt 0) {
        $txtDl = New-Object System.Windows.Controls.TextBlock
        $txtDl.Text = "Downloads: " + (Format-Downloads $mod._Downloads)
        $txtDl.Foreground = $brushSec
        $txtDl.FontSize = 11
        $metaRow.Children.Add($txtDl) | Out-Null
    }
    $left.Children.Add($metaRow) | Out-Null
    $grid.Children.Add($left) | Out-Null

    # Right side: quick action
    $right = New-Object System.Windows.Controls.StackPanel
    $right.VerticalAlignment = 'Center'
    [System.Windows.Controls.Grid]::SetColumn($right, 1)

    $btnQuick = New-Object System.Windows.Controls.Button
    $btnQuick.Padding   = [System.Windows.Thickness]::new(10, 5, 10, 5)
    $btnQuick.FontSize  = 11.5
    $btnQuick.Cursor    = [System.Windows.Input.Cursors]::Hand
    $btnQuick.Tag       = $mod

    if ($mod._Status -eq 'Update') {
        $btnQuick.Content    = 'Update'
        $btnQuick.Background = $brushGreen
        $btnQuick.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('White')
        $btnQuick.Add_Click({
            $m = $this.Tag
            Invoke-ModuleAction -Action 'Update' -Module $m
        })
    }
    elseif ($mod._Status -eq 'NotInstalled') {
        $btnQuick.Content    = 'Install'
        $btnQuick.Background = $brushAccent
        $btnQuick.Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString('White')
        $btnQuick.Add_Click({
            $m = $this.Tag
            Invoke-ModuleAction -Action 'Install' -Module $m
        })
    }
    else {
        $btnQuick.Content    = 'Details >'
        $btnQuick.Background = $brushTransp
        $btnQuick.Foreground = $brushAccent
        $btnQuick.Add_Click({
            $m = $this.Tag
            Show-Detail -mod $m
        })
    }

    $right.Children.Add($btnQuick) | Out-Null
    $grid.Children.Add($right) | Out-Null
    return $grid
}

# Populate list
function Update-ModuleList {
    param([array]$modules, [string]$emptyMsg = "No modules found")
    $ui.lstModules.Items.Clear()
    $ui.pnlBatchBar.Visibility = 'Collapsed'
    if ($modules.Count -eq 0) {
        $ui.pnlEmpty.Visibility = 'Visible'
        $ui.txtEmpty.Text = $emptyMsg
        $ui.txtCount.Text = "0 modules"
    }
    else {
        $ui.pnlEmpty.Visibility = 'Collapsed'
        $suffix = ""
        if ($modules.Count -ne 1) { $suffix = "s" }
        $ui.txtCount.Text = "$($modules.Count) module$suffix"
        foreach ($mod in $modules) {
            $card = New-ModuleCard -mod $mod
            $ui.lstModules.Items.Add($card) | Out-Null
        }
    }
    [System.Windows.Forms.Application]::DoEvents()
}

# Detail panel
function Show-Detail {
    param($mod)
    $script:selectedModule = $mod
    $ui.detailColumn.Width = [System.Windows.GridLength]::new(320)

    $ui.txtDetailName.Text    = $mod.Name
    $ui.txtDetailVersion.Text = "Version: " + $mod.Version
    if ($mod.Author) { $ui.txtDetailAuthor.Text = "Author: " + $mod.Author }
    else { $ui.txtDetailAuthor.Text = "" }

    # Scope display
    if ($mod._Scope -eq 'AllUsers') {
        $ui.txtDetailScope.Text = "Scope: AllUsers (system-wide)"
    }
    elseif ($mod._Scope -eq 'CurrentUser') {
        $ui.txtDetailScope.Text = "Scope: CurrentUser"
    }
    else {
        $ui.txtDetailScope.Text = ""
    }

    $desc = ""
    if ($mod.Description) { $desc = $mod.Description }
    elseif ($mod._Description) { $desc = $mod._Description }
    else { $desc = "No description available." }
    $ui.txtDetailDesc.Text = $desc

    if ($mod._PublishedDate) { $ui.txtDetailDate.Text = $mod._PublishedDate } else { $ui.txtDetailDate.Text = "-" }
    if ($mod._Downloads -gt 0) { $ui.txtDetailDownloads.Text = Format-Downloads $mod._Downloads } else { $ui.txtDetailDownloads.Text = "-" }
    if ($mod._License) { $ui.txtDetailLicense.Text = $mod._License } else { $ui.txtDetailLicense.Text = "-" }
    if ($mod._ProjectUri) { $ui.txtDetailUrl.Text = $mod._ProjectUri } else { $ui.txtDetailUrl.Text = "-" }
    if ($mod._Tags) { $ui.txtDetailTags.Text = $mod._Tags } else { $ui.txtDetailTags.Text = "-" }
    if ($mod._Dependencies) { $ui.txtDetailDeps.Text = $mod._Dependencies } else { $ui.txtDetailDeps.Text = "None" }

    if ($mod._Status -eq 'NotInstalled') { $ui.btnDetailInstall.Visibility = 'Visible' } else { $ui.btnDetailInstall.Visibility = 'Collapsed' }
    if ($mod._Status -eq 'Update') { $ui.btnDetailUpdate.Visibility = 'Visible' } else { $ui.btnDetailUpdate.Visibility = 'Collapsed' }
    if ($mod._Status -eq 'Installed' -or $mod._Status -eq 'Update') { $ui.btnDetailUninstall.Visibility = 'Visible' } else { $ui.btnDetailUninstall.Visibility = 'Collapsed' }
    if ($mod._InstalledLocation -and $mod._InstalledLocation -ne '') { $ui.btnOpenLocation.Visibility = 'Visible' } else { $ui.btnOpenLocation.Visibility = 'Collapsed' }
}

function Hide-Detail {
    $script:selectedModule = $null
    $ui.detailColumn.Width = [System.Windows.GridLength]::new(0)
}

# Enriches a live module object with custom properties (including scope detection)
function Add-ModuleProps {
    param($m, [string]$Status, [long]$Downloads = 0, [string]$LatestVersion = "")

    $pubDate = ""
    try { $pubDate = $m.PublishedDate.ToString('yyyy-MM-dd') } catch {}
    $tags = ""
    try { $tags = ($m.Tags -join ', ') } catch {}
    $deps = ""
    try { $deps = ($m.Dependencies.Name -join ', ') } catch {}

    # Detect scope from install path
    $detectedScope = "Unknown"
    $installPath = ""
    try {
        if ($m.InstalledLocation) {
            $installPath = $m.InstalledLocation
            $detectedScope = Get-ModuleScope -installedLocation $installPath
        }
    }
    catch {}

    $m | Add-Member -NotePropertyName '_Status'            -NotePropertyValue $Status         -Force -PassThru |
         Add-Member -NotePropertyName '_Description'       -NotePropertyValue $m.Description  -Force -PassThru |
         Add-Member -NotePropertyName '_Downloads'         -NotePropertyValue $Downloads      -Force -PassThru |
         Add-Member -NotePropertyName '_PublishedDate'     -NotePropertyValue $pubDate         -Force -PassThru |
         Add-Member -NotePropertyName '_License'           -NotePropertyValue $m.LicenseUri   -Force -PassThru |
         Add-Member -NotePropertyName '_ProjectUri'        -NotePropertyValue $m.ProjectUri   -Force -PassThru |
         Add-Member -NotePropertyName '_Tags'              -NotePropertyValue $tags            -Force -PassThru |
         Add-Member -NotePropertyName '_Dependencies'      -NotePropertyValue $deps            -Force -PassThru |
         Add-Member -NotePropertyName '_LatestVersion'     -NotePropertyValue $LatestVersion  -Force -PassThru |
         Add-Member -NotePropertyName '_Scope'             -NotePropertyValue $detectedScope  -Force -PassThru |
         Add-Member -NotePropertyName '_InstalledLocation' -NotePropertyValue $installPath    -Force -PassThru
}

# -- Data Loaders (with cache support) --

function Load-InstalledModules {
    param([switch]$Force)

    $script:currentView = 'installed'
    Set-NavActive -name 'installed'
    $ui.txtViewTitle.Text = 'Installed Modules'
    Hide-Detail

    # Try cache first (unless forced)
    if (-not $Force) {
        $cached = Load-CacheFile -path $script:cacheInstalled
        if ($cached) {
            $script:installedCache = @($cached.Modules)
            $age = Format-CacheAge -cachedAt $cached.CachedAt
            Update-ModuleList -modules $script:installedCache -emptyMsg "No modules installed via PowerShellGet."
            Set-Status -msg ("Loaded " + $script:installedCache.Count.ToString() + " module(s) from cache (" + $age + ")")
            Update-CacheAgeDisplay
            return
        }
    }

    # Live scan
    Show-Loading -msg "Scanning installed modules..."
    Set-Status -msg "Loading installed modules (live scan)..."

    try {
        $mods = Get-InstalledModule -ErrorAction SilentlyContinue | Sort-Object Name | ForEach-Object {
            Add-ModuleProps -m $_ -Status 'Installed'
        }
        $script:installedCache = @($mods)
        Update-ModuleList -modules $script:installedCache -emptyMsg "No modules installed via PowerShellGet."
        Set-Status -msg ("Found " + $script:installedCache.Count.ToString() + " installed module(s). Cache saved.")

        Save-CacheInstalled -modules $script:installedCache
        Update-CacheAgeDisplay
    }
    catch {
        Set-Status -msg ("Error: " + $_.Exception.Message)
    }
    finally {
        Hide-Loading
    }
}

function Load-Updates {
    param([switch]$Force)

    $script:currentView = 'updates'
    Set-NavActive -name 'updates'
    $ui.txtViewTitle.Text = 'Available Updates'
    Hide-Detail

    # Try cache first (unless forced)
    if (-not $Force) {
        $cached = Load-CacheFile -path $script:cacheUpdates
        if ($cached) {
            $script:updatesCache = @($cached.Modules)
            $age = Format-CacheAge -cachedAt $cached.CachedAt

            Update-ModuleList -modules $script:updatesCache -emptyMsg "All modules are up to date!"

            if ($script:updatesCache.Count -gt 0) {
                $ui.badgeUpdates.Visibility = 'Visible'
                $ui.txtBadge.Text = $script:updatesCache.Count.ToString()
                $ui.btnUpdateAll.Visibility = 'Visible'
            }
            else {
                $ui.badgeUpdates.Visibility = 'Collapsed'
                $ui.btnUpdateAll.Visibility = 'Collapsed'
            }

            Set-Status -msg ("Loaded " + $script:updatesCache.Count.ToString() + " update(s) from cache (" + $age + ")")
            Update-CacheAgeDisplay
            return
        }
    }

    # Live scan
    Show-Loading -msg "Checking for updates... (this may take a moment)"
    Set-Status -msg "Checking for module updates (live scan)..."

    try {
        if ($script:installedCache.Count -eq 0) {
            $script:installedCache = @(Get-InstalledModule -ErrorAction SilentlyContinue | Sort-Object Name | ForEach-Object {
                Add-ModuleProps -m $_ -Status 'Installed'
            })
            Save-CacheInstalled -modules $script:installedCache
        }

        $updates = @()
        $i = 0
        foreach ($mod in $script:installedCache) {
            $i++
            $ui.txtLoading.Text = "Checking $i of $($script:installedCache.Count): " + $mod.Name
            [System.Windows.Forms.Application]::DoEvents()

            try {
                $online = Find-Module -Name $mod.Name -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($online) {
                    $localVer  = [version]$mod.Version
                    $remoteVer = [version]$online.Version
                    if ($remoteVer -gt $localVer) {
                        $mod._Status        = 'Update'
                        $mod._LatestVersion = $online.Version.ToString()
                        try { $mod._Downloads = [long]$online.AdditionalMetadata.downloadCount } catch {}
                        $updates += $mod
                    }
                }
            }
            catch {}
        }

        $script:updatesCache = $updates
        Update-ModuleList -modules $script:updatesCache -emptyMsg "All modules are up to date!"

        if ($updates.Count -gt 0) {
            $ui.badgeUpdates.Visibility = 'Visible'
            $ui.txtBadge.Text = $updates.Count.ToString()
            $ui.btnUpdateAll.Visibility = 'Visible'
        }
        else {
            $ui.badgeUpdates.Visibility = 'Collapsed'
            $ui.btnUpdateAll.Visibility = 'Collapsed'
        }

        Set-Status -msg ("Found " + $updates.Count.ToString() + " update(s). Cache saved.")

        Save-CacheUpdates -modules $script:updatesCache
        Update-CacheAgeDisplay
    }
    catch {
        Set-Status -msg ("Error checking updates: " + $_.Exception.Message)
    }
    finally {
        Hide-Loading
    }
}

function Invoke-GallerySearch {
    param([string]$query)
    if ([string]::IsNullOrWhiteSpace($query)) { return }

    $script:currentView = 'search'
    Set-NavActive -name 'search'
    $ui.txtViewTitle.Text = "Search: " + $query
    Hide-Detail
    Show-Loading -msg "Searching PSGallery..."
    Set-Status -msg "Searching..."

    try {
        $installedNames = @{}
        $installedScopes = @{}
        foreach ($m in $script:installedCache) {
            $installedNames[$m.Name] = $m.Version
            $installedScopes[$m.Name] = $m._Scope
        }

        $results = Find-Module -Name "*$query*" -Repository PSGallery -ErrorAction SilentlyContinue |
                   Select-Object -First 50 | ForEach-Object {
            $isInstalled = $installedNames.ContainsKey($_.Name)
            $status = 'NotInstalled'
            $modScope = ''
            if ($isInstalled) {
                $localVer = $installedNames[$_.Name]
                $modScope = $installedScopes[$_.Name]
                if ([version]$_.Version -gt [version]$localVer) { $status = 'Update' }
                else { $status = 'Installed' }
            }

            $dlCount = 0
            try { $dlCount = [long]$_.AdditionalMetadata.downloadCount } catch {}

            $obj = Add-ModuleProps -m $_ -Status $status -Downloads $dlCount -LatestVersion $_.Version.ToString()
            # Override scope: use cached scope for installed, empty for new
            if ($modScope) { $obj._Scope = $modScope }
            else { $obj._Scope = '' }
            $obj
        }

        $script:searchResults = @($results)
        Update-ModuleList -modules $script:searchResults -emptyMsg "No modules matching that query."
        Set-Status -msg ("Found " + $script:searchResults.Count.ToString() + " result(s).")
    }
    catch {
        Set-Status -msg ("Search error: " + $_.Exception.Message)
    }
    finally {
        Hide-Loading
    }
}

# Module actions (scope-aware)
function Invoke-ModuleAction {
    param([string]$Action, $Module)

    $name      = $Module.Name
    $modScope  = $Module._Scope

    # For Install: auto-determine scope based on admin status
    if ($Action -eq 'Install') {
        $targetScope = $script:installScope
        $confirmMsg = "$Action module '$name' [$targetScope scope]?"
        $confirm = [System.Windows.MessageBox]::Show($confirmMsg, "Confirm $Action", 'YesNo', 'Question')
        if ($confirm -ne 'Yes') { return }
    }
    else {
        # Update / Uninstall: check if module is AllUsers and user is not admin
        if ($modScope -eq 'AllUsers' -and -not $isAdmin) {
            [System.Windows.MessageBox]::Show(
                "The module '$name' is installed in AllUsers scope (system-wide).`n`nTo $Action this module, please run PSGalleryManager as Administrator.",
                "Administrator Required", 'OK', 'Warning')
            return
        }

        $scopeLabel = if ($modScope) { $modScope } else { "detected" }
        $confirmMsg = "$Action module '$name' [$scopeLabel scope]?"
        $confirm = [System.Windows.MessageBox]::Show($confirmMsg, "Confirm $Action", 'YesNo', 'Question')
        if ($confirm -ne 'Yes') { return }
    }

    Show-Loading -msg "${Action}ing ${name}..."
    Set-Status -msg "${Action}ing ${name}..."

    try {
        if ($Action -eq 'Install') {
            $installParams = @{ Name = $name; Scope = $script:installScope; Force = $true; AllowClobber = $true; ErrorAction = 'Stop' }
            if ($script:supportsAcceptLicense) { $installParams['AcceptLicense'] = $true }
            Install-Module @installParams
            Set-Status -msg "Successfully installed $name."
            [System.Windows.MessageBox]::Show("Module '$name' installed successfully!", "Success", 'OK', 'Information')
        }
        elseif ($Action -eq 'Update') {
            $updateParams = @{ Name = $name; Force = $true; ErrorAction = 'Stop' }
            if ($script:supportsAcceptLicense) { $updateParams['AcceptLicense'] = $true }
            Update-Module @updateParams
            Set-Status -msg "Successfully updated $name."
            [System.Windows.MessageBox]::Show("Module '$name' updated successfully!", "Success", 'OK', 'Information')
        }
        elseif ($Action -eq 'Uninstall') {
            Uninstall-Module -Name $name -Force -AllVersions -ErrorAction Stop
            Set-Status -msg "Successfully uninstalled $name."
            [System.Windows.MessageBox]::Show("Module '$name' uninstalled.", "Done", 'OK', 'Information')
            Hide-Detail
        }

        # Invalidate caches and force refresh
        Clear-AllCache
        $script:installedCache = @()
        $script:updatesCache   = @()

        if ($script:currentView -eq 'installed') { Load-InstalledModules -Force }
        elseif ($script:currentView -eq 'updates') { Load-Updates -Force }
        elseif ($script:currentView -eq 'search') { Invoke-GallerySearch -query $ui.txtSearch.Text }
    }
    catch {
        Set-Status -msg ("Error: " + $_.Exception.Message)
        [System.Windows.MessageBox]::Show($_.Exception.Message, "Action Failed", 'OK', 'Error')
    }
    finally {
        Hide-Loading
    }
}

# -- Event wiring --

# Navigation
$ui.navInstalled.Add_Click({ Load-InstalledModules })
$ui.navUpdates.Add_Click({ Load-Updates })
$ui.navSearch.Add_Click({
    $script:currentView = 'search'
    Set-NavActive -name 'search'
    $ui.txtViewTitle.Text = 'Gallery Search'
    $ui.txtSearch.Focus()
    if ($script:searchResults.Count -gt 0) {
        Update-ModuleList -modules $script:searchResults
    }
    else {
        $ui.lstModules.Items.Clear()
        $ui.pnlBatchBar.Visibility = 'Collapsed'
        $ui.pnlEmpty.Visibility = 'Visible'
        $ui.txtEmpty.Text = "Type a keyword above and click Search Gallery"
        $ui.txtCount.Text = ''
    }
})

# Search
$ui.btnSearch.Add_Click({ Invoke-GallerySearch -query $ui.txtSearch.Text })
$ui.txtSearch.Add_KeyDown({
    if ($_.Key -eq 'Return') { Invoke-GallerySearch -query $ui.txtSearch.Text }
})

# Refresh (force live scan)
$ui.btnRefresh.Add_Click({
    $script:installedCache = @()
    $script:updatesCache   = @()
    if ($script:currentView -eq 'installed') { Load-InstalledModules -Force }
    elseif ($script:currentView -eq 'updates') { Load-Updates -Force }
    elseif ($script:currentView -eq 'search' -and $ui.txtSearch.Text) { Invoke-GallerySearch -query $ui.txtSearch.Text }
    else { Load-InstalledModules -Force }
})

# Clear cache
$ui.btnClearCache.Add_Click({
    Clear-AllCache
    $script:installedCache = @()
    $script:updatesCache   = @()
    Set-Status -msg "Cache cleared. Next load will do a live scan."
})

# Update all (scope-aware)
$ui.btnUpdateAll.Add_Click({
    if ($script:updatesCache.Count -eq 0) { return }

    # Check if any AllUsers modules exist and user is not admin
    if (-not $isAdmin) {
        $allUsersMods = @($script:updatesCache | Where-Object { $_._Scope -eq 'AllUsers' })
        if ($allUsersMods.Count -gt 0) {
            $userMods = @($script:updatesCache | Where-Object { $_._Scope -ne 'AllUsers' })
            if ($userMods.Count -eq 0) {
                [System.Windows.MessageBox]::Show(
                    "All " + $allUsersMods.Count.ToString() + " module(s) with updates are installed in AllUsers scope.`n`nPlease run PSGalleryManager as Administrator to update them.",
                    "Administrator Required", 'OK', 'Warning')
                return
            }
            else {
                $answer = [System.Windows.MessageBox]::Show(
                    $allUsersMods.Count.ToString() + " module(s) are AllUsers scope and will be skipped (requires Admin).`n`nUpdate the remaining " + $userMods.Count.ToString() + " CurrentUser module(s)?",
                    "Partial Update", 'YesNo', 'Question')
                if ($answer -ne 'Yes') { return }
                $modsToUpdate = $userMods
            }
        }
        else {
            $msg = "Update all " + $script:updatesCache.Count.ToString() + " module(s)?"
            $confirm = [System.Windows.MessageBox]::Show($msg, "Confirm Bulk Update", 'YesNo', 'Question')
            if ($confirm -ne 'Yes') { return }
            $modsToUpdate = $script:updatesCache
        }
    }
    else {
        $msg = "Update all " + $script:updatesCache.Count.ToString() + " module(s)?"
        $confirm = [System.Windows.MessageBox]::Show($msg, "Confirm Bulk Update", 'YesNo', 'Question')
        if ($confirm -ne 'Yes') { return }
        $modsToUpdate = $script:updatesCache
    }

    $i = 0
    $ok = 0
    $fail = 0
    foreach ($mod in $modsToUpdate) {
        $i++
        Show-Loading -msg ("Updating $i of " + $modsToUpdate.Count.ToString() + ": " + $mod.Name)
        try {
            $updateParams = @{ Name = $mod.Name; Force = $true; ErrorAction = 'Stop' }
            if ($script:supportsAcceptLicense) { $updateParams['AcceptLicense'] = $true }
            Update-Module @updateParams
            $ok++
        }
        catch {
            $fail++
        }
    }
    Hide-Loading

    Clear-AllCache
    $script:installedCache = @()
    $script:updatesCache   = @()

    Set-Status -msg "Bulk update complete: $ok succeeded, $fail failed."
    [System.Windows.MessageBox]::Show("Updated: $ok`nFailed: $fail", "Bulk Update Complete", 'OK', 'Information')
    Load-Updates -Force
})

# Selection changes -> update batch action bar
$ui.lstModules.Add_SelectionChanged({
    $selected = @()
    foreach ($item in $ui.lstModules.SelectedItems) {
        if ($item -and $item.Tag) { $selected += $item.Tag }
    }

    if ($selected.Count -gt 0) {
        $ui.pnlBatchBar.Visibility = 'Visible'
        $suffix = ""
        if ($selected.Count -ne 1) { $suffix = "s" }
        $ui.txtSelected.Text = $selected.Count.ToString() + " selected"

        # Show Update button if any selected have updates
        $hasUpdates = $false
        foreach ($m in $selected) {
            if ($m._Status -eq 'Update') { $hasUpdates = $true; break }
        }
        if ($hasUpdates) { $ui.btnBatchUpdate.Visibility = 'Visible' }
        else { $ui.btnBatchUpdate.Visibility = 'Collapsed' }

        # Show Uninstall button if any selected are installed
        $hasInstalled = $false
        foreach ($m in $selected) {
            if ($m._Status -eq 'Installed' -or $m._Status -eq 'Update') { $hasInstalled = $true; break }
        }
        if ($hasInstalled) { $ui.btnBatchUninstall.Visibility = 'Visible' }
        else { $ui.btnBatchUninstall.Visibility = 'Collapsed' }
    }
    else {
        $ui.pnlBatchBar.Visibility = 'Collapsed'
    }
})

# Clear selection
$ui.btnClearSelection.Add_Click({
    $ui.lstModules.UnselectAll()
    $ui.pnlBatchBar.Visibility = 'Collapsed'
})

# Batch Update
$ui.btnBatchUpdate.Add_Click({
    $selected = @()
    foreach ($item in $ui.lstModules.SelectedItems) {
        if ($item -and $item.Tag -and $item.Tag._Status -eq 'Update') { $selected += $item.Tag }
    }
    if ($selected.Count -eq 0) { return }

    # Check AllUsers modules when not admin
    $toUpdate = @()
    $skipped = @()
    foreach ($m in $selected) {
        if ($m._Scope -eq 'AllUsers' -and -not $isAdmin) {
            $skipped += $m
        }
        else {
            $toUpdate += $m
        }
    }

    if ($skipped.Count -gt 0 -and $toUpdate.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "All " + $skipped.Count.ToString() + " selected module(s) are AllUsers scope.`n`nPlease run PSGalleryManager as Administrator to update them.",
            "Administrator Required", 'OK', 'Warning')
        return
    }
    elseif ($skipped.Count -gt 0) {
        $answer = [System.Windows.MessageBox]::Show(
            $skipped.Count.ToString() + " AllUsers module(s) will be skipped (requires Admin).`n`nUpdate the remaining " + $toUpdate.Count.ToString() + " module(s)?",
            "Partial Update", 'YesNo', 'Question')
        if ($answer -ne 'Yes') { return }
    }
    else {
        $confirm = [System.Windows.MessageBox]::Show(
            "Update " + $toUpdate.Count.ToString() + " selected module(s)?",
            "Confirm Batch Update", 'YesNo', 'Question')
        if ($confirm -ne 'Yes') { return }
    }

    $i = 0; $ok = 0; $fail = 0
    foreach ($mod in $toUpdate) {
        $i++
        Show-Loading -msg ("Updating $i of " + $toUpdate.Count.ToString() + ": " + $mod.Name)
        try {
            $updateParams = @{ Name = $mod.Name; Force = $true; ErrorAction = 'Stop' }
            if ($script:supportsAcceptLicense) { $updateParams['AcceptLicense'] = $true }
            Update-Module @updateParams
            $ok++
        }
        catch { $fail++ }
    }
    Hide-Loading

    Clear-AllCache
    $script:installedCache = @()
    $script:updatesCache   = @()

    $resultMsg = "Updated: $ok"
    if ($fail -gt 0) { $resultMsg += "`nFailed: $fail" }
    if ($skipped.Count -gt 0) { $resultMsg += "`nSkipped (AllUsers): " + $skipped.Count.ToString() }
    [System.Windows.MessageBox]::Show($resultMsg, "Batch Update Complete", 'OK', 'Information')

    if ($script:currentView -eq 'installed') { Load-InstalledModules -Force }
    elseif ($script:currentView -eq 'updates') { Load-Updates -Force }
    elseif ($script:currentView -eq 'search') { Invoke-GallerySearch -query $ui.txtSearch.Text }
})

# Batch Uninstall
$ui.btnBatchUninstall.Add_Click({
    $selected = @()
    foreach ($item in $ui.lstModules.SelectedItems) {
        if ($item -and $item.Tag -and ($item.Tag._Status -eq 'Installed' -or $item.Tag._Status -eq 'Update')) {
            $selected += $item.Tag
        }
    }
    if ($selected.Count -eq 0) { return }

    # Check AllUsers modules when not admin
    $toRemove = @()
    $skipped = @()
    foreach ($m in $selected) {
        if ($m._Scope -eq 'AllUsers' -and -not $isAdmin) {
            $skipped += $m
        }
        else {
            $toRemove += $m
        }
    }

    if ($skipped.Count -gt 0 -and $toRemove.Count -eq 0) {
        [System.Windows.MessageBox]::Show(
            "All " + $skipped.Count.ToString() + " selected module(s) are AllUsers scope.`n`nPlease run PSGalleryManager as Administrator to uninstall them.",
            "Administrator Required", 'OK', 'Warning')
        return
    }
    elseif ($skipped.Count -gt 0) {
        $answer = [System.Windows.MessageBox]::Show(
            $skipped.Count.ToString() + " AllUsers module(s) will be skipped (requires Admin).`n`nUninstall the remaining " + $toRemove.Count.ToString() + " module(s)?",
            "Partial Uninstall", 'YesNo', 'Question')
        if ($answer -ne 'Yes') { return }
    }
    else {
        $nameList = ""
        foreach ($m in $toRemove) { $nameList += "`n  - " + $m.Name }
        $confirm = [System.Windows.MessageBox]::Show(
            "Uninstall " + $toRemove.Count.ToString() + " module(s)?" + $nameList,
            "Confirm Batch Uninstall", 'YesNo', 'Warning')
        if ($confirm -ne 'Yes') { return }
    }

    $i = 0; $ok = 0; $fail = 0
    foreach ($mod in $toRemove) {
        $i++
        Show-Loading -msg ("Uninstalling $i of " + $toRemove.Count.ToString() + ": " + $mod.Name)
        try {
            Uninstall-Module -Name $mod.Name -Force -AllVersions -ErrorAction Stop
            $ok++
        }
        catch { $fail++ }
    }
    Hide-Loading
    Hide-Detail

    Clear-AllCache
    $script:installedCache = @()
    $script:updatesCache   = @()

    $resultMsg = "Uninstalled: $ok"
    if ($fail -gt 0) { $resultMsg += "`nFailed: $fail" }
    if ($skipped.Count -gt 0) { $resultMsg += "`nSkipped (AllUsers): " + $skipped.Count.ToString() }
    [System.Windows.MessageBox]::Show($resultMsg, "Batch Uninstall Complete", 'OK', 'Information')

    if ($script:currentView -eq 'installed') { Load-InstalledModules -Force }
    elseif ($script:currentView -eq 'updates') { Load-Updates -Force }
    elseif ($script:currentView -eq 'search') { Invoke-GallerySearch -query $ui.txtSearch.Text }
})

# Detail actions
$ui.btnCloseDetail.Add_Click({ Hide-Detail })

$ui.btnDetailInstall.Add_Click({
    if ($script:selectedModule) {
        Invoke-ModuleAction -Action 'Install' -Module $script:selectedModule
    }
})

$ui.btnDetailUpdate.Add_Click({
    if ($script:selectedModule) {
        Invoke-ModuleAction -Action 'Update' -Module $script:selectedModule
    }
})

$ui.btnDetailUninstall.Add_Click({
    if ($script:selectedModule) {
        Invoke-ModuleAction -Action 'Uninstall' -Module $script:selectedModule
    }
})

$ui.btnOpenGallery.Add_Click({
    if ($script:selectedModule) {
        Start-Process ("https://www.powershellgallery.com/packages/" + $script:selectedModule.Name)
    }
})

$ui.btnOpenLocation.Add_Click({
    if ($script:selectedModule -and $script:selectedModule._InstalledLocation) {
        $path = $script:selectedModule._InstalledLocation
        if (Test-Path $path) {
            Start-Process explorer.exe -ArgumentList $path
        }
        else {
            [System.Windows.MessageBox]::Show("Path not found: $path", "Location Not Found", 'OK', 'Warning')
        }
    }
})

# Clickable project URL
$ui.txtDetailUrl.Add_MouseLeftButtonDown({
    $url = $ui.txtDetailUrl.Text
    if ($url -and $url -ne '-') {
        Start-Process $url
    }
})

# Initial load
$window.Add_ContentRendered({
    Update-CacheAgeDisplay
    Load-InstalledModules
})

# Show window
$window.ShowDialog() | Out-Null

}

Export-ModuleMember -Function Start-PSGalleryManager
New-Alias -Name psgm -Value Start-PSGalleryManager -Force
Export-ModuleMember -Alias psgm
