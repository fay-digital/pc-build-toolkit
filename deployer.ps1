#Requires -Version 5.1
# =============================================================================
#  PC BUILD TOOLKIT (v1.1.0)
#  - Self-elevating PowerShell + WPF
#  - Multi-source installers (winget / choco / direct / zip)
#  - Bloat removal, GPU driver auto-detect, Windows Update, Start menu grid
#  - Extraction: tar.exe -> 7-Zip standalone -> Expand-Archive fallback chain
#
#  Launch locally:  powershell -ExecutionPolicy Bypass -File .\deployer.ps1
#  Launch from web: irm https://fay.digital/pbt | iex
# =============================================================================

$SCRIPT_VERSION = 'v1.1.0'
$SCRIPT_RAW_URL = 'https://raw.githubusercontent.com/fay-digital/pc-build-toolkit/main/deployer.ps1'

# Standalone 7-Zip console binary (7zr.exe handles .7z; 7za.exe handles .zip
# including Deflate64). ~1 MB, no install required.
$SCRIPT_7ZA_URL = 'https://www.7-zip.org/a/7z2409-extra.7z'
# 7z2409-extra.7z is itself a .7z archive containing 7za.exe. We need a raw
# 7za.exe to bootstrap. Use NuGet's hosted copy which is the standard
# "get 7za.exe with just HTTP" path (same binary, stable URL, no install).
$SCRIPT_7ZA_DIRECT = 'https://raw.githubusercontent.com/mcmilk/7-Zip/main/CPP/7zip/Bundles/Alone/7za.exe'

# --- Self-elevate if not admin -----------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
            ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    if ($PSCommandPath) {
        Start-Process powershell -ArgumentList "-ExecutionPolicy Bypass -File `"$PSCommandPath`"" -Verb RunAs
    }
    exit
}

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase

# --- App catalog -------------------------------------------------------------
$script:AppCatalog = @(
    @{ Id='FinalWire.AIDA64.Extreme';        Name='AIDA64 Extreme';  Category='Diagnostics'; Source='winget' }
    @{ Id='REALiX.HWiNFO';                   Name='HWiNFO';          Category='Diagnostics'; Source='winget' }
    @{ Id='CrystalDewWorld.CrystalDiskMark'; Name='CrystalDiskMark'; Category='Benchmark';   Source='winget' }
    @{ Id='Maxon.CinebenchR23';              Name='Cinebench R23';   Category='Benchmark';   Source='winget' }
    @{ Id='3dmark-bundled';                  Name='3DMark (Steel Nomad)'; Category='Benchmark'; Source='zip'
       DownloadUrl='https://github.com/fay-digital/pc-build-toolkit/releases/download/v1.0.0/3dmark-bundle.zip'
       SetupExecutable='3dmark-setup.exe'
       SilentArgs='/S'
       UninstallRegistryMatch='3DMark' }
    @{ Id='Geeks3D.FurMark.2'; Name='FurMark 2'; Category='GPU stress';  Source='winget' }
    @{ Id='OCCT.OCCT';         Name='OCCT';      Category='Stability';   Source='winget' }
    @{ Id='CPUID.CPU-Z';       Name='CPU-Z';     Category='Info';        Source='winget' }
    @{ Id='TechPowerUp.GPU-Z'; Name='GPU-Z';     Category='Info';        Source='winget' }
    @{ Id='GIMPS.Prime95';     Name='Prime95';   Category='CPU torture'; Source='winget' }
)

$script:DefaultChecked = @(
    'FinalWire.AIDA64.Extreme','REALiX.HWiNFO','CrystalDewWorld.CrystalDiskMark',
    'Maxon.CinebenchR23','3dmark-bundled'
)

# --- Bloat list (AppX, current user) ----------------------------------------
$script:BloatList = @(
    @{ Name='Camera';              Match='Microsoft.WindowsCamera' }
    @{ Name='Clipchamp';           Match='Clipchamp.Clipchamp' }
    @{ Name='Clock';               Match='Microsoft.WindowsAlarms' }
    @{ Name='Copilot';              Match='Microsoft.Copilot' }
    @{ Name='Feedback Hub';        Match='Microsoft.WindowsFeedbackHub' }
    @{ Name='News';                Match='Microsoft.BingNews' }
    @{ Name='Outlook (new)';       Match='Microsoft.OutlookForWindows' }
    @{ Name='Power Automate';      Match='Microsoft.PowerAutomateDesktop' }
    @{ Name='Solitaire';           Match='Microsoft.MicrosoftSolitaireCollection' }
    @{ Name='Sound Recorder';      Match='Microsoft.WindowsSoundRecorder' }
    @{ Name='Start Experiences';   Match='MicrosoftWindows.Client.WebExperience' }
    @{ Name='Sticky Notes';        Match='Microsoft.MicrosoftStickyNotes' }
    @{ Name='Teams (personal)';    Match='MicrosoftTeams' }
    @{ Name='Teams (new)';         Match='MSTeams' }
    @{ Name='To Do';               Match='Microsoft.Todos' }
    @{ Name='Weather';             Match='Microsoft.BingWeather' }
    @{ Name='Web Media Extensions'; Match='Microsoft.WebMediaExtensions' }
)

# --- XAML UI -----------------------------------------------------------------
[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="PC Build Toolkit" Height="860" Width="1180" MinWidth="920" MinHeight="660"
        WindowStartupLocation="CenterScreen" Background="#12151B"
        FontFamily="Segoe UI" UseLayoutRounding="True" TextOptions.TextFormattingMode="Display">
    <Window.Resources>

        <!-- ================= Palette ================= -->
        <SolidColorBrush x:Key="BgApp"     Color="#12151B"/>
        <SolidColorBrush x:Key="BgDeep"    Color="#0D1015"/>
        <SolidColorBrush x:Key="BgPanel"   Color="#1A1E26"/>
        <SolidColorBrush x:Key="BgPanel2"  Color="#20252F"/>
        <SolidColorBrush x:Key="BgHover"   Color="#232832"/>
        <SolidColorBrush x:Key="Line"      Color="#2A2F3A"/>
        <SolidColorBrush x:Key="LineSoft"  Color="#23283132"/>
        <SolidColorBrush x:Key="TextHi"    Color="#F1F3F6"/>
        <SolidColorBrush x:Key="TextMid"   Color="#C4C9D2"/>
        <SolidColorBrush x:Key="TextMuted" Color="#8891A0"/>
        <SolidColorBrush x:Key="TextDim"   Color="#5D6572"/>
        <SolidColorBrush x:Key="Accent"    Color="#4FA3FF"/>
        <SolidColorBrush x:Key="AccentSoft" Color="#264063"/>
        <SolidColorBrush x:Key="Lime"      Color="#8AE06B"/>
        <SolidColorBrush x:Key="Amber"     Color="#E0B15B"/>
        <SolidColorBrush x:Key="Rose"      Color="#E07A6B"/>
        <FontFamily x:Key="Mono">Cascadia Mono, Consolas, Courier New</FontFamily>

        <!-- ================= Global typography ================= -->
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="{StaticResource TextHi}"/>
            <Setter Property="FontSize" Value="13"/>
        </Style>

        <!-- ================= Title bar button ================= -->
        <Style x:Key="TitleBarBtn" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{StaticResource TextMid}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Width" Value="44"/>
            <Setter Property="Height" Value="36"/>
            <Setter Property="FontFamily" Value="Segoe MDL2 Assets"/>
            <Setter Property="FontSize" Value="10"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="B" Background="{TemplateBinding Background}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="B" Property="Background" Value="#262B36"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="TitleBarBtnClose" TargetType="Button" BasedOn="{StaticResource TitleBarBtn}">
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#C42B1C"/>
                    <Setter Property="Foreground" Value="White"/>
                </Trigger>
            </Style.Triggers>
        </Style>

        <!-- ================= Sidebar nav item ================= -->
        <Style x:Key="NavItem" TargetType="RadioButton">
            <Setter Property="Foreground" Value="{StaticResource TextMid}"/>
            <Setter Property="Padding" Value="10,8"/>
            <Setter Property="Margin" Value="0,1"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Tag" Value=""/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RadioButton">
                        <Grid>
                            <Border x:Name="B" Background="Transparent" CornerRadius="6" Padding="{TemplateBinding Padding}">
                                <StackPanel Orientation="Horizontal">
                                    <Rectangle x:Name="Tick" Width="2" Height="16" Fill="Transparent" RadiusX="1" RadiusY="1" Margin="-4,0,8,0" VerticalAlignment="Center"/>
                                    <Viewbox Width="14" Height="14" Margin="0,0,10,0" VerticalAlignment="Center">
                                        <Path x:Name="Ico" Data="{TemplateBinding Tag}"
                                              Stroke="{TemplateBinding Foreground}" StrokeThickness="1.4"
                                              StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"
                                              Fill="Transparent" Width="16" Height="16"/>
                                    </Viewbox>
                                    <ContentPresenter VerticalAlignment="Center"/>
                                </StackPanel>
                            </Border>
                        </Grid>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="B" Property="Background" Value="{StaticResource BgHover}"/>
                                <Setter Property="Foreground" Value="{StaticResource TextHi}"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="B" Property="Background" Value="{StaticResource BgPanel}"/>
                                <Setter TargetName="Tick" Property="Fill" Value="{StaticResource Accent}"/>
                                <Setter Property="Foreground" Value="{StaticResource TextHi}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ================= Checkbox (custom square) ================= -->
        <Style TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource TextMid}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <Border x:Name="Row" Background="Transparent" CornerRadius="6" Padding="10,9">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                </Grid.ColumnDefinitions>
                                <Border x:Name="Box" Grid.Column="0" Width="16" Height="16" CornerRadius="4"
                                        Background="Transparent" BorderBrush="{StaticResource Line}" BorderThickness="1.4"
                                        VerticalAlignment="Center" Margin="0,0,12,0">
                                    <Path x:Name="Tick" Data="M 2.5 8 L 6 11 L 13 4" Stroke="#0F1115" StrokeThickness="2"
                                          StrokeStartLineCap="Round" StrokeEndLineCap="Round" StrokeLineJoin="Round"
                                          Stretch="None" Visibility="Collapsed"
                                          HorizontalAlignment="Center" VerticalAlignment="Center"/>
                                </Border>
                                <ContentPresenter Grid.Column="1" VerticalAlignment="Center"/>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Row" Property="Background" Value="{StaticResource BgHover}"/>
                                <Setter Property="Foreground" Value="{StaticResource TextHi}"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="Box" Property="Background" Value="{StaticResource Accent}"/>
                                <Setter TargetName="Box" Property="BorderBrush" Value="{StaticResource Accent}"/>
                                <Setter TargetName="Tick" Property="Visibility" Value="Visible"/>
                                <Setter Property="Foreground" Value="{StaticResource TextHi}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ================= Toggle switch (CheckBox retemplated, two-line label + sliding thumb) ================= -->
        <Style x:Key="ToggleSwitch" TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource TextHi}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Tag" Value=""/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <Border x:Name="Row" Background="Transparent" CornerRadius="6" Padding="10,8">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0" VerticalAlignment="Center">
                                    <TextBlock Text="{TemplateBinding Content}" Foreground="{TemplateBinding Foreground}" FontSize="13" FontWeight="Medium"/>
                                    <TextBlock Text="{TemplateBinding Tag}" Foreground="{StaticResource TextMuted}" FontSize="11" Margin="0,2,0,0" TextWrapping="Wrap"/>
                                </StackPanel>
                                <Border x:Name="Track" Grid.Column="1" Width="34" Height="19" CornerRadius="10"
                                        BorderThickness="1"
                                        VerticalAlignment="Center" Margin="14,0,0,0">
                                    <Border.Background>
                                        <SolidColorBrush Color="#1A1F2A"/>
                                    </Border.Background>
                                    <Border.BorderBrush>
                                        <SolidColorBrush Color="#2A3140"/>
                                    </Border.BorderBrush>
                                    <Grid>
                                        <Ellipse x:Name="Thumb" Width="13" Height="13"
                                                 HorizontalAlignment="Left" VerticalAlignment="Center" Margin="2,0,0,0">
                                            <Ellipse.Fill>
                                                <SolidColorBrush Color="#7D8795"/>
                                            </Ellipse.Fill>
                                        </Ellipse>
                                    </Grid>
                                </Border>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Row" Property="Background" Value="{StaticResource BgHover}"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="Track" Property="Background" Value="#4FA3FF"/>
                                <Setter TargetName="Track" Property="BorderBrush" Value="#4FA3FF"/>
                                <Setter TargetName="Thumb" Property="Fill" Value="#0B1320"/>
                                <Setter TargetName="Thumb" Property="Margin" Value="17,0,0,0"/>
                                <Trigger.EnterActions>
                                    <BeginStoryboard>
                                        <Storyboard>
                                            <ThicknessAnimation Storyboard.TargetName="Thumb" Storyboard.TargetProperty="Margin"
                                                                To="17,0,0,0" Duration="0:0:0.12"/>
                                        </Storyboard>
                                    </BeginStoryboard>
                                </Trigger.EnterActions>
                                <Trigger.ExitActions>
                                    <BeginStoryboard>
                                        <Storyboard>
                                            <ThicknessAnimation Storyboard.TargetName="Thumb" Storyboard.TargetProperty="Margin"
                                                                To="2,0,0,0" Duration="0:0:0.12"/>
                                        </Storyboard>
                                    </BeginStoryboard>
                                </Trigger.ExitActions>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ================= Toggle switch — destructive variant (amber left stripe) ================= -->
        <Style x:Key="ToggleSwitchDestructive" TargetType="CheckBox">
            <Setter Property="Foreground" Value="{StaticResource TextHi}"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Tag" Value=""/>
            <Setter Property="VerticalContentAlignment" Value="Center"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="CheckBox">
                        <Border x:Name="Row" Background="Transparent" CornerRadius="6" Padding="10,8"
                                BorderBrush="{StaticResource Amber}" BorderThickness="3,0,0,0">
                            <Grid>
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <StackPanel Grid.Column="0" VerticalAlignment="Center">
                                    <TextBlock Text="{TemplateBinding Content}" Foreground="{TemplateBinding Foreground}" FontSize="13" FontWeight="Medium"/>
                                    <TextBlock Text="{TemplateBinding Tag}" Foreground="{StaticResource Amber}" FontSize="11" Margin="0,2,0,0" TextWrapping="Wrap"/>
                                </StackPanel>
                                <Border x:Name="Track" Grid.Column="1" Width="34" Height="19" CornerRadius="10"
                                        BorderThickness="1" VerticalAlignment="Center" Margin="14,0,0,0">
                                    <Border.Background><SolidColorBrush Color="#1A1F2A"/></Border.Background>
                                    <Border.BorderBrush><SolidColorBrush Color="#2A3140"/></Border.BorderBrush>
                                    <Grid>
                                        <Ellipse x:Name="Thumb" Width="13" Height="13"
                                                 HorizontalAlignment="Left" VerticalAlignment="Center" Margin="2,0,0,0">
                                            <Ellipse.Fill><SolidColorBrush Color="#7D8795"/></Ellipse.Fill>
                                        </Ellipse>
                                    </Grid>
                                </Border>
                            </Grid>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="Row" Property="Background" Value="{StaticResource BgHover}"/>
                            </Trigger>
                            <Trigger Property="IsChecked" Value="True">
                                <Setter TargetName="Track" Property="Background" Value="#4FA3FF"/>
                                <Setter TargetName="Track" Property="BorderBrush" Value="#4FA3FF"/>
                                <Setter TargetName="Thumb" Property="Fill" Value="#0B1320"/>
                                <Setter TargetName="Thumb" Property="Margin" Value="17,0,0,0"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ================= Chip button (category filter / toggle-all) ================= -->
        <Style x:Key="ChipBtn" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{StaticResource TextMuted}"/>
            <Setter Property="BorderBrush" Value="{StaticResource Line}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="10,4"/>
            <Setter Property="FontFamily" Value="{StaticResource Mono}"/>
            <Setter Property="FontSize" Value="10.5"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Margin" Value="0,0,5,5"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="B" CornerRadius="2"
                                Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="B" Property="Background" Value="#1A1F2A"/>
                                <Setter Property="Foreground" Value="{StaticResource TextHi}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ================= Footer buttons ================= -->
        <Style x:Key="BtnGhost" TargetType="Button">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Foreground" Value="{StaticResource TextMid}"/>
            <Setter Property="BorderBrush" Value="{StaticResource Line}"/>
            <Setter Property="BorderThickness" Value="1"/>
            <Setter Property="Padding" Value="14,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="B" CornerRadius="6" Background="{TemplateBinding Background}"
                                BorderBrush="{TemplateBinding BorderBrush}"
                                BorderThickness="{TemplateBinding BorderThickness}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="B" Property="Background" Value="{StaticResource BgPanel}"/>
                                <Setter Property="Foreground" Value="{StaticResource TextHi}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="BtnDanger" TargetType="Button" BasedOn="{StaticResource BtnGhost}">
            <Setter Property="Foreground" Value="#E07A6B"/>
            <Setter Property="BorderBrush" Value="#5A2A26"/>
        </Style>
        <Style x:Key="BtnPrimary" TargetType="Button">
            <Setter Property="Background" Value="{StaticResource Accent}"/>
            <Setter Property="Foreground" Value="#0B1320"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Padding" Value="18,8"/>
            <Setter Property="FontSize" Value="13"/>
            <Setter Property="FontWeight" Value="SemiBold"/>
            <Setter Property="Cursor" Value="Hand"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border x:Name="B" CornerRadius="6" Background="{TemplateBinding Background}"
                                Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                        <ControlTemplate.Triggers>
                            <Trigger Property="IsMouseOver" Value="True">
                                <Setter TargetName="B" Property="Background" Value="#6DB4FF"/>
                            </Trigger>
                            <Trigger Property="IsEnabled" Value="False">
                                <Setter TargetName="B" Property="Background" Value="#2A3748"/>
                                <Setter Property="Foreground" Value="{StaticResource TextMuted}"/>
                            </Trigger>
                        </ControlTemplate.Triggers>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ================= Progress bar (flat, thin) ================= -->
        <Style TargetType="ProgressBar">
            <Setter Property="Height" Value="4"/>
            <Setter Property="Background" Value="{StaticResource BgPanel}"/>
            <Setter Property="Foreground" Value="{StaticResource Accent}"/>
            <Setter Property="BorderThickness" Value="0"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ProgressBar">
                        <Border CornerRadius="999" Background="{TemplateBinding Background}" ClipToBounds="True">
                            <Grid>
                                <Rectangle x:Name="PART_Track"/>
                                <Border x:Name="PART_Indicator" HorizontalAlignment="Left"
                                        Background="{TemplateBinding Foreground}" CornerRadius="999"/>
                            </Grid>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

        <!-- ================= ScrollBar (thin, dark) ================= -->
        <Style x:Key="ScrollThumb" TargetType="Thumb">
            <Setter Property="OverridesDefaultStyle" Value="True"/>
            <Setter Property="IsTabStop" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Thumb">
                        <Border Background="#2A3140" CornerRadius="3" Margin="2,2,2,2"/>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style x:Key="ScrollPageBtn" TargetType="RepeatButton">
            <Setter Property="OverridesDefaultStyle" Value="True"/>
            <Setter Property="IsTabStop" Value="False"/>
            <Setter Property="Focusable" Value="False"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="RepeatButton">
                        <Rectangle Fill="Transparent"/>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
        <Style TargetType="ScrollBar">
            <Setter Property="Background" Value="Transparent"/>
            <Setter Property="Width" Value="10"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="ScrollBar">
                        <Grid Background="{TemplateBinding Background}">
                            <Track Name="PART_Track" IsDirectionReversed="True" Focusable="False">
                                <Track.DecreaseRepeatButton>
                                    <RepeatButton Command="ScrollBar.PageUpCommand" Style="{StaticResource ScrollPageBtn}"/>
                                </Track.DecreaseRepeatButton>
                                <Track.IncreaseRepeatButton>
                                    <RepeatButton Command="ScrollBar.PageDownCommand" Style="{StaticResource ScrollPageBtn}"/>
                                </Track.IncreaseRepeatButton>
                                <Track.Thumb>
                                    <Thumb Style="{StaticResource ScrollThumb}"/>
                                </Track.Thumb>
                            </Track>
                        </Grid>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>

    </Window.Resources>

    <!-- ============ Frameless root ============ -->
    <WindowChrome.WindowChrome>
        <WindowChrome CaptionHeight="36" CornerRadius="0" GlassFrameThickness="0" ResizeBorderThickness="6" UseAeroCaptionButtons="False"/>
    </WindowChrome.WindowChrome>

    <Border Background="{StaticResource BgApp}" BorderBrush="#2A2F3A" BorderThickness="1">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="36"/>       <!-- title bar -->
            <RowDefinition Height="*"/>        <!-- body -->
        </Grid.RowDefinitions>

        <!-- ============ Title bar ============ -->
        <Grid Grid.Row="0" Background="#0F1218">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="*"/>
                <ColumnDefinition Width="Auto"/>
            </Grid.ColumnDefinitions>
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center" Margin="14,0,0,0">
                <Border Width="14" Height="14" CornerRadius="3" Margin="0,0,10,0">
                    <Border.Background>
                        <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                            <GradientStop Color="#4FA3FF" Offset="0"/>
                            <GradientStop Color="#8A6DFF" Offset="1"/>
                        </LinearGradientBrush>
                    </Border.Background>
                </Border>
                <TextBlock Foreground="{StaticResource TextMid}" FontSize="12" VerticalAlignment="Center">
                    <Run FontWeight="SemiBold" Foreground="#F1F3F6" Text="PC Build Toolkit"/>
                    <Run Text="  |  "/><Run x:Name="VersionLabel" Text="v1.1.0"/>
                    <Run Text="  |  Administrator"/>
                </TextBlock>
            </StackPanel>
            <StackPanel Grid.Column="1" Orientation="Horizontal" WindowChrome.IsHitTestVisibleInChrome="True">
                <Button x:Name="BtnMin"   Style="{StaticResource TitleBarBtn}"       Content="&#xE921;"/>
                <Button x:Name="BtnMax"   Style="{StaticResource TitleBarBtn}"       Content="&#xE922;"/>
                <Button x:Name="BtnClose" Style="{StaticResource TitleBarBtnClose}"  Content="&#xE8BB;"/>
            </StackPanel>
        </Grid>

        <!-- ============ Body: sidebar + main ============ -->
        <Grid Grid.Row="1">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="220"/>
                <ColumnDefinition Width="*"/>
            </Grid.ColumnDefinitions>

            <!-- Sidebar -->
            <Border Grid.Column="0" Background="{StaticResource BgDeep}" BorderBrush="{StaticResource Line}" BorderThickness="0,0,1,0">
                <Grid>
                    <Grid.RowDefinitions>
                        <RowDefinition Height="Auto"/>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="Auto"/>
                    </Grid.RowDefinitions>

                    <!-- Brand -->
                    <StackPanel Grid.Row="0" Margin="18,20,18,16">
                        <StackPanel Orientation="Horizontal">
                            <Border Width="28" Height="28" CornerRadius="6" Margin="0,0,10,0">
                                <Border.Background>
                                    <LinearGradientBrush StartPoint="0,0" EndPoint="1,1">
                                        <GradientStop Color="#4FA3FF" Offset="0"/>
                                        <GradientStop Color="#8A6DFF" Offset="1"/>
                                    </LinearGradientBrush>
                                </Border.Background>
                            </Border>
                            <StackPanel VerticalAlignment="Center">
                                <TextBlock Text="Build Toolkit" FontWeight="SemiBold" FontSize="13"/>
                                <TextBlock Text="fay-digital / pbt" FontFamily="{StaticResource Mono}" FontSize="10" Foreground="{StaticResource TextMuted}"/>
                            </StackPanel>
                        </StackPanel>
                    </StackPanel>

                    <Rectangle Grid.Row="0" Fill="{StaticResource Line}" Height="1" VerticalAlignment="Bottom" Margin="12,0"/>

                    <!-- Nav -->
                    <StackPanel Grid.Row="1" Margin="10,12,10,0">
                        <TextBlock Text="WORKSPACE" Foreground="{StaticResource TextDim}" FontFamily="{StaticResource Mono}" FontSize="10" Margin="10,4,0,6"/>
                        <RadioButton x:Name="NavDeploy" Style="{StaticResource NavItem}" IsChecked="True" GroupName="Nav" Content="Deploy"
                            Tag="M4,3 L13,3 L13,13 L4,13 Z M4,6 L13,6 M7,9 L10,9"/>
                        <RadioButton x:Name="NavReport" Style="{StaticResource NavItem}" GroupName="Nav" Content="Report"
                            Tag="M3,2 L13,2 L13,14 L3,14 Z M5,5 L11,5 M5,8 L11,8 M5,11 L9,11"/>
                        <RadioButton x:Name="NavLogs"   Style="{StaticResource NavItem}" GroupName="Nav" Content="Logs"
                            Tag="M3,4 L13,4 M3,8 L13,8 M3,12 L13,12"/>

                        <TextBlock Text="MACHINE" Foreground="{StaticResource TextDim}" FontFamily="{StaticResource Mono}" FontSize="10" Margin="10,18,0,6"/>
                        <Grid Margin="10,3">
                            <Grid.ColumnDefinitions><ColumnDefinition Width="20"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                            <Path Grid.Column="0" Data="M3.5,3.5 L12.5,3.5 L12.5,12.5 L3.5,12.5 Z M6,6 L10,6 L10,10 L6,10 Z M6,1 L6,3 M10,1 L10,3 M6,13 L6,15 M10,13 L10,15 M1,6 L3,6 M1,10 L3,10 M13,6 L15,6 M13,10 L15,10"
                                  Stroke="{StaticResource TextMuted}" StrokeThickness="1.3" Fill="Transparent" Width="14" Height="14" Stretch="Uniform" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <TextBlock Grid.Column="1" x:Name="SysCpu" Foreground="{StaticResource TextMid}" FontSize="12" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
                        </Grid>
                        <Grid Margin="10,3">
                            <Grid.ColumnDefinitions><ColumnDefinition Width="20"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                            <Path Grid.Column="0" Data="M1.5,5 L14.5,5 L14.5,12 L1.5,12 Z M5,8.5 A1.6,1.6 0 1 1 4.99,8.5 M11,8.5 A1.6,1.6 0 1 1 10.99,8.5"
                                  Stroke="{StaticResource TextMuted}" StrokeThickness="1.3" Fill="Transparent" Width="14" Height="14" Stretch="Uniform" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <TextBlock Grid.Column="1" x:Name="SysGpu" Foreground="{StaticResource TextMid}" FontSize="12" VerticalAlignment="Center" TextTrimming="CharacterEllipsis"/>
                        </Grid>
                        <Grid Margin="10,3">
                            <Grid.ColumnDefinitions><ColumnDefinition Width="20"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                            <Path Grid.Column="0" Data="M1.5,4.5 L14.5,4.5 L14.5,10.5 L1.5,10.5 Z M4,4.5 L4,10.5 M7,4.5 L7,10.5 M10,4.5 L10,10.5 M13,4.5 L13,10.5"
                                  Stroke="{StaticResource TextMuted}" StrokeThickness="1.3" Fill="Transparent" Width="14" Height="14" Stretch="Uniform" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <TextBlock Grid.Column="1" x:Name="SysRam" Foreground="{StaticResource TextMid}" FontSize="12" VerticalAlignment="Center"/>
                        </Grid>
                        <Grid Margin="10,3">
                            <Grid.ColumnDefinitions><ColumnDefinition Width="20"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
                            <Path Grid.Column="0" Data="M2,3 L14,3 L14,13 L2,13 Z M5,10 A0.8,0.8 0 1 1 4.99,10 M7.5,10 A0.8,0.8 0 1 1 7.49,10"
                                  Stroke="{StaticResource TextMuted}" StrokeThickness="1.3" Fill="Transparent" Width="14" Height="14" Stretch="Uniform" VerticalAlignment="Center" HorizontalAlignment="Left"/>
                            <TextBlock Grid.Column="1" x:Name="SysDisk" Foreground="{StaticResource TextMid}" FontSize="12" VerticalAlignment="Center"/>
                        </Grid>
                    </StackPanel>

                    <!-- Foot -->
                    <StackPanel Grid.Row="2" Margin="18,12,18,16">
                        <Rectangle Fill="{StaticResource Line}" Height="1" Margin="0,0,0,10"/>
                        <StackPanel Orientation="Horizontal" Margin="0,3">
                            <Ellipse x:Name="StatusWingetDot" Width="6" Height="6" Fill="{StaticResource Amber}" VerticalAlignment="Center" Margin="0,0,8,0"/>
                            <TextBlock x:Name="StatusWinget" Text="winget  checking" Foreground="{StaticResource TextMuted}" FontFamily="{StaticResource Mono}" FontSize="10.5"/>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="0,3">
                            <Ellipse x:Name="StatusDiskDot" Width="6" Height="6" Fill="{StaticResource Amber}" VerticalAlignment="Center" Margin="0,0,8,0"/>
                            <TextBlock x:Name="StatusDisk" Text="disk  -- GB free" Foreground="{StaticResource TextMuted}" FontFamily="{StaticResource Mono}" FontSize="10.5"/>
                        </StackPanel>
                    </StackPanel>
                </Grid>
            </Border>

            <!-- Main -->
            <Grid Grid.Column="1">
                <Grid.RowDefinitions>
                    <RowDefinition Height="Auto"/>
                    <RowDefinition Height="*"/>
                    <RowDefinition Height="Auto"/>
                </Grid.RowDefinitions>

                <!-- Head -->
                <Border Grid.Row="0" BorderBrush="{StaticResource Line}" BorderThickness="0,0,0,1" Padding="22,16">
                    <StackPanel>
                        <TextBlock x:Name="HeadTitle" Text="Deploy new bench run" FontSize="17" FontWeight="SemiBold"/>
                        <TextBlock x:Name="HeadSub" FontFamily="{StaticResource Mono}" FontSize="11" Foreground="{StaticResource TextMuted}" Margin="0,3,0,0"/>
                        <TextBlock x:Name="UpdateBanner" Visibility="Collapsed" FontFamily="{StaticResource Mono}"
                                   FontSize="11" Foreground="{StaticResource Amber}" Margin="0,6,0,0"/>
                    </StackPanel>
                </Border>

                <!-- Content: two cards side by side + console below -->
                <Grid x:Name="BodyGrid" Grid.Row="1" Margin="22,14,22,0">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="*"/>
                        <RowDefinition Height="220"/>
                    </Grid.RowDefinitions>

                    <!-- Deploy view: apps + tweaks cards -->
                    <Grid x:Name="DeployTop" Grid.Row="0">
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="*"/>
                        </Grid.ColumnDefinitions>

                    <!-- Applications card -->
                    <Border Grid.Column="0" Background="{StaticResource BgPanel}" CornerRadius="8"
                            BorderBrush="{StaticResource Line}" BorderThickness="1" Padding="16,14" Margin="0,0,7,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <!-- Card head -->
                            <Grid Grid.Row="0" Margin="4,0,4,10">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock Text="APPLICATIONS" FontFamily="{StaticResource Mono}" FontSize="10.5" Foreground="{StaticResource TextMid}"/>
                                <TextBlock Grid.Column="1" x:Name="AppCount" Text="" FontFamily="{StaticResource Mono}" FontSize="10.5" Foreground="{StaticResource TextMuted}"/>
                            </Grid>

                            <!-- Chip filters -->
                            <WrapPanel Grid.Row="1" Margin="4,0,4,6">
                                <Button x:Name="ChipAll"      Style="{StaticResource ChipBtn}" Content="All"/>
                                <Button x:Name="ChipDiag"     Style="{StaticResource ChipBtn}" Content="Diagnostics"/>
                                <Button x:Name="ChipBench"    Style="{StaticResource ChipBtn}" Content="Benchmark"/>
                                <Button x:Name="ChipGpu"      Style="{StaticResource ChipBtn}" Content="GPU stress"/>
                                <Button x:Name="ChipStab"     Style="{StaticResource ChipBtn}" Content="Stability"/>
                                <Button x:Name="ChipInfo"     Style="{StaticResource ChipBtn}" Content="Info"/>
                                <Button x:Name="ChipCpu"      Style="{StaticResource ChipBtn}" Content="CPU torture"/>
                            </WrapPanel>

                            <!-- App list -->
                            <ScrollViewer Grid.Row="2" VerticalScrollBarVisibility="Auto" Margin="0,4,0,0">
                                <StackPanel x:Name="AppPanel"/>
                            </ScrollViewer>

                            <!-- Toggle all -->
                            <Button Grid.Row="3" x:Name="BtnSelectAllApps" Style="{StaticResource ChipBtn}" Content="Toggle all" Margin="4,10,0,0" HorizontalAlignment="Left"/>

                            <!-- Drivers -->
                            <StackPanel Grid.Row="4" Margin="0,12,0,0">
                                <TextBlock Text="DRIVERS &amp; UPDATES" FontFamily="{StaticResource Mono}" FontSize="10.5" Foreground="{StaticResource TextDim}" Margin="4,4,0,4"/>
                                <CheckBox x:Name="OptGpuDriver"     Content="Auto-detect &amp; install GPU driver (AMD / NVIDIA)"/>
                                <CheckBox x:Name="OptWindowsUpdate" Content="Install all Windows updates (incl. optional)"/>
                            </StackPanel>
                        </Grid>
                    </Border>

                    <!-- System tweaks card -->
                    <Border Grid.Column="1" Background="{StaticResource BgPanel}" CornerRadius="8"
                            BorderBrush="{StaticResource Line}" BorderThickness="1" Padding="16,14" Margin="7,0,0,0">
                        <ScrollViewer VerticalScrollBarVisibility="Auto">
                            <StackPanel>
                                <Grid Margin="4,0,4,10">
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <TextBlock Text="SYSTEM TWEAKS" FontFamily="{StaticResource Mono}" FontSize="10.5" Foreground="{StaticResource TextMid}"/>
                                    <TextBlock Grid.Column="1" x:Name="TweakCount" Text="" FontFamily="{StaticResource Mono}" FontSize="10.5" Foreground="{StaticResource TextMuted}"/>
                                </Grid>

                                <CheckBox x:Name="TweakPowerNever"       Style="{StaticResource ToggleSwitch}" Content="Power plan &#8594; never"                Tag="Display, sleep and hibernate timeouts set to never"/>
                                <CheckBox x:Name="TweakDisableHibernate" Style="{StaticResource ToggleSwitch}" Content="Disable hibernation"                   Tag="powercfg -h off (frees hiberfil.sys from system drive)"/>
                                <CheckBox x:Name="TweakClearDownloads"   Style="{StaticResource ToggleSwitchDestructive}" Content="Clear Downloads folder"      Tag="Destructive — wipes %USERPROFILE%\Downloads recursively"/>
                                <CheckBox x:Name="TweakEmptyRecycle"     Style="{StaticResource ToggleSwitch}" Content="Empty Recycle Bin"                     Tag="All drives"/>
                                <CheckBox x:Name="TweakClearBrowser"     Style="{StaticResource ToggleSwitch}" Content="Clear browser history"                 Tag="Edge, Chrome, Firefox &#8212; close browsers first"/>
                                <CheckBox x:Name="TweakStartGrid"        Style="{StaticResource ToggleSwitch}" Content="Start menu grid layout"                Tag="Windows 11 &#8212; requires sign out / in"/>
                                <CheckBox x:Name="TweakRemoveBloat"      Style="{StaticResource ToggleSwitchDestructive}" Content="Remove Windows bloat"        Tag="17 AppX packages + Quick Assist removed per-user (irreversible)"/>

                                <TextBlock Text="OUTPUT" FontFamily="{StaticResource Mono}" FontSize="10.5" Foreground="{StaticResource TextDim}" Margin="4,14,0,4"/>
                                <CheckBox x:Name="OptKeepTemp"    Style="{StaticResource ToggleSwitch}" Content="Keep temp files"      Tag="Retains downloads and extracted archives for debug"/>
                            </StackPanel>
                        </ScrollViewer>
                    </Border>

                    </Grid><!-- /DeployTop -->

                    <!-- Console -->
                    <Border x:Name="LogConsole" Grid.Row="1" Background="#0B0E14" CornerRadius="8"
                            BorderBrush="{StaticResource Line}" BorderThickness="1" Margin="0,14,0,0">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                            </Grid.RowDefinitions>
                            <Border Grid.Row="0" Background="#10141C" BorderBrush="{StaticResource Line}" BorderThickness="0,0,0,1" Padding="14,10">
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="*"/>
                                        <ColumnDefinition Width="Auto"/>
                                    </Grid.ColumnDefinitions>
                                    <StackPanel Grid.Column="0" Orientation="Horizontal">
                                        <Ellipse Width="8" Height="8" Fill="#E07A6B" Margin="0,0,5,0"/>
                                        <Ellipse Width="8" Height="8" Fill="#E0B15B" Margin="0,0,5,0"/>
                                        <Ellipse Width="8" Height="8" Fill="#8AE06B" Margin="0,0,14,0"/>
                                        <TextBlock Text="LOG" FontFamily="{StaticResource Mono}" FontSize="10.5" Foreground="{StaticResource TextHi}" VerticalAlignment="Center"/>
                                        <TextBlock x:Name="LogPathLabel" Text="   streaming to %TEMP%\pcbt.log" FontFamily="{StaticResource Mono}" FontSize="10.5" Foreground="{StaticResource TextMuted}" VerticalAlignment="Center"/>
                                    </StackPanel>
                                    <Button x:Name="BtnClearLog" Grid.Column="1" Style="{StaticResource ChipBtn}" Content="clear" Margin="0" VerticalAlignment="Center"/>
                                </Grid>
                            </Border>
                            <ScrollViewer x:Name="LogScroll" Grid.Row="1" VerticalScrollBarVisibility="Auto" Padding="14,10">
                                <StackPanel x:Name="LogOutput"/>
                            </ScrollViewer>
                        </Grid>
                    </Border>

                    <!-- Report view (overlays body when active) -->
                    <Border x:Name="ReportPanel" Grid.Row="0" Grid.RowSpan="2"
                            Background="{StaticResource BgPanel}" CornerRadius="8"
                            BorderBrush="{StaticResource Line}" BorderThickness="1"
                            Padding="22,18" Visibility="Collapsed">
                        <Grid>
                            <Grid.RowDefinitions>
                                <RowDefinition Height="Auto"/>
                                <RowDefinition Height="*"/>
                                <RowDefinition Height="Auto"/>
                            </Grid.RowDefinitions>

                            <!-- System info -->
                            <StackPanel Grid.Row="0" Margin="0,0,0,14">
                                <TextBlock Text="SYSTEM" FontFamily="{StaticResource Mono}" FontSize="10.5"
                                           Foreground="{StaticResource TextMid}" Margin="0,0,0,8"/>
                                <Grid>
                                    <Grid.ColumnDefinitions>
                                        <ColumnDefinition Width="110"/>
                                        <ColumnDefinition Width="*"/>
                                    </Grid.ColumnDefinitions>
                                    <Grid.RowDefinitions>
                                        <RowDefinition/><RowDefinition/><RowDefinition/>
                                        <RowDefinition/><RowDefinition/><RowDefinition/>
                                    </Grid.RowDefinitions>
                                    <TextBlock Grid.Row="0" Grid.Column="0" Text="Hostname"    FontFamily="{StaticResource Mono}" FontSize="11.5" Foreground="{StaticResource TextMuted}" Margin="0,2"/>
                                    <TextBlock Grid.Row="0" Grid.Column="1" x:Name="RepHostname" FontFamily="{StaticResource Mono}" FontSize="11.5" Foreground="{StaticResource TextHi}"   Margin="0,2"/>
                                    <TextBlock Grid.Row="1" Grid.Column="0" Text="CPU"         FontFamily="{StaticResource Mono}" FontSize="11.5" Foreground="{StaticResource TextMuted}" Margin="0,2"/>
                                    <TextBlock Grid.Row="1" Grid.Column="1" x:Name="RepCpu"     FontFamily="{StaticResource Mono}" FontSize="11.5" Foreground="{StaticResource TextHi}"   Margin="0,2"/>
                                    <TextBlock Grid.Row="2" Grid.Column="0" Text="GPU"         FontFamily="{StaticResource Mono}" FontSize="11.5" Foreground="{StaticResource TextMuted}" Margin="0,2"/>
                                    <TextBlock Grid.Row="2" Grid.Column="1" x:Name="RepGpu"     FontFamily="{StaticResource Mono}" FontSize="11.5" Foreground="{StaticResource TextHi}"   Margin="0,2"/>
                                    <TextBlock Grid.Row="3" Grid.Column="0" Text="RAM"         FontFamily="{StaticResource Mono}" FontSize="11.5" Foreground="{StaticResource TextMuted}" Margin="0,2"/>
                                    <TextBlock Grid.Row="3" Grid.Column="1" x:Name="RepRam"     FontFamily="{StaticResource Mono}" FontSize="11.5" Foreground="{StaticResource TextHi}"   Margin="0,2"/>
                                    <TextBlock Grid.Row="4" Grid.Column="0" Text="Motherboard" FontFamily="{StaticResource Mono}" FontSize="11.5" Foreground="{StaticResource TextMuted}" Margin="0,2"/>
                                    <TextBlock Grid.Row="4" Grid.Column="1" x:Name="RepMobo"    FontFamily="{StaticResource Mono}" FontSize="11.5" Foreground="{StaticResource TextHi}"   Margin="0,2"/>
                                    <TextBlock Grid.Row="5" Grid.Column="0" Text="Storage"     FontFamily="{StaticResource Mono}" FontSize="11.5" Foreground="{StaticResource TextMuted}" Margin="0,2"/>
                                    <TextBlock Grid.Row="5" Grid.Column="1" x:Name="RepStorage" FontFamily="{StaticResource Mono}" FontSize="11.5" Foreground="{StaticResource TextHi}"   Margin="0,2"/>
                                </Grid>
                            </StackPanel>

                            <!-- Live actions list -->
                            <Border Grid.Row="1" Background="#0B0E14" CornerRadius="6"
                                    BorderBrush="{StaticResource Line}" BorderThickness="1">
                                <Grid>
                                    <Grid.RowDefinitions>
                                        <RowDefinition Height="Auto"/>
                                        <RowDefinition Height="*"/>
                                    </Grid.RowDefinitions>
                                    <Border Grid.Row="0" Padding="14,8"
                                            BorderBrush="{StaticResource Line}" BorderThickness="0,0,0,1">
                                        <TextBlock Text="ACTIONS" FontFamily="{StaticResource Mono}" FontSize="10.5"
                                                   Foreground="{StaticResource TextHi}"/>
                                    </Border>
                                    <ScrollViewer Grid.Row="1" Padding="12,8" VerticalScrollBarVisibility="Auto">
                                        <StackPanel x:Name="ReportActions"/>
                                    </ScrollViewer>
                                </Grid>
                            </Border>

                            <!-- Footer note -->
                            <TextBlock Grid.Row="2" x:Name="ReportFooter" Margin="0,10,0,0"
                                       FontFamily="{StaticResource Mono}" FontSize="10.5"
                                       Foreground="{StaticResource TextMuted}"/>
                        </Grid>
                    </Border>
                </Grid>

                <!-- Footer -->
                <Border Grid.Row="2" Background="#0F1218" BorderBrush="{StaticResource Line}" BorderThickness="0,1,0,0" Padding="22,14">
                    <Grid>
                        <Grid.ColumnDefinitions>
                            <ColumnDefinition Width="*"/>
                            <ColumnDefinition Width="Auto"/>
                        </Grid.ColumnDefinitions>

                        <StackPanel Grid.Column="0" VerticalAlignment="Center">
                            <Grid Margin="0,0,0,8">
                                <Grid.ColumnDefinitions>
                                    <ColumnDefinition Width="Auto"/>
                                    <ColumnDefinition Width="*"/>
                                    <ColumnDefinition Width="Auto"/>
                                </Grid.ColumnDefinitions>
                                <TextBlock x:Name="StatusText" Text="Ready." Foreground="{StaticResource TextHi}" FontSize="13" FontWeight="SemiBold"/>
                                <TextBlock x:Name="ProgressPct" Grid.Column="2" Text="" FontFamily="{StaticResource Mono}" FontSize="11" Foreground="{StaticResource TextMid}"/>
                            </Grid>
                            <ProgressBar x:Name="ProgressBar" Minimum="0" Maximum="100" Value="0"/>
                            <TextBlock x:Name="RunSummary" Visibility="Collapsed" FontFamily="{StaticResource Mono}"
                                       FontSize="10.5" Foreground="{StaticResource TextMuted}" Margin="0,7,0,0"/>
                        </StackPanel>

                        <StackPanel Grid.Column="1" Orientation="Horizontal" VerticalAlignment="Center" Margin="24,0,0,0">
                            <Button x:Name="BtnQuit"      Style="{StaticResource BtnGhost}"  Content="Quit"          Margin="0,0,8,0"/>
                            <Button x:Name="BtnUninstall" Style="{StaticResource BtnDanger}" Content="Uninstall all" Margin="0,0,8,0"/>
                            <Button x:Name="BtnRun"       Style="{StaticResource BtnPrimary}" Content="Run pipeline"/>
                        </StackPanel>
                    </Grid>
                </Border>
            </Grid>
        </Grid>
    </Grid>
    </Border>
</Window>
'@

$reader = New-Object System.Xml.XmlNodeReader $xaml
try {
    $window = [Windows.Markup.XamlReader]::Load($reader)
} catch {
    Write-Host ""
    Write-Host "===== XAML PARSE FAILED =====" -ForegroundColor Red
    $ex = $_.Exception
    $depth = 0
    while ($ex) {
        Write-Host ("[{0}] {1}: {2}" -f $depth, $ex.GetType().Name, $ex.Message) -ForegroundColor Yellow
        if ($ex -is [System.Windows.Markup.XamlParseException]) {
            Write-Host ("     Line {0}, Pos {1}" -f $ex.LineNumber, $ex.LinePosition) -ForegroundColor Yellow
        }
        $ex = $ex.InnerException
        $depth++
    }
    Write-Host "=============================" -ForegroundColor Red
    throw
}

$controls = @{}
foreach ($n in 'AppPanel','LogOutput','LogScroll','BtnRun','BtnQuit','BtnUninstall','BtnSelectAllApps',
               'StatusText','ProgressBar','ProgressPct','VersionLabel',
               'BtnMin','BtnMax','BtnClose','BtnClearLog',
               'HeadTitle','HeadSub','AppCount','TweakCount','UpdateBanner','RunSummary','LogPathLabel',
               'SysCpu','SysGpu','SysRam','SysDisk','StatusWinget','StatusDisk',
               'StatusWingetDot','StatusDiskDot',
               'ChipAll','ChipDiag','ChipBench','ChipGpu','ChipStab','ChipInfo','ChipCpu',
               'TweakPowerNever','TweakDisableHibernate','TweakClearDownloads',
               'TweakEmptyRecycle','TweakClearBrowser','TweakStartGrid','TweakRemoveBloat',
               'OptGpuDriver','OptWindowsUpdate','OptKeepTemp',
               'NavDeploy','NavReport','NavLogs',
               'BodyGrid','DeployTop','LogConsole','ReportPanel',
               'RepHostname','RepCpu','RepGpu','RepRam','RepMobo','RepStorage',
               'ReportActions','ReportFooter') {
    $controls[$n] = $window.FindName($n)
}
$controls.VersionLabel.Text = $SCRIPT_VERSION

# Title bar window controls
$controls.BtnMin.Add_Click({ $window.WindowState = 'Minimized' })
$controls.BtnMax.Add_Click({
    if ($window.WindowState -eq 'Maximized') { $window.WindowState = 'Normal' }
    else { $window.WindowState = 'Maximized' }
})
$controls.BtnClose.Add_Click({ $window.Close() })

# Header subtitle + sidebar machine info
try {
    $osCap = (Get-CimInstance Win32_OperatingSystem -EA SilentlyContinue).Caption
    if ($osCap) { $osName = ($osCap -replace '^Microsoft\s+','').Trim() } else { $osName = 'Windows' }
    $controls.HeadSub.Text = "$env:COMPUTERNAME   |   $osName   |   $(Get-Date -Format 'dd/MM/yyyy')"
    $cpu = (Get-CimInstance Win32_Processor -EA SilentlyContinue | Select-Object -First 1).Name
    $gpu = (Get-CimInstance Win32_VideoController -EA SilentlyContinue | Where-Object { $_.Name -notmatch 'Virtual|Basic' } | Select-Object -First 1).Name
    $ram = [Math]::Round(((Get-CimInstance Win32_PhysicalMemory -EA SilentlyContinue | Measure-Object Capacity -Sum).Sum / 1GB), 0)
    $vol = Get-Volume -DriveLetter C -EA SilentlyContinue
    $freeGB = if ($vol) { [Math]::Round($vol.SizeRemaining/1GB,1) } else { $null }
    if ($cpu) { $controls.SysCpu.Text = $cpu }
    if ($gpu) { $controls.SysGpu.Text = $gpu }
    if ($ram) { $controls.SysRam.Text = "${ram} GB RAM" }
    $controls.RepHostname.Text = $env:COMPUTERNAME
    if ($cpu) { $controls.RepCpu.Text = $cpu }
    if ($gpu) { $controls.RepGpu.Text = $gpu }
    if ($ram) { $controls.RepRam.Text = "${ram} GB" }
    $mb = Get-CimInstance Win32_BaseBoard -EA SilentlyContinue
    if ($mb) { $controls.RepMobo.Text = "$($mb.Manufacturer) $($mb.Product)".Trim() }
    $diskLines = (Get-CimInstance Win32_DiskDrive -EA SilentlyContinue |
                  ForEach-Object { "$($_.Model) ($([Math]::Round($_.Size/1GB,0)) GB)" }) -join "`n"
    if ($diskLines) { $controls.RepStorage.Text = $diskLines }
    if ($freeGB) {
        $controls.SysDisk.Text = "C: ${freeGB} GB free"
        $controls.StatusDisk.Text = "disk  ${freeGB} GB free"
        $diskColor = if ($freeGB -lt 10) { '#E0B15B' } else { '#8AE06B' }
        $controls.StatusDiskDot.Fill = New-Object System.Windows.Media.SolidColorBrush(
            [System.Windows.Media.ColorConverter]::ConvertFromString($diskColor))
    }
    # winget health quick check (existence only — avoids triggering a source-cache refresh
    # that can lock the source DB and cause -1978335138 failures in the install pipeline)
    $wgOk = $null -ne (Get-Command winget -ErrorAction SilentlyContinue)
    $wgColor = if ($wgOk) { '#8AE06B' } else { '#E07A6B' }
    $controls.StatusWingetDot.Fill = New-Object System.Windows.Media.SolidColorBrush(
        [System.Windows.Media.ColorConverter]::ConvertFromString($wgColor))
    $controls.StatusWinget.Text = if ($wgOk) { 'winget  healthy' } else { 'winget  not found' }
} catch {}

$sync = [hashtable]::Synchronized(@{})
$sync.Window         = $window
$sync.Log            = $controls.LogOutput
$sync.LogScroll      = $controls.LogScroll
$sync.Status         = $controls.StatusText
$sync.Progress       = $controls.ProgressBar
$sync.ProgressPct    = $controls.ProgressPct
$sync.BtnRun         = $controls.BtnRun
$sync.BtnUninst      = $controls.BtnUninstall
$sync.BtnQuit        = $controls.BtnQuit
$sync.RunSummary     = $controls.RunSummary
$sync.UpdateBanner   = $controls.UpdateBanner
$sync.LogPath        = Join-Path $env:TEMP 'pcbt.log'
$controls.LogPathLabel.Text = "   streaming to $($sync.LogPath)"
$sync.AppCatalog     = $script:AppCatalog
$sync.BloatList      = $script:BloatList
$sync.SevenZipUrl    = $SCRIPT_7ZA_DIRECT
$sync.Mode           = $null
$sync.SelectedApps   = @()
$sync.SelectedTweaks = @{}
$sync.SelectedOpts   = @{}
$sync.KeepTemp       = $false
$sync.RunResults     = @()
$sync.ReportActions  = $controls.ReportActions
$sync.ReportFooter   = $controls.ReportFooter

# Windows Update needs to run in a real powershell.exe console host (PSWindowsUpdate's
# Get-WindowsUpdate -Install fails in background runspaces with "the host program does
# not support user interaction"). The subprocess script is defined here, not inside
# $pipelineCode, because PowerShell here-strings cannot be nested - a '@ at column 0
# inside the outer here-string would prematurely terminate it.
$sync.WinUpdateScript = @'
$ErrorActionPreference = 'Continue'
$ConfirmPreference     = 'None'
$ProgressPreference    = 'SilentlyContinue'
try {
    Import-Module PSWindowsUpdate -ErrorAction Stop
    $updates = Get-WindowsUpdate -MicrosoftUpdate -AcceptAll -Install -IgnoreReboot -Confirm:$false -Verbose 4>&1
    if ($updates) {
        foreach ($u in $updates) {
            if ($u -is [System.Management.Automation.VerboseRecord]) {
                Write-Output ("verbose: " + $u.Message)
            } elseif ($u.Title) {
                $st = if ($u.Result) { $u.Result } else { 'queued' }
                Write-Output ("[{0}] {1}" -f $st, $u.Title)
            } else {
                Write-Output ([string]$u)
            }
        }
    } else {
        Write-Output "No updates available."
    }
    exit 0
} catch {
    Write-Error $_.Exception.Message
    exit 1
}
'@

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $stamp = Get-Date -Format 'HH:mm:ss'
    $line  = "[$stamp] [$Level] $Message"
    Add-Content -Path $sync.LogPath -Value $line -ErrorAction SilentlyContinue
    $logLine = $line; $logLevel = $Level
    $sync.Log.Dispatcher.Invoke([action]{
        $clr = switch ($logLevel) {
            'OK'    { '#8AE06B' }
            'WARN'  { '#E0B15B' }
            'ERROR' { '#E07A6B' }
            default { '#8891A0' }
        }
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text        = $logLine
        $tb.Foreground  = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($clr))
        $tb.FontFamily  = New-Object System.Windows.Media.FontFamily('Cascadia Mono, Consolas, Courier New')
        $tb.FontSize    = 11.5
        $tb.TextWrapping = 'Wrap'
        [void]$sync.Log.Children.Add($tb)
        $sync.LogScroll.ScrollToBottom()
    })
}

$appCheckboxes = @{}
foreach ($app in $script:AppCatalog) {
    # Row: custom checkbox with two-line content (name + mono sub-line with id / source / category)
    $cb = New-Object System.Windows.Controls.CheckBox
    $cb.IsChecked = ($script:DefaultChecked -contains $app.Id)
    $cb.Tag       = $app.Category  # used for chip filtering

    $stack = New-Object System.Windows.Controls.StackPanel
    $line1 = New-Object System.Windows.Controls.TextBlock
    $line1.Text       = $app.Name
    $line1.FontSize   = 13
    $line1.Foreground = [System.Windows.Media.Brushes]::Gainsboro
    [void]$stack.Children.Add($line1)

    $line2 = New-Object System.Windows.Controls.TextBlock
    $line2.Text       = "$($app.Id)   $($app.Source)   $($app.Category)"
    $line2.FontFamily = New-Object System.Windows.Media.FontFamily('Cascadia Mono, Consolas, Courier New')
    $line2.FontSize   = 10.5
    $line2.Foreground = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString('#8891A0')))
    $line2.Margin     = '0,2,0,0'
    [void]$stack.Children.Add($line2)

    $cb.Content = $stack
    [void]$controls.AppPanel.AddChild($cb)
    $appCheckboxes[$app.Id] = $cb
}

# Selection-count updater
$updateCounts = {
    $sel = ($appCheckboxes.Values | Where-Object { $_.IsChecked }).Count
    $controls.AppCount.Text = "$sel / $($script:AppCatalog.Count) selected"
    $tweakBoxes = @($controls.TweakPowerNever, $controls.TweakDisableHibernate, $controls.TweakClearDownloads,
                    $controls.TweakEmptyRecycle, $controls.TweakClearBrowser, $controls.TweakStartGrid, $controls.TweakRemoveBloat)
    $tw = ($tweakBoxes | Where-Object { $_.IsChecked }).Count
    $controls.TweakCount.Text = "$tw enabled"
}
foreach ($cb in $appCheckboxes.Values) {
    $cb.Add_Checked($updateCounts); $cb.Add_Unchecked($updateCounts)
}
foreach ($n in 'TweakPowerNever','TweakDisableHibernate','TweakClearDownloads',
               'TweakEmptyRecycle','TweakClearBrowser','TweakStartGrid','TweakRemoveBloat') {
    $controls[$n].Add_Checked($updateCounts); $controls[$n].Add_Unchecked($updateCounts)
}
& $updateCounts

# Chip filtering
$currentFilter = [ref]'All'
$chipBtns = @{
    'All'         = $controls.ChipAll
    'Diagnostics' = $controls.ChipDiag
    'Benchmark'   = $controls.ChipBench
    'GPU stress'  = $controls.ChipGpu
    'Stability'   = $controls.ChipStab
    'Info'        = $controls.ChipInfo
    'CPU torture' = $controls.ChipCpu
}
$applyFilter = {
    param($cat)
    $currentFilter.Value = $cat
    foreach ($kvp in $chipBtns.GetEnumerator()) {
        $isActive = ($kvp.Key -eq $cat)
        $bg = if ($isActive) { '#2A3040' } else { 'Transparent' }
        $fg = if ($isActive) { '#F1F3F6' } else { '#8891A0' }
        $kvp.Value.Background = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($bg)))
        $kvp.Value.Foreground = (New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($fg)))
    }
    foreach ($cb in $appCheckboxes.Values) {
        $cb.Visibility = if ($cat -eq 'All' -or $cb.Tag -eq $cat) { 'Visible' } else { 'Collapsed' }
    }
}
foreach ($kvp in $chipBtns.GetEnumerator()) {
    $k = $kvp.Key
    $kvp.Value.Add_Click({ & $applyFilter $k }.GetNewClosure())
}
& $applyFilter 'All'

function Invoke-PreflightChecks {
    param([int]$MinFreeGB = 10)
    Write-Log "Running pre-flight checks..."
    $problems = @()
    try {
        $ok = Test-Connection -ComputerName '1.1.1.1' -Count 1 -Quiet -ErrorAction Stop
        if ($ok) { Write-Log "Internet: reachable." 'OK' } else { $problems += "No internet connectivity." }
    } catch { $problems += "Internet check failed: $_" }
    try {
        $vol = Get-Volume -DriveLetter C -ErrorAction Stop
        $freeGB = [Math]::Round($vol.SizeRemaining / 1GB, 1)
        if ($freeGB -lt $MinFreeGB) { $problems += "Only ${freeGB} GB free on C: (need ${MinFreeGB}+ GB)." }
        else { Write-Log "Disk C: ${freeGB} GB free." 'OK' }
    } catch { Write-Log "Disk check skipped: $_" 'WARN' }
    if (Get-Command winget -ErrorAction SilentlyContinue) {
        try {
            $null = & winget source list 2>&1
            if ($LASTEXITCODE -eq 0) { Write-Log "winget sources: healthy." 'OK' }
            else { Write-Log "winget sources returned exit $LASTEXITCODE." 'WARN' }
        } catch { Write-Log "winget source check failed: $_" 'WARN' }
    } else { $problems += "winget not found on PATH." }
    if ($problems.Count -gt 0) {
        foreach ($p in $problems) { Write-Log $p 'ERROR' }
        return $false
    }
    return $true
}

function Invoke-SelfUpdateCheck {
    if (-not $PSCommandPath -or -not (Test-Path $PSCommandPath)) {
        Write-Log "Loaded from web stream; skipping self-update check."; return
    }
    try {
        $local  = Get-FileHash -Path $PSCommandPath -Algorithm SHA256
        $remote = Invoke-WebRequest -UseBasicParsing -Uri $SCRIPT_RAW_URL -TimeoutSec 10
        $bytes  = [System.Text.Encoding]::UTF8.GetBytes($remote.Content)
        $sha    = [System.Security.Cryptography.SHA256]::Create()
        $rHash  = -join ($sha.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') })
        if ($local.Hash.ToLower() -ne $rHash.ToLower()) {
            Write-Log "A newer version is available at $SCRIPT_RAW_URL" 'INFO'
            $controls.UpdateBanner.Text       = "Update available — run  irm fay.digital/pbt | iex  to get the latest version"
            $controls.UpdateBanner.Visibility = 'Visible'
        } else { Write-Log "Running latest version ($SCRIPT_VERSION)." 'OK' }
    } catch { Write-Log "Self-update check skipped: $_" 'WARN' }
}

$pipelineCode = @'
function Write-UiLog {
    param([string]$Message, [string]$Level = 'INFO')
    $stamp = Get-Date -Format 'HH:mm:ss'
    $line  = "[$stamp] [$Level] $Message"
    Add-Content -Path $sync.LogPath -Value $line -ErrorAction SilentlyContinue
    $logLine = $line; $logLevel = $Level
    $sync.Log.Dispatcher.Invoke([action]{
        $clr = switch ($logLevel) {
            'OK'    { '#8AE06B' }
            'WARN'  { '#E0B15B' }
            'ERROR' { '#E07A6B' }
            default { '#8891A0' }
        }
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text         = $logLine
        $tb.Foreground   = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString($clr))
        $tb.FontFamily   = New-Object System.Windows.Media.FontFamily('Cascadia Mono, Consolas, Courier New')
        $tb.FontSize     = 11.5
        $tb.TextWrapping = 'Wrap'
        [void]$sync.Log.Children.Add($tb)
        $sync.LogScroll.ScrollToBottom()
    })
}
function Write-UiLogReplace {
    param([string]$Message)
    $stamp = Get-Date -Format 'HH:mm:ss'
    $line  = "[$stamp] [INFO] $Message"
    $logLine = $line
    $sync.Log.Dispatcher.Invoke([action]{
        if ($sync.Log.Children.Count -gt 0) {
            $sync.Log.Children.RemoveAt($sync.Log.Children.Count - 1)
        }
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text         = $logLine
        $tb.Foreground   = New-Object System.Windows.Media.SolidColorBrush([System.Windows.Media.ColorConverter]::ConvertFromString('#8891A0'))
        $tb.FontFamily   = New-Object System.Windows.Media.FontFamily('Cascadia Mono, Consolas, Courier New')
        $tb.FontSize     = 11.5
        $tb.TextWrapping = 'Wrap'
        [void]$sync.Log.Children.Add($tb)
        $sync.LogScroll.ScrollToBottom()
    })
}
function Set-UiStatus   { param([string]$t) $sync.Status.Dispatcher.Invoke([action]{ $sync.Status.Text = $t }) }
function Set-UiProgress {
    param([double]$v)
    $sync.Progress.Dispatcher.Invoke([action]{
        $sync.Progress.Value = $v
        if ($v -le 0) { $sync.ProgressPct.Text = '' }
        else          { $sync.ProgressPct.Text = ("{0:N0}%" -f $v) }
    })
}
function Set-UiBusy {
    param([bool]$busy)
    $sync.BtnRun.Dispatcher.Invoke([action]{
        $sync.BtnRun.IsEnabled    = -not $busy
        $sync.BtnUninst.IsEnabled = -not $busy
        $sync.BtnQuit.IsEnabled   = -not $busy
    })
}
function Add-Result { param($Name, $Action, $Status, $Detail='')
    $sync.RunResults += [pscustomobject]@{ App=$Name; Action=$Action; Status=$Status; Detail=$Detail }
    $rN=$Name; $rA=$Action; $rS=$Status; $rD=$Detail
    $sync.ReportActions.Dispatcher.Invoke([action]{
        $clr = switch ($rS) {
            'OK'    { '#8AE06B' }
            'WARN'  { '#E0B15B' }
            'FAIL'  { '#E07A6B' }
            default { '#8891A0' }
        }
        $tb = New-Object System.Windows.Controls.TextBlock
        $tb.Text       = ('[{0,-4}] {1,-22} {2,-10} {3}' -f $rS, $rN, $rA, $rD)
        $tb.Foreground = New-Object System.Windows.Media.SolidColorBrush(
            [System.Windows.Media.ColorConverter]::ConvertFromString($clr))
        $tb.FontFamily = New-Object System.Windows.Media.FontFamily('Cascadia Mono, Consolas, Courier New')
        $tb.FontSize   = 11.5
        [void]$sync.ReportActions.Children.Add($tb)
    })
}

function Invoke-StreamingDownload {
    param([Parameter(Mandatory)][string]$Url, [Parameter(Mandatory)][string]$OutFile, [string]$AppName = 'file')
    Add-Type -AssemblyName System.Net.Http
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072
    $handler = New-Object System.Net.Http.HttpClientHandler
    $handler.AllowAutoRedirect = $true
    $client = New-Object System.Net.Http.HttpClient($handler)
    $client.Timeout = [TimeSpan]::FromMinutes(30)
    try {
        $response = $client.GetAsync($Url, [System.Net.Http.HttpCompletionOption]::ResponseHeadersRead).GetAwaiter().GetResult()
        if (-not $response.IsSuccessStatusCode) { throw "HTTP $([int]$response.StatusCode) $($response.ReasonPhrase)" }
        $ct = $response.Content.Headers.ContentType
        if ($ct -and $ct.MediaType -match '^(text/|application/json|application/xml)') {
            throw "Server returned $($ct.MediaType) (likely an error page)."
        }
        $totalBytes = $response.Content.Headers.ContentLength
        $totalMB    = if ($totalBytes) { [Math]::Round($totalBytes / 1MB, 1) } else { $null }
        if ($totalBytes) { Write-UiLog "Downloading $AppName ($totalMB MB)..." }
        else             { Write-UiLog "Downloading $AppName (size unknown)..." }
        $sourceStream = $response.Content.ReadAsStreamAsync().GetAwaiter().GetResult()
        $destStream   = [System.IO.File]::Create($OutFile)
        $buffer       = New-Object byte[] (1MB)
        $totalRead    = 0L
        $lastReport   = [DateTime]::UtcNow
        $startTime    = [DateTime]::UtcNow
        $firstLine    = $true
        try {
            while ($true) {
                $read = $sourceStream.Read($buffer, 0, $buffer.Length)
                if ($read -le 0) { break }
                $destStream.Write($buffer, 0, $read)
                $totalRead += $read
                $now = [DateTime]::UtcNow
                if (($now - $lastReport).TotalMilliseconds -ge 1000) {
                    $elapsed = ($now - $startTime).TotalSeconds
                    $speedMB = if ($elapsed -gt 0) { [Math]::Round(($totalRead / 1MB) / $elapsed, 1) } else { 0 }
                    $doneMB  = [Math]::Round($totalRead / 1MB, 1)
                    if ($totalBytes) {
                        $pct = ($totalRead / $totalBytes) * 100
                        $msg = "Downloading ${AppName}: $doneMB / $totalMB MB ({0:N1}%) - $speedMB MB/s" -f $pct
                        Set-UiProgress $pct
                    } else {
                        $msg = "Downloading ${AppName}: $doneMB MB - $speedMB MB/s"
                    }
                    if ($firstLine) { Write-UiLog $msg; $firstLine = $false } else { Write-UiLogReplace $msg }
                    $lastReport = $now
                }
            }
        }
        finally { $destStream.Flush(); $destStream.Close(); $sourceStream.Close() }
        if ($totalBytes -and $totalRead -ne $totalBytes) { throw "Truncated: got $totalRead bytes, expected $totalBytes." }
        Write-UiLog "Download complete: $([Math]::Round($totalRead / 1MB, 1)) MB." 'OK'
    }
    finally { $client.Dispose(); $handler.Dispose() }
}

function Test-IsZipFile {
    param([string]$Path)
    if (-not (Test-Path $Path)) { return $false }
    try {
        $fs = [System.IO.File]::OpenRead($Path)
        $hdr = New-Object byte[] 4
        $n = $fs.Read($hdr, 0, 4)
        $fs.Close()
        return ($n -eq 4 -and $hdr[0] -eq 0x50 -and $hdr[1] -eq 0x4B -and $hdr[2] -eq 0x03 -and $hdr[3] -eq 0x04)
    } catch { return $false }
}

# Fetch a standalone 7za.exe (~1 MB). Tries multiple known-stable sources
# because 7-zip.org doesn't publish a raw .exe, only .msi/.7z self-extractors.
# Returns path to 7za.exe on success, $null on total failure.
function Get-StandaloneSevenZip {
    # If the system already has 7z.exe installed, use it.
    $existing = Get-Command 7z.exe -ErrorAction SilentlyContinue
    if ($existing) { Write-UiLog "Using existing 7z.exe at $($existing.Source)." ; return $existing.Source }

    # Cache per-session so repeated zip installs don't re-download.
    if ($sync.SevenZipPath -and (Test-Path $sync.SevenZipPath)) { return $sync.SevenZipPath }

    $dest = Join-Path $env:TEMP 'pbt_7za.exe'
    $sources = @(
        # Chocolatey's CDN hosts a raw 7za.exe - stable and widely mirrored.
        'https://chocolatey.org/7za.exe',
        # GitHub mirror of 7-Zip source tree - has a compiled 7za.exe under bin/
        'https://github.com/mcmilk/7zip-zstd/releases/download/24.08-v1.5.6-R3/7z2408-extra.7z'
    )
    foreach ($url in $sources) {
        try {
            Write-UiLog "Fetching 7-Zip standalone from $url..."
            [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072
            Invoke-WebRequest -UseBasicParsing -Uri $url -OutFile $dest -TimeoutSec 60 -ErrorAction Stop
            # Sanity check: is it an EXE? (MZ header)
            $fs = [System.IO.File]::OpenRead($dest)
            $hdr = New-Object byte[] 2
            [void]$fs.Read($hdr, 0, 2)
            $fs.Close()
            if ($hdr[0] -eq 0x4D -and $hdr[1] -eq 0x5A) {
                Write-UiLog "7-Zip standalone ready at $dest." 'OK'
                $sync.SevenZipPath = $dest
                return $dest
            } else {
                Write-UiLog "Downloaded file not a PE executable; trying next source..." 'WARN'
                Remove-Item $dest -Force -ErrorAction SilentlyContinue
            }
        } catch {
            Write-UiLog "7-Zip fetch from $url failed: $_" 'WARN'
            Remove-Item $dest -Force -ErrorAction SilentlyContinue
        }
    }
    Write-UiLog "Could not obtain 7-Zip standalone from any source." 'ERROR'
    return $null
}

# Extraction chain: tar.exe first (ships with Windows, fast, no download),
# then 7-Zip (handles Deflate64/LZMA - downloaded on demand), then
# Expand-Archive (last resort, Deflate only).
function Expand-ZipArchive {
    param([Parameter(Mandatory)][string]$ZipPath, [Parameter(Mandatory)][string]$DestinationPath)
    New-Item -ItemType Directory -Path $DestinationPath -Force | Out-Null

    # --- Try tar.exe ---
    $tar = Get-Command tar.exe -ErrorAction SilentlyContinue
    if ($tar) {
        Write-UiLog "Extracting with tar.exe..."
        $stderrFile = Join-Path $env:TEMP ("pbt_tar_err_" + [Guid]::NewGuid().ToString('N').Substring(0,8) + ".log")
        $p = Start-Process -FilePath $tar.Source -ArgumentList @('-xf', "`"$ZipPath`"", '-C', "`"$DestinationPath`"") `
            -Wait -PassThru -NoNewWindow -RedirectStandardError $stderrFile
        $err = if (Test-Path $stderrFile) { Get-Content $stderrFile -Raw -ErrorAction SilentlyContinue } else { $null }
        Remove-Item $stderrFile -Force -ErrorAction SilentlyContinue
        if ($p.ExitCode -eq 0) { Write-UiLog "Extraction OK (tar.exe)." 'OK'; return }

        $errOneLine = $err -replace '\s+', ' '
        Write-UiLog "tar.exe failed (exit $($p.ExitCode)): $errOneLine" 'WARN'

        # If tar complained about compression method, clear the partial
        # extraction and fall through to 7-Zip. Otherwise, still try.
        Get-ChildItem $DestinationPath -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }

    # --- Try 7-Zip ---
    $sevenZip = Get-StandaloneSevenZip
    if ($sevenZip) {
        Write-UiLog "Extracting with 7-Zip ($sevenZip)..."
        $stderrFile = Join-Path $env:TEMP ("pbt_7z_err_" + [Guid]::NewGuid().ToString('N').Substring(0,8) + ".log")
        $stdoutFile = Join-Path $env:TEMP ("pbt_7z_out_" + [Guid]::NewGuid().ToString('N').Substring(0,8) + ".log")
        # -y: assume yes to prompts, -bb0: minimal output, -o: output dir (no space after)
        $p = Start-Process -FilePath $sevenZip -ArgumentList @('x', "-y", "-bb0", "-o`"$DestinationPath`"", "`"$ZipPath`"") `
            -Wait -PassThru -NoNewWindow -RedirectStandardError $stderrFile -RedirectStandardOutput $stdoutFile
        $err = if (Test-Path $stderrFile) { Get-Content $stderrFile -Raw -ErrorAction SilentlyContinue } else { $null }
        $out = if (Test-Path $stdoutFile) { Get-Content $stdoutFile -Raw -ErrorAction SilentlyContinue } else { $null }
        Remove-Item $stderrFile, $stdoutFile -Force -ErrorAction SilentlyContinue
        if ($p.ExitCode -eq 0) { Write-UiLog "Extraction OK (7-Zip)." 'OK'; return }
        Write-UiLog "7-Zip failed (exit $($p.ExitCode)): $($err -replace '\s+', ' ') $($out -replace '\s+', ' ')" 'WARN'
        Get-ChildItem $DestinationPath -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }

    # --- Last resort: Expand-Archive ---
    Write-UiLog "Falling back to Expand-Archive..."
    Expand-Archive -Path $ZipPath -DestinationPath $DestinationPath -Force -ErrorAction Stop
    Write-UiLog "Extraction OK (Expand-Archive)." 'OK'
}

function Start-SilentInstaller {
    param([Parameter(Mandatory)][string]$ExePath, [Parameter(Mandatory)][string]$Arguments, [Parameter(Mandatory)][string]$AppName)
    $workDir  = Split-Path -Parent $ExePath
    $stdoutFile = Join-Path $env:TEMP ("pbt_stdout_" + [Guid]::NewGuid().ToString('N').Substring(0,8) + ".log")
    $stderrFile = Join-Path $env:TEMP ("pbt_stderr_" + [Guid]::NewGuid().ToString('N').Substring(0,8) + ".log")
    Write-UiLog "Running $([IO.Path]::GetFileName($ExePath)) in '$workDir' with args '$Arguments'..."
    $p = Start-Process -FilePath $ExePath -ArgumentList $Arguments `
        -WorkingDirectory $workDir -RedirectStandardOutput $stdoutFile -RedirectStandardError $stderrFile `
        -Wait -PassThru -NoNewWindow
    $stdout = if (Test-Path $stdoutFile) { Get-Content $stdoutFile -Raw -ErrorAction SilentlyContinue } else { $null }
    $stderr = if (Test-Path $stderrFile) { Get-Content $stderrFile -Raw -ErrorAction SilentlyContinue } else { $null }
    Remove-Item $stdoutFile, $stderrFile -Force -ErrorAction SilentlyContinue
    if ($stdout -and $stdout.Trim()) { Write-UiLog "$AppName stdout: $($stdout.Trim() -replace "`r?`n", ' | ')" 'INFO' }
    if ($stderr -and $stderr.Trim()) { Write-UiLog "$AppName stderr: $($stderr.Trim() -replace "`r?`n", ' | ')" 'WARN' }
    return $p.ExitCode
}

function Test-ChocoInstalled { [bool](Get-Command choco.exe -ErrorAction SilentlyContinue) }
function Install-Chocolatey {
    Write-UiLog "Chocolatey not found. Installing..."
    Set-UiStatus "Installing Chocolatey..."
    Set-ExecutionPolicy Bypass -Scope Process -Force
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072
    Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" +
                [System.Environment]::GetEnvironmentVariable("Path","User")
    if (-not (Test-ChocoInstalled)) { throw "Chocolatey install did not expose choco.exe." }
    Write-UiLog "Chocolatey installed." 'OK'
}

function Invoke-WingetInstall { param($App)
    Write-UiLog "Installing $($App.Name) via winget ($($App.Id))..."
    $p = Start-Process winget -ArgumentList @('install','--id',$App.Id,'--silent',
        '--accept-package-agreements','--accept-source-agreements','--disable-interactivity') -Wait -PassThru -NoNewWindow
    switch ($p.ExitCode) {
        0           { Write-UiLog "$($App.Name) installed." 'OK';                   Add-Result $App.Name 'install' 'OK' }
        -1978335189 { Write-UiLog "$($App.Name) already installed." 'OK';            Add-Result $App.Name 'install' 'OK' 'already installed' }
        -1978335212 { Write-UiLog "$($App.Name) - id not in winget." 'ERROR';        Add-Result $App.Name 'install' 'FAIL' 'id not found' }
        -1978335138 { Write-UiLog "$($App.Name) - winget source error (try: winget source reset --force)." 'ERROR'
                      Add-Result $App.Name 'install' 'FAIL' 'source error 0x8A15005E' }
        default     { Write-UiLog "$($App.Name) exit $($p.ExitCode)." 'WARN';        Add-Result $App.Name 'install' 'WARN' "exit $($p.ExitCode)" }
    }
}
function Invoke-WingetUninstall { param($App)
    Write-UiLog "Uninstalling $($App.Name) via winget..."
    $p = Start-Process winget -ArgumentList @('uninstall','--id',$App.Id,'--silent',
        '--accept-source-agreements','--disable-interactivity') -Wait -PassThru -NoNewWindow
    switch ($p.ExitCode) {
        0           { Write-UiLog "$($App.Name) uninstalled." 'OK';                  Add-Result $App.Name 'uninstall' 'OK' }
        -1978335212 { Write-UiLog "$($App.Name) not installed, skipped." 'INFO';     Add-Result $App.Name 'uninstall' 'SKIP' 'not present' }
        default     { Write-UiLog "$($App.Name) exit $($p.ExitCode)." 'WARN';        Add-Result $App.Name 'uninstall' 'WARN' "exit $($p.ExitCode)" }
    }
}

function Invoke-ChocoInstall { param($App)
    if (-not (Test-ChocoInstalled)) { Install-Chocolatey }
    Write-UiLog "Installing $($App.Name) via Chocolatey..."
    $p = Start-Process choco -ArgumentList @('install',$App.Id,'-y','--no-progress','--limit-output') -Wait -PassThru -NoNewWindow
    if ($p.ExitCode -eq 0) { Write-UiLog "$($App.Name) installed." 'OK'; Add-Result $App.Name 'install' 'OK' }
    else { Write-UiLog "$($App.Name) exit $($p.ExitCode)." 'WARN'; Add-Result $App.Name 'install' 'WARN' "exit $($p.ExitCode)" }
}
function Invoke-ChocoUninstall { param($App)
    if (-not (Test-ChocoInstalled)) { Write-UiLog "Chocolatey not present, skipping $($App.Name)." 'INFO'; Add-Result $App.Name 'uninstall' 'SKIP'; return }
    Write-UiLog "Uninstalling $($App.Name) via Chocolatey..."
    $p = Start-Process choco -ArgumentList @('uninstall',$App.Id,'-y','--no-progress','--limit-output') -Wait -PassThru -NoNewWindow
    if ($p.ExitCode -eq 0) { Write-UiLog "$($App.Name) uninstalled." 'OK'; Add-Result $App.Name 'uninstall' 'OK' }
    else { Write-UiLog "$($App.Name) exit $($p.ExitCode) (may be absent)." 'INFO'; Add-Result $App.Name 'uninstall' 'SKIP' "exit $($p.ExitCode)" }
}

function Invoke-DirectInstall { param($App)
    if (-not $App.DownloadUrl) { Write-UiLog "$($App.Name): no DownloadUrl." 'ERROR'; Add-Result $App.Name 'install' 'FAIL' 'no URL'; return }
    $tmp = Join-Path $env:TEMP ("pbt_" + [IO.Path]::GetFileName($App.DownloadUrl))
    Set-UiStatus "Downloading $($App.Name)..."
    try {
        Invoke-StreamingDownload -Url $App.DownloadUrl -OutFile $tmp -AppName $App.Name
    } catch {
        Write-UiLog "$($App.Name) download failed: $_" 'ERROR'
        Add-Result $App.Name 'install' 'FAIL' 'download failed'
        if (-not $sync.KeepTemp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
        return
    }
    Set-UiStatus "Installing $($App.Name)..."
    $silent = if ($App.SilentArgs) { $App.SilentArgs } else { '/S' }
    $exit = Start-SilentInstaller -ExePath $tmp -Arguments $silent -AppName $App.Name
    if (-not $sync.KeepTemp) { Remove-Item $tmp -Force -ErrorAction SilentlyContinue }
    else { Write-UiLog "Kept installer: $tmp" 'INFO' }
    if ($exit -eq 0) { Write-UiLog "$($App.Name) installed." 'OK'; Add-Result $App.Name 'install' 'OK' }
    else { Write-UiLog "$($App.Name) installer exit $exit (0x$('{0:X8}' -f $exit))." 'WARN'; Add-Result $App.Name 'install' 'WARN' "exit $exit" }
}

function Invoke-ZipInstall { param($App)
    if (-not $App.DownloadUrl)     { Write-UiLog "$($App.Name): no DownloadUrl." 'ERROR';     Add-Result $App.Name 'install' 'FAIL' 'no URL';       return }
    if (-not $App.SetupExecutable) { Write-UiLog "$($App.Name): no SetupExecutable." 'ERROR'; Add-Result $App.Name 'install' 'FAIL' 'no setup exe'; return }
    $tmpZip = Join-Path $env:TEMP ("pbt_" + [IO.Path]::GetFileName($App.DownloadUrl))
    $tmpDir = Join-Path $env:TEMP ("pbt_extract_" + [Guid]::NewGuid().ToString('N').Substring(0,8))
    try {
        Set-UiStatus "Downloading $($App.Name)..."
        Invoke-StreamingDownload -Url $App.DownloadUrl -OutFile $tmpZip -AppName $App.Name
        if (-not (Test-IsZipFile $tmpZip)) { throw "Downloaded file is not a valid zip." }
        Set-UiStatus "Extracting $($App.Name)..."
        Write-UiLog "Extracting $($App.Name) to $tmpDir ..."
        Expand-ZipArchive -ZipPath $tmpZip -DestinationPath $tmpDir
        $setup = Get-ChildItem -Path $tmpDir -Filter $App.SetupExecutable -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $setup) { throw "'$($App.SetupExecutable)' not found in extracted content at $tmpDir" }
        Set-UiStatus "Installing $($App.Name)..."
        $silent = if ($App.SilentArgs) { $App.SilentArgs } else { '/S' }
        $exit = Start-SilentInstaller -ExePath $setup.FullName -Arguments $silent -AppName $App.Name
        if ($exit -eq 0) { Write-UiLog "$($App.Name) installed." 'OK'; Add-Result $App.Name 'install' 'OK' }
        else { $hex = '0x{0:X8}' -f $exit; Write-UiLog "$($App.Name) installer exit $exit ($hex)." 'WARN'; Add-Result $App.Name 'install' 'WARN' "exit $exit ($hex)" }
    }
    catch { Write-UiLog "$($App.Name) install error: $_" 'ERROR'; Add-Result $App.Name 'install' 'FAIL' $_.Exception.Message }
    finally {
        if ($sync.KeepTemp) { Write-UiLog "Kept zip: $tmpZip" 'INFO'; Write-UiLog "Kept extract: $tmpDir" 'INFO' }
        else { Remove-Item $tmpZip -Force -ErrorAction SilentlyContinue; Remove-Item $tmpDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
}

function Invoke-RegistryUninstall { param($App)
    $match = $App.UninstallRegistryMatch
    if (-not $match) { Write-UiLog "$($App.Name): no UninstallRegistryMatch." 'WARN'; return }
    $keys = @('HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall',
              'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')
    $found = $false
    foreach ($k in $keys) {
        if (-not (Test-Path $k)) { continue }
        Get-ChildItem $k -ErrorAction SilentlyContinue | ForEach-Object {
            $dn = $_.GetValue('DisplayName')
            if ($dn -and $dn -like "*$match*") {
                $us = $_.GetValue('QuietUninstallString'); if (-not $us) { $us = $_.GetValue('UninstallString') }
                if ($us) {
                    Write-UiLog "Uninstalling '$dn'..."
                    Start-Process -FilePath cmd.exe -ArgumentList "/c",($us + " /S") -Wait -NoNewWindow
                    Write-UiLog "$dn uninstalled." 'OK'; Add-Result $App.Name 'uninstall' 'OK' $dn; $found = $true
                }
            }
        }
    }
    if (-not $found) { Write-UiLog "$($App.Name) not found in registry, skipped." 'INFO'; Add-Result $App.Name 'uninstall' 'SKIP' 'not present' }
}

function Remove-BloatPackage {
    param([string]$Name, [string]$Match)
    $pkgs = Get-AppxPackage -Name "*$Match*" -ErrorAction SilentlyContinue
    if (-not $pkgs) { Write-UiLog "${Name}: not installed." 'INFO'; Add-Result $Name 'debloat' 'SKIP' 'not present'; return }
    foreach ($pkg in $pkgs) {
        try {
            Remove-AppxPackage -Package $pkg.PackageFullName -ErrorAction Stop
            Write-UiLog "$Name removed ($($pkg.Name))." 'OK'
            Add-Result $Name 'debloat' 'OK' $pkg.Name
        } catch {
            Write-UiLog "$Name removal failed: $_" 'WARN'
            Add-Result $Name 'debloat' 'WARN' "$_"
        }
    }
}
function Remove-QuickAssist {
    try {
        $cap = Get-WindowsCapability -Online -ErrorAction Stop | Where-Object { $_.Name -like 'App.Support.QuickAssist*' -and $_.State -eq 'Installed' }
        if (-not $cap) { Write-UiLog "Quick Assist: not installed." 'INFO'; Add-Result 'Quick Assist' 'debloat' 'SKIP' 'not present'; return }
        Write-UiLog "Removing Quick Assist capability..."
        Remove-WindowsCapability -Online -Name $cap.Name -ErrorAction Stop | Out-Null
        Write-UiLog "Quick Assist removed." 'OK'
        Add-Result 'Quick Assist' 'debloat' 'OK'
    } catch {
        Write-UiLog "Quick Assist removal failed: $_" 'WARN'
        Add-Result 'Quick Assist' 'debloat' 'WARN' "$_"
    }
}
function Invoke-RemoveBloat {
    Write-UiLog "Removing bloat (current user)..."
    foreach ($b in $sync.BloatList) { Remove-BloatPackage -Name $b.Name -Match $b.Match }
    Remove-QuickAssist
}

function Invoke-StartGridLayout {
    $key = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced'
    try {
        New-Item -Path $key -Force -ErrorAction SilentlyContinue | Out-Null
        Set-ItemProperty -Path $key -Name 'Start_Layout' -Value 1 -Type DWord -ErrorAction Stop
        Write-UiLog "Start menu set to grid layout. Sign out/in to apply." 'OK'
    } catch { Write-UiLog "Start grid setting failed: $_" 'WARN' }
}

function Invoke-WindowsUpdates {
    Write-UiLog "Preparing Windows Update..."
    Set-UiStatus "Checking for Windows updates..."
    try {
        # Module setup is host-agnostic so it runs fine in the runspace.
        if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue | Where-Object { $_.Version -ge [version]'2.8.5.201' })) {
            Write-UiLog "Installing NuGet package provider..."
            Install-PackageProvider -Name NuGet -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
        }
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
            Write-UiLog "Installing PSWindowsUpdate module..."
            # Trust PSGallery up-front so the subprocess doesn't hit a confirm prompt of its own.
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
            Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser -AllowClobber -ErrorAction Stop
        }

        # PSWindowsUpdate's Get-WindowsUpdate -Install fails in this background runspace
        # with "the host program... does not support user interaction" because the cmdlet
        # uses Read-Host internally even when -Confirm:$false and -AcceptAll are set.
        # Workaround: run it in a fresh powershell.exe, which provides a real console host.
        $stem = "pbt_winupdate_" + [Guid]::NewGuid().ToString('N').Substring(0,8)
        $tempScript = Join-Path $env:TEMP "$stem.ps1"
        $stdoutLog  = Join-Path $env:TEMP "$stem.out.log"
        $stderrLog  = Join-Path $env:TEMP "$stem.err.log"

        # Single-quoted here-string lives at top level (in $sync.WinUpdateScript) because
        # nested here-strings would break the outer $pipelineCode here-string at parse time.
        Set-Content -Path $tempScript -Value $sync.WinUpdateScript -Encoding UTF8

        Write-UiLog "Running Windows Update via powershell.exe subprocess (this can take 10+ minutes)..."
        Set-UiStatus "Installing Windows updates..."
        $p = Start-Process -FilePath powershell.exe `
            -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',"`"$tempScript`"") `
            -Wait -PassThru -NoNewWindow `
            -RedirectStandardOutput $stdoutLog `
            -RedirectStandardError $stderrLog

        # Stream subprocess stdout into the UI log line-by-line
        if (Test-Path $stdoutLog) {
            Get-Content $stdoutLog -ErrorAction SilentlyContinue | ForEach-Object {
                $line = $_.Trim()
                if ($line) { Write-UiLog "WU: $line" }
            }
        }
        if (Test-Path $stderrLog) {
            $errText = Get-Content $stderrLog -Raw -ErrorAction SilentlyContinue
            if ($errText -and $errText.Trim()) {
                Write-UiLog ("WU stderr: " + ($errText.Trim() -replace "`r?`n", ' | ')) 'WARN'
            }
        }
        Remove-Item $tempScript, $stdoutLog, $stderrLog -Force -ErrorAction SilentlyContinue

        if ($p.ExitCode -eq 0) {
            Write-UiLog "Windows Update run complete. Some updates may need a reboot." 'OK'
            Add-Result 'Windows Update' 'install' 'OK'
        } else {
            Write-UiLog "Windows Update subprocess exited with code $($p.ExitCode)." 'WARN'
            Add-Result 'Windows Update' 'install' 'WARN' "exit $($p.ExitCode)"
        }
    } catch {
        Write-UiLog "Windows Update failed: $_" 'ERROR'
        Add-Result 'Windows Update' 'install' 'FAIL' "$_"
    }
}

function Get-GpuVendor {
    $gpus = Get-CimInstance Win32_VideoController -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -notmatch 'Virtual|Basic|Parsec|Remote' }
    foreach ($g in $gpus) {
        if ($g.Name -match 'NVIDIA|GeForce|Quadro|RTX|GTX') { return 'NVIDIA' }
        if ($g.Name -match 'AMD|Radeon|RX\s')               { return 'AMD' }
    }
    return $null
}

function Get-AdrenalinDownloadUrl {
    $driversPage = 'https://www.amd.com/en/support/download/drivers.html'
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072
        $html = (Invoke-WebRequest -UseBasicParsing -Uri $driversPage -TimeoutSec 30 -UserAgent 'Mozilla/5.0').Content
        $matches = [regex]::Matches($html, 'https://[^"'']*?adrenalin[^"'']*?\.exe', 'IgnoreCase')
        if ($matches.Count -gt 0) { return $matches[0].Value }
    } catch { Write-UiLog "Adrenalin page scrape failed: $_" 'WARN' }
    return $null
}

# The NVIDIA App is not available in winget (manifest validation blocked since 2024
# due to the hardware requirement check). Scrape the official download page instead.
# URL pattern: https://us.download.nvidia.com/nvapp/client/<ver>/NVIDIA_app_v<ver>.exe
function Get-NvidiaAppDownloadUrl {
    $page = 'https://www.nvidia.com/en-us/software/nvidia-app/'
    try {
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor 3072
        $html = (Invoke-WebRequest -UseBasicParsing -Uri $page -TimeoutSec 30 -UserAgent 'Mozilla/5.0').Content
        $matches = [regex]::Matches($html, 'https://[^"'']*?NVIDIA_app[^"'']*?\.exe', 'IgnoreCase')
        if ($matches.Count -gt 0) { return $matches[0].Value }
    } catch { Write-UiLog "NVIDIA App page scrape failed: $_" 'WARN' }
    return $null
}

function Install-GpuDriver {
    $vendor = Get-GpuVendor
    if (-not $vendor) {
        Write-UiLog "No supported GPU detected. Skipping driver install." 'WARN'
        Add-Result 'GPU Driver' 'install' 'SKIP' 'no GPU detected'
        return
    }
    Write-UiLog "Detected GPU vendor: $vendor"
    if ($vendor -eq 'NVIDIA') {
        Write-UiLog "Locating current NVIDIA App installer from nvidia.com..."
        $url = Get-NvidiaAppDownloadUrl
        if (-not $url) {
            Write-UiLog "Could not find NVIDIA App download URL. Install manually from https://www.nvidia.com/en-us/software/nvidia-app/" 'ERROR'
            Add-Result 'NVIDIA App' 'install' 'FAIL' 'scrape failed'
            return
        }
        Write-UiLog "NVIDIA App URL: $url"
        # NVIDIA App installer accepts -s for silent install (NVIDIA convention,
        # same as their driver setup). The alternative /S (NSIS) is not honoured.
        Invoke-DirectInstall -App @{ Name='NVIDIA App'; DownloadUrl=$url; SilentArgs='-s' }
    }
    elseif ($vendor -eq 'AMD') {
        Write-UiLog "Locating current Adrenalin installer from amd.com..."
        $url = Get-AdrenalinDownloadUrl
        if (-not $url) {
            Write-UiLog "Could not find Adrenalin download URL. Install manually from https://www.amd.com/en/support/download/drivers.html" 'ERROR'
            Add-Result 'AMD Adrenalin' 'install' 'FAIL' 'scrape failed'
            return
        }
        Write-UiLog "Adrenalin URL: $url"
        Invoke-DirectInstall -App @{ Name='AMD Adrenalin'; DownloadUrl=$url; SilentArgs='-install' }
    }
}

function Invoke-TweakPowerNever {
    Write-UiLog "Setting power timeouts to never..."
    powercfg -change -monitor-timeout-ac 0; powercfg -change -monitor-timeout-dc 0
    powercfg -change -standby-timeout-ac 0; powercfg -change -standby-timeout-dc 0
    powercfg -change -hibernate-timeout-ac 0; powercfg -change -hibernate-timeout-dc 0
    Write-UiLog "Power timeouts set." 'OK'
}
function Invoke-TweakDisableHibernate { Write-UiLog "Disabling hibernation..."; powercfg -h off; Write-UiLog "Done." 'OK' }
function Invoke-TweakClearDownloads {
    $dl = Join-Path $env:USERPROFILE 'Downloads'
    if (Test-Path $dl) {
        Get-ChildItem $dl -Recurse -Force -EA SilentlyContinue | Remove-Item -Recurse -Force -EA SilentlyContinue
        Write-UiLog "Downloads cleared." 'OK'
    } else { Write-UiLog "Downloads folder not found." 'WARN' }
}
function Invoke-TweakEmptyRecycle {
    try { Clear-RecycleBin -Force -EA Stop; Write-UiLog "Recycle Bin emptied." 'OK' }
    catch { Write-UiLog "Recycle Bin: $_" 'WARN' }
}
function Invoke-TweakClearBrowser {
    Write-UiLog "Clearing browser data (closing browsers first)..."

    # --- Step 1: terminate browser processes so their files unlock ---
    # 'msedge' covers Edge main + content processes; same for 'chrome', 'firefox', etc.
    $browsers = @('msedge','chrome','firefox','brave','opera','vivaldi','iexplore')
    $closed = @()
    foreach ($name in $browsers) {
        $procs = Get-Process -Name $name -ErrorAction SilentlyContinue
        if ($procs) {
            $procs | Stop-Process -Force -ErrorAction SilentlyContinue
            $closed += $name
        }
    }
    if ($closed.Count -gt 0) {
        Write-UiLog "Closed: $($closed -join ', ')"
        # File handles linger briefly after Stop-Process; let the kernel release them.
        Start-Sleep -Milliseconds 1500
        # Re-check for stragglers (Chromium browsers can respawn watcher processes).
        foreach ($name in $closed) {
            $procs = Get-Process -Name $name -ErrorAction SilentlyContinue
            if ($procs) { $procs | Stop-Process -Force -ErrorAction SilentlyContinue }
        }
        Start-Sleep -Milliseconds 500
    }

    # Helper: delete a file or folder, return $true on success or if it didn't exist.
    $remove = {
        param([string]$Path)
        if (-not (Test-Path -LiteralPath $Path)) { return $true }
        try { Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop; return $true }
        catch { return $false }
    }

    # --- Step 2: Chromium-based browsers (Edge, Chrome, Brave) ---
    # Per-profile data to remove. Deliberately KEEPS:
    #   - "Login Data" / "Login Data For Account"  (saved passwords)
    #   - "Web Data" (autofill, payment methods)
    #   - "Bookmarks"
    #   - "Preferences" / "Secure Preferences" (settings)
    #   - "Extensions" folder
    $chromiumNuke = @(
        'History','History-journal','History Provider Cache',
        'Cookies','Cookies-journal',
        'Network\Cookies','Network\Cookies-journal',
        'Top Sites','Top Sites-journal','Visited Links',
        'Favicons','Favicons-journal',
        'Sessions','Session Storage','Local Storage','IndexedDB',
        'Cache','Code Cache','GPUCache','Service Worker',
        'Current Session','Current Tabs','Last Session','Last Tabs',
        'Media History','Media History-journal',
        'Shortcuts','Shortcuts-journal',
        'QuotaManager','QuotaManager-journal',
        'DawnGraphiteCache','DawnWebGPUCache'
    )
    $chromiumRoots = @(
        @{ N='Edge';   Root="$env:LOCALAPPDATA\Microsoft\Edge\User Data" }
        @{ N='Chrome'; Root="$env:LOCALAPPDATA\Google\Chrome\User Data" }
        @{ N='Brave';  Root="$env:LOCALAPPDATA\BraveSoftware\Brave-Browser\User Data" }
    )
    foreach ($b in $chromiumRoots) {
        if (-not (Test-Path -LiteralPath $b.Root)) { continue }
        $profileDirs = Get-ChildItem -LiteralPath $b.Root -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -eq 'Default' -or $_.Name -like 'Profile *' }
        if (-not $profileDirs) { continue }
        $ok = 0; $fail = 0
        foreach ($prof in $profileDirs) {
            foreach ($rel in $chromiumNuke) {
                $target = Join-Path $prof.FullName $rel
                if (-not (Test-Path -LiteralPath $target)) { continue }
                if (& $remove $target) { $ok++ } else { $fail++ }
            }
        }
        if ($fail -gt 0) {
            Write-UiLog "$($b.N): cleared $ok item(s) across $($profileDirs.Count) profile(s); $fail still locked." 'WARN'
        } else {
            Write-UiLog "$($b.N): cleared $ok item(s) across $($profileDirs.Count) profile(s)." 'OK'
        }
    }

    # --- Step 3: Firefox ---
    # NOTE: places.sqlite contains BOTH history and bookmarks; clearing it loses both.
    # This matches the prior behaviour and the user's "clear all" intent.
    # KEEPS: logins.json, key4.db (saved passwords), prefs.js, extensions/.
    $firefoxNukeRoaming = @(
        'places.sqlite','places.sqlite-wal','places.sqlite-shm',
        'cookies.sqlite','cookies.sqlite-wal','cookies.sqlite-shm',
        'formhistory.sqlite',
        'permissions.sqlite','permissions.sqlite-wal','permissions.sqlite-shm',
        'webappsstore.sqlite','webappsstore.sqlite-wal','webappsstore.sqlite-shm',
        'sessionstore.jsonlz4','sessionstore.bak','sessionstore-backups',
        'storage','storage-sync-v2.sqlite','favicons.sqlite','favicons.sqlite-wal'
    )
    $firefoxNukeLocal = @('cache2','startupCache','thumbnails','OfflineCache','jumpListCache','shader-cache')

    $ffPaths = @(
        @{ Root = Join-Path $env:APPDATA      'Mozilla\Firefox\Profiles'; List = $firefoxNukeRoaming }
        @{ Root = Join-Path $env:LOCALAPPDATA 'Mozilla\Firefox\Profiles'; List = $firefoxNukeLocal   }
    )
    foreach ($pair in $ffPaths) {
        if (-not (Test-Path -LiteralPath $pair.Root)) { continue }
        Get-ChildItem -LiteralPath $pair.Root -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $prof = $_
            $ok = 0; $fail = 0
            foreach ($rel in $pair.List) {
                $target = Join-Path $prof.FullName $rel
                if (-not (Test-Path -LiteralPath $target)) { continue }
                if (& $remove $target) { $ok++ } else { $fail++ }
            }
            if ($ok -gt 0 -or $fail -gt 0) {
                if ($fail -gt 0) { Write-UiLog "Firefox '$($prof.Name)': cleared $ok, $fail locked." 'WARN' }
                else             { Write-UiLog "Firefox '$($prof.Name)': cleared $ok item(s)." 'OK' }
            }
        }
    }
}


try {
    # BREADCRUMB 1: did the runspace even start?
    Add-Content -Path "$env:TEMP\pcbt-pipeline-trace.log" -Value "[$(Get-Date -Format 'HH:mm:ss.fff')] runspace entered" -ErrorAction SilentlyContinue
    Set-UiBusy $true
    Add-Content -Path "$env:TEMP\pcbt-pipeline-trace.log" -Value "[$(Get-Date -Format 'HH:mm:ss.fff')] Set-UiBusy returned" -ErrorAction SilentlyContinue
    Set-UiProgress 0
    Add-Content -Path "$env:TEMP\pcbt-pipeline-trace.log" -Value "[$(Get-Date -Format 'HH:mm:ss.fff')] Set-UiProgress returned" -ErrorAction SilentlyContinue
    $sync.RunResults = @()
    $apps   = @($sync.SelectedApps)
    $tweaks = $sync.SelectedTweaks
    $opts   = $sync.SelectedOpts
    $mode   = $sync.Mode
    Write-UiLog "Pipeline dispatching ($mode, $($apps.Count) apps)..."
    if ($sync.KeepTemp) { Write-UiLog "Debug mode: temp files will be kept." 'INFO' }

    if ($mode -eq 'Install') {
        Write-UiLog "=============== INSTALL RUN STARTED ==============="
        $tweakCount = (($tweaks.Values | Where-Object { $_ }) | Measure-Object).Count
        $optCount   = (($opts.Values   | Where-Object { $_ }) | Measure-Object).Count
        $total = [Math]::Max(1, $apps.Count + $tweakCount + $optCount)
        $done  = 0

        if (($apps | Where-Object { $_.Source -eq 'choco' }).Count -gt 0 -and -not (Test-ChocoInstalled)) {
            Install-Chocolatey
        }
        $i = 0
        foreach ($app in $apps) {
            $i++
            Set-UiStatus "Installing $($app.Name)  ($i of $($apps.Count))..."
            try {
                switch ($app.Source) {
                    'winget' { Invoke-WingetInstall $app }
                    'choco'  { Invoke-ChocoInstall  $app }
                    'direct' { Invoke-DirectInstall $app }
                    'zip'    { Invoke-ZipInstall    $app }
                    default  { Write-UiLog "Unknown source '$($app.Source)'." 'ERROR' }
                }
            } catch { Write-UiLog "Error on $($app.Name): $_" 'ERROR' }
            $done++; Set-UiProgress (($done / $total) * 100)
        }

        if ($opts.WindowsUpdate) {
            Set-UiStatus "Windows Update..."; Invoke-WindowsUpdates
            $done++; Set-UiProgress (($done / $total) * 100)
        }
        if ($opts.GpuDriver) {
            Set-UiStatus "GPU driver..."; Install-GpuDriver
            $done++; Set-UiProgress (($done / $total) * 100)
        }

        $tweakList = @(
            @{K='PowerNever';       L='Set power plan';    A={ Invoke-TweakPowerNever }}
            @{K='DisableHibernate'; L='Disable hibernate'; A={ Invoke-TweakDisableHibernate }}
            @{K='ClearDownloads';   L='Clear Downloads';   A={ Invoke-TweakClearDownloads }}
            @{K='EmptyRecycle';     L='Empty Recycle Bin'; A={ Invoke-TweakEmptyRecycle }}
            @{K='ClearBrowser';     L='Clear browser';     A={ Invoke-TweakClearBrowser }}
            @{K='StartGrid';        L='Start menu grid';   A={ Invoke-StartGridLayout }}
            @{K='RemoveBloat';      L='Remove bloat';      A={ Invoke-RemoveBloat }}
        )
        foreach ($t in $tweakList) {
            if ($tweaks[$t.K]) {
                Set-UiStatus "$($t.L)..."
                try { & $t.A } catch { Write-UiLog "Tweak $($t.L): $_" 'ERROR' }
                $done++; Set-UiProgress (($done / $total) * 100)
            }
        }

        Set-UiStatus "Install complete."
        Set-UiProgress 100
        Write-UiLog "=============== INSTALL RUN COMPLETE =============="
    }
    elseif ($mode -eq 'Uninstall') {
        Write-UiLog "=============== UNINSTALL RUN STARTED ============="
        $total = [Math]::Max(1, $apps.Count); $done = 0; $i = 0
        foreach ($app in $apps) {
            $i++
            Set-UiStatus "Uninstalling $($app.Name)  ($i of $($apps.Count))..."
            try {
                switch ($app.Source) {
                    'winget'         { Invoke-WingetUninstall   $app }
                    'choco'          { Invoke-ChocoUninstall    $app }
                    { $_ -in 'direct','zip' } { Invoke-RegistryUninstall $app }
                }
            } catch { Write-UiLog "Error on $($app.Name): $_" 'ERROR' }
            $done++; Set-UiProgress (($done / $total) * 100)
        }
        Set-UiStatus "Uninstall complete."
        Set-UiProgress 100
        Write-UiLog "=============== UNINSTALL RUN COMPLETE ============"
    }
    # Clean up cached 7-Zip at pipeline end unless debugging
    if (-not $sync.KeepTemp -and $sync.SevenZipPath -and (Test-Path $sync.SevenZipPath)) {
        Remove-Item $sync.SevenZipPath -Force -ErrorAction SilentlyContinue
    }
}
catch { Write-UiLog "Pipeline fatal: $_" 'ERROR' }
finally { Set-UiBusy $false }
'@

function Start-Pipeline {
    param([string]$Mode, [array]$SelectedApps, [hashtable]$SelectedTweaks, [hashtable]$SelectedOpts, [bool]$KeepTemp)
    Remove-Item "$env:TEMP\pcbt-pipeline-trace.log" -ErrorAction SilentlyContinue
    $sync.RunSummary.Visibility = 'Collapsed'
    $sync.RunSummary.Text       = ''
    $sync.ReportActions.Children.Clear()
    $sync.ReportFooter.Text = "Mode: $Mode  ·  Started: $(Get-Date -Format 'HH:mm:ss')  ·  Log: $($sync.LogPath)"
    $sync.Mode           = $Mode
    $sync.SelectedApps   = $SelectedApps
    $sync.SelectedTweaks = $SelectedTweaks
    $sync.SelectedOpts   = $SelectedOpts
    $sync.KeepTemp       = $KeepTemp
    Write-Log "Starting $Mode pipeline..."
    $rs = [RunspaceFactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'
    $rs.ThreadOptions  = 'ReuseThread'
    $rs.Open()
    $rs.SessionStateProxy.SetVariable('sync', $sync)
    $ps = [PowerShell]::Create()
    $ps.Runspace = $rs
    $null = $ps.AddScript($pipelineCode)
    Add-Content -Path "$env:TEMP\pcbt-pipeline-trace.log" -Value "[$(Get-Date -Format 'HH:mm:ss.fff')] Start-Pipeline: about to BeginInvoke (mode=$Mode)" -ErrorAction SilentlyContinue

    # Run async so the WPF UI stays responsive. The pipeline calls back into
    # the UI via $sync.<ctrl>.Dispatcher.Invoke(...) — that only works if the
    # UI thread is free to pump the dispatcher queue, hence BeginInvoke.
    # Stash handles on $sync so the DispatcherTimer tick can reach them.
    $sync.PsInstance  = $ps
    $sync.RunspaceObj = $rs
    $sync.AsyncResult = $ps.BeginInvoke()

    # Poll completion from the UI thread. This never blocks: IsCompleted is a
    # cheap flag check, and the actual work runs on the runspace thread.
    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromMilliseconds(250)
    $sync.PipelineTimer = $timer
    $timer.Add_Tick({
        if (-not $sync.AsyncResult.IsCompleted) { return }
        $sync.PipelineTimer.Stop()
        try {
            $sync.PsInstance.EndInvoke($sync.AsyncResult) | Out-Null
            Add-Content -Path "$env:TEMP\pcbt-pipeline-trace.log" -Value "[$(Get-Date -Format 'HH:mm:ss.fff')] EndInvoke completed, HadErrors=$($sync.PsInstance.HadErrors)" -ErrorAction SilentlyContinue
        } catch {
            $msg = "Pipeline threw on EndInvoke: $($_.Exception.Message)"
            Write-Log $msg 'ERROR'
            Add-Content -Path "$env:TEMP\pcbt-pipeline-trace.log" -Value $msg -ErrorAction SilentlyContinue
            if ($_.Exception.InnerException) {
                $inner = "  inner: $($_.Exception.InnerException.Message)"
                Write-Log $inner 'ERROR'
                Add-Content -Path "$env:TEMP\pcbt-pipeline-trace.log" -Value $inner -ErrorAction SilentlyContinue
            }
        }
        if ($sync.PsInstance.HadErrors) {
            foreach ($err in $sync.PsInstance.Streams.Error) {
                Write-Log "Runspace error: $err" 'ERROR'
                Add-Content -Path "$env:TEMP\pcbt-pipeline-trace.log" -Value "Runspace error: $err" -ErrorAction SilentlyContinue
                if ($err.InvocationInfo -and $err.InvocationInfo.ScriptLineNumber) {
                    $loc = "  at line $($err.InvocationInfo.ScriptLineNumber): $($err.InvocationInfo.Line.Trim())"
                    Write-Log $loc 'ERROR'
                    Add-Content -Path "$env:TEMP\pcbt-pipeline-trace.log" -Value $loc -ErrorAction SilentlyContinue
                }
            }
        }
        $sync.PsInstance.Dispose()
        $sync.RunspaceObj.Close()
        $sync.PsInstance    = $null
        $sync.RunspaceObj   = $null
        $sync.AsyncResult   = $null
        $sync.PipelineTimer = $null

        # Build post-run summary line from structured results
        $results = $sync.RunResults
        if ($results -and $results.Count -gt 0) {
            $ok   = ($results | Where-Object { $_.Status -eq 'OK'   }).Count
            $warn = ($results | Where-Object { $_.Status -eq 'WARN' }).Count
            $fail = ($results | Where-Object { $_.Status -eq 'FAIL' }).Count
            $skip = ($results | Where-Object { $_.Status -eq 'SKIP' }).Count
            $parts = @()
            if ($ok)   { $parts += "$ok ok" }
            if ($warn) { $parts += "$warn warn" }
            if ($fail) { $parts += "$fail failed" }
            if ($skip) { $parts += "$skip skipped" }
            if ($parts.Count -gt 0) {
                $sync.RunSummary.Text       = $parts -join '   |   '
                $sync.RunSummary.Visibility = 'Visible'
            }
        }
    })
    $timer.Start()
}

$controls.BtnQuit.Add_Click({ $window.Close() })
$controls.BtnClearLog.Add_Click({ $controls.LogOutput.Children.Clear() })

function Show-View {
    param([string]$View)
    switch ($View) {
        'Deploy' {
            $controls.DeployTop.Visibility   = 'Visible'
            $controls.LogConsole.Visibility  = 'Visible'
            $controls.ReportPanel.Visibility = 'Collapsed'
            $controls.BodyGrid.RowDefinitions[0].Height = New-Object System.Windows.GridLength(1, 'Star')
            $controls.BodyGrid.RowDefinitions[1].Height = New-Object System.Windows.GridLength(220)
            $controls.HeadTitle.Text = 'Deploy new bench run'
        }
        'Report' {
            $controls.DeployTop.Visibility   = 'Collapsed'
            $controls.LogConsole.Visibility  = 'Collapsed'
            $controls.ReportPanel.Visibility = 'Visible'
            $controls.HeadTitle.Text = 'Report'
        }
        'Logs' {
            $controls.DeployTop.Visibility   = 'Collapsed'
            $controls.LogConsole.Visibility  = 'Visible'
            $controls.ReportPanel.Visibility = 'Collapsed'
            $controls.BodyGrid.RowDefinitions[0].Height = New-Object System.Windows.GridLength(0)
            $controls.BodyGrid.RowDefinitions[1].Height = New-Object System.Windows.GridLength(1, 'Star')
            $controls.HeadTitle.Text = 'Logs'
        }
    }
}
$controls.NavDeploy.Add_Checked({ Show-View 'Deploy' })
$controls.NavReport.Add_Checked({ Show-View 'Report' })
$controls.NavLogs.Add_Checked({   Show-View 'Logs'   })
$controls.BtnSelectAllApps.Add_Click({
    $visible   = @($appCheckboxes.Values | Where-Object { $_.Visibility -eq 'Visible' })
    $anyUnchecked = $visible | Where-Object { -not $_.IsChecked }
    $newState  = [bool]$anyUnchecked
    $visible | ForEach-Object { $_.IsChecked = $newState }
})

$controls.BtnRun.Add_Click({
    $selectedApps = @()
    foreach ($app in $script:AppCatalog) {
        if ($appCheckboxes[$app.Id].IsChecked) { $selectedApps += $app }
    }
    $selectedTweaks = @{
        PowerNever       = [bool]$controls.TweakPowerNever.IsChecked
        DisableHibernate = [bool]$controls.TweakDisableHibernate.IsChecked
        ClearDownloads   = [bool]$controls.TweakClearDownloads.IsChecked
        EmptyRecycle     = [bool]$controls.TweakEmptyRecycle.IsChecked
        ClearBrowser     = [bool]$controls.TweakClearBrowser.IsChecked
        StartGrid        = [bool]$controls.TweakStartGrid.IsChecked
        RemoveBloat      = [bool]$controls.TweakRemoveBloat.IsChecked
    }
    $selectedOpts = @{
        GpuDriver       = [bool]$controls.OptGpuDriver.IsChecked
        WindowsUpdate   = [bool]$controls.OptWindowsUpdate.IsChecked
    }
    $anySelected = $selectedApps.Count -gt 0 -or
                   (($selectedTweaks.Values | Where-Object { $_ }).Count -gt 0) -or
                   (($selectedOpts.Values   | Where-Object { $_ }).Count -gt 0)
    if (-not $anySelected) {
        [System.Windows.MessageBox]::Show('Select at least one app, tweak, or option.','Nothing to do','OK','Information') | Out-Null
        return
    }
    if (-not (Invoke-PreflightChecks)) {
        [System.Windows.MessageBox]::Show('Pre-flight checks failed. See log.','Cannot run','OK','Error') | Out-Null
        return
    }
    Start-Pipeline -Mode 'Install' -SelectedApps $selectedApps -SelectedTweaks $selectedTweaks `
                   -SelectedOpts $selectedOpts `
                   -KeepTemp     ([bool]$controls.OptKeepTemp.IsChecked)
})

$controls.BtnUninstall.Add_Click({
    $r = [System.Windows.MessageBox]::Show(
        "Uninstall every app in the catalog that is currently installed?`n`nApps not present will be skipped.",
        'Confirm uninstall', 'YesNo', 'Warning')
    if ($r -ne [System.Windows.MessageBoxResult]::Yes) { return }
    Start-Pipeline -Mode 'Uninstall' -SelectedApps $script:AppCatalog -SelectedTweaks @{} -SelectedOpts @{} `
                   -KeepTemp ([bool]$controls.OptKeepTemp.IsChecked)
})

Write-Log "PC Build Toolkit $SCRIPT_VERSION ready. Log: $($sync.LogPath)"
Invoke-SelfUpdateCheck

$window.ShowDialog() | Out-Null
