#requires -Version 3

#region Add Assemblies
Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms, WindowsFormsIntegration
$code = @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace System
{
	public class IconExtractor
	{

	 public static Icon Extract(string file, int number, bool largeIcon)
	 {
	  IntPtr large;
	  IntPtr small;
	  ExtractIconEx(file, number, out large, out small, 1);
	  try
	  {
	   return Icon.FromHandle(largeIcon ? large : small);
	  }
	  catch
	  {
	   return null;
	  }

	 }
	 [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
	 private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);

	}
}
"@
Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing

# Mahapps Library
if (Test-Path -Path "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\MahApps.Metro.dll")
{
    [System.Reflection.Assembly]::LoadFrom("$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\MahApps.Metro.dll")    | out-null
    [System.Reflection.Assembly]::LoadFrom("$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\System.Windows.Interactivity.dll") | out-null
}

if (Test-Path -Path "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\MahApps.Metro.dll")
{
    [System.Reflection.Assembly]::LoadFrom("${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\MahApps.Metro.dll")    | out-null
    [System.Reflection.Assembly]::LoadFrom("${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\System.Windows.Interactivity.dll") | out-null
}

#endregion

#region GUI and Variables
### Main Window ###
[xml]$xaml = @"
<Controls:MetroWindow 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
        Title="ConfigMgr Task Sequence Monitor" Height="685" Width="1347" WindowStartupLocation="CenterScreen" ResizeMode="CanResizeWithGrip">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Colors.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/Cobalt.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/BaseLight.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <Grid VerticalAlignment="Stretch" HorizontalAlignment="Stretch" Margin="0,0,2,0" Width="auto">
        <Grid.RowDefinitions>
            <RowDefinition Height="503*" />
            <RowDefinition Height="10" />
            <RowDefinition Height="141*" />
        </Grid.RowDefinitions>
        <GroupBox Header="MDT" HorizontalAlignment="Left" Margin="10,96,0,0" VerticalAlignment="Top" Height="110" Width="1228">
            <Grid HorizontalAlignment="Left" Height="72" Margin="0,0,-2,-3" VerticalAlignment="Top" Width="1218">
                <Grid.RowDefinitions>
                    <RowDefinition Height="0*"/>
                    <RowDefinition/>
                </Grid.RowDefinitions>
                <Label Content="MDT Integrated?" HorizontalAlignment="Left" Margin="0,5,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
                <CheckBox x:Name="MDTIntegrated" Content="" HorizontalAlignment="Left" Margin="104,10,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
                <Label x:Name="DeploymentStatusLabel" Content="Deployment Status:" HorizontalAlignment="Left" Margin="136,5,0,0" Grid.RowSpan="2" VerticalAlignment="Top" IsEnabled="False"/>
                <TextBox x:Name="DeploymentStatus" HorizontalAlignment="Left" Height="23" Margin="254,8,0,0" Grid.RowSpan="2" TextWrapping="Wrap" VerticalAlignment="Top" Width="212" IsEnabled="False" IsReadOnly="True" VerticalContentAlignment="Center"/>
                <Label x:Name="PercentCompleteLabel" Content="Percent Complete:" HorizontalAlignment="Left" Margin="731,39,0,0" Grid.RowSpan="2" VerticalAlignment="Top" IsEnabled="False"/>
                <TextBox x:Name="PercentComplete" HorizontalAlignment="Left" Height="23" Margin="843,39,0,0" Grid.RowSpan="2" TextWrapping="Wrap" VerticalAlignment="Top" Width="61" IsEnabled="False" IsReadOnly="True" VerticalContentAlignment="Center"/>
                <ProgressBar x:Name="ProgressBar" HorizontalAlignment="Left" Height="32" Margin="956,29,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="252" IsEnabled="False" Minimum="0" Maximum="100" Background="#FFE6E6E6" Foreground="#FF46726A">
                    <ProgressBar.Effect>
                        <DropShadowEffect/>
                    </ProgressBar.Effect>
                </ProgressBar>
                <Label x:Name="CurrentStepLabel" Content="Current Step:" HorizontalAlignment="Left" Margin="470,5,0,0" Grid.RowSpan="2" VerticalAlignment="Top" IsEnabled="False"/>
                <TextBox x:Name="CurrentStep" HorizontalAlignment="Left" Height="23" Margin="549,8,0,0" Grid.RowSpan="2" TextWrapping="Wrap" VerticalAlignment="Top" Width="61" IsEnabled="False" IsReadOnly="True" VerticalContentAlignment="Center"/>
                <Label x:Name="StepNameLabel" Content="StepName:" HorizontalAlignment="Left" Margin="615,5,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Height="26" Width="68" IsEnabled="False"/>
                <TextBox x:Name="StepName" HorizontalAlignment="Left" Height="23" Margin="683,8,0,0" Grid.RowSpan="2" TextWrapping="Wrap" VerticalAlignment="Top" Width="254" IsEnabled="False" IsReadOnly="True" VerticalContentAlignment="Center"/>
                <Label x:Name="StartLabel" Content="Start:" HorizontalAlignment="Left" Margin="0,39,0,0" Grid.RowSpan="2" VerticalAlignment="Top" IsEnabled="False"/>
                <Label x:Name="EndLabel" Content="End:" HorizontalAlignment="Left" Margin="281,39,0,0" Grid.RowSpan="2" VerticalAlignment="Top" IsEnabled="False"/>
                <Label x:Name="ElapsedLabel" Content="Elapsed:" HorizontalAlignment="Left" Margin="557,39,0,0" Grid.RowSpan="2" VerticalAlignment="Top" IsEnabled="False"/>
                <TextBox x:Name="MDTStartTime" HorizontalAlignment="Left" Height="23" Margin="42,39,0,0" Grid.RowSpan="2" TextWrapping="Wrap" VerticalAlignment="Top" Width="220" IsEnabled="False" VerticalContentAlignment="Center" IsReadOnly="True"/>
                <TextBox x:Name="MDTEndTime" HorizontalAlignment="Left" Height="23" Margin="319,39,0,0" Grid.RowSpan="2" TextWrapping="Wrap" VerticalAlignment="Top" Width="220" IsEnabled="False" VerticalContentAlignment="Center" IsReadOnly="True"/>
                <TextBox x:Name="MDTElapsedTime" HorizontalAlignment="Left" Height="23" Margin="615,39,0,0" Grid.RowSpan="2" TextWrapping="Wrap" VerticalAlignment="Top" Width="92" IsEnabled="False" VerticalContentAlignment="Center" IsReadOnly="True"/>
                <Label x:Name="ProgressLabel" Content="Deployment Progress:" HorizontalAlignment="Left" Margin="956,0,0,0" Grid.RowSpan="2" VerticalAlignment="Top" IsEnabled="False"/>
            </Grid>
        </GroupBox>
        <GroupBox Header="ConfigMgr" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Height="81" Width="1317">
            <Grid HorizontalAlignment="Left" Height="46" Margin="0,0,-2,-1" VerticalAlignment="Top" Width="1299">
                <Label Content="Task Sequence:" HorizontalAlignment="Left" Margin="0,7,0,0" VerticalAlignment="Top" Width="96"/>
                <ComboBox x:Name="TaskSequence" HorizontalAlignment="Left" Margin="96,7,0,0" VerticalAlignment="Top" Width="356" Height="26"/>
                <Label Content="Time Period &#xD;&#xA;(Hours):" HorizontalAlignment="Left" Margin="457,0,0,0" VerticalAlignment="Top" Width="73" Height="46"/>
                <TextBox x:Name="TimePeriod" HorizontalAlignment="Left" Height="26" Margin="535,7,0,0" TextWrapping="Wrap" Text="24" VerticalAlignment="Top" Width="48" TextAlignment="Center" VerticalContentAlignment="Center"/>
                <Label Content="Errors &#xD;&#xA;Only:" HorizontalAlignment="Left" Margin="855,0,0,0" VerticalAlignment="Top" Width="41" Height="46"/>
                <CheckBox x:Name="ErrorsOnly" Content="" HorizontalAlignment="Left" Margin="901,18,0,0" VerticalAlignment="Top"/>
                <Label Content="ComputerName:" HorizontalAlignment="Left" Margin="588,8,0,0" VerticalAlignment="Top"/>
                <ComboBox x:Name="ComputerName" HorizontalAlignment="Left" Margin="690,7,0,0" VerticalAlignment="Top" Width="160" Height="26"/>
                <Label Content="Refresh Period &#xD;&#xA;(Minutes):" HorizontalAlignment="Left" Margin="1073,0,0,0" VerticalAlignment="Top" Height="46"/>
                <TextBox x:Name="RefreshPeriod" HorizontalAlignment="Left" Height="26" Margin="1165,8,0,0" TextWrapping="Wrap" Text="1" VerticalAlignment="Top" Width="33" TextAlignment="Center" VerticalContentAlignment="Center"/>
                <Button x:Name="RefreshNow" Content="Refresh Now!" HorizontalAlignment="Left" Margin="1204,7,0,0" VerticalAlignment="Top" Width="95" Height="29"/>
                <Label Content="Error Count:" HorizontalAlignment="Left" Margin="921,10,0,0" VerticalAlignment="Top"/>
                <TextBox x:Name="ErrorCount" HorizontalAlignment="Left" Height="26" Margin="1000,8,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="37" VerticalContentAlignment="Center" HorizontalContentAlignment="Center" IsReadOnly="True"/>
            </Grid>
        </GroupBox>
        <Button x:Name="SettingsButton" Content="Settings" HorizontalAlignment="Left" Margin="1243,96,0,0" VerticalAlignment="Top" Width="83" Height="29"/>
        <Button x:Name="ReportButton" Content="Generate &#xA;  Report" HorizontalAlignment="Left" Margin="1243,130,0,0" VerticalAlignment="Top" Width="83" Height="38" HorizontalContentAlignment="Center"/>
        <DataGrid x:Name="DataGrid" AutoGenerateColumns="False" HorizontalAlignment="Stretch" Margin="10,222,10,0" VerticalAlignment="Stretch" Height="Auto" IsReadOnly="True" HorizontalGridLinesBrush="#FF297566" VerticalGridLinesBrush="#FF489183">
            <DataGrid.Columns>
                <DataGridTemplateColumn Width="SizeToCells">
                    <DataGridTemplateColumn.CellTemplate>
                        <DataTemplate>
                            <Image Source="{Binding Path=Icon}" Width="15" Height="15" />
                        </DataTemplate>
                    </DataGridTemplateColumn.CellTemplate>
                </DataGridTemplateColumn>
                <DataGridTextColumn Header="ComputerName" Binding="{Binding Path=ComputerName}" />
                <DataGridTextColumn Header="GUID" Binding="{Binding Path=GUID}" Visibility="Hidden"/>
                <DataGridTextColumn Header="ExecutionTime" Binding="{Binding Path=ExecutionTime}" />
                <DataGridTextColumn Header="Step" Binding="{Binding Path=Step}" />
                <DataGridTextColumn Header="ActionName" Binding="{Binding Path=ActionName}" />
                <DataGridTextColumn Header="GroupName" Binding="{Binding Path=GroupName}" />
                <DataGridTextColumn Header="LastStatusMsgName" Binding="{Binding Path=LastStatusMsgName}" />
                <DataGridTextColumn Header="ExitCode" Binding="{Binding Path=ExitCode}"/>
                <DataGridTextColumn Header="Record" Binding="{Binding Path=Record}" Visibility="Hidden"/>
            </DataGrid.Columns>
        </DataGrid>
        <GridSplitter Grid.Row="1" HorizontalAlignment="Stretch" Width="auto" Height="8" Background="White" ToolTip="Resize" />
        <TextBox x:Name="ActionOutput" Grid.Row="2" HorizontalAlignment="Stretch" Margin="10,26,10,10" TextWrapping="Wrap" VerticalAlignment="Stretch" ScrollViewer.VerticalScrollBarVisibility="Auto" IsReadOnly="True"/>
        <Label Content="Action Output:" Grid.Row="2" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" Height="26" Width="88"/>
    </Grid>
</Controls:MetroWindow>
"@

$hash = [hashtable]::Synchronized(@{})
$reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml)
$hash.Window = [Windows.Markup.XamlReader]::Load( $reader )
$global:PSInstances = @()
$Global:Timezones = @()

$hash.TaskSequence = $hash.Window.FindName('TaskSequence')
$hash.TimePeriod = $hash.Window.FindName('TimePeriod')
$hash.ErrorsOnly = $hash.Window.FindName('ErrorsOnly')
$hash.ComputerName = $hash.Window.FindName('ComputerName')
$hash.RefreshPeriod = $hash.Window.FindName('RefreshPeriod')
$hash.RefreshNow = $hash.Window.FindName('RefreshNow')
$hash.DataGrid = $hash.Window.FindName('DataGrid')
$hash.ActionOutput = $hash.Window.FindName('ActionOutput')
$hash.MDTIntegrated = $hash.Window.FindName('MDTIntegrated')
$hash.DeploymentStatus = $hash.Window.FindName('DeploymentStatus')
$hash.CurrentStep = $hash.Window.FindName('CurrentStep')
$hash.StepName = $hash.Window.FindName('StepName')
$hash.PercentComplete = $hash.Window.FindName('PercentComplete')
$hash.MDTStartTime = $hash.Window.FindName('MDTStartTime')
$hash.MDTEndTime = $hash.Window.FindName('MDTEndTime')
$hash.MDTElapsedTime = $hash.Window.FindName('MDTElapsedTime')
$hash.SettingsButton = $hash.Window.FindName('SettingsButton')
$hash.ReportButton = $hash.Window.FindName('ReportButton')
$hash.ErrorCount = $hash.Window.FindName('ErrorCount')

$hash.DeploymentStatusLabel = $hash.Window.FindName('DeploymentStatusLabel')
$hash.CurrentStepLabel = $hash.Window.FindName('CurrentStepLabel')
$hash.StepNameLabel = $hash.Window.FindName('StepNameLabel')
$hash.PercentCompleteLabel = $hash.Window.FindName('PercentCompleteLabel')
$hash.StartLabel = $hash.Window.FindName('StartLabel')
$hash.EndLabel = $hash.Window.FindName('EndLabel')
$hash.ElapsedLabel = $hash.Window.FindName('ElapsedLabel')
$hash.ProgressLabel = $hash.Window.FindName('ProgressLabel')

$hash.ProgressBar = $hash.Window.FindName('ProgressBar')
if (Test-Path -Path "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
{
    #$hash.Window.Icon = "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
}
if (Test-Path -Path "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
{
    #$hash.Window.Icon = "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
}


### Settings Window ###
[xml]$xaml2 = @"
<Controls:MetroWindow 
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:Controls="clr-namespace:MahApps.Metro.Controls;assembly=MahApps.Metro"
        Title="Settings" Height="256.128" Width="520.986" WindowStartupLocation="CenterScreen" ResizeMode="NoResize">
    <Window.Resources>
        <ResourceDictionary>
            <ResourceDictionary.MergedDictionaries>
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Controls.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Fonts.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Colors.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/Cobalt.xaml" />
                <ResourceDictionary Source="pack://application:,,,/MahApps.Metro;component/Styles/Accents/BaseLight.xaml" />
            </ResourceDictionary.MergedDictionaries>
        </ResourceDictionary>
    </Window.Resources>
    <Grid>
        <TabControl HorizontalAlignment="Left" Height="212" Margin="10,10,0,0" VerticalAlignment="Top" Width="497">
            <TabItem x:Name="SettingsTab" Header="Settings">
                <Grid Background="#FFE5E5E5" Margin="0,0,0,-2">
                    <Label Content="SQL Server:" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
                    <Label Content="Database:" HorizontalAlignment="Left" Margin="10,41,0,0" VerticalAlignment="Top"/>
                    <Label Content="MDT Monitoring URL:" HorizontalAlignment="Left" Margin="10,72,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="SQLServer" HorizontalAlignment="Left" Height="23" Margin="140,10,0,0" TextWrapping="Wrap" Text="&lt;SQLServer\Instance&gt;" VerticalAlignment="Top" Width="206" VerticalContentAlignment="Center"/>
                    <TextBox x:Name="Database" HorizontalAlignment="Left" Height="23" Margin="140,41,0,0" TextWrapping="Wrap" Text="&lt;Database&gt;" VerticalAlignment="Top" Width="87" VerticalContentAlignment="Center"/>
                    <TextBox x:Name="MDTURL" HorizontalAlignment="Left" Height="23" Margin="140,72,0,0" TextWrapping="Wrap" Text="http://&lt;MDTServer&gt;:9801/MDTMonitorData/Computers" VerticalAlignment="Top" Width="341" VerticalContentAlignment="Center"/>
                    <Button x:Name="ConnectSQL" Content="Connect SQL" HorizontalAlignment="Left" Margin="381,10,0,0" VerticalAlignment="Top" Width="100" Height="29"/>
                    <Label x:Name="Runasadmin" Content="Note:  Please run the application as administrator to save these&#xD;&#xA;settings to the registry!" HorizontalAlignment="Left" Margin="140,125,0,0" VerticalAlignment="Top" Height="43" Visibility="Hidden"/>
                    <Label Content="Display Date/Time in:" HorizontalAlignment="Left" Margin="10,103,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="DTFormat" HorizontalAlignment="Left" Margin="140,103,0,0" VerticalAlignment="Top" Width="206"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="ReportTab" Header="Summary Report">
                <Grid Background="#FFE5E5E5">
                    <Label Content="Task Sequence:" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
                    <Label Content="Start Date:" HorizontalAlignment="Left" Margin="10,41,0,0" VerticalAlignment="Top"/>
                    <Label Content="End Date:" HorizontalAlignment="Left" Margin="10,72,0,0" VerticalAlignment="Top"/>
                    <ComboBox x:Name="TSList" HorizontalAlignment="Left" Margin="106,14,0,0" VerticalAlignment="Top" Width="375" IsReadOnly="True"/>
                    <DatePicker x:Name="StartDate" HorizontalAlignment="Left" Margin="106,43,0,0" VerticalAlignment="Top" Width="114"/>
                    <DatePicker x:Name="EndDate" HorizontalAlignment="Left" Margin="106,74,0,0" VerticalAlignment="Top" Width="114"/>
                    <Button x:Name="GenerateReport" Content="Generate Report" HorizontalAlignment="Left" Margin="10,112,0,0" VerticalAlignment="Top" Width="118" Height="30"/>
                    <Label x:Name="Working" Content="" HorizontalAlignment="Left" Margin="146,112,0,0" VerticalAlignment="Top" Height="30" VerticalContentAlignment="Center" FontStyle="Italic"/>
                    <Label Content="Create an HTML report summarizing the&#xD;&#xA;task sequence deployments executed in &#xD;&#xA;a given time period." HorizontalAlignment="Left" Margin="243,43,0,0" VerticalAlignment="Top" Width="238" Height="65"/>
                    <ProgressBar x:Name="ReportProgress" HorizontalAlignment="Left" Height="20" Margin="224,116,0,0" VerticalAlignment="Top" Width="238" Minimum="0" Maximum="100" Visibility="Hidden"/>
                </Grid>
            </TabItem>
            <TabItem Header="About" HorizontalAlignment="Left" VerticalAlignment="Top">
                <Grid Background="#FFE5E5E5" Margin="0,0,0,-4">
                    <RichTextBox HorizontalAlignment="Left" Height="171" VerticalAlignment="Top" Width="491" IsDocumentEnabled="True" IsReadOnly="True" IsReadOnlyCaretVisible="True">
                        <FlowDocument>
                            <Paragraph>
                                <Run Text="ConfigMgr Task Sequence Monitor" FontFamily="Calibri" FontSize="18"/>
                                <Run FontFamily="Calibri" FontSize="13" Text="is a WPF application coded in PowerShell.  It enables you to monitor or review task sequence executions in System Center Configuration Manager, and where MDT integration is enabled, link data from MDT with Configuration Manager for enhanced monitoring of ZTI OS deployments."/>
                            </Paragraph>
                            <Paragraph>
                                <Run FontFamily="Calibri" FontSize="13" Text="Documentation can be found on my blog:" />
                                <Hyperlink x:Name="Link1" NavigateUri="http://smsagent.wordpress.com/tools/configmgr-task-sequence-monitor/">smsagent.wordpress.com</Hyperlink>
                                <LineBreak />
                                <Run FontFamily="Calibri" FontSize="13" Text="by Trevor Jones" />
                            </Paragraph>
                            <Paragraph>
                                <Run FontFamily="Calibri" FontSize="13" Text="Version 1.6" />
                            </Paragraph>
                        </FlowDocument>
                    </RichTextBox>
                </Grid>
            </TabItem>
        </TabControl>
    </Grid>
</Controls:MetroWindow>

"@
$reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml2)
$hash.Window2 = [Windows.Markup.XamlReader]::Load( $reader )
$hash.SQLServer = $hash.Window2.FindName('SQLServer')
$hash.Database = $hash.Window2.FindName('Database')
$hash.MDTURL = $hash.Window2.FindName('MDTURL')
$hash.ConnectSQL = $hash.Window2.FindName('ConnectSQL')
$hash.TSList = $hash.Window2.FindName('TSList')
$hash.StartDate = $hash.Window2.FindName('StartDate')
$hash.EndDate = $hash.Window2.FindName('EndDate')
$hash.GenerateReport = $hash.Window2.FindName('GenerateReport')
$hash.SettingsTab = $hash.Window2.FindName('SettingsTab')
$hash.ReportTab = $hash.Window2.FindName('ReportTab')
$hash.Tabs = $hash.Window2.FindName('Tabs')
$hash.Working = $hash.Window2.FindName('Working')
$hash.Runasadmin = $hash.Window2.FindName('Runasadmin')
$hash.ReportProgress = $hash.Window2.FindName('ReportProgress')
$hash.Link1 = $hash.Window2.FindName('Link1')
$hash.DTFormat = $hash.Window2.FindName('DTFormat')

if (Test-Path -Path "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
{
    #$hash.Window2.Icon = "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
    $Hash.Window2.ShowInTaskbar = $true
}
if (Test-Path -Path "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
{
    #$hash.Window2.Icon = "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
    $Hash.Window2.ShowInTaskbar = $true
}

$script:SQLServer = $hash.SQLServer.Text
$Script:Database = $hash.Database.Text
#endregion

#region Icons and Runspacepool
# Output SystemIcons to bmps
$icons = @()
$global:greentickiconpath = "$env:temp\GreenTick.bmp"
$icons += $greentickiconpath 
$global:redcrossiconpath = "$env:temp\RedCross.bmp"
$icons += $redcrossiconpath

if (!(Test-Path $greentickiconpath))
{
    $global:greentickicon = [System.IconExtractor]::Extract('comres.dll',8,$true).ToBitmap()
    $greentickicon.save("$greentickiconpath")
}
if (!(Test-Path $redcrossiconpath))
{
    $global:redcrossicon = [System.IconExtractor]::Extract('comres.dll',10,$true).ToBitmap()
    $redcrossicon.save("$redcrossiconpath")
}

$script:RunspacePool = [runspacefactory]::CreateRunspacePool()
$RunspacePool.ApartmentState = 'STA'
$RunspacePool.ThreadOptions = 'ReUseThread'
$RunspacePool.Open()
#endregion

#region Functions

Function Get-DateTimeFormat 
{
    if ([System.TimeZone]::CurrentTimeZone.IsDaylightSavingTime($(Get-Date)))
    {
        $TimeZone = [System.TimeZone]::CurrentTimeZone.DaylightName
    }
    Else 
    {
        $TimeZone = [System.TimeZone]::CurrentTimeZone.StandardName
    }

    #$Global:Timezones = @()
    $obj = New-Object -TypeName psobject -Property @{
        TimeZone = 'UTC'
    }
    $Global:Timezones = [Array]$Timezones + $obj
    $obj = New-Object -TypeName psobject -Property @{
        TimeZone = $TimeZone
    }
    $Global:Timezones = [Array]$Timezones + $obj
}

Function Get-TaskSequenceList 
{
    # Set variables
    $script:SQLServer = $hash.SQLServer.Text
    $Script:Database = $hash.Database.Text

    # If SQLinstance not populated, ask for connection
    if ($SQLServer -eq '<SQLServer\Instance>')
    {
        $hash.ActionOutput.Text = 'No SQL Server defined.  Click Settings, and set the SQL Server, database and MDT URL if applicable.'
        return
    }

    # Connect to SQL server
    try
    {
        $connectionString = "Server=$SQLServer;Database=$Database;Integrated Security=SSPI;"
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open()
        $hash.ActionOutput.Text = 'Connected to SQL Server database.  Select a Task Sequence.'
    }
    catch 
    {
        $hash.ActionOutput.Text = '[ERROR} Could not connect to SQL Server database!'
        return
    }
    # Run SQL query
    $Query = "
        SELECT DISTINCT summ.SoftwareName AS 'Task Sequence'
        FROM vDeploymentSummary summ
        WHERE (summ.FeatureType=7)
        ORDER BY summ.SoftwareName
    "
    $command = $connection.CreateCommand()
    $command.CommandText = $Query
    $result = $command.ExecuteReader()
    $table = New-Object -TypeName 'System.Data.DataTable'
    $table.Load($result)
    $connection.Close()

    # Load data into psobject            
    $global:Views = @()
    Foreach ($Row in $table.Rows)
    {
        $obj = New-Object -TypeName psobject
        Add-Member -InputObject $obj -MemberType NoteProperty -Name 'TS' -Value $Row.'Task Sequence'
        $global:Views = [Array]$Views + $obj
    }

    # Output to Task Sequence combobox
    $hash.Window.Dispatcher.Invoke(
        [action]{
            $hash.TaskSequence.ItemsSource = [Array]$Views.TS
    })
}

Function Get-TaskSequenceData 
{
    param ($hash,$RunspacePool)

    $code = 
    {
        param($hash,$SQLServer,$Database,$TimePeriod,$ErrorsOnly,$ComputerName,$TS,$MDTIntegrated,$URL,$DTFormat)

        # Notify of data retrieval         
        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.ActionOutput.Text = 'Retrieving data...'
                $hash.DataGrid.ItemsSource = ''
        })

        # Set variable values
        if ($MyGUID)
        {
            Remove-Variable -Name MyGuid
        }
        if ($Unknowns) # put after display ####
        {
            Remove-Variable -Name Unknowns
        }

        if ($ErrorsOnly -eq 'True')
        {
            $ExitCode = 0
        }
        else 
        {
            $ExitCode = 999999999999999999999999
        }
        
        if ($ComputerName -eq '-All-' -or $ComputerName -eq '' -or $ComputerName -eq $Null)
        {
            $SQLComputerName = '%'
        }
        Else 
        {
            $SQLComputerName = $ComputerName
        }
        
        $greentickiconpath = "$env:temp\GreenTick.bmp"
        $redcrossiconpath = "$env:temp\RedCross.bmp"
        
        # Connect to SQL server
        try
        {
            $connectionString = "Server=$SQLServer;Database=$Database;Integrated Security=SSPI;"
            $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
            $connection.ConnectionString = $connectionString
            $connection.Open()
        }
        catch 
        {
            $MyError = $_.Exception.Message
            $hash.Window.Dispatcher.Invoke(
                [action]{
                    $hash.ActionOutput.Text = "[ERROR} Could not connect to SQL Server database! $MyError"
            })
            return
        }
        
        if ($MDTIntegrated -eq 'True')
        {
            # Get Unknown Computers from ConfigMgr database if there are any
            $Query = "
                Select Distinct Name0,
                SMBIOS_GUID0 as 'GUID'
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and DATEDIFF(hour,ExecutionTime,GETDATE()) < $TimePeriod
                --and Name0 like '$SQLComputerName'
                and ExitCode <> $ExitCode
                ORDER BY Name0 Desc
            "
            $command = $connection.CreateCommand()
            $command.CommandText = $Query
            $result = $command.ExecuteReader()
            $table = New-Object -TypeName 'System.Data.DataTable'
            $table.Load($result)

        
            # Gather unknowns into PS object    
            $UnknownComputers = @()
            Foreach ($Row in $table.Rows | Where-Object -FilterScript {
                    $_.Name0 -eq 'Unknown'
            })
            {
                $obj = New-Object -TypeName psobject
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ComputerName' -Value $Row.Name0
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GUID' -Value $Row.GUID
                $UnknownComputers += $obj
            }


            # If there are unknowns, get computername from MDT
            if ($UnknownComputers.Count -ge 1)
            {
                $URL1 = $URL.Replace('Computers','ComputerIdentities')

                # Get ID numbers and Identifiers (GUIDs)
                function GetMDTIDs 
                { 
                    param ($URL)
                    $Data = Invoke-RestMethod -Uri $URL
                    foreach($property in ($Data.content.properties) ) 
                    { 
                        New-Object -TypeName PSObject -Property @{
                            ID         = $($property.ID.'#text')
                            Identifier = $($property.Identifier)
                        }
                    } 
                }
                
                # Filter out only the GUIDs
                $MDTIDs = GetMDTIDs -URL $URL1 |
                Select-Object -Property * |
                Where-Object -FilterScript {
                    $_.Identifier -like '*-*'
                } |
                Sort-Object -Property ID
                $MDTComputerIDs = @()
                Foreach ($Computer in $UnknownComputers)
                {
                    $MDTComputerID = $MDTIDs |
                    Where-Object -FilterScript {
                        $_.Identifier -eq $Computer.GUID
                    } |
                    Select-Object -Property ID, Identifier 
                    $MDTComputerIDs += $MDTComputerID
                }
                
                # Get ComputerNames from MDT
                function GetMDTComputerNames 
                { 
                    param ($URL)
                    $Data = Invoke-RestMethod -Uri $URL
                    foreach($property in ($Data.content.properties) ) 
                    { 
                        New-Object -TypeName PSObject -Property @{
                            Name = $($property.Name)
                            ID   = $($property.ID.'#text')
                        } 
                    } 
                } 
                
                # Filter out the computer names from the IDs
                $MDTComputers = GetMDTComputerNames -URL $URL |
                Select-Object -Property * |
                Sort-Object -Property ID

                $ResolvedComputerNames = @()
                Foreach ($MDTComputerID in $MDTComputerIDs)
                {
                    $MDTComputerName = $MDTComputers |
                    Where-Object -FilterScript {
                        $_.ID -eq $MDTComputerID.ID
                    } |
                    Select-Object -ExpandProperty Name
                    $GUID = $MDTIDs |
                    Where-Object -FilterScript {
                        $_.ID -eq $MDTComputerID.ID
                    } |
                    Select-Object -ExpandProperty Identifier
                    $obj = New-Object -TypeName PSObject
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name ComputerName -Value $MDTComputerName
                    Add-Member -InputObject $obj -MemberType NoteProperty -Name GUID -Value $GUID
                    $ResolvedComputerNames += $obj
                }
            }
            foreach ($Computer in $ResolvedComputerNames)
            {
                if ($ComputerName -eq $Computer.ComputerName)
                {
                    $MyGUID = $Computer.GUID
                }
            }
        }


        # Get TS execution data from ConfigMgr
        if ($MyGUID)
        {
            $Query = "
                Select Distinct Name0 as 'Computer Name',
                sys.SMBIOS_GUID0 as 'GUID',
                Name as 'Task Sequence',
                ExecutionTime,
                Step,
                ActionName,
                GroupName,
                tes.LastStatusMsgName,
                ExitCode,
                ActionOutput
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_RA_System_MACAddresses mac on tes.ResourceID = mac.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and DATEDIFF(hour,ExecutionTime,GETDATE()) < $TimePeriod
                --and Name0 like '$SQLComputerName'
                and sys.SMBIOS_GUID0 = '$MyGUID'
                and ExitCode <> $ExitCode
                ORDER BY ExecutionTime Desc
            "
            $ErrQuery = "
                Select Count(Name0) as 'Count'
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_RA_System_MACAddresses mac on tes.ResourceID = mac.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and DATEDIFF(hour,ExecutionTime,GETDATE()) < $TimePeriod
                --and Name0 like '$SQLComputerName'
                and sys.SMBIOS_GUID0 = '$MyGUID'
                and ExitCode <> 0
            "
        }

        if (!$MyGUID)
        {
            $Query = "
                Select Distinct Name0 as 'Computer Name',
                sys.SMBIOS_GUID0 as 'GUID',
                Name as 'Task Sequence',
                ExecutionTime,
                Step,
                ActionName,
                GroupName,
                tes.LastStatusMsgName,
                ExitCode,
                ActionOutput
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_RA_System_MACAddresses mac on tes.ResourceID = mac.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and DATEDIFF(hour,ExecutionTime,GETDATE()) < $TimePeriod
                and Name0 like '$SQLComputerName'
                and ExitCode <> $ExitCode
                ORDER BY ExecutionTime Desc
            "
            $ErrQuery = "
                Select Count(Name0) as 'Count'
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_RA_System_MACAddresses mac on tes.ResourceID = mac.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and DATEDIFF(hour,ExecutionTime,GETDATE()) < $TimePeriod
                and Name0 like '$SQLComputerName'
                and ExitCode <> 0
            "
        }
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $result = $command.ExecuteReader()
        $table = New-Object -TypeName 'System.Data.DataTable'
        $table.Load($result)
        $command = $connection.CreateCommand()
        $command.CommandText = $ErrQuery
        $erresult = $command.ExecuteReader()
        $errtable = New-Object -TypeName 'System.Data.DataTable'
        $errtable.Load($erresult)
        $connection.Close()

        if ($table.rows.Count -lt 1)
        {
            $hash.Window.Dispatcher.Invoke(
                [action]{
                    $hash.ActionOutput.Text = 'No results.'
            })
            return
        }

        # Gather results into psobject            
        $global:Results = @()
        $i = 0
        Foreach ($Row in $table.Rows)
        {
            $obj = New-Object -TypeName psobject
            $i ++
            if ($Row.ExitCode -eq 0)
            {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Icon' -Value $greentickiconpath
            }
            if ($Row.ExitCode -ne 0)
            {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Icon' -Value $redcrossiconpath
            }
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ComputerName' -Value $Row.'Computer Name'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GUID' -Value $Row.'GUID'
            if ($DTFormat -eq 'UTC')
            {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ExecutionTime' -Value $Row.'ExecutionTime'
            }
            Else 
            {
                $extime = [System.TimeZone]::CurrentTimeZone.ToLocalTime($($Row.'ExecutionTime' | Get-Date))
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ExecutionTime' -Value $extime
            }
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Step' -Value $Row.'Step'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ActionName' -Value $Row.'ActionName'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GroupName' -Value $Row.'GroupName'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'LastStatusMsgName' -Value $Row.'LastStatusMsgName'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ExitCode' -Value $Row.'ExitCode'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ActionOutput' -Value $Row.'ActionOutput'
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Record' -Value $i
            $Results += $obj
        }
        if ($Results.Count -eq 1)
        {
            $obj = New-Object -TypeName psobject
            $i ++
            if ($Row.ExitCode -eq 0)
            {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Icon' -Value $greentickiconpath
            }
            if ($Row.ExitCode -ne 0)
            {
                Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Icon' -Value $redcrossiconpath
            }
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ComputerName' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GUID' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ExecutionTime' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Step' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ActionName' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GroupName' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'LastStatusMsgName' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ExitCode' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ActionOutput' -Value ' '
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'Record' -Value ' '
            $Results += $obj
        }

        $FilteredResults = $Results | Select-Object -Property Icon, ComputerName, GUID, ExecutionTime, Step, ActionName, GroupName, LastStatusMsgName, ExitCode, Record

        if (!$MyGUID -and $ComputerName -in ('-ALL-', '', $Null))
        {
            if ($FilteredResults.ComputerName -match 'unknown')
            {
                foreach ($Computer in $ResolvedComputerNames)
                {
                    $Unknowns = $FilteredResults | Where-Object -FilterScript {
                        $_.ComputerName -eq 'Unknown' -and $_.GUID -eq $Computer.GUID
                    }
                    $i = -1
                    do
                    {
                        $i ++
                        $Unknowns[$i].ComputerName = $Computer.ComputerName
                    }
                    until ($i -eq ($Unknowns.Count -1))
                }
            }
        }

        if ($MyGUID)
        {
            $Unknowns = $FilteredResults | Where-Object -FilterScript {
                $_.ComputerName -eq 'Unknown' -and $_.GUID -eq $MyGUID
            }
            if ($Unknowns)
            {
                $i = -1
                do
                {
                    $i ++
                    $Unknowns[$i].ComputerName = $ComputerName
                }
                until ($i -eq ($Unknowns.Count -1))
            }
        }
        
        # Display results in datagrid         
        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.DataGrid.ItemsSource = $FilteredResults
                $hash.ErrorCount.Text = $errtable.Count
                $hash.ActionOutput.Text = 'Click any step to see the action output.'
        })
    }

    # Set variables from Hash table
    $SQLServer = $hash.SQLServer.Text
    $Database = $hash.Database.Text
    $TimePeriod = $hash.TimePeriod.Text
    $ErrorsOnly = $hash.ErrorsOnly.IsChecked
    $ComputerName = $hash.ComputerName.SelectedItem
    $TS = $hash.TaskSequence.SelectedItem
    $MDTIntegrated = $hash.MDTIntegrated.IsChecked
    $URL = $hash.MDTURL.Text
    $DTFormat = $hash.DTFormat.SelectedItem

    # Create PS instance in runspace pool and execute
    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($SQLServer).AddArgument($Database).AddArgument($TimePeriod).AddArgument($ErrorsOnly).AddArgument($ComputerName).AddArgument($TS).AddArgument($MDTIntegrated).AddArgument($URL).AddArgument($DTFormat)

    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

Function Populate-ActionOutput 
{
    param ($hash,$RunspacePool)

    $code = 
    {
        param($hash,$Record)
        $msg = $Results |
        Select-Object -Property * |
        Where-Object -FilterScript {
            $_.Record -eq $Record
        }
        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.ActionOutput.Text = $msg.ActionOutput
        })
    }

    # Set variables from Hash table
    $Record = $hash.DataGrid.SelectedItem.Record

    # Create PS instance in runspace pool and execute
    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($Record)
    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

Function Populate-ComputerNames 
{
    param ($hash,$RunspacePool)

    $code = 
    {
        param($hash,$SQLServer,$Database,$TimePeriod,$ErrorsOnly,$TS,$MDTIntegrated,$URL)
        
        # Set variable values
        if ($ErrorsOnly -eq 'True')
        {
            $ExitCode = 0
        }
        else 
        {
            $ExitCode = 999999999999999999999999
        }

        # Connect to SQL Server
        $connectionString = "Server=$SQLServer;Database=$Database;Integrated Security=SSPI;"
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open()

        # Run SQL query
        $Query = "
            Select Distinct Name0,
            SMBIOS_GUID0 as 'GUID'
            from vSMS_TaskSequenceExecutionStatus tes
            inner join v_R_System sys on tes.ResourceID = sys.ResourceID
            inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
            where tsp.Name = '$TS'
            and DATEDIFF(hour,ExecutionTime,GETDATE()) < $TimePeriod
            --and Name0 like '$SQLComputerName'
            and ExitCode <> $ExitCode
            ORDER BY Name0 Desc
        "
        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $result = $command.ExecuteReader()
        $table = New-Object -TypeName 'System.Data.DataTable'
        $table.Load($result)
        $connection.Close()
         
        # Gather results into PS object    
        $PCResults = @()
        Foreach ($Row in $table.Rows)
        {
            $obj = New-Object -TypeName psobject
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'ComputerName' -Value $Row.Name0
            Add-Member -InputObject $obj -MemberType NoteProperty -Name 'GUID' -Value $Row.GUID
            $PCResults += $obj
        }
        
        # For each 'Unknown' computer in the list, get the PC name from MDT
        if ($MDTIntegrated -eq 'true')
        {
            $URL1 = $URL.Replace('Computers','ComputerIdentities')

            # Get ID numbers and Identifiers (GUIDs)
            function GetMDTIDs 
            { 
                param ($URL)
                $Data = Invoke-RestMethod -Uri $URL
                foreach($property in ($Data.content.properties) ) 
                { 
                    New-Object -TypeName PSObject -Property @{
                        ID         = $($property.ID.'#text')
                        Identifier = $($property.Identifier)
                    }
                } 
            }
                
            # Filter out only the GUIDs
            $MDTIDs = GetMDTIDs -URL $URL1 |
            Select-Object -Property * |
            Where-Object -FilterScript {
                $_.Identifier -like '*-*'
            } |
            Sort-Object -Property ID
            $UnknownComputers = $PCResults | Where-Object -FilterScript {
                $_.ComputerName -eq 'Unknown'
            }
            $MDTComputerIDs = @()
            Foreach ($Computer in $UnknownComputers)
            {
                $MDTComputerID = $MDTIDs |
                Where-Object -FilterScript {
                    $_.Identifier -eq $Computer.GUID
                } |
                Select-Object -Property ID 
                $MDTComputerIDs += $MDTComputerID
            }
                
            # Get ComputerNames from MDT
            function GetMDTComputerNames 
            { 
                param ($URL)
                $Data = Invoke-RestMethod -Uri $URL
                foreach($property in ($Data.content.properties) ) 
                { 
                    New-Object -TypeName PSObject -Property @{
                        Name = $($property.Name)
                        ID   = $($property.ID.'#text')
                    } 
                } 
            } 
                
            # Filter out the computer names from the IDs
            $MDTComputers = GetMDTComputerNames -URL $URL |
            Select-Object -Property * |
            Sort-Object -Property ID

            $AdditionalComputerNames = @()
            Foreach ($MDTComputerID in $MDTComputerIDs)
            {
                $MDTComputerName = $MDTComputers |
                Where-Object -FilterScript {
                    $_.ID -eq $MDTComputerID.ID
                } |
                Select-Object -ExpandProperty Name
                $AdditionalComputerNames += $MDTComputerName.ToUpper()
            }
                
            $ConfigMgrList = $PCResults |
            Select-Object -Property ComputerName |
            Where-Object -FilterScript {
                $_.ComputerName -ne 'Unknown'
            }
            $FinalComputerNameList = @()
            $FinalComputerNameList += $ConfigMgrList.ComputerName
            $FinalComputerNameList += $AdditionalComputerNames
        }  
        
        # Add a wildcard option and add only ConfigMgr results if MDT not enabled
        if ($MDTIntegrated -eq $false)
        {
            $FinalComputerNameList = @()
            $PCResults = $PCResults | Select-Object -ExpandProperty ComputerName
            $FinalComputerNameList += $PCResults
        }
        $FinalComputerNameList += '-All-'
         
  
        # Display results in ComputerName comboxbox     
        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.ComputerName.ItemsSource = [Array]$FinalComputerNameList
        })
    }

    # Set variables from Hash table
    $SQLServer = $hash.SQLServer.Text
    $Database = $hash.Database.Text
    $TimePeriod = $hash.TimePeriod.Text
    $ErrorsOnly = $hash.ErrorsOnly.IsChecked
    $TS = $hash.TaskSequence.SelectedItem
    $MDTIntegrated = $hash.MDTIntegrated.IsChecked
    $URL = $hash.MDTURL.Text

    # Create PS instance in runspace pool and execute
    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($SQLServer).AddArgument($Database).AddArgument($TimePeriod).AddArgument($ErrorsOnly).AddArgument($TS).AddArgument($MDTIntegrated).AddArgument($URL)
    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

Function Enable-MDT 
{
    #Dispose-PSInstances
    $hash.DeploymentStatus.IsEnabled = 'True'
    $hash.CurrentStep.IsEnabled = 'True'
    $hash.StepName.IsEnabled = 'True'
    $hash.PercentComplete.IsEnabled = 'True'
    $hash.MDTStartTime.IsEnabled = 'True'
    $hash.MDTEndTime.IsEnabled = 'True'
    $hash.MDTElapsedTime.IsEnabled = 'True'

    $hash.DeploymentStatusLabel.IsEnabled = 'True'
    $hash.CurrentStepLabel.IsEnabled = 'True'
    $hash.StepNameLabel.IsEnabled = 'True'
    $hash.PercentCompleteLabel.IsEnabled = 'True'
    $hash.StartLabel.IsEnabled = 'True'
    $hash.EndLabel.IsEnabled = 'True'
    $hash.ElapsedLabel.IsEnabled = 'True'
    $hash.ProgressLabel.IsEnabled = 'True'
}

Function Disable-MDT 
{
    #Dispose-PSInstances
    $hash.DeploymentStatus.IsEnabled = $false
    $hash.CurrentStep.IsEnabled = $false
    $hash.StepName.IsEnabled = $false
    $hash.PercentComplete.IsEnabled = $false
    $hash.MDTStartTime.IsEnabled = $false
    $hash.MDTEndTime.IsEnabled = $false
    $hash.MDTElapsedTime.IsEnabled = $false

    $hash.DeploymentStatusLabel.IsEnabled = $false
    $hash.CurrentStepLabel.IsEnabled = $false
    $hash.StepNameLabel.IsEnabled = $false
    $hash.PercentCompleteLabel.IsEnabled = $false
    $hash.StartLabel.IsEnabled = $false
    $hash.EndLabel.IsEnabled = $false
    $hash.ElapsedLabel.IsEnabled = $false
    $hash.ProgressLabel.IsEnabled = $false
}

Function Get-MDTData 
{
    param ($hash,$RunspacePool)

    $code = 
    {
        param($hash,$URL,$ComputerName,$IsMDTIntegrated,$DTFormat)

        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.DeploymentStatus.Text = ''
                $hash.CurrentStep.Text = ''
                $hash.StepName.Text = ''
                $hash.PercentComplete.Text = ''
                $hash.MDTStartTime.Text = ''
                $hash.MDTEndTime.Text = ''
                $hash.MDTElapsedTime.Text = ''
                $hash.ProgressBar.Value = 0
        })

        if ($IsMDTIntegrated -eq $true -and $ComputerName -ne '-All-' -and $ComputerName -ne '' -and $ComputerName -ne $Null)
        {
            function GetMDTData2 
            { 
                param ($URL)
                $Data = Invoke-RestMethod -Uri $URL
        
                foreach($property in ($Data.content.properties) ) 
                { 
                    New-Object -TypeName PSObject -Property @{
                        Name             = $($property.Name)
                        PercentComplete  = $($property.PercentComplete.'#text')
                        CurrentStep      = $($property.CurrentStep.'#text')
                        StepName         = $($property.StepName)
                        Warnings         = $($property.Warnings.'#text')
                        Errors           = $($property.Errors.'#text')
                        DeploymentStatus = $( 
                            Switch ($property.DeploymentStatus.'#text') { 
                                1 
                                {
                                    'Active/Running'
                                } 
                                2 
                                {
                                    'Failed'
                                } 
                                3 
                                {
                                    'Successfully completed'
                                } 
                                Default 
                                {
                                    'Unknown'
                                } 
                            } 
                        )
                        StartTime        = $($property.StartTime.'#text') -replace 'T', ' '
                        EndTime          = $($property.EndTime.'#text') -replace 'T', ' '
                    } 
                } 
            } 
            try 
            {
                $MDT = GetMDTData2 -URL $URL | 
                Select-Object -Property Name, DeploymentStatus, PercentComplete, CurrentStep, StepName, Warnings, Errors, StartTime, EndTime | 
                Sort-Object -Property Name | 
                Where-Object -FilterScript {
                    $_.Name -eq $ComputerName
                }
                        
                if ($MDT)
                {
                    # Calculate times
                    $MDTServer = $URL.Split('//')[2].Split(':')[0]
                    $Start = $MDT.StartTime | Get-Date
                    if ($DTFormat -ne 'UTC')
                    {
                        $Start = [System.TimeZone]::CurrentTimeZone.ToLocalTime($Start)
                    }
                    if (!$MDT.EndTime)
                    {
                        $MDTDate = Invoke-Command -ComputerName $MDTServer -ScriptBlock {
                            (Get-Date).ToUniversalTime()
                        }
                        if ($DTFormat -ne 'UTC')
                        {
                            $MDTDate = [System.TimeZone]::CurrentTimeZone.ToLocalTime($MDTDate)
                        }
                        $Elapsed = $MDTDate - $Start
                        $Elapsed = "$($Elapsed.Hours)h $($Elapsed.Minutes)m $($Elapsed.Seconds)s"
                        $Elapsed
                    }
                    if ($MDT.EndTime)
                    {
                        $End = $MDT.EndTime | Get-Date
                        if ($DTFormat -ne 'UTC')
                        {
                            $End = [System.TimeZone]::CurrentTimeZone.ToLocalTime($End)
                        }
                        $Elapsed = $End - $Start
                        $Elapsed = "$($Elapsed.Hours)h $($Elapsed.Minutes)m $($Elapsed.Seconds)s"
                        $Elapsed
                    }

                    $hash.Window.Dispatcher.Invoke(
                        [action]{
                            $hash.DeploymentStatus.Text = $MDT.DeploymentStatus
                            $hash.CurrentStep.Text = $MDT.CurrentStep
                            $hash.StepName.Text = $MDT.StepName
                            $hash.PercentComplete.Text = $MDT.PercentComplete
                            $hash.ProgressBar.Value = $MDT.PercentComplete
                            $hash.MDTStartTime.Text = $Start
                            $hash.MDTEndTime.Text = $End
                            $hash.MDTElapsedTime.Text = $Elapsed
                    })
                }
                Else 
                {
                    $hash.Window.Dispatcher.Invoke(
                        [action]{
                            $hash.DeploymentStatus.Text = 'No data found'
                    })
                }
            }
            catch
            {
                $hash.Window.Dispatcher.Invoke(
                    [action]{
                        $hash.ActionOutput.Text = '[ERROR] Could not connect to MDT Web Service'
                })
            }
        }
    }

    # Set variables from Hash table
    $ComputerName = $hash.ComputerName.SelectedItem
    $IsMDTIntegrated = $hash.MDTIntegrated.IsChecked
    $URL = $hash.MDTURL.Text
    $DTFormat = $hash.DTFormat.SelectedItem

    # Create PS instance in runspace pool and execute
    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($URL).AddArgument($ComputerName).AddArgument($IsMDTIntegrated).AddArgument($DTFormat)
    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

function Dispose-PSInstances 
{
    foreach ($PSinstance in $PSInstances)
    {
        if ($PSinstance.InvocationStateInfo.State -eq 'Completed')
        {
            $PSinstance.Dispose()
        }
    }
}

Function Create-Timer 
{
    $global:Timer = New-Object -TypeName System.Windows.Forms.Timer
    $timer.Interval = [int]$hash.RefreshPeriod.Text * 60000
}

Function Start-Timer 
{
    if ($timer)
    {
        $timer.Start()
    }
}

Function Stop-Timer 
{
    if ($timer)
    {
        $timer.Stop()
    }
}

Function Update-Registry 
{
    param($hash)
    # Test whether running as admin first
    If (([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator'))
    {
        if (Test-Path -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor')
        {
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name SQLServer -Value $hash.SQLServer.Text
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name Database -Value $hash.Database.Text
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name MDTURL -Value $hash.MDTURL.Text
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name DTFormat -Value $hash.DTFormat.SelectedItem
        }

        if (Test-Path -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor')
        {
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name SQLServer -Value $hash.SQLServer.Text
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name Database -Value $hash.Database.Text
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name MDTURL -Value $hash.MDTURL.Text
            Set-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name DTFormat -Value $hash.DTFormat.SelectedItem
        }
    }
}

Function Read-Registry 
{
    if (Test-Path -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor')
    {
        $regsql = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name SQLServer | Select-Object -ExpandProperty SQLServer
        $regdb = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name Database | Select-Object -ExpandProperty Database
        $regmdt = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name MDTURL | Select-Object -ExpandProperty MDTURL
        $regdtformat = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Wow6432Node\SMSAgent\ConfigMgr Task Sequence Monitor' -Name DTFormat | Select-Object -ExpandProperty DTFormat
    }

    if (Test-Path -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor')
    {
        $regsql = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name SQLServer | Select-Object -ExpandProperty SQLServer
        $regdb = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name Database | Select-Object -ExpandProperty Database
        $regmdt = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name MDTURL | Select-Object -ExpandProperty MDTURL
        $regdtformat = Get-ItemProperty -Path 'HKLM:\SOFTWARE\SMSAgent\ConfigMgr Task Sequence Monitor' -Name DTFormat | Select-Object -ExpandProperty DTFormat
    }
    if ($regsql -ne $Null -and $regsql -ne '')
    {
        $hash.SQLServer.Text = $regsql
    }
    if ($regdb -ne $Null -and $regdb -ne '')
    {
        $hash.Database.Text = $regdb
    }
    if ($regmdt -ne $Null -and $regmdt -ne '')
    {
        $hash.MDTURL.Text = $regmdt
    }

    if ($regdtformat -ne $Null -and $regdtformat -ne '')
    {
        if ($regdtformat -eq 'UTC')
        {
            if (!$CurrentDateTimeF)
            {
                $Global:CurrentDateTimeF = 'UTC'
            }
        }
    }
}

Function Generate-Report 
{
    param ($hash,$RunspacePool)

    $code = 
    {
        param($hash,$SQLServer,$Database,$StartDate,$EndDate,$TS,$DTFormat)
        $Results = @()

        if ($DTFormat -ne 'UTC')
        {
            [datetime]$StartDate = $StartDate.ToUniversalTime()
            [datetime]$EndDate = $EndDate.ToUniversalTime()
        }

        # Set dates to ISO standard format for SQL Server
        $SQLStart = $StartDate | Get-Date -Format s
        $SQLEnd = $EndDate | Get-Date -Format s

        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.Working.Content = 'Working...'
                $hash.ReportProgress.Visibility = 'Visible'
                $hash.ReportProgress.Value = 10
        })

        $connectionString = "Server=$SQLServer;Database=$Database;Integrated Security=SSPI;"
        $connection = New-Object -TypeName System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open()

        # Find all resourceID for TS steps between the selected dates
        $Query = "
            select distinct tes.ResourceID
            from vSMS_TaskSequenceExecutionStatus tes
            --inner join v_R_System sys on tes.ResourceID = sys.ResourceID
            inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
            where tsp.Name = '$TS'
            and tes.ExecutionTime >= '$SQLStart'
            and tes.ExecutionTime <= '$SQLEnd'
        "

        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $reader = $command.ExecuteReader()
        $table = New-Object -TypeName 'System.Data.DataTable'
        $table.Load($reader)

        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.ReportProgress.Value = 20
        })

        foreach ($ResourceID in $table.Rows.ResourceID)
        {
            $Query = "
                Select (select top(1) convert(datetime,ExecutionTime,121)
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and tes.ExecutionTime >= '$SQLStart' 
                and tes.ExecutionTime <= '$SQLEnd'
                and LastStatusMsgName = 'The task sequence execution engine started execution of a task sequence'
                and Step = 0
                and tes.ResourceID = $ResourceID
                order by ExecutionTime desc) as 'Start',
                (select top(1) convert(datetime,ExecutionTime,121)
                from vSMS_TaskSequenceExecutionStatus tes
                inner join v_R_System sys on tes.ResourceID = sys.ResourceID
                inner join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
                where tsp.Name = '$TS'
                and tes.ExecutionTime >= '$SQLStart'
                and tes.ExecutionTime <= '$SQLEnd'
                and LastStatusMsgName = 'The task sequence execution engine successfully completed a task sequence'
                and tes.ResourceID = $ResourceID
                order by ExecutionTime desc) as 'Finish',
                (Select name0 from v_R_System sys where sys.ResourceID = $ResourceID) as 'ComputerName',
                (select Model0 from v_GS_Computer_System comp where comp.ResourceID = $ResourceID) as 'Model'
            "
            $command = $connection.CreateCommand()
            $command.CommandText = $Query
            $reader = $command.ExecuteReader()
            $table = New-Object -TypeName 'System.Data.DataTable'
            $table.Load($reader)


            if ($table.rows[0].Start.GetType().Name -eq 'DBNull')
            {
                $Start = ''
            }
            Else 
            {
                if ($DTFormat -eq 'UTC')
                {
                    $Start = $table.rows[0].Start
                }
                Else 
                {
                    $Start = [System.TimeZone]::CurrentTimeZone.ToLocalTime($($table.rows[0].Start | Get-Date))
                }
            }

            if ($table.rows[0].Finish.GetType().Name -eq 'DBNull')
            {
                $Finish = ''
            }
            Else 
            {
                if ($DTFormat -eq 'UTC')
                {
                    $Finish = $table.rows[0].Finish
                }
                Else 
                {
                    $Finish = [System.TimeZone]::CurrentTimeZone.ToLocalTime($($table.rows[0].Finish | Get-Date))
                }
            }


            #$table
            if ($Start -eq '' -or $Finish -eq '')
            {
                $diff = $Null
            }
            else 
            {
                $diff = $Finish-$Start
            }


            $PC = New-Object -TypeName psobject
            Add-Member -InputObject $PC -MemberType NoteProperty -Name ComputerName -Value $table.rows[0].ComputerName
            Add-Member -InputObject $PC -MemberType NoteProperty -Name StartTime -Value $Start
            Add-Member -InputObject $PC -MemberType NoteProperty -Name FinishTime -Value $Finish
            if ($Start -eq '' -or $Finish -eq '')
            {
                Add-Member -InputObject $PC -MemberType NoteProperty -Name DeploymentTime -Value ''
            }
            else
            {
                Add-Member -InputObject $PC -MemberType NoteProperty -Name DeploymentTime -Value $("$($diff.hours)" + ' hours ' + "$($diff.minutes)" + ' minutes')
            }
            Add-Member -InputObject $PC -MemberType NoteProperty -Name Model -Value $table.rows[0].Model
            $Results += $PC
        }

        $Results = $Results | Sort-Object -Property ComputerName

        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.ReportProgress.Value = 50
        })

        $Query = "
            select sys.Name0 as 'ComputerName',
            tsp.Name 'Task Sequence',
            comp.Model0 as Model,
            tes.ExecutionTime,
            tes.Step,
            tes.GroupName,
            tes.ActionName,
            tes.LastStatusMsgName,
            tes.ExitCode,
            tes.ActionOutput
            from vSMS_TaskSequenceExecutionStatus tes
            left join v_R_System sys on tes.ResourceID = sys.ResourceID
            left join v_TaskSequencePackage tsp on tes.PackageID = tsp.PackageID
            left join v_GS_COMPUTER_SYSTEM comp on tes.ResourceID = comp.ResourceID
            where tsp.Name = '$TS'
            and tes.ExecutionTime >= '$SQLStart'
            and tes.ExecutionTime <= '$SQLEnd'
            and tes.ExitCode not in (0,-2147467259)
            Order by tes.ExecutionTime desc
        "

        $command = $connection.CreateCommand()
        $command.CommandText = $Query
        $reader = $command.ExecuteReader()
        $table = New-Object -TypeName 'System.Data.DataTable'
        $table.Load($reader)

        if ($DTFormat -ne 'UTC')
        {
            $newdates = foreach ($item in $table.rows.ExecutionTime)
            {
                [System.TimeZone]::CurrentTimeZone.ToLocalTime($item)
            }
            $i = -1
            $table.rows.ExecutionTime | ForEach-Object -Process {
                $i ++
                $table.Rows[$i].ExecutionTime = $newdates[$i]
            }
        }

        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.ReportProgress.Value = 80
        })

        #Convert dates if necessary
        if ($DTFormat -ne 'UTC')
        {
            $StartDate = [System.TimeZone]::CurrentTimeZone.ToLocalTime($StartDate)
            $EndDate = [System.TimeZone]::CurrentTimeZone.ToLocalTime($EndDate)
        }         

        # Create html email
        $style = @"
<style>
body {
    color:#012E34;
    font-family:Calibri,Tahoma;
    font-size: 10pt;
}
h1 {
    text-align:center;
}
h2 {
    border-top:1px solid #666666;
}
 
 
th {
    font-weight:bold;
    color:#012E34;
    background-color:#69969C;
}
.odd  { background-color:#012E34; }
.even { background-color:#012E34; }
</style>
"@


        $HEaders = @"
<H1>Task Sequence Execution Summary Report</H1>
<H3>Starting Date: $StartDate</H3>
<H3>End Date: $EndDate</H3>
<H3>Task Sequence: $TS</H3>
<H3>TimeZone for Date/Time: $DTFormat</H3>
"@

        $body1 = $Results | 
        Select-Object -Property ComputerName, StartTime, FinishTime , DeploymentTime, Model |
        ConvertTo-Html -Head $style -Body "<H2>Task Sequence Executions ($($Results.Count))</H2>" | 
        Out-String

        $body2 = $table | 
        Select-Object -Property ComputerName, 'Task Sequence', Model, ExecutionTime, Step, GroupName, ActionName, LastStatusMsgName, ExitCode |
        ConvertTo-Html -Head $style -Body "<H2>Task Sequence Execution Errors ($($table.Rows.Count))</H2>" | 
        Out-String

        $Body = $HEaders + $body1 + $body2

        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.Working.Content = ''
                $hash.ReportProgress.Value = 100
        })

        $Body | Out-File -FilePath $env:temp\TSReport.htm -Force
        Invoke-Item -Path $env:temp\TSReport.htm


        # Close the connection
        $connection.Close()
    }

    # Set variables from Hash table
    $SQLServer = $hash.SQLServer.Text
    $Database = $hash.Database.Text
    $DTFormat = $hash.DTFormat.SelectedItem
    #if ($DTFormat -eq "UTC")
    #   {
    [datetime]$StartDate = $hash.StartDate.Text | Get-Date -Format "MM'/'dd'/'yyyy HH':'mm':'ss"
    [datetime]$EndDate = $hash.EndDate.Text | Get-Date -Format "MM'/'dd'/'yyyy HH':'mm':'ss"
    #  }
    #Else {
    #       [datetime]$StartDate = [System.TimeZone]::CurrentTimeZone.ToLocalTime($($hash.StartDate.Text | Get-Date -Format "MM'/'dd'/'yyyy HH':'mm':'ss"))
    #      [datetime]$EndDate = [System.TimeZone]::CurrentTimeZone.ToLocalTime($($hash.EndDate.Text | Get-Date -Format "MM'/'dd'/'yyyy HH':'mm':'ss"))
    # }

    $EndDate = $EndDate.AddDays(1).AddSeconds(-1)
    $TS = $hash.TSList.SelectedItem

    # Create PS instance in runspace pool and execute
    $PSinstance = [powershell]::Create().AddScript($code).AddArgument($hash).AddArgument($SQLServer).AddArgument($Database).AddArgument($StartDate).AddArgument($EndDate).AddArgument($TS).AddArgument($DTFormat)
    $PSInstances += $PSinstance
    $PSinstance.RunspacePool = $RunspacePool
    $PSinstance.BeginInvoke()
}

Function Clear-MDT 
{
    param ($hash)

    $hash.Window.Dispatcher.Invoke(
        [action]{
            $hash.DeploymentStatus.Text = ''
            $hash.CurrentStep.Text = ''
            $hash.StepName.Text = ''
            $hash.PercentComplete.Text = ''
            $hash.ProgressBar.Value = 0
            $hash.MDTStartTime.Text = ''
            $hash.MDTEndTime.Text = ''
            $hash.MDTElapsedTime.Text = ''
    })
}

#endregion

#region Event Handlers

$hash.Window.Add_ContentRendered({
        #Disable-MDT
        Read-Registry
        Get-DateTimeFormat
        Get-TaskSequenceList
})

$hash.TaskSequence.Add_SelectionChanged({
        $Count = $hash.ComputerName.Items.Count
        $hash.Window.Dispatcher.Invoke(
            [action]{
                $hash.ComputerName.SelectedIndex = ($Count -1)
        })
        Dispose-PSInstances
        Clear-MDT -hash $hash
        Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
        Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
        Get-MDTData -hash $hash -RunspacePool $RunspacePool
        Stop-Timer
        Create-Timer
        $timer.add_Tick({
                Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
                Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
                Get-MDTData -hash $hash -RunspacePool $RunspacePool
        })
        Start-Timer
        $Global:CurrentTS = $hash.TaskSequence.SelectedItem
})

$hash.ErrorsOnly.Add_Checked({
        Dispose-PSInstances
        Stop-Timer
        Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
        Start-Timer
})

$hash.MDTIntegrated.Add_Checked({
        Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
        Enable-MDT
})

$hash.MDTIntegrated.Add_Unchecked({
        Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
        Disable-MDT
})

$hash.ErrorsOnly.Add_Unchecked({
        Dispose-PSInstances
        Stop-Timer
        Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
        Start-Timer
})

$hash.DataGrid.Add_SelectionChanged({
        Dispose-PSInstances
        Populate-ActionOutput -hash $hash -RunspacePool $RunspacePool
})

$hash.RefreshNow.Add_Click({
        Dispose-PSInstances
        Stop-Timer
        Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
        Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
        Get-MDTData -hash $hash -RunspacePool $RunspacePool
        $timer.Interval = [int]$hash.RefreshPeriod.Text * 60000
        Start-Timer
})

$hash.ComputerName.Add_SelectionChanged({
        if ($hash.TaskSequence.SelectedItem -eq $CurrentTS)
        {
            Dispose-PSInstances
            Stop-Timer
            Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
            Get-MDTData -hash $hash -RunspacePool $RunspacePool
            Start-Timer
        }
})

$hash.TimePeriod.Add_KeyDown({
        if ($_.Key -eq 'Return')
        {
            Dispose-PSInstances
            Stop-Timer
            Get-TaskSequenceData -hash $hash -RunspacePool $RunspacePool
            Populate-ComputerNames -hash $hash -RunspacePool $RunspacePool
            Get-MDTData -hash $hash -RunspacePool $RunspacePool
            $timer.Interval = [int]$hash.RefreshPeriod.Text * 60000
            Start-Timer
        }
})

$hash.RefreshPeriod.Add_TextChanged({
        Stop-Timer
        $timer.Interval = [int]$hash.RefreshPeriod.Text * 60000
        Start-Timer
})

$hash.SettingsButton.Add_Click({
        $reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml2)
        $hash.Window2 = [Windows.Markup.XamlReader]::Load( $reader )
        $hash.SQLServer = $hash.Window2.FindName('SQLServer')
        $hash.Database = $hash.Window2.FindName('Database')
        $hash.MDTURL = $hash.Window2.FindName('MDTURL')
        $hash.ConnectSQL = $hash.Window2.FindName('ConnectSQL')
        $hash.TSList = $hash.Window2.FindName('TSList')
        $hash.StartDate = $hash.Window2.FindName('StartDate')
        $hash.EndDate = $hash.Window2.FindName('EndDate')
        $hash.GenerateReport = $hash.Window2.FindName('GenerateReport')
        $hash.SettingsTab = $hash.Window2.FindName('SettingsTab')
        $hash.ReportTab = $hash.Window2.FindName('ReportTab')
        $hash.Tabs = $hash.Window2.FindName('Tabs')
        $hash.Runasadmin = $hash.Window2.FindName('Runasadmin')
        $hash.Working = $hash.Window2.FindName('Working')
        $hash.ReportProgress = $hash.Window2.FindName('ReportProgress')
        $hash.Link1 = $hash.Window2.FindName('Link1')
        $hash.DTFormat = $hash.Window2.FindName('DTFormat')
        if (Test-Path -Path "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
        {
            #$hash.Window2.Icon = "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
        }
        if (Test-Path -Path "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
        {
            #$hash.Window2.Icon = "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
        }

        Read-Registry
        $hash.SettingsTab.Focus()
        If (!(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')))
        {
            $hash.Runasadmin.Visibility = 'Visible'
        }
        $hash.TSList.ItemsSource = $Views.TS
        $hash.DTFormat.ItemsSource = $Timezones.TimeZone
        If ($CurrentDateTimeF -eq 'UTC')
        {
            $hash.DTFormat.SelectedIndex = 0
        }
        Else 
        {
            $hash.DTFormat.SelectedIndex = 1
        }

        $hash.SQLServer.Add_GotMouseCapture({
                if ($hash.SQLServer.Text -eq '<SQLServer\Instance>')
                {
                    $hash.SQLServer.Text = ''
                }
        })

        $hash.SQLServer.Add_GotKeyboardFocus({
                if ($hash.SQLServer.Text -eq '<SQLServer\Instance>')
                {
                    $hash.SQLServer.Text = ''
                }
        })

        $hash.Database.Add_GotMouseCapture({
                if ($hash.Database.Text -eq '<Database>')
                {
                    $hash.Database.Text = ''
                }
        })

        $hash.Database.Add_GotKeyboardFocus({
                if ($hash.Database.Text -eq '<Database>')
                {
                    $hash.Database.Text = ''
                }
        })
        $hash.ConnectSQL.Add_Click({
                Update-Registry -hash $hash
                Get-TaskSequenceList
        })

        $hash.GenerateReport.Add_Click({
                Generate-Report -hash $hash -RunspacePool $RunspacePool
        })

        $hash.Link1.Add_Click({
                Start-Process -FilePath 'http://smsagent.wordpress.com/tools/configmgr-task-sequence-monitor/'
        })

        $hash.DTFormat.Add_SelectionChanged({
                $Global:CurrentDateTimeF = $hash.DTFormat.SelectedItem
        })

        $hash.Window2.Add_Closed({
                Update-Registry -hash $hash
        })

        $Null = $hash.Window2.ShowDialog()
})

$hash.ReportButton.Add_Click({
        $reader = (New-Object -TypeName System.Xml.XmlNodeReader -ArgumentList $xaml2)
        $hash.Window2 = [Windows.Markup.XamlReader]::Load( $reader )
        $hash.SQLServer = $hash.Window2.FindName('SQLServer')
        $hash.Database = $hash.Window2.FindName('Database')
        $hash.MDTURL = $hash.Window2.FindName('MDTURL')
        $hash.ConnectSQL = $hash.Window2.FindName('ConnectSQL')
        $hash.TSList = $hash.Window2.FindName('TSList')
        $hash.StartDate = $hash.Window2.FindName('StartDate')
        $hash.EndDate = $hash.Window2.FindName('EndDate')
        $hash.GenerateReport = $hash.Window2.FindName('GenerateReport')
        $hash.SettingsTab = $hash.Window2.FindName('SettingsTab')
        $hash.ReportTab = $hash.Window2.FindName('ReportTab')
        $hash.Tabs = $hash.Window2.FindName('Tabs')
        $hash.Runasadmin = $hash.Window2.FindName('Runasadmin')
        $hash.Working = $hash.Window2.FindName('Working')
        $hash.ReportProgress = $hash.Window2.FindName('ReportProgress')
        $hash.Link1 = $hash.Window2.FindName('Link1')
        $hash.DTFormat = $hash.Window2.FindName('DTFormat')
        if (Test-Path -Path "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
        {
            #$hash.Window2.Icon = "$env:ProgramFiles\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
        }
        if (Test-Path -Path "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico")
        {
            #$hash.Window2.Icon = "${env:ProgramFiles(x86)}\SMSAgent\ConfigMgr Task Sequence Monitor\Grid.ico"
        }

        Read-Registry
        $hash.ReportTab.Focus()
        If (!(([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')))
        {
            $hash.Runasadmin.Visibility = 'Visible'
        }
        $hash.TSList.ItemsSource = $Views.TS
        $hash.DTFormat.ItemsSource = $Timezones.TimeZone
        If ($CurrentDateTimeF -eq 'UTC')
        {
            $hash.DTFormat.SelectedIndex = 0
        }
        Else 
        {
            $hash.DTFormat.SelectedIndex = 1
        }

        $hash.SQLServer.Add_GotMouseCapture({
                if ($hash.SQLServer.Text -eq '<SQLServer\Instance>')
                {
                    $hash.SQLServer.Text = ''
                }
        })

        $hash.SQLServer.Add_GotKeyboardFocus({
                if ($hash.SQLServer.Text -eq '<SQLServer\Instance>')
                {
                    $hash.SQLServer.Text = ''
                }
        })

        $hash.Database.Add_GotMouseCapture({
                if ($hash.Database.Text -eq '<Database>')
                {
                    $hash.Database.Text = ''
                }
        })

        $hash.Database.Add_GotKeyboardFocus({
                if ($hash.Database.Text -eq '<Database>')
                {
                    $hash.Database.Text = ''
                }
        })

        $hash.ConnectSQL.Add_Click({
                Update-Registry -hash $hash
                Get-TaskSequenceList
        })

        $hash.GenerateReport.Add_Click({
                Generate-Report -hash $hash -RunspacePool $RunspacePool
        })

        $hash.Link1.Add_Click({
                Start-Process -FilePath 'http://smsagent.wordpress.com/tools/configmgr-task-sequence-monitor/'
        })

        $hash.DTFormat.Add_SelectionChanged({
                $Global:CurrentDateTimeF = $hash.DTFormat.SelectedItem
        })

        $hash.Window2.Add_Closed({
                Update-Registry -hash $hash
        })

        $Null = $hash.Window2.ShowDialog()
})

$hash.Window.Add_Closed({
        Stop-Timer
        Dispose-PSInstances
        $RunspacePool.close()
        $RunspacePool.Dispose()
})

# Stop process on closing, #comment our for development
$hash.window.Add_Closing({[System.Windows.Forms.Application]::Exit(); Stop-Process $pid})
#endregion


# Make PowerShell Disappear #comment our for development
$windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' 
$asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru 
$null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0)
    
#$app = [Windows.Application]::new()
$app = New-Object Windows.Application
$app.Run($Hash.Window)

