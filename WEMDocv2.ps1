<#
.SYNOPSIS
This script documents the Citrix Workspace Environment Management Solution.

Data is gathered using Arjan Mensch's Citrix.WEMSDK PowerShell Module

Output is sent to both Word and HTML format using the PSCribo PowerShell Module

.DESCRIPTION

.PARAMETER DBServer
Mandatory parameter specifying your SQL Database Server name or instance

.PARAMETER DBName
Mandatory parameter specifying your Citrix WEM Database Name

.PARAMETER Site
Specifies the WEM Configuration Set to document via Site ID. Defaults to Site ID 1 (Default Site)

If you do not know your site ID, use the -listAllConfigSets parameter

.PARAMETER ListAllConfigSets
Optional Parameter. Creates an initial connection to the WEM Database and lists all Configuration Sets

.PARAMETER CompanyName
Optional Parameter used to personalise the Document Output for a particular Customer Name

.PARAMETER Detailed
Optional Parameter which will output an appendix with full details for applications,rules etc

.PARAMETER OutputLocation
Optional Parameter allowing you so specify a custom output directory. Defaults to ~\Desktop

.EXAMPLE
Documents WEM based on the default SQL instance found on SERVER against the Database named WEM. Default Site (1) is used.

.\DocumentWEM_V2.ps1 -DBServer SERVER -DBName CitrixWEM
.EXAMPLE
Documents WEM based on the SQL instance named INSTANCE found on SERVER against the Database named WEM. Default Site (1) is used.

.\DocumentWEM_V2.ps1 -DBServer SERVER\INSTANCE -DBName CitrixWEM
.EXAMPLE
Lists all Config Sets in WEM based on the SQL instance named INSTANCE found on SERVER against the Database named WEM.

.\DocumentWEM_V2.ps1 -DBServer SERVER\INSTANCE -DBName CitrixWEM -ListAllConfigSets
.EXAMPLE
Documents WEM based on the SQL instance named INSTANCE found on SERVER against the Database named WEM. Site 2 is used.

.\DocumentWEM_V2.ps1 -DBServer SERVER\INSTANCE -DBName CitrixWEM -Site 2
.EXAMPLE
Documents WEM based on the SQL instance named INSTANCE found on SERVER against the Database named WEM. Default Site (1) is used. A detailed Appendix is added

.\DocumentWEM_V2.ps1 -DBServer SERVER\INSTANCE -DBName CitrixWEM -Detailed
.EXAMPLE
Documents WEM based on the SQL instance named INSTANCE found on SERVER against the Database named WEM. Default Site (1) is used. A detailed Appendix is added. The Output Location is specified

.\DocumentWEM_V2.ps1 -DBServer SERVER\INSTANCE -DBName CitrixWEM -Detailed -OutPutLocation "C:\Temp\Output"
.EXAMPLE
Documents WEM based on the SQL instance named INSTANCE found on SERVER against the Database named WEM. A detailed Appendix is added. Default Site (1) is used. a Company Name is Added

.\DocumentWEM_V2.ps1 -DBServer SERVER\INSTANCE -DBName CitrixWEM -Detailed -CompanyName "KindonEnterprises"

.NOTES
Credits as follows:
Arjan Mensch. For being the PowerShell master in relation to WEM https://github.com/msfreaks 
Aaron Parker. Functions, Parsing and basic powershell guidance throughout the project https://github.com/aaronparker/
Iain Brighton. For PSCribo https://github.com/iainbrighton/PScribo

.LINK

#>
# ============================================================================
# Parameters
# ============================================================================
Param(
    [Parameter(Mandatory = $true)]
    [string]$DBServer,

    [Parameter(Mandatory = $true)]
    [string]$DBName,

    [Parameter(Mandatory = $false)]
    [int]$Site = 1,

    [Parameter(Mandatory = $false)]
    [string]$OutputLocation = "~\Desktop",

    [Parameter(Mandatory = $false)]
    [switch]$Detailed,

    [Parameter(Mandatory = $false)]
    [switch]$ListAllConfigSets,
    
    [Parameter(Mandatory = $false)]
    [string]$CompanyName
)

#region Translation
# ============================================================================
# Translation Details: SQL DB Record Name -> Console Description
# ============================================================================
$DescriptionTable = @{
    #region Environmental Settings
    # Environmental Settings -> Start Menu
    # Environmental Settings Management
    processEnvironmentalSettings                            = "Process Environmental Settings"
    processEnvironmentalSettingsForAdmins                   = "Exclude Administrators"
    
    # User Interface: Start Menu
    HideCommonPrograms                                      = "Hide Common Programs"
    RemoveRunFromStartMenu                                  = "Remove Run From Start Menu"
    HideAdministrativeTools                                 = "Hide Administrative Tools"
    HideHelp                                                = "Hide Help"
    HideFind                                                = "Hide Find"
    HideWindowsUpdate                                       = "Hide Windows Update"
    LockTaskbar                                             = "Lock Taskbar"
    HideSystemClock                                         = "Hide System Clock"
    HideDevicesandPrinters                                  = "Hide Devices and Printers"
    HideTurnOff                                             = "Hide Turn Off Computer"
    ForceLogoff                                             = "Force Logoff Button"
    Turnoffnotificationareacleanup                          = "Turn Off Notification Area Cleanup"
    TurnOffpersonalizedmenus                                = "Turn Off Personalized Menus"
    ClearRecentprogramslist                                 = "Clear Recent Programs List"
    
    # User Interface: Appearance
    SetSpecificThemeFile                                    = "Set Specific Theme File"
    SpecificThemeFileValue                                  = "Specific Theme File"
    SetVisualStyleFile                                      = "Set Visual Style File"
    VisualStyleFileValue                                    = "Visual Style File"
    SetWallpaper                                            = "Set Wallpaper"
    Wallpaper                                               = "Wallpaper"
    WallpaperStyle                                          = "Wallpaper Style"
    SetDesktopBackGroundColor                               = "Set Desktop BackGround Color"
    DesktopBackGroundColor                                  = "Desktop BackGround Color"
    
    # Environmental Settings -> Desktop
    # User Interface: Desktop
    NoMyComputerIcon                                        = "Hide My Computer Icon"
    NoRecycleBinIcon                                        = "Hide Recycle Bin Icon"
    NoMyDocumentsIcon                                       = "Hide My Documents Icon"
    BootToDesktopInsteadOfStart                             = "Go To Desktop Instead Of Start"
    NoPropertiesMyComputer                                  = "Disable System Properties"
    NoPropertiesRecycleBin                                  = "Disable Recycle Bin Properties"
    NoPropertiesMyDocuments                                 = "Disable My Documents Properties"
    HideNetworkIcon                                         = "Hide Network Icon"
    HideNetworkConnections                                  = "Hide Network Connections"
    DisableTaskMgr                                          = "Disable Task Manager"
    
    # User Interface: Edge UI
    DisableTLcorner                                         = "Disable Switcher"
    DisableCharmsHint                                       = "Disable Charms Hint"
    
    # Environmental Settings -> Windows Explorer
    # User Interface: Explorer
    DisableRegistryEditing                                  = "Prevent Access to Registry Editing Tools"
    DisableSilentRegedit                                    = "Disable Silent Regedit"
    DisableCmd                                              = "Prevent Access to the Command Prompt"
    DisableCmdScripts                                       = "Disable Cmd Scripts"
    RemoveContextMenuManageItem                             = "Remove Context Menu Manage Item"
    NoNetConnectDisconnect                                  = "Remove Network Context Menu Item"
    HideLibrairiesInExplorer                                = "Hide Libraries In Explorer" #Typo in DB value - Not me!
    HideNetworkInExplorer                                   = "Hide Network Icon In Explorer"
    HideControlPanel                                        = "Hide Programs Control Panel"
    NoNtSecurity                                            = "Disable Windows Security"
    NoViewContextMenu                                       = "Disable Explorer Context Menu"
    NoTrayContextMenu                                       = "Disable Taskbar Context Menu"
    
    #Drive Restrictions
    HideSpecifiedDrivesFromExplorer                         = "Hide Specified Drives From Explorer"
    ExplorerHiddenDrives                                    = "Hidden Drives"
    RestrictSpecifiedDrivesFromExplorer                     = "Restrict Specified Drives From Explorer"
    ExplorerRestrictedDrives                                = "Restricted Drives"
    
    # Environmental Settings -> Windows Explorer
    # User Interface: Control Panel
    NoProgramsCPL                                           = "Hide Control Panel"
    RestrictCpl                                             = "Show Only Specified Control Panel Applets"
    RestrictCplList                                         = "Allowed Control Panel Applets"
    DisallowCpl                                             = "Hide Specified Control Panel Applets"
    DisallowCplList                                         = "Hideen Control Panel Applets"
    
    # Environmental Settings -> Known Folders Management
    # Known Folders Restrictions
    DisabledKnownFolders                                    = "Disable Specified Known Folders"
    DisableSpecifiedKnownFolders                            = "Disabled Known Folders"
    
    # Environmental Settings -> SBC / HVD Tuning
    DisableDragFullWindows                                  = "Disable Drag Full Windows"
    DisableCursorBlink                                      = "Disable Cursor Blink"
    EnableAutoEndTasks                                      = "Enable Auto End Tasks"
    WaitToKillAppTimeout                                    = "WaitToKillApp Timeout"
    SetCursorBlinkRate                                      = "Set Cursor Blink Rate"
    CursorBlinkRateValue                                    = "Cursor Blink Rate"
    SetMenuShowDelay                                        = "Set Menu Show Delay"
    MenuShowDelayValue                                      = "Menu Show Delay"
    SetInteractiveDelay                                     = "Set Interactive Delay"
    InteractiveDelayValue                                   = "Interactive Delay"
    DisableSmoothScroll                                     = "Disable Smooth Scroll"
    DisableMinAnimate                                       = "Disable MinAnimate"
    #endregion
    #region Advanced Settings
    # Advanced Settings -> Configuration -> Main Configuration
    # Agent Actions
    processVUEMApps                                         = "Process Applications"
    processVUEMPrinters                                     = "Process Printers"
    processVUEMNetDrives                                    = "Process Network Drives"
    processVUEMVirtualDrives                                = "Process Virtual Drives"
    processVUEMRegValues                                    = "Process Registry Values"
    processVUEMEnvVariables                                 = "Process Environment Variables"
    processVUEMPorts                                        = "Process Ports"
    processVUEMIniFilesOps                                  = "Process Ini File Operations"
    processVUEMExtTasks                                     = "Process External Tasks"
    processVUEMFileSystemOps                                = "Process File System Operations"
    processVUEMUserDSNs                                     = "Process DSNS"
    processVUEMFileAssocs                                   = "Process File Associations"
    
    # Agent Service Actions
    LaunchVUEMAgentOnLogon                                  = "Launch Agent at Logon"
    LaunchVUEMAgentOnReconnect                              = "Launch Agent on Reconnect"
    ProcessVUEMAgentLaunchForAdmins                         = "Launch Agent for Admins"
    VUEMAgentType                                           = "Agent Type"
    EnableVirtualDesktopCompatibility                       = "Enable (virtual) Desktop Compatibility"
    ExecuteOnlyCmdAgentInPublishedApplications              = "Execute Only CMD Agent In Published Applications"
    
    # Shortcut Deletion at startup
    DeleteDesktopShortcuts                                  = "Delete Desktop Shortcuts"
    DeleteStartMenuShortcuts                                = "Delete Start Menu Shortcuts"
    DeleteQuickLaunchShortcuts                              = "Delete Quick Launch Shortcuts"
    DeleteTaskBarPinnedShortcuts                            = "Delete TaskBar Pinned Shortcuts"
    DeleteStartMenuPinnedShortcuts                          = "Delete Start Menu Pinned Shortcuts"
    
    # Drives Deletion at Startup
    DeleteNetworkDrives                                     = "Delete Network Drives"
    
    # Printers Deletion at Startup
    DeleteNetworkPrinters                                   = "Delete Network Printers on Startup"
    PreserveAutocreatedPrinters                             = "Preserve Autocreated Printers"
    PreserveSpecificPrinters                                = "Preserve Specific Printers"
    SpecificPreservedPrinters                               = "Specific Preserved Printer List"
    
    # Advanced Settings -> Configuration -> Agent Options
    # Agent Logs
    EnableAgentLogging                                      = "Enable Agent Logging"
    AgentLogFile                                            = "Log File"
    AgentDebugMode                                          = "Debug Mode"
    
    # Offline Mode Settings
    OfflineModeEnabled                                      = "Enable Offline Mode"
    UseCacheEvenIfOnline                                    = "Use Cache Even If Online"
    
    #Refresh Settings
    RefreshEnvironmentSettings                              = "Refresh Environment Settings"
    RefreshSystemSettings                                   = "Refresh System Settings"
    RefreshOnEnvironmentalSettingChange                     = "Refresh On Environmental Setting Change"
    RefreshDesktop                                          = "Refresh Desktop"
    RefreshAppearance                                       = "Refresh Appearance"
    
    #Asynchronous Processing
    aSyncVUEMPrintersProcessing                             = "aSync Printers Processing"
    aSyncVUEMNetDrivesProcessing                            = "aSync Network Drives Processing"
    aSyncVUEMAppsProcessing                                 = "" #<- Doesnt Exist in Console
    aSyncVUEMPortsProcessing                                = "" #<- Doesnt Exist in Console
    aSyncVUEMRegValuesProcessing                            = "" #<- Doesnt Exist in Console
    aSyncVUEMFileSystemOpsProcessing                        = "" #<- Doesnt Exist in Console
    aSyncVUEMIniFilesOpsProcessing                          = "" #<- Doesnt Exist in Console
    aSyncVUEMFileAssocsProcessing                           = "" #<- Doesnt Exist in Console
    aSyncVUEMExtTasksProcessing                             = "" #<- Doesnt Exist in Console
    aSyncVUEMUserDSNsProcessing                             = "" #<- Doesnt Exist in Console
    aSyncVUEMEnvVariablesProcessing                         = "" #<- Doesnt Exist in Console
    aSyncVUEMVirtualDrivesProcessing                        = "" #<- Doesnt Exist in Console
    
    #Extra Features
    InitialEnvironmentCleanUp                               = "Initial Environment CleanUp"
    InitialDesktopUICleaning                                = "Initial Desktop UI CleanUp"
    checkAppShortcutExistence                               = "Check Application Existence"
    appShortcutExpandEnvironmentVariables                   = "Expand App Variables"
    AgentEnableCrossDomainsUserGroupsSearch                 = "Enable Cross Domains User Groups Search"
    AgentBrokerServiceTimeoutValue                          = "Broker Service Timeout (ms)"
    AgentDirectoryServiceTimeoutValue                       = "Directory Service Timeout (ms)"
    AgentNetworkResourceCheckTimeoutValue                   = "Network Resource Timeout (ms)"
    AgentMaxDegreeOfParallelism                             = "Agent Max Degree Of Parallelism"
    ConnectionStateChangeNotificationEnabled                = "Enable Notifications"
    
    # Advanced Settings -> Configuration -> Advanced Options
    # Agent Actions Enforce Execution
    enforceProcessVUEMApps                                  = "Enforce Applications Processing"
    enforceProcessVUEMPrinters                              = "Enforce Printers Processing"
    enforceProcessVUEMNetDrives                             = "Enforce Network Drives Processing"
    enforceProcessVUEMVirtualDrives                         = "Enforce Virtual Drives Processing"
    enforceProcessVUEMEnvVariables                          = "Enforce Environment Variables Processing"
    enforceProcessVUEMPorts                                 = "Enforce Ports Processing"
    enforceProcessVUEMFileSystemOps                         = "" #<- Doesnt Exist in Console
    enforceProcessVUEMFileAssocs                            = "" #<- Doesnt Exist in Console
    enforceProcessVUEMUserDSNs                              = "" #<- Doesnt Exist in Console
    enforceProcessVUEMRegValues                             = "" #<- Doesnt Exist in Console
    enforceProcessVUEMIniFilesOps                           = "" #<- Doesnt Exist in Console
    enforceProcessVUEMExtTasks                              = "" #<- Doesnt Exist in Console
    
    # Unassigned Actions Revert Processing
    revertUnassignedVUEMApps                                = "Revert Unassigned Applications"
    revertUnassignedVUEMPrinters                            = "Revert Unassigned Printers"
    revertUnassignedVUEMNetDrives                           = "Revert Unassigned Network Drives"
    revertUnassignedVUEMVirtualDrives                       = "Revert Unassigned Virtual Drives"
    revertUnassignedVUEMRegValues                           = "Revert Unassigned Registry Values"
    revertUnassignedVUEMEnvVariables                        = "Revert Unassigned Ports"
    revertUnassignedVUEMPorts                               = "Revert Unassigned Ports"
    revertUnassignedVUEMIniFilesOps                         = "Revert Unassigned Ini Files Operations"
    revertUnassignedVUEMExtTasks                            = "Revert Unassigned External Tasks"
    revertUnassignedVUEMFileSystemOps                       = "Revert Unassigned File System Operations"
    revertUnassignedVUEMUserDSNs                            = "Revert Unassigned User DSNs"
    revertUnassignedVUEMFileAssocs                          = "Revert Unassigned File Associations"
    
    # Automatic Refresh (UI Agent Only)
    EnableUIAgentAutomaticRefresh                           = "Enable Automatic Refresh"
    UIAgentAutomaticRefreshDelay                            = "Refresh Delay (min)"
    
    # Advanced Settings -> Configuration -> Reconnection Actions
    processVUEMAppsOnReconnect                              = "Process Applications"
    processVUEMPrintersOnReconnect                          = "Process Printers"
    processVUEMNetDrivesOnReconnect                         = "Process Network Drives"
    processVUEMVirtualDrivesOnReconnect                     = "Process Virtual Drives"
    processVUEMRegValuesOnReconnect                         = "Process Registry Values"
    processVUEMEnvVariablesOnReconnect                      = "Process Environment Variables"
    processVUEMPortsOnReconnect                             = "Process Ports"
    processVUEMIniFilesOpsOnReconnect                       = "Process Ini File Operations"
    processVUEMExtTasksOnReconnect                          = "Process External Tasks"
    processVUEMFileSystemOpsOnReconnect                     = "Process File System Operations"
    processVUEMUserDSNsOnReconnect                          = "Process User DSNs"
    processVUEMFileAssocsOnReconnect                        = "Process File Associations"
    
    # Advanced Settings -> Configuration -> Advanced Processing
    enforceVUEMAppsFiltersProcessing                        = "Enforce Applications Filters Processing"
    enforceVUEMPrintersFiltersProcessing                    = "Enforce Printers Filters Processing"
    enforceVUEMNetDrivesFiltersProcessing                   = "Enforce Network Drives Filters Processing"
    enforceVUEMVirtualDrivesFiltersProcessing               = "Enforce Virtual Drives Filters Processing"
    enforceVUEMRegValuesFiltersProcessing                   = "Enforce Registry Values Filters Processing"
    enforceVUEMEnvVariablesFiltersProcessing                = "Enforce Environment Variables Filters Processing"
    enforceVUEMPortsFiltersProcessing                       = "Enforce Ports Filters Processing"
    enforceVUEMIniFilesOpsFiltersProcessing                 = "Enforce Ini File Operations Filters Processing"
    enforceVUEMExtTasksFiltersProcessing                    = "Enforce External Tasks Filters Processing"
    enforceVUEMFileSystemOpsFiltersProcessing               = "Enforce File System Operations Filters Processing"
    enforceVUEMUserDSNsFiltersProcessing                    = "Enforce User DSNs Filters Processing"
    enforceVUEMFileAssocsFiltersProcessing                  = "Enforce File Associations Filters Processing"
    
    # Advanced Settings -> Configuration -> Service Options
    # Agent Service Advanced Options
    VUEMAgentCacheRefreshDelay                              = "Agent Cache Refresh Delay (min)"
    VUEMAgentSQLSettingsRefreshDelay                        = "SQL Settings Refresh Delay (min)"
    VUEMAgentDesktopsExtraLaunchDelay                       = "Agent Extra Launch Delay (ms)"
    AgentServiceDebugMode                                   = "Enable Debug mode"
    byPassie4uinitCheck                                     = "byPass ie4uinit Check"
    
    # Agent Launch Exclusions
    AgentLaunchExcludeGroups                                = "Do not launch VUEM agent for specifed Groups"
    AgentLaunchExcludedGroups                               = "Excluded Groups"
    AgentLaunchIncludeGroups                                = "Launch VUEM agent for specifed Groups"
    AgentLaunchIncludedGroups                               = "Included Groups"
    
    # Advanced Settings -> Configuration -> Agent Switch
    AgentSwitchFeatureToggle                                = ""
    SwitchtoServiceAgent                                    = "Switch to Service Agent"
    CloudConnectors                                         = "Configure Citrix Cloud Connectors"
    UseGPO                                                  = "Skip Citrix Cloud Connector Configuration"
    
    # Advanced Settings -> UI Agent Personalization -> UI Agent Options
    # Branding
    UIAgentSplashScreenBackGround                           = "Custom Background Image Path"
    UIAgentLoadingCircleColor                               = "Loading Circle Color"
    UIAgentLbl1TextColor                                    = "Text Label Color"
    UIAgentSkinName                                         = "UI Agent Skin"
    HideUIAgentSplashScreen                                 = "Hide Agent Splashscreen"
    HideUIAgentSplashScreenOnReconnect                      = "Hide Agent Splashscreen on Reconnection"
    
    # Published Applications Behavior
    HideUIAgentIconInPublishedApplications                  = "Hide Agent Icon In Published Applications"
    HideUIAgentSplashScreenInPublishedApplications          = "Hide Agent Splash Screen In Published Applications"
    
    # User Interaction
    AgentExitForAdminsOnly                                  = "Only Admins can Close Agent"
    AgentAllowUsersToManagePrinters                         = "Allow Users To Manage Printers"
    AgentAllowUsersToManageApplications                     = "Allow Users To Manage Applications"
    AgentPreventExitForAdmins                               = "Prevent Admins from Closing Agent"
    AgentEnableApplicationsShortcuts                        = "Enable Applications Shortcuts"
    DisableAdministrativeRefreshFeedback                    = "Disable Administrative Refresh Feedback"
    
    # Advanced Settings -> UI Agent Personalization -> Helpdesk Options
    # Help & Custom Links
    UIAgentHelpLink                                         = "Help Link Action"
    UIAgentCustomLink                                       = "Custom Link Action"
    
    # Screen Capture Options
    AgentAllowScreenCapture                                 = "Enable Screen Capture"
    AgentScreenCaptureEnableSendSupportEmail                = "Enable Send to Support Option"
    AgentScreenCaptureSupportEmailAddress                   = "Support Email Address"
    MailSMTPToAddress                                       = ""
    MailCustomSubject                                       = "Custom Subject"
    AgentScreenCaptureSupportEmailTemplate                  = "Email Template"
    MailEnableUseSMTP                                       = "Use SMTP to send Email"
    MailSMTPServer                                          = "SMTP Server"
    MailSMTPPort                                            = "SMTP Port"
    MailEnableSMTPSSL                                       = "Require SSL"
    MailSMTPFromAddress                                     = "From Address"
    MailEnableUseSMTPCredentials                            = "Use SMTP Credentials"
    MailSMTPUser                                            = "User Name"
    MailSMTPPassword                                        = "Password"
    
    # Power Saving
    AgentShutdownAfterEnabled                               = "Shut down at Specified time (HH:MM)"
    AgentShutdownAfter                                      = "Shut down time"
    AgentShutdownAfterIdleEnabled                           = "Shut down When Idle (seconds)"
    AgentShutdownAfterIdleTime                              = "Idle Time"
    AgentSuspendInsteadOfShutdown                           = "Suspend Instead Of Shutdown"
    #endregion
    #region System Optimization
    # System Optimization -> CPU Management
    # Spikes Protection
    EnableCPUSpikesProtection                               = "Enable CPU Spike Protection"
    AutoCPUSpikeProtectionSelected                          = "Auto Prevent CPU Spikes" 
    SpikesProtectionCPUUsageLimitPercent                    = "CPU Usage Limit (%)"
    SpikesProtectionCPUUsageLimitSampleTime                 = "Limit Sample Time (s)"
    SpikesProtectionIdlePriorityConstraintTime              = "Idle Priority Time (s)"
    SpikesProtectionCPUCoreLimit                            = "Enable CPU Core Usage Limit"
    SpikesProtectionLimitCPUCoreNumber                      = "Limit CPU Core Usage"
    EnableIntelligentCpuOptimization                        = "Enable Intelligent CPU Optimization"
    EnableIntelligentIoOptimization                         = "Enable Intelligent I/O Optimization"
    ExcludeProcessesFromCPUSpikesProtection                 = "Exclude Specified Processes"
    CPUSpikesProtectionExcludedProcesses                    = "Excluded Processes"
    
    # CPU Priority
    EnableProcessesCpuPriority                              = "Enable Process Priority"
    ProcessesCpuPriorityList                                = "Process Priority List"
    
    # CPU Affinity
    EnableProcessesAffinity                                 = "Enable Process Affinity" 
    ProcessesAffinityList                                   = "Process Affinity List"
    
    # CPU Clamping
    EnableProcessesClamping                                 = "Enable Process Clamping"
    ProcessesClampingList                                   = "Process Clamping List"
    
    # System Optimization -> Memory Management
    # Memory Management
    EnableMemoryWorkingSetOptimization                      = "Enable Working Set Optimization"
    MemoryWorkingSetOptimizationIdleSampleTime              = "Idle Sample Time (min)"
    MemoryWorkingSetOptimizationIdleStateLimitPercent       = "Idle State Limit (percent) (if value = enabled then 1%)"
    ExcludeProcessesFromMemoryWorkingSetOptimization        = "Exclude Specified Processes"
    MemoryWorkingSetOptimizationExcludedProcesses           = "Excluded Processes"
    
    # System Optimization -> I/O Priority
    # I/O Priority Process List
    EnableProcessesIoPriority                               = "Enable Processes I/O Priority"
    ProcessesIoPriorityList                                 = "Process List"
    
    # System Optimization -> Fast LogOff
    # Settings
    EnableFastLogoff                                        = "Enable Fast Logoff"
    ExcludeGroupsFromFastLogoff                             = "Exclude Specified Groups"
    FastLogoffExcludedGroups                                = "Excluded Groups"
    #endregion
    #region Security
    # Security 
    # Process Security
    EnableProcessesManagement                               = "Enable Processes Management"
    
    EnableProcessesBlackListing                             = "Enable Process BlackList"
    ProcessesManagementBlackListedProcesses                 = "BlackListed Processes"
    ProcessesManagementBlackListExcludeLocalAdministrators  = "Exclude Local Administrators"
    ProcessesManagementBlackListExcludeSpecifiedGroups      = "Exclude Specified Groups"
    ProcessesManagementBlackListExcludedSpecifiedGroupsList = "Excluded Groups"
    
    EnableProcessesWhiteListing                             = "Enable Process Whitelist"
    ProcessesManagementWhiteListedProcesses                 = "Whitelisted Processes"
    ProcessesManagementWhiteListExcludeLocalAdministrators  = "Exclude Local Administrators"
    ProcessesManagementWhiteListExcludeSpecifiedGroups      = "Exclude Specified Groups"
    ProcessesManagementWhiteListExcludedSpecifiedGroupsList = "Excluded Groups"
    
    # App Locker
    AppLockerControllerManagement                           = ""
    AppLockerControllerReplaceModeOn                        = "Process AppLocker Rules in Replace Mode"
    
    EnableProcessesAppLocker                                = "Process Application Security Rules"
    EnableDLLRuleCollection                                 = "Process DLL Rules"
    CollectionExeEnforcementState                           = "Executable Rule Enforcement State"
    CollectionMsiEnforcementState                           = "Windows Installer Rule Enforcement State"
    CollectionScriptEnforcementState                        = "Scripts Rule Enforcement State"
    CollectionAppxEnforcementState                          = "Packaged Rule Enforcement State"
    CollectionDllEnforcementState                           = "DLL Rule Enforcement State"
    #endregion
    #region Policies and Profiles
    #region USV
    # Policies and Profiles -> User State Virtualization -> Roaming Profiles Configuration
    # Process USV
    processUSVConfiguration                                 = "Process User State Virtualization Configuration"
    processUSVConfigurationForAdmins                        = "Exclude Administrators"
    #region Microsoft Roaming Profiles
    # Windows Roaming Profiles Settings
    SetWindowsRoamingProfilesPath                           = "Set Windows Roaming Profiles Path"
    WindowsRoamingProfilesPath                              = "Roaming Profile Path"
    
    # RDS Roaming Profiles Settings
    SetRDSRoamingProfilesPath                               = "Set RDS Roaming Profiles Path"
    RDSRoamingProfilesPath                                  = "RDS Roaming Profile Path"
    SetRDSHomeDrivePath                                     = "Set RDS Home Drive Path"
    RDSHomeDrivePath                                        = "RDS Home Drive Path"
    RDSHomeDriveLetter                                      = "RDS Home Drive Letter"
    
    # Policies and Profiles -> User State Virtualization -> Roaming Profiles Advanced Configuration
    SetRoamingProfilesFoldersExclusions                     = "Enable Folders Exclusions"
    RoamingProfilesFoldersExclusions                        = "Excluded Folders"
    DeleteRoamingCachedProfiles                             = "Delete Cached Copies of Roaming Profiles"
    AddAdminGroupToRUP                                      = "Add the Administrators Security Group to Roaming User Profiles"
    CompatibleRUPSecurity                                   = "Do Not Check for User Ownership of Roaming Profile Folders"
    DisableSlowLinkDetect                                   = "Do Not Detect Slow Network Connections"     
    SlowLinkProfileDefault                                  = "Wait for Remote User Profile"
    #endregion
    #region Folder Redirection
    # Policies and Profiles -> User State Virtualization -> Folder Redirection
    #Folder Redirection
    processFoldersRedirectionConfiguration                  = "Process Folder Redirection Configuration"
    
    # Folder Redirection Process Settings
    processDesktopRedirection                               = "Redirect Desktop"
    processPersonalRedirection                              = "Redirect Documents"
    processPicturesRedirection                              = "Redirect Pictures"
    processMusicRedirection                                 = "Redirect Music"
    processVideoRedirection                                 = "Redirect Videos"
    processStartMenuRedirection                             = "Redirect Start Menu"
    processFavoritesRedirection                             = "Redirect Favorites"
    processAppDataRedirection                               = "Redirect AppData (Roaming)"
    processContactsRedirection                              = "Redirect Contacts"
    processDownloadsRedirection                             = "Redirect Downloads"
    processLinksRedirection                                 = "Redirect Links"
    processSearchesRedirection                              = "Redirect Searches"
    DeleteLocalRedirectedFolders                            = "Delete Local Redirected Folders"
    DesktopRedirectedPath                                   = "Desktop Path"
    PersonalRedirectedPath                                  = "Documents Path"
    PicturesRedirectedPath                                  = "Pictures Path"
    MyPicturesFollowsDocuments                              = "Pictures Follows Documents"
    MusicRedirectedPath                                     = "Music Path"
    MyMusicFollowsDocuments                                 = "Music Follows Documents"
    VideoRedirectedPath                                     = "Videos Path"
    MyVideoFollowsDocuments                                 = "Videos Follows Documents"
    StartMenuRedirectedPath                                 = "Start Menu Path"
    FavoritesRedirectedPath                                 = "Favorites Path"
    AppDataRedirectedPath                                   = "AppData Path"
    ContactsRedirectedPath                                  = "Contacts Path"
    DownloadsRedirectedPath                                 = "Downloads Path"
    LinksRedirectedPath                                     = "Links Path"
    SearchesRedirectedPath                                  = "Searches Path"
    #endregion
    #endregion
    #region Citrix Profile Management
    # Policies and Profiles -> Citrix Profile Management Settings -> Main Citrix Profile Management Settings
    # Citrix profile Management
    UPMManagementEnabled                                    = "Enable Profile Management Configuration"
    
    # Profile Management
    ServiceActive                                           = "Enable Profile Management"
    SetProcessedGroups                                      = "Set Processed Groups"
    ProcessedGroupsList                                     = "Processed Groups List"
    SetExcludedGroups                                       = "Set Excluded Groups"
    ExcludedGroupsList                                      = "Excluded Groups List"
    ProcessAdmins                                           = "Process logons of local administrators"
    SetPathToUserStore                                      = "Set path to user store"
    PathToUserStore                                         = "User Store Path"
    MigrateUserStore                                        = "Migrate User Store"
    MigrateUserStorePath                                    = "Path to the previous user store"
    PSMidSessionWriteBack                                   = "Enable active write back"
    PSMidSessionWriteBackReg                                = "Enable active write back registry"
    OfflineSupport                                          = "Enable offline profile support"
    
    # Policies and Profiles -> Citrix Profile Management Settings -> Profile Handling
    # Profile Handling
    DeleteCachedProfilesOnLogoff                            = "Delete locally cached profiles on logoff"
    SetProfileDeleteDelay                                   = " Set delay before deleting cached profiles"
    ProfileDeleteDelay                                      = "cached profile deletion Delay in seconds"
    SetMigrateWindowsProfilesToUserStore                    = "Enable migration of existing profiles"
    MigrateWindowsProfilesToUserStore                       = "Type of user profiles to be migrated if the user store is empty `n 1: local and roaming `n 2: local `n 3: roaming `n 4: none"
    AutomaticMigrationEnabled                               = "Automatic migration of existing application profiles"
    SetLocalProfileConflictHandling                         = "Enable local profile conflict handling"
    LocalProfileConflictHandling                            = "local profile conflict handling"
    SetTemplateProfilePath                                  = "Enable Template Profile"
    TemplateProfilePath                                     = "Template Profile Path"
    TemplateProfileOverridesLocalProfile                    = "Template profile overrides local profile"
    TemplateProfileOverridesRoamingProfile                  = "Template profile overrides local profile"
    TemplateProfileIsMandatory                              = "Template profile used as Citrix mandatory profile for all logons"
    
    # Policies and Profiles -> Citrix Profile Management Settings -> Advanced Settings
    # Advanced Settings
    SetLoadRetries                                          = "Set number of retires when accessing locked files"
    LoadRetries                                             = "Number of profile load retries"
    XenAppOptimizationEnabled                               = "Enable application profiler"
    XenAppOptimizationPath                                  = "path to application profiler"
    SetUSNDBPath                                            = "Set directory of the MFT cache file"
    USNDBPath                                               = "MFT cache file Absolute Path"
    ProcessCookieFiles                                      = "Process Internet cookie files on logoff"
    DeleteRedirectedFolders                                 = "Delete redirected folders"
    DisableDynamicConfig                                    = "Disable automatic configuration"
    LogoffRatherThanTempProfile                             = "Log off user if a problem is encountered"
    CEIPEnabled                                             = "Customer experience improvement program"
    OutlookSearchRoamingEnabled                             = "Enabled search index roaming for Microsoft Outlook ssers"
    SearchBackupRestoreEnabled                              = "Outlook search index database - backup and restore"
    
    # Policies and Profiles -> Citrix Profile Management Settings -> Log Settings
    # Log Settings
    LoggingEnabled                                          = "Enable Logging"
    SetLogLevels                                            = "Configure Log Setings"
    LogLevels                                               = "Log Level Settings (Check Citrix UPM Doco)"
    SetMaxLogSize                                           = "Set maximum size of the log file"
    MaxLogSize                                              = "Maximum size in bytes"
    SetPathToLogFile                                        = "Set path to log file"
    PathToLogFile                                           = "Path to log file"
    
    # Policies and Profiles -> Citrix Profile Management Settings -> Registry
    # Registry
    LastKnownGoodRegistry                                   = "NTUSER.DAT Backup"
    EnableDefaultExclusionListRegistry                      = "Enable Default Exclusion List"
    ExclusionDefaultRegistry01                              = "Registry Default Exclusion"
    ExclusionDefaultRegistry02                              = "Registry Default Exclusion"
    ExclusionDefaultRegistry03                              = "Registry Default Exclusion"
    SetExclusionListRegistry                                = "Enable Registry Exclusions"
    ExclusionListRegistry                                   = "Registry Exclusions"
    SetInclusionListRegistry                                = "Enable Registry Inclusions"
    InclusionListRegistry                                   = "Registry Inclusions"
    
    # Policies and Profiles -> Citrix Profile Management Settings -> File System
    # File System
    EnableLogonExclusionCheck                               = "Enable Logon Exclusion Check"
    LogonExclusionCheck                                     = "Logon exclusion Check Setting"
    EnableDefaultExclusionListDirectories                   = "Enable Default Exclusion List - Directories"
    ExclusionDefaultDir01                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir02                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir03                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir04                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir05                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir06                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir07                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir09                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir08                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir10                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir11                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir12                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir13                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir14                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir15                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir16                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir17                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir18                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir19                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir20                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir21                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir22                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir23                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir24                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir25                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir26                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir27                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir28                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir29                                   = "Default Citrix Exclusion"
    ExclusionDefaultDir30                                   = "Default Citrix Exclusion"
    SetSyncExclusionListFiles                               = "Enable File Exclusions"
    SyncExclusionListFiles                                  = "File Exclusion List"
    SetSyncExclusionListDir                                 = "Enable Folder exclusions"
    SyncExclusionListDir                                    = "Folder Exclusion List"
    
    # Policies and Profiles -> Citrix Profile Management Settings -> Synchronization
    # Synchronization
    SetSyncDirList                                          = "Enable Directory Synchronization"
    SyncDirList                                             = "Sync Directory List"
    SetSyncFileList                                         = "Enable File Synchronization"
    SyncFileList                                            = "Sync File List"
    SetMirrorFoldersList                                    = "Enable Folder Mirroring"
    MirrorFoldersList                                       = "Mirror Folders List"
    SetProfileContainerList                                 = "Enable Profile Container"
    ProfileContainerList                                    = "Profile Container List"
    SetLargeFileHandlingList                                = "Enable Large File Handling"
    LargeFileHandlingList                                   = "Large File Handling List"
    
    # Policies and Profiles -> Citrix Profile Management Settings -> Streamed User Profiles
    # Streamed User Profiles
    PSEnabled                                               = "Enable Profile Streaming"
    PSAlwaysCache                                           = "Always Cache"
    PSAlwaysCacheSize                                       = "Cache files this size or larger (megabyte)"
    SetPSPendingLockTimeout                                 = "Set timeout for pending area lock files"
    PSPendingLockTimeout                                    = "Timeout for pending area lock files (days)"
    SetPSUserGroupsList                                     = "Set streamed user profile groups"
    PSUserGroupsList                                        = "Streamed user profile groups"
    EnableStreamingExclusionList                            = "Enable Profile Streaming Exclusion List - Directories"
    StreamingExclusionList                                  = "Streaming Excluded Directories List"
    
    # Policies and Profiles -> Citrix Profile Management Settings -> Cross Platform Settings
    # Cross Platform Settings
    CPEnabled                                               = "Enable cross-platform settings"
    SetCPUserGroupList                                      = "Set Cross platform settings groups"
    CPUserGroupList                                         = "Cross platform settings groups"
    SetCPSchemaPath                                         = "Set path to cross-platform definitions"
    CPSchemaPath                                            = "Path to cross-platform definitions"
    SetCPPath                                               = "Set path to cross-platform settings store"
    CPPath                                                  = "Path to cross-platform settings store"
    CPMigrationFromBaseProfileToCPStore                     = "Enable source for creating cross-platform settings"
    #endregion
    #endregion
    #region Transformer Settings
    # Transformer Settings -> General -> General Settings
    # General Settings
    IsKioskEnabled                                          = "Enable Transformer"
    GeneralStartUrl                                         = "Web Interface URL"
    
    # Appearance
    GeneralTitle                                            = "Custom Title"
    
    GeneralWindowMode                                       = "Enable Window Mode"
    GeneralClockEnabled                                     = "Display Clock"
    GeneralClockUses12Hours                                 = "Show 12 Hour-Clock"
    GeneralEnableLanguageSelect                             = "Allow Language Selection"
    GeneralEnableAppPanel                                   = "Enable Application Panel"
    GeneralAutoHideAppPanel                                 = "Auto Hide Application Panel"
    GeneralShowNavigationButtons                            = "Show Navigation Buttons"
    # Change Unlock Password
    GeneralUnlockPassword                                   = "Unlock Password"
    
    # Transformer Settings -> General -> Site Settings
    # Site Settings
    SitesIsListEnabled                                      = "Enable Site List"
    SitesNamesAndLinks                                      = "Site List"
    # Transformer Settings -> General -> Tool Settings
    # Tool Settings
    ToolsEnabled                                            = "Enable Tools List"
    ToolsAppsList                                           = "Tools List"
    # Transformer Settings -> Advanced -> Process Launcher
    # Process Launcher Settings
    ProcessLauncherEnabled                                  = "Enable Process Launcher"
    ProcessLauncherApplication                              = "Process Command Line"
    ProcessLauncherArgs                                     = "Process Arguments"
    ProcessLauncherClearLastUsernameVMWare                  = "Clear Last Username for VMWare View"
    ProcessLauncherEnableVMWareViewMode                     = "Enable VMWare View Mode"
    ProcessLauncherEnableMicrosoftRdsMode                   = "Enabled Microsoft RDS Mode"
    ProcessLauncherEnableCitrixMode                         = "Enable Citrix Mode"
    # Transformer Settings -> Advanced -> Advanced and Administration Settings
    # Advanced Settings
    AdvancedFixBrowserRendering                             = "Fix Browser Rendering"
    AdvancedLogOffScreenRedirection                         = "Log Off Screen Redirection"
    AdvancedSuppressScriptErrors                            = "Supress Script Errors"
    AdvancedFixSslSites                                     = "Fix SSL Sites"
    AdvancedHideKioskWhileCitrixSession                     = "Hide Kiosk While Citrix Session"
    AdvancedAlwaysShowAdminMenu                             = "Always Show Admin Menu"
    AdvancedHideTaskbar                                     = "Hide Taskbar & Start Button"
    AdvancedLockAltTab                                      = "Lock Alt-Tab"
    AdvancedFixZOrder                                       = "Fix Z-Order"
    SetCitrixReceiverFSOMode                                = "Lock Citrix Desktop Viewer"
    #AdvancedShowWifiSettings = "" <- Not in Console
    #AdvancedLockCtrlAltDel = "" <- not in Console
    
    # Administration Settings
    AdministrationHideDisplaySettings                       = "Hide Display Settings"
    AdministrationHideKeyboardSettings                      = "Hide Keyboard Settings"
    AdministrationHideMouseSettings                         = "Hide Mouse Settings"
    AdministrationHideVolumeSettings                        = "Hide Volume Details"
    AdministrationHideClientDetails                         = "Hide Client Details"
    AdministrationDisableProgressBar                        = "Disable Progress Bar"
    AdministrationHideWindowsVersion                        = "Hide Windows Version"
    AdministrationHideHomeButton                            = "Hide Home Button"
    AdministrationHidePrinterSettings                       = "Hide Printer Settings"
    AdministrationPreLaunchReceiver                         = "Pre-Launch Receiver"
    AdministrationDisableUnlock                             = "Disable Unlock"
    AdministrationHideLogOffOption                          = "Hide Log Off Option"
    AdministrationHideRestartOption                         = "Hide Restart Option"
    AdministrationHideShutdownOption                        = "Hide Shutdown Option"
    AdministrationIgnoreLastLanguage                        = "Ignore Last Language"
    # Transformer Settings -> Advanced -> Logon/Logoff & Power Settings
    # Autologon Options
    AutologonEnable                                         = "Enable Autologon Mode"
    AutologonUserName                                       = "User Name"
    AutologonPassword                                       = "Password"
    AutologonDomain                                         = "Domain/PC"
    AutologonRegistryForce                                  = "Autologon Force"
    AutologonRegistryIgnoreShiftOverride                    = "Ignore Shift Override"
    # Desktop Mode Options
    DesktopModeLogOffWebPortal                              = "Log Off Web Portal When a Session is Launched"
    # End of Session Options
    EndSessionOption                                        = "Action to take when the remote session ends"
    # Power Options
    PowerShutdownAfterSpecifiedTime                         = "Shut down at Specified Time (HH:MM)"
    PowerShutdownAfterIdleTime                              = "Shut down When Idel (Seconds)"
    PowerDontCheckBattery                                   = "Don't Check Battery Status"
    #endregion
    #region Monitoring
    # Monitoring
    BusinessDayStartHour                                    = "Business Day Start (hour)"
    BusinessDayEndHour                                      = "Business Day End (hour)"
    EnableWorkDaysFiltering                                 = "Enable Work Days Filtering"
    WorkDaysFilter                                          = "Enabled Work Days"
    ReportsBootTimeMinimum                                  = "Boot Time Minimum Value"
    ReportsLoginTimeMinimum                                 = "Login Time Minimum Value"
    
    EnableApplicationReportsWindows2K3XPCompliance          = ""
    LocalDatabaseRetentionPeriod                            = ""
    EnableUserExperienceMonitoring                          = ""
    ExcludedProcessesFromApplicationReports                 = ""
    LocalDataUploadFrequency                                = ""
    EnableProcessActivityMonitoring                         = ""
    ExcludeProcessesFromApplicationReports                  = ""
    EnableSystemMonitoring                                  = ""
    EnableStrictPrivacy                                     = ""
    EnableGlobalSystemMonitoring                            = ""
    #endregion
    #region Advanced Hidden Params
    # Advanced Parameters
    ADSearchForestBlacklist                                 = ""
    AgentSiteIdCacheOverdueTime                             = ""
    ActionGroupsToggle                                      = ""
    VersionInfo                                             = ""
    GlobalLicenseServerPort                                 = ""
    GlobalLicenseServer                                     = ""
    DisplayUPMStatusToggle                                  = ""
    ProfileContainerToggle                                  = ""
    AgentDomainCacheOverdueTime                             = ""
    #endregion
}
#endregion

#region functions
# ============================================================================
# Functions
# ============================================================================
function CheckModuleExists {
    param (
        $Module
    )

    if (Get-Module -Name $Module) {
        Write-Verbose "Module $Module Exists" -Verbose
    }
    else {
        try {
            Write-Verbose "$Module Module is not installed, attempting to install" -Verbose
            Install-Module -Name $Module -Force
            Import-Module -Name $Module -Force
            Write-Verbose "Success! Module $Module installed" -Verbose
        }
        catch {
            Write-Warning "$Module module failed to install. Please install the module manually" -Verbose
            Write-Warning "Ensure running Script elevated to install module" -Verbose
            Break
        }
    }
}

function CountAndReportAssignments {
    if ($Count -ne 0) {
        Paragraph "There are $($Count) $($AssignmentType) Assignments. The following $($AssignmentType) Assignments are in place"
    }
    else {
        Paragraph "There are no $($AssignmentType) Assignments"
    }
}

function CountAndReportActions {
    if ($Count -ne "0") {
        Paragraph "There are $($Count) $($ActionType) Actions. The following $($ActionType) Actions have been defined"
    }
    else {
        Paragraph "There are no $($ActionType) Actions defined"
    }
}

Function Convert-Hashtable {
    Param(
        [Parameter()]
        $Hashtable
    )
    if (($Hashtable).Count -gt 1) {
        ForEach ($item in $Hashtable.GetEnumerator()) {
            $Name = $null
            ForEach ($Record in $DescriptionTable.GetEnumerator()) {
                if ($Item.Name -eq $Record.Name) { $Name = $Record.Value } 
                $PSObject = [PSCustomObject] @{
                    Name        = $Item.Key
                    Description = $Name
                    State       = if ($Item.Value -eq 0) { "Disabled" } elseif ($Item.Value -eq 1) { "Enabled" } elseif ($Item.Value -ne 0 -and $Item.Value -ne 1) { $Item.Value }
                } 
            }
            Write-Output -InputObject $PSObject
        }
    }
    elseif (($Hashtable).Count -eq 1) {
        ForEach ($item in $Hashtable) {
            $Name = $null
            ForEach ($Record in $DescriptionTable.GetEnumerator()) {
                if ($Item.Name -eq $Record.Name) { $Name = $Record.Value } 
                $PSObject = [PSCustomObject] @{
                    Name        = $Item.Key
                    Description = $Name
                    State       = if ($Item.Value -eq 0) { "Disabled" } elseif ($Item.Value -eq 1) { "Enabled" } elseif ($Item.Value -ne 0 -and $Item.Value -ne 1) { $Item.Value }
                } 
            }
            Write-Output -InputObject $PSObject
        }
    
    }
}
function StandardOutput {
    param (
        [Parameter()] $OutputObject,
        [Parameter()] [int] $Col1 = 40,
        [Parameter()] [int] $Col2 = 40,
        [Parameter()] [int] $Col3 = 20
    )
    $OutputObject = Convert-Hashtable -Hashtable $OutputObject
    $OutputObject | Table -Columns Name, Description, State -Headers Setting, Description, Value -ColumnWidths $Col1, $col2, $Col3
    BlankLine
}

function WriteDoc {
    [cmdletbinding()]
    param()

    Document "Citrix WEM Documentation" {
        # ============================================================================
        # Cover Page and ToC
        # ============================================================================
        Section -Name "Citrix WEM Documentation" -Style Heading1 -ExcludeFromTOC  {
            Paragraph (Get-Date -Format d)
            BlankLine
            if ($CompanyName) {
                Paragraph "For $CompanyName"
            }
            PageBreak
        }

        TOC -Name 'Table of Contents'
        PageBreak
        #region Config Sets
        # ============================================================================
        # WEM Config Sets
        # ============================================================================
        Section -Name "WEM Configuration Sets" -Style Heading1 {
            $WEMConfigSets = Get-WEMConfiguration -Connection $Connection -Verbose
            Paragraph "There are ($($WEMConfigSets.Count) Configuration Sets found in the deployment):"
            $WEMConfigSets | Table -Columns Name, Description, State -Headers 'Name', 'Description', 'State'
            BlankLine
            Paragraph "The following documentation outlines the $($WEMSite.Name) Configuration Set"
        }
        PageBreak
        #endregion
        #region Actions
        # ============================================================================
        # WEM Actions
        # ============================================================================
        Section -Name "WEM Actions" -Style Heading1 {
            Section -Name "Actions - Action Groups" -Style Heading2 {
                $WEMActionGroups = Get-WEMActionGroup -Connection $Connection -IdSite $Site -Verbose | Select-Object Name, Description, @{Name = 'Actions'; Expression = { $_.Actions -join '; ' } }, State
                $Count = ($WEMActionGroups | Measure-Object).Count
                $ActionType = "Action Group"
                CountAndReportActions
                if ($Count -ne 0) {
                    $WEMActionGroups | Table -Columns Name, Actions, State
                }
            }
            Section -Name "Actions - Applications" -Style Heading2 {
                $WEMApplications = Get-WEMApplication -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMApplications | Measure-Object).Count
                $ActionType = "Application"
                CountAndReportActions
                if ($Count -ne 0) {
                    $WEMApplications | Table -Columns Name, Description, Type -ColumnWidths 40,40,20
                }
            }
            Section -Name "Actions - Printers" -Style Heading2 {
                $WEMPrinters = Get-WEMPrinter -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMPrinters | Measure-Object).Count
                $ActionType = "Printer"
                CountAndReportActions
                if ($Count -ne 0) {
                    $WEMPrinters | Table -Columns Name, TargetPath, ActionType -ColumnWidths 40,40,20
                }
            }
            Section -Name "Actions - Network Drives" -Style Heading2 {
                $WEMNetworkDrives = Get-WEMNetDrive -Connection $Connection -IdSite $site -Verbose
                $Count = ($WEMNetworkDrives | Measure-Object).Count
                $ActionType = "Network Drive"
                CountAndReportActions
                if ($Count -ne 0) {
                    $WEMNetworkDrives | Table -Columns Name, Description, TargetPath
                }
            }
            Section -Name "Actions - Virtual Drives" -Style Heading2 {
                $WEMVirtualDrives = Get-WEMVirtualDrive -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMVirtualDrives | Measure-Object).Count
                $ActionType = "Virtual Drive"
                CountAndReportActions
                if ($Count -ne 0) {
                    $WEMVirtualDrives | Table -Columns Name, Description, TargetPath
                }
            }
            Section -Name "Actions - Registry Values" -Style Heading2 {
                $WEMRegistryValues = Get-WEMRegistryEntry -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMRegistryValues | Measure-Object).Count
                $ActionType = "Registry Value"
                CountAndReportActions
                if ($Count -ne 0) {            
                    $WEMRegistryValues | Table -Columns Name, Description, ActionType -ColumnWidths 40,40,20
                }
            }
            Section -Name "Actions - Environment Variables" -Style Heading2 {
                $WEMEnvironmentVariables = Get-WEMEnvironmentVariable -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMEnvironmentVariables | Measure-Object).Count
                $ActionType = "Environment Variable"
                CountAndReportActions
                if ($Count -ne 0) {
                    $WEMEnvironmentVariables | Table -Columns Name, VariableName, VariableValue
                }
            }
            Section -Name "Actions - Ports" -Style Heading2 {
                $WEMPorts = Get-WEMPort -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMPorts | Measure-Object).Count
                $ActionType = "Port"
                CountAndReportActions
                if ($Count -ne 0) {
                    $WEMPorts | Table -Columns Name, Description, PortName, TargetPath
                }
            }
            Section -Name "Actions - Ini File Operations" -Style Heading2 {
                $WEMIniFiles = Get-WEMIniFileOperation -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMIniFiles | Measure-Object).Count
                $ActionType = "Ini File Operation"
                CountAndReportActions
                if ($Count -ne 0) {
                    $WEMIniFiles | Table -Columns Name, ActionType, TargetPath, TargetName, TargetValue
                }
            }
            Section -Name "Actions - External Tasks" -Style Heading2 {
                $WEMExternalTasks = Get-WEMExternalTask -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMExternalTasks | Measure-Object).Count
                $ActionType = "External Task"
                CountAndReportActions  
                if ($Count -ne 0) {          
                    $WEMExternalTasks | Table -Columns Name, TargetPath, TargetArguments -ColumnWidths 30,30,40
                }
            }
            Section -Name "Actions - File System Operations" -Style Heading2 {
                $WEMFileSystemObjects = Get-WEMFileSystemOp -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMFileSystemObjects | Measure-Object).Count
                $ActionType = "File System Operation"
                CountAndReportActions            
                if ($Count -ne 0) {
                    $WEMFileSystemObjects | Table -Columns Name, SourcePath, ActionType -ColumnWidths 30,40,30
                }
            }
            Section -Name "Actions - User DSNs" -Style Heading2 {
                $WEMUserDSNs = Get-WEMUserDSN -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMUserDSNs | Measure-Object).Count
                $ActionType = "User DSN"
                CountAndReportActions
                if ($Count -ne 0) { 
                    $WEMUserDSNs | Table -Columns Name, TargetName, ActionType -ColumnWidths 30,40,30
                }
            }
            Section -Name "Actions - File Associations" -Style Heading2 {
                $WEMFileAssocs = Get-WEMFileAssoc -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMFileAssocs | Measure-Object).Count
                $ActionType = "File Association"
                CountAndReportActions
                if ($Count -ne 0) { 
                    $WEMFileAssocs | Table -Columns Name, FileExtension, ProgramId, TargetPath
                }
            }
        }
        PageBreak
        #endregion
        #region Filters
        # ============================================================================
        # WEM Filters
        # ============================================================================
        Section -Name "WEM Filters" -Style Heading1 {
            Section -Name "WEM Conditions" -Style Heading2 {
                $WEMConditions = Get-WEMCondition -Connection $Connection -IdSite $Site -Verbose
                Paragraph "The following Conditions have been defined"
                $WEMConditions | Table  -List -Columns Name, Description, State, Type, TestValue, TestResult
            }
            Section -Name "WEM Rules" -Style Heading2 {
                #$WEMRules = Get-WEMRule -Connection $Connection -IdSite $Site -Verbose
                $WEMRules = Get-WEMRule -Connection $Connection -IdSite $Site -Verbose | Select-Object Name, Description, @{Name = 'Conditions'; Expression = { $_.Conditions -join '; ' } }, State
                Paragraph "The following Rules have been defined"
                $WEMRules | Table -Columns Name, Conditions
            }
        }
        PageBreak
        #endregion
        #region Assignments
        # ============================================================================
        # WEM Assignments
        # ============================================================================
        Section -Name "WEM Action Assignments" -Style Heading1 {
            Section -Name "Assignments - Action Groups" -Style Heading2 {
                $WEMActionGroupAssignments = Get-WEMActionGroupAssignment -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMActionGroupAssignments | Measure-Object).Count
                $AssignmentType = "Action Group"
                CountAndReportAssignments
                if ($Count -ne 0) {
                    $WEMActionGroupAssignments | Table -Columns AssignedObject, ADObject, Rule
                }
            }
            Section -Name "Assignments - Applications" -Style Heading2 {
                $WEMApplicationAssignments = Get-WEMAppAssignment -Connection $Connection -IdSite $Site -Verbose | Select-Object AssignedObject, ADObject, Rule, @{Name = 'AssignmentProperties'; Expression = { $_.AssignmentProperties -join '; ' } }
                $Count = ($WEMApplicationAssignments | Measure-Object).Count
                $AssignmentType = "Applications"
                CountAndReportAssignments
                if ($Count -ne 0) {
                    $WEMApplicationAssignments | Table -Columns AssignedObject, ADObject, Rule, AssignmentProperties
                }
            }
            Section -Name "Assignments - Printers" -Style Heading2 {
                $WEMPrinterAssignments = Get-WEMPrinterAssignment -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMPrinterAssignments | Measure-Object).Count
                $AssignmentType = "Printer"
                CountAndReportAssignments
                if ($Count -ne 0) {
                    $WEMPrinterAssignments | Table -Columns AssignedObject, ADObject, Rule
                }
            }
            Section -Name "Assignments - Network Drives" -Style Heading2 {
                $WEMNetworkDriveAssignments = Get-WEMNetDriveAssignment -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMNetworkDriveAssignments | Measure-Object).Count
                $AssignmentType = "Network Drive"
                CountAndReportAssignments
                if ($Count -ne 0) {
                    $WEMNetworkDriveAssignments | Table -Columns AssignedObject, ADObject, Rule, AssignmentProperties
                }
            }
            Section -Name "Assignments - Virtual Drives" -Style Heading2 {
                $WEMVirtualDriveAssignments = Get-WEMVirtualDriveAssignment -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMVirtualDriveAssignments | Measure-Object).Count
                $AssignmentType = "Virtual Drive"
                CountAndReportAssignments
                if ($Count -ne 0) {
                    $WEMVirtualDriveAssignments | Table -Columns AssignedObject, ADObject, Rule, AssignmentProperties
                }
            }
            Section -Name "Assignments - Registry Values" -Style Heading2 {
                $WEMRegistryValueAssignments = Get-WEMRegistryEntryAssignment -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMRegistryValueAssignments | Measure-Object).Count
                $AssignmentType = "Registry Value"
                CountAndReportAssignments
                if ($Count -ne 0) {
                    $WEMRegistryValueAssignments | Table -Columns AssignedObject, ADObject, Rule
                }
            }
            Section -Name "Assignments - Environment Variables" -Style Heading2 {
                $WEMEnvironmentVariableAssignments = Get-WEMEnvironmentVariableAssignment -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMEnvironmentVariableAssignments | Measure-Object).Count
                $AssignmentType = "Environment Variable"
                CountAndReportAssignments
                if ($Count -ne 0) {
                    $WEMEnvironmentVariableAssignments | Table -Columns AssignedObject, ADObject, Rule
                }
            }
            Section -Name "Assignments - Ports" -Style Heading2 {
                $WEMPortAssignments = Get-WEMPortAssignment -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMPortAssignments | Measure-Object).Count
                $AssignmentType = "Port"
                CountAndReportAssignments
                if ($Count -ne 0) {
                    $WEMPortAssignments | Table -Columns AssignedObject, ADObject, Rule
                }
            }
            Section -Name "Assignments - Ini File Operations" -Style Heading2 {
                $WEMIniFileAssignments = Get-WEMIniFileOperationAssignment -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMIniFileAssignments | Measure-Object).Count
                $AssignmentType = "Ini File Operation"
                CountAndReportAssignments
                if ($Count -ne 0) {
                    $WEMIniFileAssignments | Table -Columns AssignedObject, ADObject, Rule
                }
            }
            Section -Name "Assignments - External Tasks" -Style Heading2 {
                $WEMExternalTaskAssignments = Get-WEMExternalTaskAssignment -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMExternalTaskAssignments | Measure-Object).Count
                $AssignmentType = "External Task"
                CountAndReportAssignments
                if ($Count -ne 0) {
                    $WEMExternalTaskAssignments | table -Columns AssignedObject, ADObject, Rule
                }
            }
            Section -Name "Assignments - File System Operations" -Style Heading2 {
                $WEMFileSystemObjectAssignments = Get-WEMFileSystemOpAssignment -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMFileSystemObjectAssignments | Measure-Object).Count
                $AssignmentType = "File System Operations"
                CountAndReportAssignments
                if ($Count -ne 0) {
                    $WEMFileSystemObjectAssignments | Table -Columns AssignedObject, ADObject, Rule
                }
            }
            Section -Name "Assignments - User DSNs" -Style Heading2 {
                $WEMUserDSNAssignments = Get-WEMUserDSNAssignment -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMUserDSNAssignments | Measure-Object).Count
                $AssignmentType = "User DSN"
                CountAndReportAssignments
                if ($Count -ne 0) {
                    $WEMUserDSNAssignments | Table -Columns AssignedObject, ADObject, Rule
                }
            }
            Section -Name "Assignments - File Associations" -Style Heading2 {
                $WEMFileAssocAssignments = Get-WEMFileAssocAssignment -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMFileAssocAssignments | Measure-Object).Count
                $AssignmentType = "File Association"
                CountAndReportAssignments
                if ($Count -ne 0) {
                    $WEMFileAssocAssignments | Table -Columns AssignedObject, ADObject, Rule
                }
            }
        }
        PageBreak
        #endregion
        #region System Optimization
        # ============================================================================
        # WEM System Optimization
        # ============================================================================
        Section -Name "WEM System Optimization" -Style Heading1 {
            $WEMSystemOptimization = Get-WEMSystemOptimization -Connection $Connection -IdSite $Site -Verbose
            Section -Name "CPU Management" -Style Heading2 {
                # Spikes Protection
                $SpikesProtectionSettingsList = @("EnableCPUSpikesProtection", 
                    "AutoCPUSpikeProtectionSelected", 
                    "SpikesProtectionCPUUsageLimitPercent",
                    "SpikesProtectionCPUUsageLimitSampleTime",
                    "SpikesProtectionIdlePriorityConstraintTime",
                    "SpikesProtectionCPUCoreLimit",
                    "SpikesProtectionLimitCPUCoreNumber",
                    "CPUSpikesProtectionExcludedProcesses",
                    "EnableIntelligentCpuOptimization",
                    "EnableIntelligentIoOptimization",
                    "ExcludeProcessesFromCPUSpikesProtection"
                )
                $WEMSpikesProtection = $WEMSystemOptimization.GetEnumerator() | Where-Object { $_.Key -in $SpikesProtectionSettingsList } | Sort-Object -Property Key
                Paragraph -Style Heading2 "CPU Spikes Protection"
                Paragraph "The following configurations relate to CPU Spikes Protection Settings"
                StandardOutput -OutputObject $WEMSpikesProtection
        
                # CPU Priority
                $CPUPrioritySettingsList = @("EnableProcessesCpuPriority",
                    "ProcessesCpuPriorityList"
                )
                $WEMCPUPriority = $WEMSystemOptimization.GetEnumerator() | Where-Object { $_.Key -in $CPUPrioritySettingsList } | Sort-Object -Property Key
                Paragraph -Style Heading2 "CPU Priority"
                Paragraph "The following configurations relate to CPU Priority Settings"
                StandardOutput -OutputObject $WEMCPUPriority

                # CPU Affinity
                $CPUAffinitySettingsList = @("EnableProcessesAffinity", 
                    "ProcessesAffinityList"
                )
                $WEMCPUAffinity = $WEMSystemOptimization.GetEnumerator() | Where-Object { $_.Key -in $CPUAffinitySettingsList } | Sort-Object -Property Key
                Paragraph -Style Heading2 "CPU Affinity"
                Paragraph "The following configurations relate to CPU Affinity Settings"
                StandardOutput -OutputObject $WEMCPUAffinity

                # CPU Clamping
                $CPUClampingSettingsList = @("EnableProcessesClamping",
                    "ProcessesClampingList"
                )
                $WEMCPUClamping = $WEMSystemOptimization.GetEnumerator() | Where-Object { $_.Key -in $CPUClampingSettingsList } | Sort-Object -Property Key
                Paragraph -Style Heading2 "CPU Clamping"
                Paragraph "The following configurations relate to CPU Clamping Settings"
                StandardOutput -OutputObject $WEMCPUClamping
            }
            Section -Name "Memory Management" -Style Heading2 {
                # Memory Management
                $WEMMemoryManagementSettingsList = @("EnableMemoryWorkingSetOptimization",
                    "ExcludeProcessesFromMemoryWorkingSetOptimization",
                    "MemoryWorkingSetOptimizationExcludedProcesses",
                    "MemoryWorkingSetOptimizationIdleStateLimitPercent",
                    "MemoryWorkingSetOptimizationIdleSampleTime"
                )
                $WEMMemoryManagement = $WEMSystemOptimization.GetEnumerator() | Where-Object { $_.Key -in $WEMMemoryManagementSettingsList } | Sort-Object -Property Key
                Paragraph -Style Heading2 "Memory Management"
                Paragraph "The following configurations relate to Working Set Optimizatoin Settings"
                StandardOutput -OutputObject $WEMMemoryManagement
            }
            Section -Name "IO Management" -Style Heading2 {
                # IO Management
                $WEMIOManagementSettinglesList = @("EnableProcessesIoPriority",
                    "ProcessesIoPriorityList"
                )
                $WEMIOManagement = $WEMSystemOptimization.GetEnumerator() | Where-Object { $_.Key -in $WEMIOManagementSettinglesList } | Sort-Object -Property Key
                Paragraph -Style Heading2 "I/O Management"
                Paragraph "The following configurations relate to I/O Management Settings"
                StandardOutput -OutputObject $WEMIOManagement
            }
            Section -Name "Fast Logoff" -Style Heading2 {
                # Fast LogOff
                $WEMFastLogOffSettingsList = @("EnableFastLogoff",
                    "ExcludeGroupsFromFastLogoff",
                    "FastLogoffExcludedGroups"
                )
                $WEMFastLogOff = $WEMSystemOptimization.GetEnumerator() | Where-Object { $_.Key -in $WEMFastLogOffSettingsList } | Sort-Object -Property Key
                Paragraph -Style Heading2 "Fast LogOff"
                Paragraph "The following configurations relate to FastLogOff Settings"
                StandardOutput -OutputObject $WEMFastLogOff
            }
        }
        PageBreak
        #endregion
        #region Policies and Profiles
        # ============================================================================
        # WEM Policies and Profiles
        # ============================================================================
        Section -Name "WEM Policies and Profiles" -Style Heading1 {
            $WEMEnvironmentalSettings = Get-WEMEnvironmentalSettings -Connection $Connection -IdSite $Site -Verbose
            Section -Name "Environmental Settings" -Style Heading2 {
                Paragraph "The following Environmental Settings are in place"
                # Environmental Settings Management
                Section -Name "Start Menu" -Style Heading2 {
                    Paragraph -Style Heading3 "Environmental Settings Management"
                    $SettingsList = @("processEnvironmentalSettings",
                        "processEnvironmentalSettingsForAdmins"
                    )
                    $Settings = $WEMEnvironmentalSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                    StandardOutput -OutputObject $Settings

                    # Start Menu
                    Paragraph -Style Heading3 "User Interface: Start Menu"
                    $SettingsList = @("HideCommonPrograms",
                        "RemoveRunFromStartMenu",
                        "HideAdministrativeTools",
                        "HideHelp",
                        "HideFind",
                        "HideWindowsUpdate",
                        "LockTaskbar",
                        "HideSystemClock",
                        "HideDevicesandPrinters",
                        "HideTurnOff",
                        "ForceLogoff",
                        "Turnoffnotificationareacleanup",
                        "TurnOffpersonalizedmenus",
                        "ClearRecentprogramslist"
                    )
                    $Settings = $WEMEnvironmentalSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                    StandardOutput -OutputObject $Settings

                    # Appearance
                    Paragraph -Style Heading3 "User Interface: Appearance"
                    $SettingsList = @("SetSpecificThemeFile",
                        "SpecificThemeFileValue",
                        "SetVisualStyleFile",
                        "VisualStyleFileValue",
                        "SetWallpaper",
                        "Wallpaper",
                        "WallpaperStyle",
                        "SetDesktopBackGroundColor",
                        "DesktopBackGroundColor"
                    )
                    $Settings = $WEMEnvironmentalSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                    StandardOutput -OutputObject $Settings

                }
                Section -Name "Desktop" -Style Heading2 {
                    # Desktop
                    Paragraph -Style Heading3 "User Interface: Desktop"
                    $SettingsList = @("NoMyComputerIcon",
                        "NoRecycleBinIcon",
                        "NoMyDocumentsIcon",
                        "BootToDesktopInsteadOfStart",
                        "NoPropertiesMyComputer",
                        "NoPropertiesRecycleBin",
                        "NoPropertiesMyDocuments",
                        "HideNetworkIcon",
                        " HideNetworkConnections",
                        "DisableTaskMgr"
                    )
                    $Settings = $WEMEnvironmentalSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                    StandardOutput -OutputObject $Settings

                    #Edge UI
                    Paragraph -Style Heading3 "User Interface: Edge UI" 
                    $SettingsList = @("DisableTLcorner",
                        "DisableCharmsHint"
                    )
                    $Settings = $WEMEnvironmentalSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                    StandardOutput -OutputObject $Settings
                }
                Section -Name "Windows Explorer" -Style Heading2 {
                    # Explorer
                    Paragraph -Style Heading3 "User Interface: Explorer"
                    $SettingsList = @("DisableRegistryEditing",
                        "DisableSilentRegedit",
                        "DisableCmd",
                        "DisableCmdScripts",
                        "RemoveContextMenuManageItem",
                        "NoNetConnectDisconnect",
                        "HideLibrairiesInExplorer",
                        "HideNetworkInExplorer",
                        "HideControlPanel",
                        "NoNtSecurity",
                        "NoViewContextMenu",
                        "NoTrayContextMenu"
                    )
                    $Settings = $WEMEnvironmentalSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                    StandardOutput -OutputObject $Settings

                    # Drive Restrictions
                    Paragraph -Style Heading3 "Drive Restrictions"
                    $SettingsList = @("HideSpecifiedDrivesFromExplorer",
                        "ExplorerHiddenDrives",
                        "RestrictSpecifiedDrivesFromExplorer",
                        "ExplorerRestrictedDrives"
                    )
                    $Settings = $WEMEnvironmentalSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                    StandardOutput -OutputObject $Settings
                }
                Section -Name "Control Panel" -Style Heading2 {
                    # Control Panel
                    $SettingsList = @("NoProgramsCPL",
                        "RestrictCpl",
                        "RestrictCplList",
                        "DisallowCpl",
                        "DisallowCplList"
                    )
                    $Settings = $WEMEnvironmentalSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                    StandardOutput -OutputObject $Settings

                }
                Section -Name "Known Folders Management" -Style Heading2 {
                    # Known Folders Management
                    $SettingsList = @("DisabledKnownFolders",
                        "DisableSpecifiedKnownFolders"
                    )
                    $Settings = $WEMEnvironmentalSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                    StandardOutput -OutputObject $Settings
        
                }
                Section -Name "SBC / HVD Tuning" -Style Heading2 {
                    # SBC / HVD Tuning
                    $SettingsList = @("DisableDragFullWindows",
                        "DisableCursorBlink",
                        "EnableAutoEndTasks",
                        "WaitToKillAppTimeout",
                        "SetCursorBlinkRate",
                        "CursorBlinkRateValue",
                        "SetMenuShowDelay",
                        "MenuShowDelayValue",
                        "SetInteractiveDelay",
                        "InteractiveDelayValue",
                        "DisableSmoothScroll",
                        "DisableMinAnimate"
                    )
                    $Settings = $WEMEnvironmentalSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                    StandardOutput -OutputObject $Settings    
                }
            }
            $WEMUSVConfiguration = Get-WEMUSVSettings -Connection $Connection -IdSite $Site -Verbose
            Section -Name "USV - Processing Settings" -Style Heading2 {
                Paragraph "The following Microsoft USV Configurations are in place"
                BlankLine
                Paragraph "Global USV Processing Settings"
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -like "processUSVConfiguration" -or $_.Key -Like "processUSVConfigurationForAdmins" } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "USV - Microsoft Profiles" -Style Heading2 {
                $SettingsList = @("DisableSlowLinkDetect",
                    "RDSHomeDriveLetter",
                    "SlowLinkProfileDefault",
                    "RDSHomeDrivePath",
                    "DeleteRoamingCachedProfiles",
                    "SetRDSHomeDrivePath",
                    "RoamingProfilesFoldersExclusions",
                    "SetRoamingProfilesFoldersExclusions",
                    "CompatibleRUPSecurity",
                    "SetRDSRoamingProfilesPath",
                    "SetWindowsRoamingProfilesPath",
                    "RDSRoamingProfilesPath",
                    "WindowsRoamingProfilesPath",
                    "AddAdminGroupToRUP"
                )
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                Paragraph "The following Microsoft Roaming Profile Configurations are in place"
                StandardOutput -OutputObject $Settings -Col1 30 -Col2 30 -Col3 40
            }
            Section -Name "USV - Folder Redirection Configuration" -Style Heading2 {
                # Folder Redirection Configuration
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -like "processFoldersRedirectionConfiguration" -or $_.Key -Like "DeleteLocalRedirectedFolders" }
                Paragraph -Style Heading3 "Folder Redirection - Configuration"
                Paragraph "The following Settings Outline the Folder Redirection Configuration"
                StandardOutput -OutputObject $Settings

                # Redirection Settings
                $SettingsList = @("processDesktopRedirection",
                    "processPersonalRedirection",
                    "processPicturesRedirection",
                    "processMusicRedirection",
                    "processVideoRedirection",
                    "processStartMenuRedirection",
                    "processFavoritesRedirection",
                    "processAppDataRedirection",
                    "processContactsRedirection",
                    "processDownloadsRedirection",
                    "processLinksRedirection",
                    "processSearchesRedirection"
                )
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                Paragraph -Style Heading3 "Folder Redirection - Processing Settings"
                Paragraph "The following Settings Outline the Folder Redirection Processing Settings"
                StandardOutput -OutputObject $Settings

                # Desktop Settings
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -like "DesktopRedirectedPath" }
                Paragraph -Style Heading3 "Folder Redirection - Desktop"
                Paragraph "The following Settings Outline the Desktop Folder Redirection Settings"
                StandardOutput -OutputObject $Settings -Col1 30 -Col2 25 -Col3 45

                # Documents Settings
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -like "PersonalRedirectedPath" }
                Paragraph -Style Heading3 "Folder Redirection - Documents"
                Paragraph "The following Settings Outline the Documents Folder Redirection Settings"
                StandardOutput -OutputObject $Settings -Col1 30 -Col2 25 -Col3 45

                # Pictures Settings
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -like "PicturesRedirectedPath" -or $_.Key -Like "MyPicturesFollowsDocuments" }
                Paragraph -Style Heading3 "Folder Redirection - Pictures"
                Paragraph "The following Settings Outline the Pictures Folder Redirection Settings"
                StandardOutput -OutputObject $Settings -Col1 30 -Col2 25 -Col3 45

                # Music
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -like "MusicRedirectedPath" -or $_.Key -Like "MyMusicFollowsDocuments" } | Sort-Object -Descending
                Paragraph -Style Heading3 "Folder Redirection - Music"
                Paragraph "The following Settings Outline the Music Folder Redirection Settings"
                StandardOutput -OutputObject $Settings -Col1 30 -Col2 25 -Col3 45

                # Videos
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -like "VideoRedirectedPath" -or $_.Key -Like "MyVideoFollowsDocuments" }
                Paragraph -Style Heading3 "Folder Redirection - Videos"
                Paragraph "The following Settings Outline the Videos Folder Redirection Settings"
                StandardOutput -OutputObject $Settings -Col1 30 -Col2 25 -Col3 45

                # Start
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -like "StartMenuRedirectedPath" }
                Paragraph -Style Heading3 "Folder Redirection - Start Menu"
                Paragraph "The following Settings Outline the Start Menu Folder Redirection Settings"
                StandardOutput -OutputObject $Settings -Col1 30 -Col2 25 -Col3 45

                # Favorites
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -like "FavoritesRedirectedPath" }
                Paragraph -Style Heading3 "Folder Redirection - Favorites"
                Paragraph "The following Settings Outline the Favorites Folder Redirection Settings"
                StandardOutput -OutputObject $Settings -Col1 30 -Col2 25 -Col3 45

                # AppData
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -like "AppDataRedirectedPath" }
                Paragraph -Style Heading3 "Folder Redirection - AppData"
                Paragraph "The following Settings Outline the AppData Folder Redirection Settings"
                StandardOutput -OutputObject $Settings -Col1 30 -Col2 25 -Col3 45

                # Contacts
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -like "ContactsRedirectedPath" }
                Paragraph -Style Heading3 "Folder Redirection - Contacts"
                Paragraph "The following Settings Outline the Contacts Folder Redirection Settings"
                StandardOutput -OutputObject $Settings -Col1 30 -Col2 25 -Col3 45

                # Downloads
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -like "DownloadsRedirectedPath" }
                Paragraph -Style Heading3 "Folder Redirection - Downloads"
                Paragraph "The following Settings Outline the Downloads Folder Redirection Settings"
                StandardOutput -OutputObject $Settings -Col1 30 -Col2 25 -Col3 45

                # Links
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -like "LinksRedirectedPath" }
                Paragraph -Style Heading3 "Folder Redirection - Links"
                Paragraph "The following Settings Outline the Links Folder Redirection Settings"
                StandardOutput -OutputObject $Settings -Col1 30 -Col2 25 -Col3 45

                # Searches
                $Settings = $WEMUSVConfiguration.GetEnumerator() | Where-Object { $_.Key -like "SearchesRedirectedPath" }
                Paragraph -Style Heading3 "Folder Redirection - Searches"
                Paragraph "The following Settings Outline the Searches Folder Redirection Settings"
                StandardOutput -OutputObject $Settings -Col1 30 -Col2 25 -Col3 45
            }
            $WEMCitrixUPM = Get-WEMUPMSettings -Connection $Connection -IdSite $Site -Verbose
            Section -Name "Citrix Profile Management" -Style Heading2 {
                Paragraph "The following Citrix Profile Management Configuration is in place"
                # Citrix Profile Management Enabled
                BlankLine
                Paragraph -Style Heading3 "UPM - Profile Management Configuration"
                $SettingsList = @("UPMManagementEnabled")
                $Settings = $WEMCitrixUPM.GetEnumerator() | Where-Object { $_.Key -in $SettingsList }
                StandardOutput -OutputObject $Settings

                # Citrix Profile Management
                Paragraph -Style Heading3 "UPM - Citrix profile Management"
                $SettingsList = @("ServiceActive",
                    "SetProcessedGroups",
                    "ProcessedGroupsList",
                    "SetExcludedGroups",
                    "ExcludedGroupsList",
                    "ProcessAdmins",
                    "SetPathToUserStore",
                    "PathToUserStore",
                    "MigrateUserStore",
                    "MigrateUserStorePath",
                    "PSMidSessionWriteBack",
                    "PSMidSessionWriteBackReg",
                    "OfflineSupport"
                )
                $Settings = $WEMCitrixUPM.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings -Col1 35 -Col2 35 -Col3 30

                # Profile Handling
                Paragraph -Style Heading3 "UPM - Profile Handling"
                $SettingsList = @("DeleteCachedProfilesOnLogoff",
                    "SetProfileDeleteDelay",
                    "ProfileDeleteDelay",
                    "SetMigrateWindowsProfilesToUserStore",
                    "MigrateWindowsProfilesToUserStore",
                    "AutomaticMigrationEnabled",
                    "SetLocalProfileConflictHandling",
                    "LocalProfileConflictHandling",
                    "SetTemplateProfilePath",
                    "TemplateProfilePath",
                    "TemplateProfileOverridesLocalProfile",
                    "TemplateProfileOverridesRoamingProfile",
                    "TemplateProfileIsMandatory"
                )
                $Settings = $WEMCitrixUPM.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings -Col1 35 -Col2 35 -Col3 30

                # Advanced Settings
                Paragraph -Style Heading3 "UPM - Advanced Settings"
                $SettingsList = @("SetLoadRetries",
                    "LoadRetries",
                    "XenAppOptimizationEnabled",
                    "XenAppOptimizationPath",
                    "SetUSNDBPath",
                    "USNDBPath",
                    "ProcessCookieFiles",
                    "DeleteRedirectedFolders",
                    "DisableDynamicConfig",
                    "LogoffRatherThanTempProfile",
                    "CEIPEnabled",
                    "OutlookSearchRoamingEnabled",
                    "SearchBackupRestoreEnabled"
                )
                $Settings = $WEMCitrixUPM.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings -Col1 35 -Col2 35 -Col3 30

                # Log Settings
                Paragraph -Style Heading3 "UPM - Log Settings"
                $SettingsList = @("LoggingEnabled",
                    "SetLogLevels",
                    "LogLevels",
                    "SetMaxLogSize",
                    "MaxLogSize",
                    "SetPathToLogFile",
                    "PathToLogFile"
                )
                $Settings = $WEMCitrixUPM.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings -Col1 35 -Col2 35 -Col3 30

                # Registy
                Paragraph -Style Heading3 "UPM - Registry"
                $SettingsList = @("LastKnownGoodRegistry",
                    "EnableDefaultExclusionListRegistry",
                    "ExclusionDefaultRegistry01",
                    "ExclusionDefaultRegistry02",
                    "ExclusionDefaultRegistry03",
                    "SetExclusionListRegistry",
                    "ExclusionListRegistry",
                    "SetInclusionListRegistry",
                    "InclusionListRegistry"
                )
                $Settings = $WEMCitrixUPM.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings -Col1 35 -Col2 35 -Col3 30

                # File System
                Paragraph -Style Heading3 "UPM - File System"
                $SettingsList = @("EnableLogonExclusionCheck",
                    "LogonExclusionCheck",
                    "EnableDefaultExclusionListDirectories",
                    "ExclusionDefaultDir01",
                    "ExclusionDefaultDir02",
                    "ExclusionDefaultDir03",
                    "ExclusionDefaultDir04",
                    "ExclusionDefaultDir05",
                    "ExclusionDefaultDir06",
                    "ExclusionDefaultDir07",
                    "ExclusionDefaultDir09",
                    "ExclusionDefaultDir08",
                    "ExclusionDefaultDir10",
                    "ExclusionDefaultDir11",
                    "ExclusionDefaultDir12",
                    "ExclusionDefaultDir13",
                    "ExclusionDefaultDir14",
                    "ExclusionDefaultDir15",
                    "ExclusionDefaultDir16",
                    "ExclusionDefaultDir17",
                    "ExclusionDefaultDir18",
                    "ExclusionDefaultDir19",
                    "ExclusionDefaultDir20",
                    "ExclusionDefaultDir21",
                    "ExclusionDefaultDir22",
                    "ExclusionDefaultDir23",
                    "ExclusionDefaultDir24",
                    "ExclusionDefaultDir25",
                    "ExclusionDefaultDir26",
                    "ExclusionDefaultDir27",
                    "ExclusionDefaultDir28",
                    "ExclusionDefaultDir29",
                    "ExclusionDefaultDir30",
                    "SetSyncExclusionListFiles",
                    "SyncExclusionListFiles",
                    "SetSyncExclusionListDir",
                    "SyncExclusionListDir"
                )
                $Settings = $WEMCitrixUPM.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings -Col1 30 -Col2 30 -Col3 40

                # Synchronization
                Paragraph -Style Heading3 "UPM - Synchronization"
                $SettingsList = @("SetSyncDirList",
                    "SyncDirList",
                    "SetSyncFileList",
                    "SyncFileList",
                    "SetMirrorFoldersList",
                    "MirrorFoldersList",
                    "SetProfileContainerList",
                    "ProfileContainerList",
                    "SetLargeFileHandlingList",
                    "LargeFileHandlingList"
                )
                $Settings = $WEMCitrixUPM.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings -Col1 30 -Col2 30 -Col3 40

                # Streamed User profiles
                Paragraph -Style Heading3 "UPM - Streamed User profiles"
                $SettingsList = @("PSEnabled",
                    "PSAlwaysCache",
                    "PSAlwaysCacheSize",
                    "SetPSPendingLockTimeout",
                    "PSPendingLockTimeout",
                    "SetPSUserGroupsList",
                    "PSUserGroupsList",
                    "EnableStreamingExclusionList",
                    "StreamingExclusionList"
                )
                $Settings = $WEMCitrixUPM.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings -Col1 35 -Col2 35 -Col3 30

                # Streamed User profiles
                Paragraph -Style Heading3 "UPM - Cross Platform Settings"
                $SettingsList = @("CPEnabled",
                    "SetCPUserGroupList",
                    "CPUserGroupList",
                    "SetCPSchemaPath",
                    "CPSchemaPath",
                    "SetCPPath",
                    "CPPath",
                    "CPMigrationFromBaseProfileToCPStore"
                )
                $Settings = $WEMCitrixUPM.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings -Col1 35 -Col2 35 -Col3 30
            }
        }
        PageBreak
        #endregion
        #region Security
        # ============================================================================
        # WEM Security
        # ============================================================================
        Section -Name "WEM Security" -Style Heading1 {
            $WEMSystemOptimization = Get-WEMSystemOptimization -Connection $Connection -IdSite $Site -Verbose
        
            Section -Name "Process Management" -Style Heading2 {
                #Security Settings
                Paragraph "The following configurations relate to Process Management Settings"
                $SettingsList = @("EnableProcessesManagement")
                $Settings = $WEMSystemOptimization.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # Process Blacklist
                Paragraph -Style Heading2 "Process Blacklist"
                Paragraph "The following configurations relate to Process Blacklist Settings"
                $SettingsList = @("EnableProcessesBlackListing",
                    "ProcessesManagementBlackListedProcesses",
                    "ProcessesManagementBlackListExcludeLocalAdministrators",
                    "ProcessesManagementBlackListExcludeSpecifiedGroups",
                    "ProcessesManagementBlackListExcludedSpecifiedGroupsList"
                )
                $Settings = $WEMSystemOptimization.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                #Process Whielist
                Paragraph -Style Heading2 "Process Whitelist"
                Paragraph "The following configurations relate to Process Whitelist Settings"
                $SettingsList = @("EnableProcessesWhiteListing",
                    "ProcessesManagementWhiteListedProcesses",
                    "ProcessesManagementWhiteListExcludeLocalAdministrators",
                    "ProcessesManagementWhiteListExcludeSpecifiedGroups",
                    "ProcessesManagementWhiteListExcludedSpecifiedGroupsList"
                )
                $Settings = $WEMSystemOptimization.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "Application Security" -Style Heading2 {   
                #AppLocker Basics
                Paragraph "The following configurations relate to AppLocker Settings"
                $SettingsList = @("AppLockerControllerManagement",
                    "AppLockerControllerReplaceModeOn"
                )
                $Settings = $WEMSystemOptimization.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                #AppLocker Settings
                $AppLockerProcessSettings = Get-WEMAppLockerSettings -Connection $Connection -IdSite $Site -Verbose
                Paragraph "AppLocker Processing Settings"
                $SettingsList = @("CollectionExeEnforcementState",
                    "EnableDLLRuleCollection",
                    "CollectionDllEnforcementState",
                    "CollectionMsiEnforcementState",
                    "CollectionScriptEnforcementState",
                    "EnableProcessesAppLocker",
                    "CollectionAppxEnforcementState"
                )
                $Settings = $AppLockerProcessSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
        }
        PageBreak
        #endregion
        #region Active Directory Objects
        # ============================================================================
        # WEM Active Directory Objects
        # ============================================================================
        Section -Name "WEM Active Directory Objects" -Style Heading1 {
            Section -Name "Computer Objects Assigned" -Style Heading2 {
                $WEMComputers = Get-WEMADAgentObject -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMComputers | Measure-Object).Count
                Paragraph "The following Computer Objects have been Assigned to this Configuration Set"
                if ($Count -ne 0) {
                    $WEMComputers | Table -Columns Name, Type, Priority, State
                }
                BlankLine
            }
            Section -Name "User and Group Objects Defined" -Style Heading2 {
                $WEMUsers = Get-WEMADUserObject -Connection $Connection -IdSite $Site -Verbose
                $Count = ($WEMUsers | Measure-Object).Count
                Paragraph "The following user and Group Objects have been added to this Configuration Set"
                if ($Count -ne 0) {
                    $WEMUsers | Table -Columns Name, Type, Description, Priority
                }
            }
        }
        PageBreak
        #endregion
        #region Transformer Settings
        # ============================================================================
        # WEM Transformer Settings
        # ============================================================================
        Section -Name "WEM Transformer Settings" -Style Heading1 {
            $WEMTransformerSettings = Get-WEMTransformerSettings -Connection $Connection -IdSite $site -Verbose
            Section -Name "General - General Settings" -Style Heading2 {
                # General Settings
                $SettingsList = @("IsKioskEnabled",
                    "GeneralStartUrl"
                )
                Paragraph -Style Heading3 "General Settings"
                $Settings = $WEMTransformerSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            
                # Appearance
                $SettingsList = @("GeneralTitle",
                    "GeneralWindowMode",
                    "GeneralClockEnabled",
                    "GeneralClockUses12Hours",
                    "GeneralEnableLanguageSelect",
                    "GeneralEnableAppPanel",
                    "GeneralAutoHideAppPanel",
                    "GeneralShowNavigationButtons"
                )
                Paragraph -Style Heading3 "Appearance"
                $Settings = $WEMTransformerSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # Change Unlock password
                $SettingsList = @("GeneralUnlockPassword")
                Paragraph -Style Heading3 "Change Unlock Password"
                $Settings = $WEMTransformerSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "General - Site Settings" -Style Heading2 {
                # Site Settings
                $SettingsList = @("SitesIsListEnabled",
                    "SitesNamesAndLinks"
                )
                Paragraph -Style Heading3 "Site Settings"
                $Settings = $WEMTransformerSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings           
            }
            Section -Name "General - Tool Settings" -Style Heading2 {
                #Tool Settings
                $SettingsList = @("ToolsEnabled",
                    "ToolsAppsList"
                )
                Paragraph -Style Heading3 "Tool Settings"
                $Settings = $WEMTransformerSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "Advanced - Process Launcher" -Style Heading2 {
                #Process Launcher
                $SettingsList = @("ProcessLauncherEnabled",
                    "ProcessLauncherApplication",
                    "ProcessLauncherArgs",
                    "ProcessLauncherClearLastUsernameVMWare",
                    "ProcessLauncherEnableVMWareViewMode",
                    "ProcessLauncherEnableMicrosoftRdsMode",
                    "ProcessLauncherEnableCitrixMode"
                )
                Paragraph -Style Heading3 "Process Launcher Settings"
                $Settings = $WEMTransformerSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "Advanced - Advanced & Administration Settings" -Style Heading2 {
                # Advanced Settings
                $SettingsList = @("AdvancedFixBrowserRendering",
                    "AdvancedLogOffScreenRedirection",
                    "AdvancedSuppressScriptErrors",
                    "AdvancedFixSslSites",
                    "AdvancedHideKioskWhileCitrixSession",
                    "AdvancedAlwaysShowAdminMenu",
                    "AdvancedHideTaskbar",
                    "AdvancedLockAltTab",
                    "AdvancedFixZOrder",
                    "SetCitrixReceiverFSOMode",
                    "AdvancedShowWifiSettings",
                    "AdvancedLockCtrlAltDel"
                )
                Paragraph -Style Heading3 "Advanced Settings"
                $Settings = $WEMTransformerSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # Administration Settings
                $SettingsList = @("AdministrationHideDisplaySettings",
                    "AdministrationHideKeyboardSettings",
                    "AdministrationHideMouseSettings",
                    "AdministrationHideVolumeSettings",
                    "AdministrationHideClientDetails",
                    "AdministrationDisableProgressBar",
                    "AdministrationHideWindowsVersion",
                    "AdministrationHideHomeButton",
                    "AdministrationHidePrinterSettings",
                    "AdministrationPreLaunchReceiver",
                    "AdministrationDisableUnlock",
                    "AdministrationHideLogOffOption",
                    "AdministrationHideRestartOption",
                    "AdministrationHideShutdownOption",
                    "AdministrationIgnoreLastLanguage"
                )
                Paragraph -Style Heading3 "Administration Settings"
                $Settings = $WEMTransformerSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "Advanced - Logon/Logoff & Power Settings" -Style Heading2 {
                # Autologon Options
                $SettingsList = @("AutologonEnable",
                    "AutologonUserName",
                    "AutologonPassword",
                    "AutologonDomain",
                    "AutologonRegistryForce",
                    "AutologonRegistryIgnoreShiftOverride"
                )
                Paragraph -Style Heading3 "Autologon Options"
                $Settings = $WEMTransformerSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # Desktop Mode Options
                $SettingsList = @("DesktopModeLogOffWebPortal")
                Paragraph -Style Heading3 "Desktop Mode Options"
                $Settings = $WEMTransformerSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # End of Session Options
                $SettingsList = @("EndSessionOption")
                Paragraph -Style Heading3 "End of Session Options"
                $Settings = $WEMTransformerSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # Power Options
                $SettingsList = @("PowerShutdownAfterSpecifiedTime",
                    "PowerShutdownAfterIdleTime",
                    "PowerDontCheckBattery"
                )
                Paragraph -Style Heading3 "Power Options"
                $Settings = $WEMTransformerSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
        }
        PageBreak
        #endregion
        #region Advanced Settings
        # ============================================================================
        # WEM Advanced Settings
        # ============================================================================
        Section -Name "WEM Advanced Settings" -Style Heading1 {
            $WEMAgentSettings = Get-WEMAgentSettings -Connection $Connection -IdSite 1 -Verbose
            Section -Name "Configuration - Main Configuration" -Style Heading2 {
                # Agent Actions
                Paragraph "Agent Actions" -Style Heading3
                $SettingsList = @("processVUEMApps",
                    "processVUEMPrinters",
                    "processVUEMNetDrives",
                    "processVUEMVirtualDrives",
                    "processVUEMRegValues",
                    "processVUEMEnvVariables",
                    "processVUEMPorts",
                    "processVUEMIniFilesOps",
                    "processVUEMExtTasks",
                    "processVUEMFileSystemOps",
                    "processVUEMUserDSNs",
                    "processVUEMFileAssocs"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                #Agent Service Actions
                Paragraph "Agent Service Actions" -Style Heading3
                $SettingsList = @("LaunchVUEMAgentOnLogon",
                    "LaunchVUEMAgentOnReconnect",
                    "ProcessVUEMAgentLaunchForAdmins",
                    "VUEMAgentType",
                    "EnableVirtualDesktopCompatibility",
                    "ExecuteOnlyCmdAgentInPublishedApplications"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "Cleanup Actions" -Style Heading2 {
                # Shortcuts deletions
                Paragraph "Shortcuts Deletion at Startup" -Style Heading3
                $SettingsList = @("DeleteDesktopShortcuts",
                    "DeleteStartMenuShortcuts",
                    "DeleteQuickLaunchShortcuts",
                    "DeleteTaskBarPinnedShortcuts",
                    "DeleteStartMenuPinnedShortcuts"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # Network Drive Deletions
                Paragraph "Drive Deletion at Startup" -Style Heading3
                $SettingsList = @("DeleteNetworkDrives")
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # Printer Deletion at Startup
                Paragraph "Printers Deletion at Startup" -Style Heading3
                $SettingsList = @("DeleteNetworkPrinters",
                    "PreserveAutocreatedPrinters",
                    "PreserveSpecificPrinters",
                    "SpecificPreservedPrinters"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "Configuration - Agent Options" -Style Heading2 {
                # Agent Logs
                Paragraph "Agent Logs" -Style Heading3
                $SettingsList = @("EnableAgentLogging",
                    "AgentLogFile",
                    "AgentDebugMode"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # OFfline Mode Settings
                Paragraph "Offline Mode Settings" -Style Heading3
                $SettingsList = @("OfflineModeEnabled",
                    "UseCacheEvenIfOnline"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # Refresh Settings
                Paragraph "Refresh Settings" -Style Heading3
                $SettingsList = @("RefreshEnvironmentSettings",
                    "RefreshSystemSettings",
                    "RefreshOnEnvironmentalSettingChange",
                    "RefreshDesktop",
                    "RefreshAppearance"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # Asynchronous Processing
                Paragraph "Asynchronous Processing" -Style Heading3
                $SettingsList = @("aSyncVUEMPrintersProcessing",
                    "aSyncVUEMNetDrivesProcessing"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # Extra Features
                Paragraph "Extra Features" -Style Heading3
                $SettingsList = @("InitialEnvironmentCleanUp",
                    "InitialDesktopUICleaning",
                    "checkAppShortcutExistence",
                    "appShortcutExpandEnvironmentVariables",
                    "AgentEnableCrossDomainsUserGroupsSearch",
                    "AgentBrokerServiceTimeoutValue",
                    "AgentDirectoryServiceTimeoutValue",
                    "AgentNetworkResourceCheckTimeoutValue",
                    "AgentMaxDegreeOfParallelism"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # Connection State Change Notification
                Paragraph "Connection State Change Notification" -style Heading3
                $SettingsList = @("ConnectionStateChangeNotificationEnabled")
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "Configuration - Advanced Options" -Style Heading2 {
                # Agent Actions Enforce Execution
                Paragraph "Agents Actions Enforce Execution" -Style Heading3
                $SettingsList = @("enforceProcessVUEMApps",
                    "enforceProcessVUEMPrinters",
                    "enforceProcessVUEMNetDrives",
                    "enforceProcessVUEMVirtualDrives",
                    "enforceProcessVUEMEnvVariables",
                    "enforceProcessVUEMPorts"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # Unassigned Actions Revert Processing
                Paragraph "Unassigned Actions Revert Processing" -Style Heading3 
                $SettingsList = @("revertUnassignedVUEMApps",
                    "revertUnassignedVUEMPrinters",
                    "revertUnassignedVUEMNetDrives",
                    "revertUnassignedVUEMVirtualDrives",
                    "revertUnassignedVUEMRegValues",
                    "revertUnassignedVUEMEnvVariables",
                    "revertUnassignedVUEMPorts",
                    "revertUnassignedVUEMIniFilesOps",
                    "revertUnassignedVUEMExtTasks",
                    "revertUnassignedVUEMFileSystemOps",
                    "revertUnassignedVUEMUserDSNs",
                    "revertUnassignedVUEMFileAssocs"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings 

                # Automatic Refresh (UI Agent Only)
                Paragraph "Automatic Refresh (UI Agent Only)" -Style Heading3
                $SettingsList = @("EnableUIAgentAutomaticRefresh",
                    "UIAgentAutomaticRefreshDelay"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "Configuration - Reconnection Actions" -Style Heading2 {
                # Advanced Settings -> Configuration -> Reconnection Actions
                Paragraph -Name "Actions Processing upon Reconnection" -Style Heading3
                $SettingsList = @("processVUEMAppsOnReconnect",
                    "processVUEMPrintersOnReconnect",
                    "processVUEMNetDrivesOnReconnect",
                    "processVUEMVirtualDrivesOnReconnect",
                    "processVUEMRegValuesOnReconnect",
                    "processVUEMEnvVariablesOnReconnect",
                    "processVUEMPortsOnReconnect",
                    "processVUEMIniFilesOpsOnReconnect",
                    "processVUEMExtTasksOnReconnect",
                    "processVUEMFileSystemOpsOnReconnect",
                    "processVUEMUserDSNsOnReconnect",
                    "processVUEMFileAssocsOnReconnect"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "Configuration - Advanced Processing" -Style Heading2 {
                # Advanced Settings -> Configuration -> Advanced Processing
                Paragraph "Filters Processing Enforcement" -Style Heading3
                $SettingsList = @("enforceVUEMAppsFiltersProcessing",
                    "enforceVUEMPrintersFiltersProcessing",
                    "enforceVUEMNetDrivesFiltersProcessing",
                    "enforceVUEMVirtualDrivesFiltersProcessing",
                    "enforceVUEMRegValuesFiltersProcessing",
                    "enforceVUEMEnvVariablesFiltersProcessing",
                    "enforceVUEMPortsFiltersProcessing",
                    "enforceVUEMIniFilesOpsFiltersProcessing",
                    "enforceVUEMExtTasksFiltersProcessing",
                    "enforceVUEMFileSystemOpsFiltersProcessing",
                    "enforceVUEMUserDSNsFiltersProcessing",
                    "enforceVUEMFileAssocsFiltersProcessing"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "Configuration - Service Options" -Style Heading2 {
                # Agent Service Advanced Options
                Paragraph "Agent Service Advanced Options" -Style Heading3
                $SettingsList = @("VUEMAgentCacheRefreshDelay",
                    "VUEMAgentSQLSettingsRefreshDelay",
                    "VUEMAgentDesktopsExtraLaunchDelay",
                    "AgentServiceDebugMode",
                    "byPassie4uinitCheck"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings 

                # Agent Launch Exclusions
                Paragraph "Agent Launch Exclusions" -Style Heading3
                $SettingsList = @("AgentLaunchExcludeGroups",
                    "AgentLaunchExcludedGroups",
                    "AgentLaunchIncludeGroups",
                    "AgentLaunchIncludedGroups"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "Configuration - Console Settings" -Style Heading2 {
                $WEMAdvancedParams = Get-WEMParameters -Connection $Connection -IdSite $Site -Verbose

                # Forbidden Drives
                Paragraph "Forbidden Drives" -Style Heading3
                $SettingsList = @("excludedDriveletters",
                    "AllowDriveLetterReuse"
                )
                $Settings = $WEMAdvancedParams.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key -Descending
                StandardOutput -OutputObject $Settings
            }
            Section -Name "Configuration - StoreFront" -Style Heading2 {
                $WEMStoreFrontSettings = Get-WEMStorefrontSetting -Connection $Connection -IdSite $site -Verbose
                Paragraph "StoreFront Settings" -Style Heading3
                if ($null -ne $WEMStoreFrontSettings) {
                    $WEMStoreFrontSettings | Table -Columns StorefrontUrl, Description, State -Headers Setting, Description, Value
                    BlankLine
                }
            }
            Section -Name "Configuration - Agent Switch" -Style Heading2 {            
                # Switch to Service Agent
                Paragraph "Switch to Service Agent" -Style Heading3
                $SettingsList = @("AgentSwitchFeatureToggle",
                    "SwitchtoServiceAgent",
                    "CloudConnectors",
                    "UseGPO"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "UI Agent Personalization - UI Agent Options" -Style Heading2 {            
                # Branding
                Paragraph "Branding" -Style Heading2
                $SettingsList = @("UIAgentSplashScreenBackGround",
                    "UIAgentLoadingCircleColor",
                    "UIAgentLbl1TextColor",
                    "UIAgentSkinName",
                    "HideUIAgentSplashScreen",
                    "HideUIAgentSplashScreenOnReconnect"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # Published Applications Behavior
                Paragraph "Published Applications Behavior" -Style Heading2
                $SettingsList = @("HideUIAgentIconInPublishedApplications",
                    "HideUIAgentSplashScreenInPublishedApplications"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # User Interaction
                Paragraph "User Interaction" -Style Heading2
                $SettingsList = @("AgentExitForAdminsOnly",
                    "AgentAllowUsersToManagePrinters",
                    "AgentAllowUsersToManageApplications",
                    "AgentPreventExitForAdmins",
                    "AgentEnableApplicationsShortcuts",
                    "DisableAdministrativeRefreshFeedback"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "UI Agent Personalization - Helpdesk Options" -Style Heading2 {            
                # Help & Custom Links
                Paragraph "Help & Custom Links" -Style Heading2
                $SettingsList = @("UIAgentHelpLink",
                    "UIAgentCustomLink"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # Screen Capture Options
                Paragraph "Screen Capture Options" -Style Heading2
                $SettingsList = @("AgentAllowScreenCapture",
                    "AgentScreenCaptureEnableSendSupportEmail",
                    "AgentScreenCaptureSupportEmailAddress",
                    "MailSMTPToAddress",
                    "MailCustomSubject",
                    "AgentScreenCaptureSupportEmailTemplate",
                    "MailEnableUseSMTP",
                    "MailSMTPServer",
                    "MailSMTPPort",
                    "MailEnableSMTPSSL",
                    "MailSMTPFromAddress",
                    "MailEnableUseSMTPCredentials",
                    "MailSMTPUser",
                    "MailSMTPPassword"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
            Section -Name "UI Agent Personalization - Power Saving" -Style Heading2 {
                Paragraph "Power Options" -Style Heading2
                # Power Saving
                $SettingsList = @("AgentShutdownAfterEnabled",
                    "AgentShutdownAfter",
                    "AgentShutdownAfterIdleEnabled",
                    "AgentShutdownAfterIdleTime",
                    "AgentSuspendInsteadOfShutdown"
                )
                $Settings = $WEMAgentSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings
            }
        }
        PageBreak
        #endregion
        #region Administration
        # ============================================================================
        # WEM Administration
        # ============================================================================
        Section -Name "WEM Administration" -Style Heading1 {
            $GlobalAdmins = Get-WEMAdministrator -Connection $Connection -Verbose | Where-Object { $_.Permissions -like "Global Admin *" }
            $SiteAdmins = Get-WEMAdministrator -Connection $Connection -IdSite $Site -Verbose

            Paragraph "The following Global Administrators have been defined within the WEM Environment"
            $GlobalAdmins | Table -Columns Name, Type, Permissions, Description, State
            BlankLine

            $Count = ($SiteAdmins | Measure-Object).Count
            if ($Count -ne 0) {
                Paragraph "The following Site Specific Administrators have been defined with the WEM Site"
                $SiteAdmins | Table -Columns Name, Type, Permissions, Description, State
            }
            BlankLine
        }
        PageBreak
        #endregion
        #region Monitoring
        # ============================================================================
        # WEM Monitoring
        # ============================================================================
        Section -Name "WEM Monitoring" -Style Heading1 {
            $WEMMonitoringSettings = Get-WEMSystemMonitoringSettings -Connection $Connection -IdSite $Site -Verbose
            Section -Name "Configuration" -Style Heading2 {
                Paragraph "Advanced Settings" -Style Heading2
                $SettingsList = @("BusinessDayEndHour",
                    "ReportsBootTimeMinimum",
                    "BusinessDayStartHour",
                    "EnableWorkDaysFiltering",
                    "WorkDaysFilter",
                    "ReportsLoginTimeMinimum"
                )
                $Settings = $WEMMonitoringSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings

                # Advanced Monitoring - Database Only
                Paragraph "Advanced Settings - Database Only" -Style Heading2

                $SettingsList = @("EnableApplicationReportsWindows2K3XPCompliance",
                    "EnableGlobalSystemMonitoring",
                    "EnableProcessActivityMonitoring",
                    "EnableUserExperienceMonitoring",
                    "EnableSystemMonitoring",
                    "ExcludedProcessesFromApplicationReports",
                    "ExcludeProcessesFromApplicationReports",
                    "LocalDatabaseRetentionPeriod",
                    "LocalDataUploadFrequency",
                    "EnableStrictPrivacy"
                )
                $Settings = $WEMMonitoringSettings.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
                StandardOutput -OutputObject $Settings -col1 30 -Col2 30 -Col3 40
            }
        }
        PageBreak
        #endregion
        #region WEM Advanced Options
        # ============================================================================
        # WEM Advanced Options
        # ============================================================================
        Section -Name "WEM Advanced Options" -Style Heading1 {
            $WEMAdvancedParams = Get-WEMParameters -Connection $Connection -IdSite $Site -Verbose
            Paragraph "The following Advanced Options exist within the environment, though are not always visbile in the WEM Console"
            $SettingsList = @("ADSearchForestBlacklist",
                "AgentSiteIdCacheOverdueTime",
                "ActionGroupsToggle",
                "VersionInfo",
                "GlobalLicenseServerPort",
                "GlobalLicenseServer",
                "DisplayUPMStatusToggle",
                "ProfileContainerToggle",
                "AgentDomainCacheOverdueTime"
            )
            $Settings = $WEMAdvancedParams.GetEnumerator() | Where-Object { $_.Key -in $SettingsList } | Sort-Object -Property Key
            StandardOutput -OutputObject $Settings
        }
        PageBreak
        #endregion
        #region Appendix
        # ============================================================================
        # Detailed Appendix
        # ============================================================================
        if ($Detailed.IsPresent) {
            Write-Verbose "Detailed output requested" -Verbose
            # ============================================================================
            # Appendix - Actions
            # ============================================================================
            Section -Name "Detailed Appendix - Actions" -Style Heading1 {
                Section -Name "Actions - Applications" -Style Heading2 {
                    $WEMApplications = Get-WEMApplication -Connection $Connection -IdSite $Site -Verbose
                    Paragraph "Detailed Configurations for all WEM Application actions are outlined below"
                    BlankLine
                    foreach ($App in $WEMApplications) {
                        Paragraph -Style Heading3 "$($app.Name)"
                        $App | Table -List -Columns Name, DisplayName, Description, State, Type, StartMenuTarget, TargetPath, Parameters, WorkingDirectory, WindowStyle, HotKey, IconLocation, SelfHealingEnabled, EnforceIconLocation, EnforceIconXValue, EnforceIconYValue, DoNotShowInSelfService, CreateShortcutInUserFavoritesFolder
                        BlankLine
                    }
                }
                Section -Name "Actions - Registry Values" -Style Heading2 {
                    $WEMRegistryValues = Get-WEMRegistryEntry -Connection $Connection -IdSite $Site -Verbose
                    Paragraph "Detailed Configurations for all WEM Registry Value actions are outlined below"
                    BlankLine
                    foreach ($RegValue in $WEMRegistryValues) {
                        Paragraph -Style Heading3 "$($RegValue.Name)"
                        $RegValue | Table -List -Columns Name, Description, State, ActionType, TargetPath, TargetName, TargetType, TargetValue, RunOnce
                        BlankLine
                    }
                }
                Section -Name "Actions - Ini File Ops" -Style Heading2 {
                    $WEMIniFiles = Get-WEMIniFileOperation -Connection $Connection -IdSite $Site -Verbose
                    Paragraph "Detailed Configurations for all Ini File Ops actions are outlined below"
                    BlankLine
                    foreach ($IniValue in $WEMIniFiles) {
                        Paragraph -Style Heading3 "$($IniValue.Name)"
                        $IniValue | Table -List -Columns Name, Description, State, ActionType, TargetPath, TargetName, TargetValue, RunOnce
                    }
                }
                Section -Name "Actions - External Tasks" -Style Heading2 {
                    $WEMExternalTasks = Get-WEMExternalTask -Connection $Connection -IdSite $Site -Verbose
                    Paragraph "Detailed Configurations for all WEM External Task actions are outlined below"
                    BlankLine
                    foreach ($Task in $WEMExternalTasks) {
                        Paragraph -Style Heading3 "$($Task.Name)"
                        $Task | Table -List -Columns Name, Description, State, ActionType, TargetPath, TargetArguments, RunHidden, WaitForFinish, TimeOut, ExecutionOrder, RunOnce, ExecuteOnlyAtLogon
                        BlankLine
                    }
                }
                Section -Name "Actions - File System Operations" -Style Heading2 {
                    $WEMFileSystemObjects = Get-WEMFileSystemOp -Connection $Connection -IdSite $Site -Verbose
                    Paragraph "Detailed Configurations for all WEM File System Operation actions are outlined below"
                    BlankLine
                    foreach ($FSO in $WEMFileSystemObjects) {
                        Paragraph -Style Heading3 "$($FSO.Name)"
                        $FSO | Table -List -Columns Name, Description, State, ActionType, SourcePath, TargetPath, TargetOverwrite, RunOnce, ExecutionOrder
                        BlankLine
                    }
                }
                Section -Name "Actions - User DSNs" -Style Heading2 {
                    $WEMUserDSNs = Get-WEMUserDSN -Connection $Connection -IdSite $Site -Verbose
                    Paragraph "Detailed Configurations for all WEM User DSN actions are outlined below"
                    BlankLine
                    foreach ($DSN in $WEMUserDSNs) {
                        Paragraph -Style Heading3 "$($DSN.Name)"
                        $DSN | Table -List -Columns Name, Description, State, ActionType, TargetName, TargetDriverName, TargetServerName, TargetDatabaseName, UseExternalCredentials, ExternalUsername, ExternalPassword, RunOnce
                        BlankLine
                    }
                }
                Section -Name "Actions - File Associations" -Style Heading2 {
                    $WEMFileAssocs = Get-WEMFileAssoc -Connection $Connection -IdSite $Site -Verbose
                    Paragraph "Detailed Configurations for all WEM File Association actions are outlined below"
                    BlankLine
                    foreach ($FileAssoc in $WEMFileAssocs) {
                        Paragraph -Style Heading3 "$($FileAssoc.Name)"
                        $FileAssoc | Table -List -Columns Name, Description, State, ActionType, FileExtension, ProgramId, Action, IsDefault, TargetPath, TargetCommand, TargetOverwrite, RunOnce
                        BlankLine
                    }
                }
            }
            # ============================================================================
            # Appendix - Filters
             # ============================================================================
            Section -Name "Detailed Appendix - Filters" -Style Heading1 {
                Section -Name "Conditions" -Style Heading2 {
                    $WEMConditions = Get-WEMCondition -Connection $Connection -IdSite $Site -Verbose
                    Paragraph "Detailed Configurations for all WEM Conditions are outlined below"
                    BlankLine
                    foreach ($Condition in $WEMConditions) {
                        Paragraph -Style Heading3 "$($Condition.Name)"
                        $Condition | Table -List -Columns Name, Description, State, Type, TestValue, TestResult
                        BlankLine
                    }
                }
                Section -Name "Rules" -Style Heading2 {
                    $WEMRules = Get-WEMRule -Connection $Connection -IdSite $Site -Verbose | Select-Object Name, Description, State, @{Name = 'Conditions'; Expression = { $_.Conditions -join '; ' } }
                    Paragraph "Detailed Configurations for all WEM Rules are outlined below"
                    BlankLine
                    foreach ($Rule in $WEMRules) {
                        Paragraph -Style Heading3 "$($Rule.Name)"
                        $Rule | Table -List -Columns Name, Description, State, Conditions
                        BlankLine
                    }
                }
            }
        }
        #endregion
    } | Export-Document -Path $OutputLocation -Format Word, HTML -Verbose
}
#endregion

# ============================================================================
# Execute  the Script
# ============================================================================
Import-Module C:\users\JKindon\Documents\GitHub\Citrix.WEMSDK\Citrix.WEMSDK.psd1 -Force # <- This will change once released into PS Gallery

CheckModuleExists -Module "PScribo"
#CheckModuleExists -Module "Citrix.WEMSDK"

if ($DBServer) {
    Write-Verbose "Selected DBName Server: $DBServer" -Verbose
}

if ($DBName) {
    Write-Verbose "Selected DBName Name: $DBName" -Verbose
}

if ($ListAllConfigSets.IsPresent) {
    Write-Verbose "Listing all Configuration Sets" -Verbose
    # Create a Connection Object to list sites
    $Connection = New-WEMDatabaseConnection -Server $DBServer -Database $DBName -Verbose
    Get-WEMConfiguration -Connection $Connection | Format-Table
    break
}

if ($OutputLocation) {
    Write-Verbose "Selected to output report to $OutputLocation" -Verbose
    if (Test-Path $OutputLocation) {
        Write-Verbose "Confirming $OutputLocation Exists, Using as Output Location" -Verbose
    }
    else {
        try {
            Write-Warning "$OutputLocation does not exist, Attempting to create" -Verbose
            New-Item -Path $OutputLocation -ItemType Directory
        }
        catch {
            Write-Warning "Failed to Create $OutputLocation Directory" -Verbose
            Break
        }
    }
}

if ($CompanyName) {
    Write-Verbose "Company Name is: $CompanyName" -Verbose
}

if ($Site) {
    Write-Verbose "Selected Site ID: $Site" -Verbose
    $Connection = New-WEMDatabaseConnection -Server $DBServer -Database $DBName -Verbose
    $WEMSite = Get-WEMConfiguration -Connection $Connection -IdSite $Site -Verbose
}

WriteDoc
