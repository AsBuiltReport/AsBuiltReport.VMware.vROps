function Invoke-AsBuiltReport.VMware.vROps {
    <#
    .SYNOPSIS  
        PowerShell script to document the configuration of vRealize Operations Manager (vROps) in Word/HTML/XML/Text formats
    .DESCRIPTION
        Documents the configuration of vROps in Word/HTML/XML/Text formats using PScribo.
        vROps code provided by andydvmware's PowervROps PowerShell Module.
    .NOTES
        Version:        1.0.0
        Author:         Tim Williams
        Twitter:        @ymmit85
        Github:         ymmit85
        Credits:        andydvmware - PowervROPs PowerShell Module
                        Iain Brighton (@iainbrighton) - PScribo module
    .LINK
        https://github.com/AsBuiltReport/AsBuiltReport.VMware.vROps
        https://github.com/andydvmware/PowervROps
        https://github.com/ymmit85/PowervROps
        https://github.com/iainbrighton/PScribo
    #>

    #region Configuration Settings
    ###############################################################################################
    #                                    CONFIG SETTINGS                                          #
    ###############################################################################################

    param (
        [String[]] $Target,
        [String]$StylePath,
        [String] $Username,
        [String] $Password,
        [PSCredential] $Credential
    )

    # Import JSON Configuration for Options and InfoLevel
    $InfoLevel = $ReportConfig.InfoLevel
    $Options = $ReportConfig.Options

    # If custom style not set, use vROps style
    if (!$StyleName) {
        & "$PSScriptRoot\..\..\AsBuiltReport.VMware.vROps.Style.ps1"
    }
    #endregion Configuration Settings

    #region Script Body
    ###############################################################################################
    #                                       SCRIPT BODY                                           #
    ###############################################################################################
    $TextInfo = (Get-Culture).TextInfo


    # Connect to vROps using supplied credentials
    foreach ($vropshost in $Target) {
        Section -Style Heading1 $vropshost {
            #region Global Settings
            if ($InfoLevel.GlobalSettings -ge 1) {
                $globalSettings = $(getGlobalsettings -resthost $vropshost -credential $Credential).keyValues
                if ($globalSettings) {
                    Section -Style Heading2 -Name 'Global Settings' {
                        $globalSettings = $globalSettings | Select-Object @{l = 'Setting'; e = { $_.key } }, @{l = 'Value'; e = { $_.values } }
                        $globalSettings | Table -Name 'Global Settings'
                    }
                }
            }
            #endregion Global Settings

            #region Service Status
            if ($InfoLevel.ServiceStatus -ge 1) {
                $serviceStatus = $(getServicesInfo -resthost $vropshost -credential $Credential).service
                if ($serviceStatus) {
                    Section -Style Heading2 -Name 'Service Status' {
                        $serviceStatus = $serviceStatus | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'Health'; e = { $_.Health } }, @{l = 'Details'; e = { $_.details } }
                        $serviceStatus | Table -Name 'Service Status'
                    }
                }
            }
            #endregion Service Status

            #region Authentication
            if ($InfoLevel.Authentication -ge 1) {
                $AuthSources = $(getAuthSources -resthost $vropshost -credential $Credential).sources
                if ($AuthSources) {
                    Section -Style Heading2 -Name 'Authentication' {
                        Paragraph -Style Heading3 -Name 'AD Auth Sources' 
                            $AuthSources = $AuthSources | Where-Object { $_.sourcetype.name -like '*ACTIVE_DIRECTORY*' } | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'ID'; e = { $_.ID } }
                            $AuthSources | Table -Name 'AD Auth Sources'
                    }
                }
            }
            #endregion Authentication

            #region Roles
            if ($InfoLevel.Roles -ge 1) {
                $roles = $(getRoles -resthost $vropshost -credential $Credential).userRoles
                if ($roles) {
                    Section -Style Heading2 -Name 'Roles' {
                        Paragraph -Style Heading3 -Name 'System Roles' 
                            $roleSystem = $roles | Where-Object { $_.'system-created' -like 'True' } | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'Description'; e = { $_.description } }, @{l = 'Display Name'; e = { $_.displayName } }   
                            $roleSystem | Table -Name 'System Roles'
                        BlankLine

                        Paragraph -Style Heading3 -Name 'Custom Roles' 
                            $roleSystem = $roles | Where-Object { $_.'system-created' -like 'False' } | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'Description'; e = { $_.description } }, @{l = 'Display Name'; e = { $_.displayName } } 
                            $roleSystem | Table -Name 'Custom Roles'
                        BlankLine
                    }
                }
            }
            #endregion Roles

            #region Groups
            $users = $(getUsers -resthost $vropshost -credential $Credential).users
            $groups = $(getUserGroups -resthost $vropshost -credential $Credential).userGroups


            if ($InfoLevel.Groups -ge 1) {
                Section -Style Heading2 -Name 'Groups' {
                    if ($InfoLevel.Groups -eq 1) {

                    Section -Style Heading3 -Name 'System Groups' {
                        $groupsSystem = $groups | Where-Object { !($_.authSourceId) } | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'Description'; e = { $_.description } }, @{l = 'Roles'; e = { $_.roleNames } }
                        $groupsSystem | Table -Name 'System Groups'
                    }

                    Section -Style Heading2 -Name 'Imported Groups' {
                        $groupsImported = $groups | Where-Object { $_.authSourceId } | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'Description'; e = { $_.description } }, @{l = 'Roles'; e = { $_.roleNames } }
                        $groupsImported | Table -Name 'Imported Groups'
                    }
                    }   elseif ($InfoLevel.Groups -ge 2)    {

                        foreach ($g in $groups | Where-Object { !($_.authSourceId) }) {
                            Paragraph -Style Heading3 -Name 'System Groups' 
                                $groupsSystem = $g | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'Description'; e = { $_.description } }, @{l = 'Roles'; e = { $_.roleNames } }
                                $groupsSystem | Table -Name 'System Groups' -List -ColumnWidths 20, 80
                                BlankLine
                                    Paragraph -Style Heading3 -Name 'Users'
                                    BlankLine
                                        $usersInGroup = @()
                                        if ($g.userIds){
                                            foreach ($c in $g.userIds) {
                                                $usersInGroup += $users | Where-Object { $_.id -eq $c } | Select-Object @{l = 'Username'; e = { $_.username } }, @{l = 'First Name'; e = { $_.firstName } }, @{l = 'Last Name'; e = { $_.lastName } }
                                            }
                                        $usersInGroup | Table -Name 'Users in System Groups'
                                        BlankLine
                                        }

                        }
                            foreach ($g in $groups | Where-Object { $_.authSourceId }) {
                                Paragraph -Style Heading3 -Name 'Imported Groups' 
                                    $groupsSystem = $g | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'Description'; e = { $_.description } }, @{l = 'Roles'; e = { $_.roleNames } }
                                    $groupsSystem | Table -Name  'Imported Groups' -List -ColumnWidths 20, 80
                                    BlankLine
                                        Paragraph -Style Heading3 -Name 'Users'
                                        BlankLine
                                        $usersInGroup = @()
                                        if ($g.userIds){
                                            foreach ($c in $g.userIds) {
                                                $usersInGroup += $users | Where-Object { $_.id -eq $c } | Select-Object @{l = 'Username'; e = { $_.username } }, @{l = 'First Name'; e = { $_.firstName } }, @{l = 'Last Name'; e = { $_.lastName } }
                                            }
                                            $usersInGroup | Table -Name 'Users in Imported Groups'
                                            BlankLine
                                        }

                        }
                    }
                }
            }

            #endregion Groups
            #region Users
            if ($InfoLevel.Users -ge 1) {
                if (!($users)) {
                    $users = $(getUsers -resthost $vropshost -credential $Credential).users
                }
                Section -Style Heading2 -Name 'User Accounts' {
                    Paragraph -Style Heading3 -Name 'System Users' 
                        $systemUsers = $users | Where-Object { $_.distinguishedName -like '' } | Select-Object @{l = 'Username'; e = { $_.username } }, @{l = 'First Name'; e = { $_.firstName } }, @{l = 'Last Name'; e = { $_.lastName } }, @{l = 'Enabled'; e = { $_.enabled } }, @{l = 'Roles'; e = { $($_.rolenames) -join ', ' } }
                        $systemUsers | Table -Name 'System Users' -List -ColumnWidths 20, 80
                        BlankLine
                    
                    if ($InfoLevel.Users -ge 2) {
                        Paragraph -Style Heading3 -Name 'Imported Users' 
                            $importedUsers = $users | Where-Object { $_.'distinguishedName' -notlike '' } | Select-Object @{l = 'Username'; e = { $_.username } }, @{l = 'First Name'; e = { $_.firstName } }, @{l = 'Last Name'; e = { $_.lastName } }, @{l = 'Distinguished Name'; e = { $_.distinguishedName } }, @{l = 'Enabled'; e = { $_.enabled } }
                            $importedUsers | Table -Name 'Imported Users' -List -ColumnWidths 20, 80
                            BlankLine
                    }
                }
            }
            #endregion Users

            #region Remote Collectors
            if ($InfoLevel.RemoteCollectors -ge 1) {
                Section -Style Heading2 -Name 'Cluster Management' {
                    $collectors = $(getCollectors -resthost $vropshost -credential $Credential).collector
                    $localNodes = $collectors | Where-Object { $_.local -like '*True*' }
                    $remoteCollectors = $collectors | Where-Object { $_.local -like '*False*' }

                    if ($localNodes) {
                        Paragraph -Style Heading3 -Name 'Local Nodes' 
                        BlankLine
                            $localNodes = $localNodes | Sort-Object ID | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'ID'; e = { $_.id } }, @{l = 'State'; e = { $_.State } }, @{l = 'Hostname'; e = { $_.hostname } }, @{l = 'Last Heartbeat'; e = { (convertEpoch -epochTime $_.lastHeartbeat) } }
                            $localNodes | Table -Name 'Local Nodes'
                            BlankLine
                    }

                    if ($remoteCollectors) {
                        Paragraph -Style Heading3 -Name 'Remote Nodes'
                        BlankLine
                            $remoteCollectors = $remoteCollectors | Sort-Object ID | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'ID'; e = { $_.id } }, @{l = 'State'; e = { $_.State } }, @{l = 'Hostname'; e = { $_.hostname } }, @{l = 'Last Heartbeat'; e = { (convertEpoch -epochTime $_.lastHeartbeat) } }
                            $remoteCollectors | Table -Name 'Remote Nodes'
                        BlankLine
                    }
                }
            }

            if ($InfoLevel.RemoteCollectors -ge 1) {
                $collectorGroups = $(getCollectorGroups -resthost $vropshost -credential $Credential).collectorGroups
                if ($collectorGroups) {
                    Section -Style Heading2 -Name 'Remote Collector Groups' {
                        foreach ($rcGroup in $collectorGroups) {
                            Paragraph -Style Heading3 -Name $($rcGroup).name
                            BlankLine
                                $Group = $rcGroup | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'Description'; e = { $_.Description } }
                                $Group | Table -Name 'Remote Collector Groups'
                                BlankLine

                                if ($InfoLevel.RemoteCollectors -ge 2) {
                                    foreach ($rcId in $($rcGroup).collectorId) {
                                        $rcNames = $collectors | Where-Object { $_.id -like $rcId }
                                        if ($rcNames) {
                                            Paragraph -Style Heading3 -Name "Members" 
                                            BlankLine
                                                $rcNames = $rcNames | Select-Object @{l = 'Name'; e = { $_.name } },@{l = 'ID'; e = { $_.id } }, @{l = 'Hostname'; e = { $_.hostName } }
                                                $rcNames | Table -Name 'Members'
                                        }
                                    }
                                }
                        }
                    }
                }
            }
            #endregion Remote Collectors

            #region Adapters
            if ($InfoLevel.Adapters -ge 1) {
                $AdapterInstance = $(getAdapterInstances -resthost $vropshost -credential $Credential).adapterInstancesInfoDto 
                $collectors = $(getCollectors -resthost $vropshost -credential $Credential).collector
                if ($AdapterInstance) {
                    Section -Style Heading2 -Name 'Adapters' {
                        foreach ($adapterKind in $AdapterInstance.resourcekey.adapterKindKey | Sort-Object -Unique) {
                            Paragraph -Style Heading3 -Name "$adapterKind"
                            BlankLine
                                $AdapterInstances = $AdapterInstance | Where-Object { $_.resourcekey.adapterKindKey -like $adapterKind }

                                foreach ($adapter in $AdapterInstances) {
                                    $rc = $collectors | Where-Object { $_.id -like $adapter.collectorId }
                                    $adapter = $adapter | Select-Object @{l = 'Name'; e = { $_.resourceKey.name } }, @{l = 'Resource Kind'; e = { $_.resourceKey.resourceKindKey } }, @{l = 'Description'; e = { $_.description } }, @{l = 'Message from Adapter'; e = { $_.messageFromAdapterInstance } }, @{l = 'Collector Node'; e = { $rc.name } }, @{l = 'Last Heartbeat'; e = { (convertEpoch -epochTime $_.lastHeartbeat) } }, @{l = 'Last Collected'; e = { (convertEpoch -epochTime $_.lastCollected) } }, @{l = 'Metrics Collected'; e = { $_.numberOfMetricsCollected } }, @{l = 'Resources Collected'; e = { $_.numberOfResourcesCollected } } 
                                    $adapter | Table -Name 'Adapters' -List -ColumnWidths 20, 80
                                    BlankLine
                                }
                        }
                    }
                }
            }
            #endregion Adapters

            #region Alerts
            if ($InfoLevel.Alerts -ge 1) {
                $alerts = $(getAlertDefinitions -resthost $vropshost -credential $Credential).alertDefinitions | Where-Object { $_.name -like "*$($Options.AlertFilter)*" }
                if ($alerts) {

                    $alertSeverity = $alerts.states.severity | Sort-Object | Get-Unique -AsString
                    Section -Style Heading2 -Name 'Alerts' {
                        foreach ($s in $Options.AlertSeverityFilter) {
                        Section -Name "$s Alerts" -Style Heading3 {
                        foreach ($a in $alerts | where {$_.states.severity -eq $s}) {
                                Paragraph -Name "Alert: $($a.name)" -Style Heading3
                                BlankLine
                                $alertDetail = $a | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'ID'; e = { $_.id } }, @{l = 'Description'; e = { $_.description } }, @{l = 'Severity'; e = { $TextInfo.ToTitleCase(($_.states.severity).ToLower()) } }, @{l = 'Adapter Kind Key'; e = { $TextInfo.ToTitleCase(($_.adapterKindKey).ToLower()) } }, @{l = 'Resource Kind Key'; e = { $_.resourceKindKey } }, @{l = 'Wait Cycles'; e = { $_.waitCycles } }, @{l = 'Cancel Cycles'; e = { $_.cancelCycles } }
                                $alertDetail | Sort-Object -Property name| Table -Name 'Alerts' -List -ColumnWidths 20, 80
                                BlankLine

                                if ($InfoLevel.Alerts -ge 2) {
                                    $symp = $a.states.'base-symptom-set'.symptomDefinitionIds
                                    foreach ($s in $symp) {
                                        $sympHashTable = @()

                                        $symDef = $(getSymptomDefinitions -resthost $vropshost -credential $Credential -symptomdefinitionid $s).symptomDefinitions

                                        if ($symDef.state.condition.type -contains 'CONDITION_MESSAGE_EVENT' ) {
                                                Paragraph -Name Symptom -Style Heading3
                                                BlankLine
                                                $sympHashTable += [PSCustomObject]@{
                                                    'Name' = $symDef.name
                                                    'Id' = $symDef.Id
                                                    'AdapterKindKey' = $TextInfo.ToTitleCase(($symDef.adapterKindKey).ToLower())
                                                    'ResourceKindKey' = $symDef.resourceKindKey
                                                    'Wait Cycles' = $symDef.waitCycles
                                                    'Cancel Cycles' = $symDef.cancelCycles
                                                    'Type' = $TextInfo.ToTitleCase(($symDef.state.condition.type).ToLower())
                                                    'Event Type' = $symDef.state.condition.eventType
                                                    'Message' = $symDef.state.condition.message
                                                    'Operator' = $TextInfo.ToTitleCase(($symDef.state.condition.operator).ToLower())
                                                }
                                                $sympHashTable | Table -Name 'Symptoms' -List -ColumnWidths 20, 80
                                                BlankLine

                                        } elseif ($symDef.state.condition.type -contains 'CONDITION_HT' ) {

                                                Paragraph -Name Symptom -Style Heading3
                                                BlankLine

                                                $sympHashTable += [PSCustomObject]@{
                                                    'Name' = $symDef.name
                                                    'Id' = $symDef.Id
                                                    'AdapterKindKey' = $TextInfo.ToTitleCase(($symDef.adapterKindKey).ToLower())
                                                    'ResourceKindKey' = $symDef.resourceKindKey
                                                    'Wait Cycles' = $symDef.waitCycles
                                                    'Cancel Cycles' = $symDef.cancelCycles
                                                    'Severity' = $TextInfo.ToTitleCase(($symDef.state.severity).ToLower())
                                                    'Type' = $TextInfo.ToTitleCase(($symDef.state.condition.type).ToLower())
                                                    'Key' = $symDef.state.condition.key
                                                    'Operator' = $TextInfo.ToTitleCase(($symDef.state.condition.operator).ToLower())
                                                    'Value' = $TextInfo.ToTitleCase(($symDef.state.condition.value).ToLower())
                                                    'Value Type' = $TextInfo.ToTitleCase(($symDef.state.condition.valueType).ToLower())
                                                    'Instanced' = $symDef.state.condition.instanced
                                                    'Threshold Type' = $TextInfo.ToTitleCase(($symDef.state.condition.thresholdType).ToLower())
                                                }
                                                $sympHashTable | Table -Name 'Symptoms' -List -ColumnWidths 20, 80
                                                BlankLine

                                        } elseif ($symDef.state.condition.type -contains 'CONDITION_PROPERTY_STRING' ) {

                                                Paragraph -Name Symptom -Style Heading3
                                                BlankLine

                                                $sympHashTable += [PSCustomObject]@{
                                                    'Name' = $symDef.name
                                                    'Id' = $symDef.Id
                                                    'AdapterKindKey' = $TextInfo.ToTitleCase(($symDef.adapterKindKey).ToLower())
                                                    'ResourceKindKey' = $symDef.resourceKindKey
                                                    'Wait Cycles' = $symDef.waitCycles
                                                    'Cancel Cycles' = $symDef.cancelCycles
                                                    'Severity' = $TextInfo.ToTitleCase(($symDef.state.severity).ToLower())
                                                    'Type' = $TextInfo.ToTitleCase(($symDef.state.condition.type).ToLower())
                                                    'String Value' = $symDef.state.condition.stringValue
                                                    'Key' = $symDef.state.condition.key
                                                    'Operator' = $TextInfo.ToTitleCase(($symDef.state.condition.operator).ToLower())
                                                    'Threshold Type' = $TextInfo.ToTitleCase(($symDef.state.condition.thresholdType).ToLower())
                                                }
                                                $sympHashTable | Table -Name 'Symptoms' -List -ColumnWidths 20, 80
                                                BlankLine

                                        } elseif ($symDef.state.condition.type -contains 'CONDITION_FAULT' ) {

                                                Paragraph -Name Symptom -Style Heading3
                                                BlankLine
                                                $sympHashTable += [PSCustomObject]@{
                                                    'Name' = $symDef.name
                                                    'Id' = $symDef.Id
                                                    'AdapterKindKey' = $TextInfo.ToTitleCase(($symDef.adapterKindKey).ToLower())
                                                    'ResourceKindKey' = $symDef.resourceKindKey
                                                    'Wait Cycles' = $symDef.waitCycles
                                                    'Cancel Cycles' = $symDef.cancelCycles
                                                    'Severity' = $TextInfo.ToTitleCase(($symDef.state.severity).ToLower())
                                                    'Type' = $TextInfo.ToTitleCase(($symDef.state.condition.type).ToLower())
                                                    'Fault Key' = $symDef.state.condition.faultKey
                                                }
                                                $sympHashTable | Table -Name 'Symptoms' -List -ColumnWidths 20, 80
                                                BlankLine

                                        } elseif ($symDef.state.condition.type -contains 'CONDITION_PROPERTY_NUMERIC' ) {

                                                Paragraph -Name Symptom -Style Heading3
                                                BlankLine

                                                $sympHashTable += [PSCustomObject]@{
                                                    'Name' = $symDef.name
                                                    'Id' = $symDef.Id
                                                    'AdapterKindKey' = $TextInfo.ToTitleCase(($symDef.adapterKindKey).ToLower())
                                                    'ResourceKindKey' = $symDef.resourceKindKey
                                                    'Wait Cycles' = $symDef.waitCycles
                                                    'Cancel Cycles' = $symDef.cancelCycles
                                                    'Severity' = $TextInfo.ToTitleCase(($symDef.state.severity).ToLower())
                                                    'Type' = $TextInfo.ToTitleCase(($symDef.state.condition.type).ToLower())
                                                    'Value' = $symDef.state.condition.value
                                                    'Operator' = $symDef.state.condition.operator
                                                    'Key' = $symDef.state.condition.key
                                                    'Threshold Type' = $symDef.state.condition.thresholdType
                                                }
                                                $sympHashTable | Table -Name 'Symptoms' -List -ColumnWidths 20, 80
                                                BlankLine
                                        }
                                    }
                                }
                        }
                    }
                    }
                    }
                }
            }
            #endregion Alerts

            #region Super Metrics
            if ($InfoLevel.SuperMetrics -ge 1) {
                $superMetrics = $(getSuperMetrics -resthost $vropshost -credential $Credential).supermetrics
                if ($superMetrics) {
                    Section -Style Heading2 -Name 'Super Metrics' {
                        $superMetrics = $superMetrics | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'ID'; e = { $_.ID } }, @{l = 'Formula'; e = { $_.formula } }
                        $superMetrics | Table -Name "Super Metrics" -List -ColumnWidths 20, 80
                    }
                }
            }
            #endregion Super Metrics

            #region Custom Groups
            if ($InfoLevel.CustomGroups -ge 1) {
                if ($customGroups) {
                    $customGroups = $(getCustomGroups -resthost $vropshost -credential $Credential).values
                    if ($customGroups) {
                        Section -Style Heading2 -Name 'Custom Groups' {
                            $customGroups = $customGroups | Select-Object @{l = 'Name'; e = { $_.resourceKey.name } }, @{l = 'Adapter Kind'; e = { $_.resourceKey.adapterKindKey } }, @{l = 'Resource Kind'; e = { $_.resourceKey.resourceKindKey } }
                            $customGroups | Table -Name "Custom Groups"
                        }
                    }
                }
            }
            #endregion Custom Groups

            #region Reports
            if ($InfoLevel.Reports -ge 1) {
                $reports = $(getReportDefinitions -resthost $vropshost -credential $Credential).reportDefinitions
                if ($reports) {
                    Section -Style Heading2 -Name 'Reports' {
                        $reports = $reports | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'Description'; e = { $_.description } }, @{l = 'Owner'; e = { $_.owner } }, @{l = 'Subject'; e = { $_.subject -join ", " } }, @{l = 'Active'; e = { $_.active } }
                        $reports | Table -Name "Reports" -List -ColumnWidths 20, 80
                    }
                }
            }
            #endregion Reports
        }
    }
    #endregion Script Body
}