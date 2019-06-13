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
                        Section -Style Heading2 -Name 'AD Auth Sources' {
                            $AuthSources = $AuthSources | Where-Object { $_.sourcetype.name -like '*ACTIVE_DIRECTORY*' } | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'ID'; e = { $_.ID } }
                            $AuthSources | Table -Name 'AD Auth Sources'
                        }
                    }
                }
            }
            #endregion Authentication

            #region Roles
            if ($InfoLevel.Roles -ge 1) {
                $roles = $(getRoles -resthost $vropshost -credential $Credential).userRoles
                if ($roles) {
                    Section -Style Heading2 -Name 'Roles' {
                        Section -Style Heading2 -Name 'System Roles' {
                            $roleSystem = $roles | Where-Object { $_.'system-created' -like 'True' } | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'Description'; e = { $_.description } }, @{l = 'Display Name'; e = { $_.displayName } }   
                            $roleSystem | Table -Name 'System Roles' -List -ColumnWidths 25, 75
                        }

                        Section -Style Heading2 -Name 'Custom Roles' {
                            $roleSystem = $roles | Where-Object { $_.'system-created' -like 'False' } | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'Description'; e = { $_.description } }, @{l = 'Display Name'; e = { $_.displayName } } 
                            $roleSystem | Table -Name 'Custom Roles' -List -ColumnWidths 25, 75
                        }
                    }
                }
            }
            #endregion Roles

            #region Groups
            if ($InfoLevel.Groups -ge 1) {
                $users = $(getUsers -resthost $vropshost -credential $Credential).users
                $groups = $(getUserGroups -resthost $vropshost -credential $Credential).userGroups

                Section -Style Heading2 -Name 'Groups' {
                    if ($groups) {
                        foreach ($g in ($groups | Where-Object { !($_.authSourceId) })) {
                            Section -Style Heading3 -Name 'System Groups' {
                                $groupsSystem = $g | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'Description'; e = { $_.description } }
                                $groupsSystem | Table -Name 'System Groups' -List -ColumnWidths 25, 75
                            }
                            if ($InfoLevel.Groups -ge 2) {
                                Section -Style Heading3 -Name 'Users' {
                                    $usersInGroup = @()
                                    foreach ($c in $g.userIds) {
                                        $usersInGroup = $users | Where-Object { $_.id -eq $c } | Select-Object @{l = 'Username'; e = { $_.username } }, @{l = 'First Name'; e = { $_.firstName } }, @{l = 'Last Name'; e = { $_.lastName } }, @{l = 'Distinguished Name'; e = { $_.distinguishedName } }
                                        $usersInGroup | Table -Name 'Users' -List -ColumnWidths 25, 75
                                    }
                                }
                            }
                        }

                        foreach ($g in $groups | Where-Object { $_.authSourceId }) {
                            Section -Style Heading2 -Name 'Imported Groups' {
                                $groupsImported = $g | Where-Object { $_.authSourceId } | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'Description'; e = { $_.description } }
                                $groupsImported | Table -Name 'Imported Groups' -List -ColumnWidths 25, 75
                            }
                            if ($InfoLevel.Groups -ge 2) { 
                                Section -Style Heading3 -Name 'Users in Group' {
                                    $usersInGroup = @()
                                    foreach ($c in $g.userIds) {
                                        $usersInGroup = $users | Where-Object { $_.id -eq $c } | Select-Object @{l = 'Username'; e = { $_.username } }, @{l = 'First Name'; e = { $_.firstName } }, @{l = 'Last Name'; e = { $_.lastName } }#, @{l = 'Distinguished Name'; e = {$_.distinguishedName}}
                                        $usersInGroup | Table -Name 'Users in Group' -List -ColumnWidths 25, 75
                                    }
                                }
                            }
                        }
                    }
                }
            }
            #endregion Groups

            #region Users
            if ($InfoLevel.Users -ge 1) {
                if ($users) {
                    Section -Style Heading2 -Name 'User Accounts' {
                        Section -Style Heading3 -Name 'System Users' {
                            $systemUsers = $users | Where-Object { $_.distinguishedName -like '' } | Select-Object @{l = 'Username'; e = { $_.username } }, @{l = 'First Name'; e = { $_.firstName } }, @{l = 'Last Name'; e = { $_.lastName } }, @{l = 'Enabled'; e = { $_.enabled } }, @{l = 'Roles'; e = { $($_.rolenames) -join ', ' } }
                            $systemUsers | Table -Name 'System Users' -List -ColumnWidths 25, 75
                        }

                        Section -Style Heading3 -Name 'Imported Users' {
                            $importedUsers = $users | Where-Object { $_.'distinguishedName' -notlike '' } | Select-Object @{l = 'Username'; e = { $_.username } }, @{l = 'First Name'; e = { $_.firstName } }, @{l = 'Last Name'; e = { $_.lastName } }, @{l = 'Distinguished Name'; e = { $_.distinguishedName } }, @{l = 'Enabled'; e = { $_.enabled } }
                            $importedUsers | Table -Name 'Imported Users' -List -ColumnWidths 25, 75
                        }
                    }
                }
            }
            #endregion Users

            #region Remote Collectors
            if ($InfoLevel.RemoteCollectors -ge 1) {
                $collectorGroups = $(getCollectorGroups -resthost $vropshost -credential $Credential).collectorGroups
                if ($collectorGroups) {
                    Section -Style Heading2 -Name 'Remote Collector Groups' {
                        foreach ($rcGroup in $collectorGroups) {
                            Section -Style Heading3 -Name $($rcGroup).name {
                                $Group = $rcGroup | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'Description'; e = { $_.Description } }
                                $Group | Table -Name 'Remote Collector Groups' -ColumnWidths 25, 75

                                if ($InfoLevel.RemoteCollectors -ge 2) {
                                    Section -Style Heading4 -Name "Members" {
                                        foreach ($rcId in $($rcGroup).collectorId) {
                                            $rcNames = $(getCollectors -resthost $vropshost -credential $Credential).collector | Where-Object { $_.id -like $rcId }
                                            $rcNames = $rcNames | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'Hostname'; e = { $_.hostName } }
                                            $rcNames | Table -Name 'Members' -List -ColumnWidths 25, 75
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if ($InfoLevel.RemoteCollectors -ge 1) {
                Section -Style Heading2 -Name 'Remote Collectors' {
                    $collectors = $(getCollectors -resthost $vropshost -credential $Credential).collector
                    $localNodes = $collectors | Where-Object { $_.local -like '*True*' }
                    $remoteCollectors = $collectors | Where-Object { $_.local -like '*False*' }

                    if ($localNodes) {
                        Section -Style Heading3 -Name 'Local Nodes' {
                            $localNodes = $localNodes | Sort-Object ID | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'ID'; e = { $_.id } }, @{l = 'State'; e = { $_.State } }, @{l = 'Hostname'; e = { $_.hostname } }, @{l = 'Last Heartbeat'; e = { (convertEpoch -epochTime $_.lastHeartbeat) } }
                            $localNodes | Table -Name 'Local Nodes' -List -ColumnWidths 25, 75
                        }
                    }

                    if ($remoteCollectors) {
                        Section -Style Heading3 -Name 'Remote Nodes' {
                            $remoteCollectors = $remoteCollectors | Sort-Object ID | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'ID'; e = { $_.id } }, @{l = 'State'; e = { $_.State } }, @{l = 'Hostname'; e = { $_.hostname } }, @{l = 'Last Heartbeat'; e = { (convertEpoch -epochTime $_.lastHeartbeat) } }
                            $remoteCollectors | Table -Name 'Remote Nodes' -List -ColumnWidths 25, 75
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
                            Section -Style Heading3 -Name "$adapterKind" {
                                $AdapterInstances = $AdapterInstance | Where-Object { $_.resourcekey.adapterKindKey -like $adapterKind }

                                foreach ($adapter in $AdapterInstances) {
                                    $rc = $collectors | Where-Object { $_.id -like $adapter.collectorId }
                                    $adapter = $adapter | Select-Object @{l = 'Name'; e = { $_.resourceKey.name } }, @{l = 'Resource Kind'; e = { $_.resourceKey.resourceKindKey } }, @{l = 'Description'; e = { $_.description } }, @{l = 'Message from Adapter'; e = { $_.messageFromAdapterInstance } }, @{l = 'Collector Node'; e = { $rc.name } }, @{l = 'Last Heartbeat'; e = { (convertEpoch -epochTime $_.lastHeartbeat) } }, @{l = 'Last Collected'; e = { (convertEpoch -epochTime $_.lastCollected) } }, @{l = 'Metrics Collected'; e = { $_.numberOfMetricsCollected } }, @{l = 'Resources Collected'; e = { $_.numberOfResourcesCollected } } 
                                    $adapter | Table -Name 'Adapters' -List -ColumnWidths 25, 75
                                }
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
                    Section -Style Heading2 -Name 'Alerts' {
                        foreach ($a in $alerts) {
                            Section -Style Heading3 -Name $a.name {
                                $alertDetail = $a | Select-Object @{l = 'Name'; e = { $_.name } }, @{l = 'ID'; e = { $_.id } }, @{l = 'description'; e = { $_.description } }, @{l = 'adapterKindKey'; e = { $_.adapterKindKey } }, @{l = 'resourceKindKey'; e = { $_.resourceKindKey } }, @{l = 'waitCycles'; e = { $_.waitCycles } }, @{l = 'cancelCycles'; e = { $_.cancelCycles } }
                                $alertDetail | Table -Name 'Alerts' -List -ColumnWidths 25, 75

                                if ($InfoLevel.Symptoms -ge 2) {
                                    $symp = $a.states.'base-symptom-set'.symptomDefinitionIds
                                    foreach ($s in $symp) {
                                        $sympHashTable = @()
                                        $symDef = $(getSymptomDefinitions -resthost $vropshost -credential $Credential -symptomdefinitionid $s).symptomDefinitions
                                        if ($symDef.state.condition.type -contains 'CONDITION_MESSAGE_EVENT' ) {
                                            Section -Style Heading4 -Name "Symptom: $($symDef.name)" {

                                                $sympHashTable += [PSCustomObject]@{
                                                    'Name' = $symDef.name
                                                    'Id' = $symDef.Id
                                                    'AdapterKindKey' = $symDef.adapterKindKey
                                                    'ResourceKindKey' = $symDef.resourceKindKey
                                                    'Wait Cycles' = $symDef.waitCycles
                                                    'Cancel Cycles' = $symDef.cancelCycles
                                                    'Type' = $symDef.state.condition.type
                                                    'Event Type' = $symDef.state.condition.eventType
                                                    'Message' = $symDef.state.condition.message
                                                    'Operator' = $symDef.state.condition.operator
                                                }
                                                $sympHashTable | Table -Name 'Symptoms' -List -ColumnWidths 25, 75
                                            }
                                        } elseif ($symDef.state.condition.type -contains 'CONDITION_HT' ) {

                                            Section -Style Heading4 -Name "Symptom: $($symDef.name)" {
                                                $sympHashTable += [PSCustomObject]@{
                                                    'Name' = $symDef.name
                                                    'Id' = $symDef.Id
                                                    'AdapterKindKey' = $symDef.adapterKindKey
                                                    'ResourceKindKey' = $symDef.resourceKindKey
                                                    'Wait Cycles' = $symDef.waitCycles
                                                    'Cancel Cycles' = $symDef.cancelCycles
                                                    'Severity' = $symDef.state.severity
                                                    'Type' = $symDef.state.condition.type
                                                    'Key' = $symDef.state.condition.key
                                                    'Operator' = $symDef.state.condition.operator
                                                    'Value' = $symDef.state.condition.value
                                                    'Value Type' = $symDef.state.condition.valueType
                                                    'Instanced' = $symDef.state.condition.instanced
                                                    'Threshold Type' = $symDef.state.condition.thresholdType
                                                }
                                                $sympHashTable | Table -Name 'Symptoms' -List -ColumnWidths 25, 75
                                            } 
                                        } elseif ($symDef.state.condition.type -contains 'CONDITION_PROPERTY_STRING' ) {

                                            Section -Style Heading4 -Name "Symptom: $($symDef.name)" {
                                                $sympHashTable += [PSCustomObject]@{
                                                    'Name' = $symDef.name
                                                    'Id' = $symDef.Id
                                                    'AdapterKindKey' = $symDef.adapterKindKey
                                                    'ResourceKindKey' = $symDef.resourceKindKey
                                                    'Wait Cycles' = $symDef.waitCycles
                                                    'Cancel Cycles' = $symDef.cancelCycles
                                                    'Severity' = $symDef.state.severity
                                                    'Type' = $symDef.state.condition.type
                                                    'String Value' = $symDef.state.condition.stringValue
                                                    'Key' = $symDef.state.condition.key
                                                    'Operator' = $symDef.state.condition.operator
                                                    'Threshold Type' = $symDef.state.condition.thresholdType
                                                }
                                                $sympHashTable | Table -Name 'Symptoms' -List -ColumnWidths 25, 75
                                            }
                                        } elseif ($symDef.state.condition.type -contains 'CONDITION_FAULT' ) {

                                            Section -Style Heading4 -Name "Symptom: $($symDef.name)" {
                                                $sympHashTable += [PSCustomObject]@{
                                                    'Name' = $symDef.name
                                                    'Id' = $symDef.Id
                                                    'AdapterKindKey' = $symDef.adapterKindKey
                                                    'ResourceKindKey' = $symDef.resourceKindKey
                                                    'Wait Cycles' = $symDef.waitCycles
                                                    'Cancel Cycles' = $symDef.cancelCycles
                                                    'Severity' = $symDef.state.severity
                                                    'Type' = $symDef.state.condition.type
                                                    'Fault Key' = $symDef.state.condition.faultKey
                                                }
                                                $sympHashTable | Table -Name 'Symptoms' -List -ColumnWidths 25, 75
                                            }
                                        } elseif ($symDef.state.condition.type -contains 'CONDITION_PROPERTY_NUMERIC' ) {

                                            Section -Style Heading4 -Name "Symptom: $($symDef.name)" {
                                                $sympHashTable += [PSCustomObject]@{
                                                    'Name' = $symDef.name
                                                    'Id' = $symDef.Id
                                                    'AdapterKindKey' = $symDef.adapterKindKey
                                                    'ResourceKindKey' = $symDef.resourceKindKey
                                                    'Wait Cycles' = $symDef.waitCycles
                                                    'Cancel Cycles' = $symDef.cancelCycles
                                                    'Severity' = $symDef.state.severity
                                                    'Type' = $symDef.state.condition.type
                                                    'Value' = $symDef.state.condition.value
                                                    'Operator' = $symDef.state.condition.operator
                                                    'Key' = $symDef.state.condition.key
                                                    'Threshold Type' = $symDef.state.condition.thresholdType
                                                }
                                                $sympHashTable | Table -Name 'Symptoms' -List -ColumnWidths 25, 75

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
                        $superMetrics | Table -Name "Super Metrics" -List -ColumnWidths 25, 75
                    }
                }
            }
            #endregion Super Metrics

            #region Custom Groups
            if ($InfoLevel.CustomGroups -ge 1) {
                $customGroups = $(getCustomGroups -resthost $vropshost -credential $Credential).values
                if ($customGroups) {
                    Section -Style Heading2 -Name 'Custom Groups' {
                        $customGroups = $customGroups | Select-Object @{l = 'Name'; e = { $_.resourceKey.name } }, @{l = 'Adapter Kind'; e = { $_.resourceKey.adapterKindKey } }, @{l = 'Resource Kind'; e = { $_.resourceKey.resourceKindKey } }
                        $customGroups | Table -Name "Custom Groups" -List -ColumnWidths 25, 75
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
                        $reports | Table -Name "Reports" -List -ColumnWidths 25, 75
                    }
                }
            }
            #endregion Reports
        }       
    }
    #endregion Script Body
}