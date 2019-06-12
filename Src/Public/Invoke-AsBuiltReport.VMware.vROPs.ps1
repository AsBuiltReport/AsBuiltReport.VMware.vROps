function Invoke-AsBuiltReport.VMware.vROPs {
    <#
    .SYNOPSIS  
        PowerShell script to document the configuration of vRealize Operations Manager (vROPs) in Word/HTML/XML/Text formats
    .DESCRIPTION
        Documents the configuration of vROPs in Word/HTML/XML/Text formats using PScribo.
        vROPs code provided by andydvmware's PowervROPs PowerShell Module.
    .NOTES
        Version:        1.0
        Author:         Tim Williams
        Twitter:        @ymmit85
        Github:         ymmti85
        Credits:        andydvmware - PowervROPs powerShell Module
                        Iain Brighton (@iainbrighton) - PScribo module
    .LINK
        https://github.com/tpcarman/As-Built-Report
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

    $Options

    # If custom style not set, use vROPs style
    if (!$StyleName) {
        & "$PSScriptRoot\..\..\AsBuiltReport.VMware.vROPs.Style.ps1"
    }

        # Connect to vROPs using supplied credentials
        foreach ($vropshost in $Target) {
        #$token = acquireToken -resthost $vropshost -cr -authSource local
        #endregion Configuration Settings

        #region Script Body
        ###############################################################################################
        #                                       SCRIPT BODY                                           #
        ###############################################################################################
        if ($InfoLevel.GlobalSettings -ge 1) {
            Section -Style Heading1 -Name 'Global Settings' {
                $globalSettings = $(getGlobalsettings -resthost $vropshost -credential $Credential).keyValues
                if ($globalSettings) {
                    $globalSettings = $globalSettings |  select-object @{l = 'Setting'; e = {$_.key}}, @{l = 'Value'; e = {$_.values}}
                    $globalSettings | Table -List -ColumnWidths 25, 75
                    BlankLine
                }
            }
        }

        if ($InfoLevel.ServiceStatus -ge 1) {
            Section -Style Heading1 -Name 'Service Status' {
                $serviceStatus = $(getServicesInfo -resthost $vropshost -credential $Credential).service
                if ($serviceStatus) {
                    $serviceStatus = $serviceStatus |  select-object @{l = 'Name'; e = {$_.name}}, @{l = 'Health'; e = {$_.Health}}, @{l = 'Details'; e = {$_.details}}
                    $serviceStatus | Table -List -ColumnWidths 25, 75
                    BlankLine
                }
            }
        }

        if ($InfoLevel.Authentication -ge 1) {
            Section -Style Heading1 -Name 'Authentication' {
                $AuthSources = $(getAuthSources -resthost $vropshost -credential $Credential).sources
                if ($AuthSources) {
                    Section -Style Heading2 -Name 'AD Auth Sources' {
                        $AuthSources = $AuthSources | where {$_.sourcetype.name -like '*ACTIVE_DIRECTORY*'} | select-object @{l = 'Name'; e = {$_.name}}, @{l = 'ID'; e = {$_.ID}}
                        $AuthSources | Table -List -ColumnWidths 25, 75
                        BlankLine
                    }
                }
            }
        }

        if ($InfoLevel.Roles -ge 1) {
            Section -Style Heading1 -Name 'Roles' {
                $roles = $(getRoles -resthost $vropshost -credential $Credential).userRoles
                if ($roles) {

                    Section -Style Heading2 -Name 'System Roles' {
                    $roleSystem = $roles | where {$_.'system-created' -like 'True'} |select-object @{l = 'Name'; e = {$_.name}}, @{l = 'Description'; e = {$_.description}}, @{l = 'Display Name'; e = {$_.displayName}}   
                    $roleSystem | Table -List -ColumnWidths 25, 75
                    BlankLine
                    }

                    Section -Style Heading2 -Name 'Custom Roles' {
                        $roleSystem = $roles | where {$_.'system-created' -like 'False'} |select-object @{l = 'Name'; e = {$_.name}}, @{l = 'Description'; e = {$_.description}}, @{l = 'Display Name'; e = {$_.displayName}} 
                        $roleSystem | Table -List -ColumnWidths 25, 75
                        BlankLine
                    }
                }
            }
        }
        if ($InfoLevel.Groups -ge 1) {
            $users = $(getUsers -resthost $vropshost -credential $Credential).users
            $groups = $(getUserGroups -resthost $vropshost -credential $Credential).userGroups

            Section -Style Heading1 -Name 'Groups' {
                if ($groups) {
                    foreach ($g in ($groups | where {!($_.authSourceId)})) {
                        Section -Style Heading2 -Name 'System Groups' {
                            $groupsSystem = $g  |select-object @{l = 'Name'; e = {$_.name}}, @{l = 'Description'; e = {$_.description}}
                            $groupsSystem | Table -List -ColumnWidths 25, 75
                            BlankLine
                        }
                        if ($InfoLevel.Groups -ge 2) {
                            Section -Style Heading3 -Name 'Users' {
                                $usersInGroup = @()
                                foreach ($c in $g.userIds) {
                                    $usersInGroup = $users | where {$_.id -eq $c} | select-object @{l = 'Username'; e = {$_.username}}, @{l = 'First Name'; e = {$_.firstName}}, @{l = 'Last Name'; e = {$_.lastName}}, @{l = 'Distinguished Name'; e = {$_.distinguishedName}}
                                    $usersInGroup | Table -List -ColumnWidths 25, 75
                                    BlankLine
                                }
                            }
                        }
                    } 

                    foreach ($g in $groups | where {$_.authSourceId}) {
                        Section -Style Heading2 -Name 'Imported Groups' {
                            $groupsImported = $g | where {$_.authSourceId} |select-object @{l = 'Name'; e = {$_.name}}, @{l = 'Description'; e = {$_.description}}
                            $groupsImported | Table -List -ColumnWidths 25, 75
                            BlankLine
                        }
                            if ($InfoLevel.Groups -ge 2) { 
                                Section -Style Heading3 -Name 'Users in Group' {
                                    $usersInGroup = @()
                                    foreach ($c in $g.userIds) {
                                        $usersInGroup = $users | where {$_.id -eq $c} | select-object @{l = 'Username'; e = {$_.username}}, @{l = 'First Name'; e = {$_.firstName}}, @{l = 'Last Name'; e = {$_.lastName}}#, @{l = 'Distinguished Name'; e = {$_.distinguishedName}}
                                        $usersInGroup | Table -List -ColumnWidths 25, 75
                                        BlankLine
                                    }
                                }
                            }
                    }
                }
            }
        }

        if ($InfoLevel.Users -ge 1) {
            Section -Style Heading1 -Name 'User Accounts' {
                if ($users) {
                    Section -Style Heading2 -Name 'System Users' {
                        $systemUsers = $users | where {$_.distinguishedName -like ''} | select-object @{l = 'Username'; e = {$_.username}}, @{l = 'First Name'; e = {$_.firstName}}, @{l = 'Last Name'; e = {$_.lastName}}, @{l = 'Enabled'; e = {$_.enabled}}, @{l = 'Roles'; e = {$($_.rolenames) -join ', '}}
                        $systemUsers | Table -List -ColumnWidths 25, 75
                        BlankLine
                    }

                    Section -Style Heading2 -Name 'Imported Users' {
                        $importedUsers = $users | where {$_.'distinguishedName' -notlike ''} | select-object @{l = 'Username'; e = {$_.username}}, @{l = 'First Name'; e = {$_.firstName}}, @{l = 'Last Name'; e = {$_.lastName}}, @{l = 'Distinguished Name'; e = {$_.distinguishedName}}, @{l = 'Enabled'; e = {$_.enabled}}
                        $importedUsers | Table -List -ColumnWidths 25, 75
                        BlankLine
                    }
                }
            }
        }

        if ($InfoLevel.RemoteCollectors -ge 1) {
            $collectorGroups= $(getCollectorGroups -resthost $vropshost -credential $Credential).collectorGroups
            if ($collectorGroups) {
                Section -Style Heading2 -Name 'Remote Collector Groups' {
                    foreach ($rcGroup in $collectorGroups) {
                        Section -Style Heading3 -Name $($rcGroup).name {
                            $Group = $rcGroup |  select-object @{l = 'Name'; e = {$_.name}}, @{l = 'Description'; e = {$_.Description}}
                            $Group | Table -List -ColumnWidths 25, 75
                            BlankLine

                            if ($InfoLevel.RemoteCollectors -ge 2) {
                                Section -Style Heading4 -Name "Group Members" {
                                    foreach ($rcId in $($rcGroup).collectorId) {
                                        $rcNames = $(getCollectors -resthost $vropshost -credential $Credential).collector | where {$_.id -like $rcId}
                                        $rcNames = $rcNames | Select-Object @{l = 'Name'; e = {$_.name}},@{l = 'Hostname'; e = {$_.hostName}}
                                        $rcNames | Table -List -ColumnWidths 25, 75
                                        BlankLine
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        if ($InfoLevel.RemoteCollectors -ge 1) {
            Section -Style Heading1 -Name 'Remote Collectors' {
                $collectors= $(getCollectors -resthost $vropshost -credential $Credential).collector
                $localNodes = $collectors  | where {$_.local -like '*True*'}
                $remoteCollectors = $collectors  | where {$_.local -like '*False*'}

                if ($localNodes) {
                    Section -Style Heading2 -Name 'Local Nodes' {
                        $localNodes = $localNodes | sort-object ID | select-object @{l = 'Name'; e = {$_.name}},@{l = 'ID'; e = {$_.id}},  @{l = 'State'; e = {$_.State}}, @{l = 'Hostname'; e = {$_.hostname}}, @{l = 'Last Heartbeat'; e = {(convertEpoch -epochTime $_.lastHeartbeat)}}
                        $localNodes | Table -List -ColumnWidths 25, 75
                        BlankLine
                    }
                }

                if ($remoteCollectors) {
                    Section -Style Heading2 -Name 'Remote Collectors' {
                        $remoteCollectors = $remoteCollectors | sort-object ID | select-object @{l = 'Name'; e = {$_.name}},@{l = 'ID'; e = {$_.id}},  @{l = 'State'; e = {$_.State}}, @{l = 'Hostname'; e = {$_.hostname}}, @{l = 'Last Heartbeat'; e = {(convertEpoch -epochTime $_.lastHeartbeat)}}
                        $remoteCollectors | Table -List -ColumnWidths 25, 75
                        BlankLine
                    }
                }
            }
        }

        if ($InfoLevel.Adapters -ge 1) {
            Section -Style Heading1 -Name 'Adapters' {
                $AdapterInstance = $(getAdapterInstances -resthost $vropshost -credential $Credential).adapterInstancesInfoDto 
                $collectors= $(getCollectors -resthost $vropshost -credential $Credential).collector
                if ($AdapterInstance) {
                    foreach ($adapterKind in $AdapterInstance.resourcekey.adapterKindKey | sort -Unique) {
                        Section -Style Heading2 -Name "$adapterKind" {
                        $AdapterInstances = $AdapterInstance | where {$_.resourcekey.adapterKindKey -like $adapterKind}

                            foreach ($adapter in $AdapterInstances) {
                                $rc = $collectors | where {$_.id -like $adapter.collectorId}
                                $adapter = $adapter | select-object @{l = 'Name'; e = {$_.resourceKey.name}}, @{l = 'Resource Kind'; e = {$_.resourceKey.resourceKindKey}}, @{l = 'Description'; e = {$_.description}}, @{l = 'Message from Adapter'; e = {$_.messageFromAdapterInstance}}, @{l = 'Collector Node'; e = {$rc.name}}, @{l = 'Last Heartbeat'; e = {(convertEpoch -epochTime $_.lastHeartbeat)}}, @{l = 'Last Collected'; e = {(convertEpoch -epochTime $_.lastCollected)}}, @{l = 'Metrics Collected'; e = {$_.numberOfMetricsCollected}}, @{l = 'Resources Collected'; e = {$_.numberOfResourcesCollected}} 
                                $adapter | Table -List -ColumnWidths 25, 75
                                BlankLine
                            }
                        }
                    }
                }
            }
        }

        if ($InfoLevel.Alerts -ge 1) {
            $alerts = $(getAlertDefinitions -resthost $vropshost -credential $Credential).alertDefinitions | where {$_.name -like "*$($Options.AlertFilter)*"}
            if ($alerts) {
                Section -Style Heading1 -Name 'Alerts' {
                    foreach ($a in $alerts){
                        Section -Style Heading2 -Name $a.name {
                            $alertDetail = $a | select-object @{l = 'Name'; e = {$_.name}},@{l = 'ID'; e = {$_.id}},@{l = 'description'; e = {$_.description}},@{l = 'adapterKindKey'; e = {$_.adapterKindKey}},@{l = 'resourceKindKey'; e = {$_.resourceKindKey}},@{l = 'waitCycles'; e = {$_.waitCycles}},@{l = 'cancelCycles'; e = {$_.cancelCycles}}
                            $alertDetail | Table -List -ColumnWidths 25, 75
                            BlankLine

                            if ($InfoLevel.Alerts -ge 2) {
                                $symp = $a.states.'base-symptom-set'.symptomDefinitionIds
                                foreach ($s in $symp) {
                                    $sympHashTable = @()
                                    $symDef = $(getSymptomDefinitions -resthost $vropshost -credential $Credential -symptomdefinitionid $s).symptomDefinitions
                                    if ($symDef.state.condition.type -contains 'CONDITION_MESSAGE_EVENT' ) {
                                        Section -Style Heading3 -Name "Symptom: $($symDef.name)" {

                                        $sympHashTable += [PSCustomObject]@{
                                            'Name' = $symDef.name
                                            'Id' = $symDef.Id
                                            'adapterKindKey' = $symDef.adapterKindKey
                                            'resourceKindKey' = $symDef.resourceKindKey
                                            'Wait Cycles' = $symDef.waitCycles
                                            'cancelCycles' = $symDef.cancelCycles
                                            'type' = $symDef.state.condition.type
                                            'eventType' = $symDef.state.condition.eventType
                                            'message' = $symDef.state.condition.message
                                            'operator' = $symDef.state.condition.operator
                                        }
                                        $sympHashTable | Table -List -ColumnWidths 25, 75
                                        BlankLine
                                        }
                                    } elseif ($symDef.state.condition.type -contains 'CONDITION_HT' ) {

                                        Section -Style Heading4 -Name "Symptom: $($symDef.name)" {
                                            $sympHashTable += [PSCustomObject]@{
                                                'Name' = $symDef.name
                                                'Id' = $symDef.Id
                                                'adapterKindKey' = $symDef.adapterKindKey
                                                'resourceKindKey' = $symDef.resourceKindKey
                                                'Wait Cycles' = $symDef.waitCycles
                                                'cancelCycles' = $symDef.cancelCycles
                                                'severity' = $symDef.state.severity
                                                'type' = $symDef.state.condition.type
                                                'key' = $symDef.state.condition.key
                                                'operator' = $symDef.state.condition.operator
                                                'value' = $symDef.state.condition.value
                                                'valueType' = $symDef.state.condition.valueType
                                                'instanced' = $symDef.state.condition.instanced
                                                'thresholdType' = $symDef.state.condition.thresholdType
                                            }
                                                $sympHashTable | Table -List -ColumnWidths 25, 75
                                                BlankLine

                                        } 
                                    }   elseif ($symDef.state.condition.type -contains 'CONDITION_PROPERTY_STRING' ) {

                                        Section -Style Heading4 -Name "Symptom: $($symDef.name)" {
                                            $sympHashTable += [PSCustomObject]@{
                                                'Name' = $symDef.name
                                                'Id' = $symDef.Id
                                                'adapterKindKey' = $symDef.adapterKindKey
                                                'resourceKindKey' = $symDef.resourceKindKey
                                                'Wait Cycles' = $symDef.waitCycles
                                                'cancelCycles' = $symDef.cancelCycles
                                                'severity' = $symDef.state.severity
                                                'type' = $symDef.state.condition.type
                                                'stringValue' = $symDef.state.condition.stringValue
                                                'key' = $symDef.state.condition.key
                                                'operator' = $symDef.state.condition.operator
                                                'thresholdType' = $symDef.state.condition.thresholdType
                                            }
                                                $sympHashTable | Table -List -ColumnWidths 25, 75
                                                BlankLine
                                        }
                                    } elseif ($symDef.state.condition.type -contains 'CONDITION_FAULT' ) {

                                        Section -Style Heading4 -Name "Symptom: $($symDef.name)" {
                                            $sympHashTable += [PSCustomObject]@{
                                                'Name' = $symDef.name
                                                'Id' = $symDef.Id
                                                'adapterKindKey' = $symDef.adapterKindKey
                                                'resourceKindKey' = $symDef.resourceKindKey
                                                'Wait Cycles' = $symDef.waitCycles
                                                'cancelCycles' = $symDef.cancelCycles
                                                'severity' = $symDef.state.severity
                                                'type' = $symDef.state.condition.type
                                                'faultKey' = $symDef.state.condition.faultKey
                                                #'faultEvents' = $symDef.state.condition.faultEvents
                                            }
                                                $sympHashTable | Table -List -ColumnWidths 25, 75
                                                BlankLine
                                        }
                                    } elseif ($symDef.state.condition.type -contains 'CONDITION_PROPERTY_NUMERIC' ) {

                                        Section -Style Heading4 -Name "Symptom: $($symDef.name)" {
                                            $sympHashTable += [PSCustomObject]@{
                                                'Name' = $symDef.name
                                                'Id' = $symDef.Id
                                                'adapterKindKey' = $symDef.adapterKindKey
                                                'resourceKindKey' = $symDef.resourceKindKey
                                                'Wait Cycles' = $symDef.waitCycles
                                                'cancelCycles' = $symDef.cancelCycles
                                                'severity' = $symDef.state.severity
                                                'type' = $symDef.state.condition.type
                                                'value' = $symDef.state.condition.value
                                                'operator' = $symDef.state.condition.operator
                                                'key' = $symDef.state.condition.key
                                                'thresholdType' = $symDef.state.condition.thresholdType
                                            }
                                                $sympHashTable | Set-Style -Style Warning

                                        }
                                    } <#else {

                                        Section -Style Heading4 -Name "Symptom: $($symDef.name)" {
                                            $sympHashTable += [PSCustomObject]@{
                                                'Name' = $symDef.name
                                                'Id' = $symDef.Id
                                                'adapterKindKey' = $symDef.adapterKindKey
                                                'resourceKindKey' = $symDef.resourceKindKey
                                                'Wait Cycles' = $symDef.waitCycles
                                                'cancelCycles' = $symDef.cancelCycles
                                                'type' = $symDef.state.condition.type

                                            } 
                                            $sympHashTable | Table -List -ColumnWidths 25, 75
                                            BlankLine
                                        }
                                    }#>
                                }
                            }
                        }
                    }
                }
            }
        }

        if ($InfoLevel.SuperMetrics -ge 1) {
                Section -Style Heading1 -Name 'Super Metrics' {
                $superMetrics = $(getSuperMetrics -resthost $vropshost -credential $Credential).supermetrics
                if ($superMetrics) {
                    $superMetrics = $superMetrics |  select-object @{l = 'Name'; e = {$_.name}}, @{l = 'ID'; e = {$_.ID}}, @{l = 'Formula'; e = {$_.formula}}
                    $superMetrics | Table -List -ColumnWidths 25, 75
                    BlankLine
                }
            }
        }

        if ($InfoLevel.CustomGroups -ge 1) {
            Section -Style Heading1 -Name 'Custom Groups' {
                $customGroups = $(getCustomGroups -resthost $vropshost -credential $Credential).values
                if ($customGroups) {
                    $customGroups = $customGroups |  select-object @{l = 'Name'; e = {$_.resourceKey.name}}, @{l = 'Adapter Kind'; e = {$_.resourceKey.adapterKindKey}}, @{l = 'Resource Kind'; e = {$_.resourceKey.resourceKindKey}}
                    $customGroups | Table -List -ColumnWidths 25, 75
                    BlankLine
                }
            }
        }

        if ($InfoLevel.Reports -ge 1) {
            Section -Style Heading1 -Name 'Reports' {
                $reports = $(getReportDefinitions -resthost $vropshost -credential $Credential).reportDefinitions
                if ($reports) {
                    $reports = $reports |  select-object @{l = 'Name'; e = {$_.name}}, @{l = 'Description'; e = {$_.description}}, @{l = 'Owner'; e = {$_.owner}}, @{l = 'Subject'; e = {$_.subject -join ", " }}, @{l = 'Active'; e = {$_.active}}
                    $reports | Table -List -ColumnWidths 25, 75
                    BlankLine
                }
            }
        }


        #endregion Script Body
    }
}