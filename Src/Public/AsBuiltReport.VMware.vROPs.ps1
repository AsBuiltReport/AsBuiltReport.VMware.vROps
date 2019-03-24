#requires -Modules @{ModuleName="PScribo";ModuleVersion="0.7.23"},PowervROPs
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

    # If custom style not set, use vROPs style
    if (!$StyleName) {
        & "$PSScriptRoot\..\..\Styles\VMware.ps1"
    }


        # Connect to vROPs using supplied credentials
        foreach ($vropshost in $Target) {
        $token = acquireToken -resthost $vropshost -username $username -password $password -authSource $options.authsource
        
        #endregion Configuration Settings

        #region Script Body
        ###############################################################################################
        #                                       SCRIPT BODY                                           #
        ###############################################################################################


        if ($InfoLevel.Authentication -ge 1) {
            Section -Style Heading1 -Name 'Authentication' {
                $AuthSources = $(getAuthSources -resthost $vropshost -token $token).sources
                if ($AuthSources) {
                    Section -Style Heading2 -Name 'AD Auth Sources' {
                        $AuthSources = $AuthSources | where {$_.sourcetype.name -like '*ACTIVE_DIRECTORY*'} | select-object @{l = 'Name'; e = {$_.name}}, @{l = 'ID'; e = {$_.ID}}
                        $AuthSources | Table -List -ColumnWidths 25, 75
                    }
                }
            }
        }

        if ($InfoLevel.Roles -ge 1) {
            Section -Style Heading1 -Name 'Roles' {
                $roles = $(getRoles -resthost $vropshost -token $token).userRoles
                if ($roles) {

                    Section -Style Heading2 -Name 'System Roles' {
                    $roleSystem = $roles | where {$_.'system-created' -like 'True'} |select-object @{l = 'Name'; e = {$_.name}}, @{l = 'Description'; e = {$_.description}}, @{l = 'Display Name'; e = {$_.displayName}}   
                    $roleSystem | Table -List -ColumnWidths 25, 75
                    }

                    Section -Style Heading2 -Name 'Custom Roles' {
                        $roleSystem = $roles | where {$_.'system-created' -like 'False'} |select-object @{l = 'Name'; e = {$_.name}}, @{l = 'Description'; e = {$_.description}}, @{l = 'Display Name'; e = {$_.displayName}} 
                        $roleSystem | Table -List -ColumnWidths 25, 75
                    }
                }
            }
        }
        if ($InfoLevel.Groups -ge 1) {
            $users = $(getUsers -resthost $vropshost -token $token).users
            $groups = $(getUserGroups -resthost $vropshost -token $token).userGroups

            Section -Style Heading1 -Name 'Groups' {
                if ($groups) {
                    foreach ($g in ($groups | where {!($_.authSourceId)})) {
                        Section -Style Heading2 -Name 'System Groups' {
                            $groupsSystem = $g  |select-object @{l = 'Name'; e = {$_.name}}, @{l = 'Description'; e = {$_.description}}
                            $groupsSystem | Table -List -ColumnWidths 25, 75
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
                    }

                    Section -Style Heading2 -Name 'Imported Users' {
                        $importedUsers = $users | where {$_.'distinguishedName' -notlike ''} | select-object @{l = 'Username'; e = {$_.username}}, @{l = 'First Name'; e = {$_.firstName}}, @{l = 'Last Name'; e = {$_.lastName}}, @{l = 'Distinguished Name'; e = {$_.distinguishedName}}, @{l = 'Enabled'; e = {$_.enabled}}
                        $importedUsers | Table -List -ColumnWidths 25, 75
                    }
                }
            }
        }

        if ($InfoLevel.ServiceStatus -ge 1) {
            Section -Style Heading1 -Name 'Service Status' {
                $serviceStatus = $(getServicesInfo -resthost $vropshost -token $token).service
                if ($serviceStatus) {
                    $serviceStatus = $serviceStatus |  select-object @{l = 'Name'; e = {$_.name}}, @{l = 'Health'; e = {$_.Health}}, @{l = 'Details'; e = {$_.details}}
                    $serviceStatus | Table -List -ColumnWidths 25, 75
                }
            }
        }

        if ($InfoLevel.RemoteCollectors -ge 1) {
            Section -Style Heading1 -Name 'Remote Collectors' {
                $collectors= $(getCollectors -resthost $vropshost -token $token).collector
                $localNodes = $collectors  | where {$_.local -like '*True*'}
                $remoteCollectors = $collectors  | where {$_.local -like '*False*'}

                if ($localNodes) {
                    Section -Style Heading2 -Name 'Local Nodes' {
                        $localNodes = $localNodes | sort-object ID | select-object @{l = 'Name'; e = {$_.name}},@{l = 'ID'; e = {$_.id}},  @{l = 'State'; e = {$_.State}}, @{l = 'Hostname'; e = {$_.hostname}}
                        $localNodes | Table -List -ColumnWidths 25, 75
                    }
                }

                if ($remoteCollectors) {
                    Section -Style Heading2 -Name 'Remote Collectors' {
                        $remoteCollectors = $remoteCollectors | sort-object ID | select-object @{l = 'Name'; e = {$_.name}},@{l = 'ID'; e = {$_.id}},  @{l = 'State'; e = {$_.State}}, @{l = 'Hostname'; e = {$_.hostname}}
                        $remoteCollectors | Table -List -ColumnWidths 25, 75
                    }
                }
            }
        }

        if ($InfoLevel.RemoteCollectors -ge 1) {
            $collectorGroups= $(getCollectorGroups -resthost $vropshost -token $token).collectorGroups
            if ($collectorGroups) {
                Section -Style Heading2 -Name 'Remote Collector Groups' {
                    foreach ($rcGroup in $collectorGroups) {
                        Section -Style Heading3 -Name $($rcGroup).name {
                            $Group = $rcGroup |  select-object @{l = 'Name'; e = {$_.name}}, @{l = 'Description'; e = {$_.Description}}
                            $Group | Table -List -ColumnWidths 25, 75

                            if ($InfoLevel.RemoteCollectors -ge 2) {
                                Section -Style Heading4 -Name "Group Members" {
                                    foreach ($rcId in $($rcGroup).collectorId) {
                                        $rcNames = $(getCollectors -resthost $vropshost -token $token).collector | where {$_.id -like $rcId}
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

        if ($InfoLevel.Adapters -ge 1) {
            Section -Style Heading1 -Name 'Adapters' {
                $AdapterInstance = $(getAdapterInstances -resthost $vropshost -token $token).adapterInstancesInfoDto 
                $collectors= $(getCollectors -resthost $vropshost -token $token).collector
                if ($AdapterInstance) {
                    foreach ($adapterKind in $AdapterInstance.resourcekey.adapterKindKey | sort -Unique) {
                        Section -Style Heading2 -Name "$adapterKind" {
                        $AdapterInstances = $AdapterInstance | where {$_.resourcekey.adapterKindKey -like $adapterKind}

                            foreach ($adapter in $AdapterInstances) {
                                $rc = $collectors | where {$_.id -like $adapter.collectorId}
                                $adapter = $adapter | select-object @{l = 'Name'; e = {$_.resourceKey.name}}, @{l = 'Resource Kind'; e = {$_.resourceKey.resourceKindKey}}, @{l = 'Description'; e = {$_.description}}, @{l = 'Message from Adapter'; e = {$_.messageFromAdapterInstance}}, @{l = 'Collector Node'; e = {$rc.name}}
                                $adapter | Table -List -ColumnWidths 25, 75
                                BlankLine
                            }
                        }
                    }
                }
            }
        }

        if ($InfoLevel.Alerts -ge 1) {
            $alerts = $(getAlertDefinitions -resthost $vropshost -token $token).alertDefinitions | where {$_.name -like "*$($Options.AlertFilter)*"}
            if ($alerts) {
                Section -Style Heading1 -Name 'Alerts' {
                    foreach ($a in $alerts){
                        Section -Style Heading2 -Name $a.name {
                            $alertDetail = $a | select-object @{l = 'Name'; e = {$_.name}},@{l = 'ID'; e = {$_.id}},@{l = 'description'; e = {$_.description}},@{l = 'adapterKindKey'; e = {$_.adapterKindKey}},@{l = 'resourceKindKey'; e = {$_.resourceKindKey}},@{l = 'waitCycles'; e = {$_.waitCycles}},@{l = 'cancelCycles'; e = {$_.cancelCycles}}
                            $alertDetail | Table -List -ColumnWidths 25, 75
                            if ($InfoLevel.Alerts -ge 2) {

                                $symp = $a.states.'base-symptom-set'.symptomDefinitionIds
                                foreach ($s in $symp) {
                                    $sympHashTable = @()
                                    $symDef = $(getSymptomDefinitions -resthost $vropshost -token $token -symptomdefinitionid $s).symptomDefinitions
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
                                    } else {

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
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        if ($InfoLevel.SuperMetrics -ge 1) {
                Section -Style Heading1 -Name 'Super Metrics' {
                $superMetrics = $(getSuperMetrics -resthost $vropshost -token $token).supermetrics
                if ($superMetrics) {
                    $superMetrics = $superMetrics |  select-object @{l = 'Name'; e = {$_.name}}, @{l = 'ID'; e = {$_.ID}}, @{l = 'Formula'; e = {$_.formula}}
                    $superMetrics | Table -List -ColumnWidths 25, 75
                }
            }
        }

        if ($InfoLevel.CustomGroups -ge 1) {
            Section -Style Heading1 -Name 'Custom Groups' {
                $customGroups = $(getCustomGroups -resthost $vropshost -token $token).values.resourceKey
                if ($customGroups) {
                    $customGroups = $customGroups |  select-object @{l = 'Name'; e = {$_.name}}, @{l = 'Adapter Kind'; e = {$_.adapterKindKey}}, @{l = 'Resource Kind'; e = {$_.resourceKindKey}}
                    $customGroups | Table -List -ColumnWidths 25, 75
                }
            }
        }

        #endregion Script Body
    }
}