Function Add-CMDevicesToDeviceCollection {
    param(
        [string]$SCCMServer,
        [string]$SiteCode,
        [string]$ComputerName,
        [string]$TextFilePath,
        [string]$CSVFilePath,
        [string]$CSVColumnName,
        [string]$CommaSeperatedList,
        [string]$Query,
        [string]$QueryName,
        [string]$CollectionName,
        [switch]$Force,
        [string]$LimitingCollection,
        [ValidateSet("Both", "Continuous", "Manual", "None", "Periodic")]
        $RefreshType
    )

    $ProviderMachineName = "$SCCMServer"
    $initParams = @{}

    # Import the ConfigurationManager.psd1 module 
    if ((Get-Module ConfigurationManager) -eq $null) {
        Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams
    }

    # Connect to the site's drive if it is not already present
    if ((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
        New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
    }

    # Set the current location to be the site code.
    Set-Location "$($SiteCode):\" @initParams

    # Handle $ComputerName parameter
    if ($ComputerName) {
        $DeviceCollection = Get-CMDeviceCollection -Name $CollectionName

        if ($DeviceCollection) {
            $Device = (Get-CMDevice -Name $ComputerName -Fast).ResourceID

            if ($Device) {
                try {
                    Add-CMDeviceCollectionDirectMembershipRule -CollectionName $CollectionName -ResourceId $Device -ErrorAction Stop
                    Write-Host "Successfully added $ComputerName to $CollectionName"
                }
                catch {
                    Write-Host "Failed to add $ComputerName to $CollectionName"
                }
            }
            else {
                Write-Error "$ComputerName not found. Please ensure the correct name was entered and try again."
            }
        }
        else {
            # If force is used and LimitingCollection & RefreshType are specified
            if ($Force -and $LimitingCollection -and $RefreshType) {
                try {
                    New-CMDeviceCollection -Name $CollectionName -LimitingCollectionName $LimitingCollection -RefreshType $RefreshType | Out-Null
                    Start-Sleep 5
                    $Device = (Get-CMDevice -Name $ComputerName -Fast).ResourceID

                    if ($Device) {
                        Add-CMDeviceCollectionDirectMembershipRule -CollectionName $CollectionName -ResourceId $Device -ErrorAction Stop | Out-Null
                    }
                    else {
                        Write-Host "Created $CollectionName device collection, but $ComputerName not found to add."
                    }
                }
                catch {
                    Write-Error "Failed to create $CollectionName or failed to add $ComputerName to $CollectionName"
                }
            }
            elseif ($Force -and ((!$LimitingCollection) -or (!$RefreshType))) {
                Write-Error "Force parameter is used, but either LimitingCollection or RefreshType is not specified. Please specify and try again."
            }
            else {
                Write-Error "$CollectionName device collection was not found in CM. If you would like to create this device collection, please use the -Force parameter."
            }
        }
    }

    # Handle $TextFilePath parameter
    if ($TextFilePath) {
        $DeviceCollection = Get-CMDeviceCollection -Name $CollectionName

        if ($DeviceCollection) {
            if (Test-Path $TextFilePath) {
                $Computers = Get-Content -Path $TextFilePath
                foreach ($Computer in $Computers) {
                    $Device = (Get-CMDevice -Name $Computer -Fast).ResourceID
                    if ($Device) {
                        try {
                            Add-CMDeviceCollectionDirectMembershipRule -CollectionName $CollectionName -ResourceId $Device -ErrorAction Stop
                            Write-Host "Successfully added $Computer to $CollectionName"
                        }
                        catch {
                            Write-Host "Failed to add $Computer to $CollectionName"
                        }
                    }
                    else {
                        Write-Error "$Computer not found. Please ensure the correct name was entered and try again."
                    }
                }
            }
            else {
                Write-Error "Text file path $TextFilePath does not exist."
            }
        }
        else {
            # If force is used and LimitingCollection & RefreshType are specified
            if ($Force -and $LimitingCollection -and $RefreshType) {
                try {
                    New-CMDeviceCollection -Name $CollectionName -LimitingCollectionName $LimitingCollection -RefreshType $RefreshType | Out-Null
                    Start-Sleep 5
                    if (Test-Path $TextFilePath) {
                        $Computers = Get-Content -Path $TextFilePath
                        foreach ($Computer in $Computers) {
                            $Device = (Get-CMDevice -Name $Computer -Fast).ResourceID
                            if ($Device) {
                                Add-CMDeviceCollectionDirectMembershipRule -CollectionName $CollectionName -ResourceId $Device -ErrorAction Stop | Out-Null
                            }
                            else {
                                Write-Host "Created $CollectionName device collection, but $Computer not found to add."
                            }
                        }
                    }
                    else {
                        Write-Error "Text file path $TextFilePath does not exist."
                    }
                }
                catch {
                    Write-Error "Failed to create $CollectionName or failed to add computers from $TextFilePath to $CollectionName"
                }
            }
            elseif ($Force -and ((!$LimitingCollection) -or (!$RefreshType))) {
                Write-Error "Force parameter is used, but either LimitingCollection or RefreshType is not specified. Please specify and try again."
            }
            else {
                Write-Error "$CollectionName device collection was not found in CM. If you would like to create this device collection, please use the -Force parameter."
            }
        }
    }


    # Handle $CSVFilePath parameter
if ($CSVFilePath) {
    if (!$CSVColumnName) {
        Write-Error "CSVColumnName must be specified when using CSVFilePath."
    } 
    else {
        $CSVData = Import-Csv -Path $CSVFilePath
        $Computers = $CSVData.$CSVColumnName

        $DeviceCollection = Get-CMDeviceCollection -Name $CollectionName

        if ($DeviceCollection) {
            foreach ($Computer in $Computers) {
                $Device = (Get-CMDevice -Name $Computer -Fast).ResourceID
                if ($Device) {
                    try {
                        Add-CMDeviceCollectionDirectMembershipRule -CollectionName $CollectionName -ResourceId $Device -ErrorAction Stop
                        Write-Host "Successfully added $Computer to $CollectionName"
                    }
                    catch {
                        Write-Host "Failed to add $Computer to $CollectionName"
                    }
                }
                else {
                    Write-Error "$Computer not found. Please ensure the correct name was entered and try again."
                }
            }
        } 
        else {
            # If force is used and LimitingCollection & RefreshType are specified
            if ($Force -and $LimitingCollection -and $RefreshType) {
                try {
                    New-CMDeviceCollection -Name $CollectionName -LimitingCollectionName $LimitingCollection -RefreshType $RefreshType | Out-Null
                    Start-Sleep 5
                    foreach ($Computer in $Computers) {
                        $Device = (Get-CMDevice -Name $Computer -Fast).ResourceID
                        if ($Device) {
                            Add-CMDeviceCollectionDirectMembershipRule -CollectionName $CollectionName -ResourceId $Device -ErrorAction Stop | Out-Null
                        }
                        else {
                            Write-Host "Created $CollectionName device collection, but $Computer not found to add."
                        }
                    }
                }
                catch {
                    Write-Error "Failed to create $CollectionName or failed to add computers from $CSVFilePath to $CollectionName"
                }
            }
            elseif ($Force -and ((!$LimitingCollection) -or (!$RefreshType))) {
                Write-Error "Force parameter is used, but either LimitingCollection or RefreshType is not specified. Please specify and try again."
            }
            else {
                Write-Error "$CollectionName device collection was not found in CM. If you would like to create this device collection, please use the -Force parameter."
            }
        }
    }
}

# Handle $Query parameter
if ($Query) {
    if (!$QueryName) {
        Write-Error "QueryName must be specified when using Query."
    } 
    else {
        $DeviceCollection = Get-CMDeviceCollection -Name $CollectionName

        if ($DeviceCollection) {
            try {
                Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionName -QueryExpression $Query -RuleName $QueryName -ErrorAction Stop
                Write-Host "Successfully added query $QueryName to $CollectionName"
            }
            catch {
                Write-Error "Failed to add query $QueryName to $CollectionName"
            }
        } 
        else {
            # If force is used and LimitingCollection & RefreshType are specified
            if ($Force -and $LimitingCollection -and $RefreshType) {
                try {
                    New-CMDeviceCollection -Name $CollectionName -LimitingCollectionName $LimitingCollection -RefreshType $RefreshType | Out-Null
                    Start-Sleep 5
                    Add-CMDeviceCollectionQueryMembershipRule -CollectionName $CollectionName -QueryExpression $Query -RuleName $QueryName -ErrorAction Stop
                    Write-Host "Created $CollectionName device collection and added query $QueryName"
                }
                catch {
                    Write-Error "Failed to create $CollectionName or failed to add query $QueryName to $CollectionName"
                }
            }
            elseif ($Force -and ((!$LimitingCollection) -or (!$RefreshType))) {
                Write-Error "Force parameter is used, but either LimitingCollection or RefreshType is not specified. Please specify and try again."
            }
            else {
                Write-Error "$CollectionName device collection was not found in CM. If you would like to create this device collection, please use the -Force parameter."
            }
        }
    }
}

# Handle $CommaSeparatedList parameter
if ($CommaSeparatedList) {
    $Computers = $CommaSeparatedList -split ',' | ForEach-Object { $_.Trim() }

    $DeviceCollection = Get-CMDeviceCollection -Name $CollectionName

    if ($DeviceCollection) {
        foreach ($Computer in $Computers) {
            $Device = (Get-CMDevice -Name $Computer -Fast).ResourceID
            if ($Device) {
                try {
                    Add-CMDeviceCollectionDirectMembershipRule -CollectionName $CollectionName -ResourceId $Device -ErrorAction Stop
                    Write-Host "Successfully added $Computer to $CollectionName"
                }
                catch {
                    Write-Host "Failed to add $Computer to $CollectionName"
                }
            }
            else {
                Write-Error "$Computer not found. Please ensure the correct name was entered and try again."
            }
        }
    } 
    else {
        # If force is used and LimitingCollection & RefreshType are specified
        if ($Force -and $LimitingCollection -and $RefreshType) {
            try {
                New-CMDeviceCollection -Name $CollectionName -LimitingCollectionName $LimitingCollection -RefreshType $RefreshType | Out-Null
                Start-Sleep 5
                foreach ($Computer in $Computers) {
                    $Device = (Get-CMDevice -Name $Computer -Fast).ResourceID
                    if ($Device) {
                        Add-CMDeviceCollectionDirectMembershipRule -CollectionName $CollectionName -ResourceId $Device -ErrorAction Stop | Out-Null
                    }
                    else {
                        Write-Host "Created $CollectionName device collection, but $Computer not found to add."
                    }
                }
            }
            catch {
                Write-Error "Failed to create $CollectionName or failed to add computers from $CommaSeparatedList to $CollectionName"
            }
        }
        elseif ($Force -and ((!$LimitingCollection) -or (!$RefreshType))) {
            Write-Error "Force parameter is used, but either LimitingCollection or RefreshType is not specified. Please specify and try again."
        }
        else {
            Write-Error "$CollectionName device collection was not found in CM. If you would like to create this device collection, please use the -Force parameter."
        }
    }
}

}
