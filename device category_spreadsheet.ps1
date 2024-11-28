<#
    Author: Adouglastech
    Created: 11/18/2024
    Script Purpose:
    This script connects to the Microsoft Graph API to process a raw device list and validate device IDs. It then updates the device categories based on the information provided in a spreadsheet.

    Usage Instructions:
    1. Ensure you have the necessary permissions and packages to connect to the Microsoft Graph API.
    2. Before running the script, export a raw device list from intune and prepare the spreadsheet with the necessary updates (device IDs and categories).
    3. Modify any required variables in the script, such as path to the log file.
    4. Run the script in a PowerShell session with appropriate privileges.

    Requirements:
    - PowerShell 7.0 or above
    - Microsoft.Graph PowerShell module
    - Authentication credentials for Microsoft Graph API
    
    Notes:
    - Make sure your xlsx spreadsheet is formatted with the correct colums ( DeviceID | DeviceName | NewCategory | and are named appropriately ).
    - This script is intended to streamline the device management process within Intune.
	- If you plan to use the script interace to point to the path of the xlsx file remove the ""'s
#>

$ErrorActionPreference = "silentlycontinue"

# Import module to handle Excel files
Import-Module ImportExcel

# Authenticate to Microsoft Graph using Device Code Authentication
Try {
    Connect-MgGraph -Scopes "DeviceManagementManagedDevices.ReadWrite.All", "DeviceManagementConfiguration.ReadWrite.All" -DeviceCode -ErrorAction Stop
}
Catch {
    Write-Host "An error occurred during authentication:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Add-Content -Path "$env:USERPROFILE\Desktop\DeviceCategoryChangeLog.txt" -Value "An error occurred during authentication: $($_.Exception.Message)"
    pause
    exit
}

# Define color functions for better readability
function Green { process { Write-Host $_ -ForegroundColor Green } }
function Red { process { Write-Host $_ -ForegroundColor Red } }
function Yellow { process { Write-Host $_ -ForegroundColor Yellow } }

# Function to change the device category
function Change-DeviceCategory {
    param(
        [Parameter(Mandatory)]
        [string]$DeviceID,

        [Parameter(Mandatory)]
        [string]$NewCategoryID,

        [Parameter(Mandatory)]
        [string]$DeviceName
    )

    # Log the intended change for tracking purposes
    Write-Host "Attempting to change category for device: $DeviceName (ID: $DeviceID) to category ID: $NewCategoryID" -ForegroundColor Yellow
    Add-Content -Path "$env:USERPROFILE\Desktop\DeviceCategoryChangeLog.txt" -Value "Attempting to change category for device: $DeviceName (ID: $DeviceID) to category ID: $NewCategoryID"

    $body = @{ "@odata.id" = "https://graph.microsoft.com/beta/deviceManagement/deviceCategories/$NewCategoryID" }
    Try {
        $response = Invoke-MgGraphRequest -Method PUT -Uri "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$DeviceID/deviceCategory/`$ref" -Body ($body | ConvertTo-Json) -ErrorAction Stop
        
        if ($response -eq $null) {
            Write-Host "Category change request sent successfully for device: $DeviceID." -ForegroundColor Green
            Add-Content -Path "$env:USERPROFILE\Desktop\DeviceCategoryChangeLog.txt" -Value "Category change request sent successfully for device: $DeviceName (ID: $DeviceID)"
        }
        else {
            Write-Host "Unexpected response during category change for device: $DeviceID." -ForegroundColor Red
            Write-Output $response
            Add-Content -Path "$env:USERPROFILE\Desktop\DeviceCategoryChangeLog.txt" -Value "Unexpected response during category change for device: $DeviceName (ID: $DeviceID): $($response | ConvertTo-Json)"
        }
    }
    Catch {
        Write-Host "An error occurred while changing category for device: $DeviceID" -ForegroundColor Red
        Write-Host $_.Exception.Message -ForegroundColor Red
        Add-Content -Path "$env:USERPROFILE\Desktop\DeviceCategoryChangeLog.txt" -Value "An error occurred while changing category for device: $DeviceName (ID: $DeviceID): $($_.Exception.Message)"
    }
}

# Function to process devices from spreadsheet
function Process-DevicesFromSpreadsheet {
    param(
        [Parameter(Mandatory)]
        [string]$SpreadsheetPath,

        [Parameter(Mandatory)]
        [int]$MaxRetries
    )

    # Load data from spreadsheet
    Try {
        $DeviceData = Import-Excel -Path $SpreadsheetPath -ErrorAction Stop
    }
    Catch {
        Write-Host "Failed to load spreadsheet. Please check the file path and format." -ForegroundColor Red
        Add-Content -Path "$env:USERPROFILE\Desktop\DeviceCategoryChangeLog.txt" -Value "Failed to load spreadsheet. Error: $($_.Exception.Message)"
        exit
    }

    foreach ($Device in $DeviceData) {
        $DeviceID = $Device.DeviceID.Trim()
        $DeviceName = $Device.DeviceName.Trim()
        $NewCategory = $Device.NewCategory.Trim()

        if (-not $DeviceID -or -not $DeviceName -or -not $NewCategory) {
            Write-Host "Incomplete data for a device. Skipping..." -ForegroundColor Red
            Add-Content -Path "$env:USERPROFILE\Desktop\DeviceCategoryChangeLog.txt" -Value "Incomplete data for a device. Skipping entry: $($Device | ConvertTo-Json)"
            continue
        }

        # Attempt to get category ID from display name
        Try {
            $CategoryDetails = Get-MgDeviceManagementDeviceCategory -Filter "displayName eq '$NewCategory'" | Select-Object -First 1 -Property Id, DisplayName
            $NewCategoryID = $CategoryDetails.Id
            if (-not $NewCategoryID) {
                Write-Host "Category ID not found for category name: $NewCategory" -ForegroundColor Red
                Add-Content -Path "$env:USERPROFILE\Desktop\DeviceCategoryChangeLog.txt" -Value "Category ID not found for category name: $NewCategory for device: $DeviceName (ID: $DeviceID)"
                continue
            }
        }
        Catch {
            Write-Host "Failed to retrieve category ID for category name: $NewCategory" -ForegroundColor Red
            Add-Content -Path "$env:USERPROFILE\Desktop\DeviceCategoryChangeLog.txt" -Value "Failed to retrieve category ID for category name: $NewCategory for device: $DeviceName (ID: $DeviceID). Error: $($_.Exception.Message)"
            continue
        }

        # Change the device category
        Change-DeviceCategory -DeviceID $DeviceID -NewCategoryID $NewCategoryID -DeviceName $DeviceName

        # Wait for the category assignment to complete, with a maximum retry count
        if ($MaxRetries -gt 0) {
            $RetryCount = 0
            do {
                $DeviceDetails = Get-MgDeviceManagementManagedDevice -DeviceId $DeviceID | Select-Object -Property Id, DeviceCategoryDisplayName
                $DeviceCategoryCurrent = $DeviceDetails.DeviceCategoryDisplayName
                if ($DeviceCategoryCurrent -ne $NewCategory) {
                    Write-Host "Please wait..." -ForegroundColor Yellow
                    Start-Sleep -Seconds 10
                    $RetryCount++
                }
            } Until ($DeviceCategoryCurrent -eq $NewCategory -or $RetryCount -ge $MaxRetries)

            if ($DeviceCategoryCurrent -eq $NewCategory) {
                Write-Host "Category of device '$DeviceName' (ID: $DeviceID) is changed to $NewCategory" -ForegroundColor Green
                Add-Content -Path "$env:USERPROFILE\Desktop\DeviceCategoryChangeLog.txt" -Value "Category of $DeviceName (ID: $DeviceID) is changed to $NewCategory"
            }
            else {
                Write-Host "Category change verification timed out for device: $DeviceName. Please verify manually." -ForegroundColor Red
                Add-Content -Path "$env:USERPROFILE\Desktop\DeviceCategoryChangeLog.txt" -Value "Category change verification timed out for device: $DeviceName (ID: $DeviceID). Please verify manually."
            }
        }
        else {
            Write-Host "Max retries set to 0, skipping verification for device: $DeviceName (ID: $DeviceID)." -ForegroundColor Yellow
            Add-Content -Path "$env:USERPROFILE\Desktop\DeviceCategoryChangeLog.txt" -Value "Max retries set to 0, skipping verification for device: $DeviceName (ID: $DeviceID)."
        }
    }
}

# Prompt user for the spreadsheet path
$SpreadsheetPath = Read-Host -Prompt 'Enter the path to the spreadsheet containing device information'

# Process devices from the spreadsheet
$MaxRetries = Read-Host -Prompt 'Enter the maximum number of retries (default is 5)'
if (-not $MaxRetries) { $MaxRetries = 5 }
Process-DevicesFromSpreadsheet -SpreadsheetPath $SpreadsheetPath -MaxRetries $MaxRetries

# Prevent the PowerShell window from closing immediately
Write-Host "Script execution completed. Press any key to exit..." -ForegroundColor Green
Add-Content -Path "$env:USERPROFILE\Desktop\DeviceCategoryChangeLog.txt" -Value "Script execution completed."
[System.Console]::ReadKey()
