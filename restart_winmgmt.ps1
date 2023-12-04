# =================================================================
# Powershell script to restart Windows Management Instrumentation service
# Workaround for issue where WmiPRVSE.exe consumes excessive CPU time when Synapse is running
# IMPORTANT: Must be run as administrator
# =================================================================


# =================================================================
# HOW TO USE:
# 1. Save script with .ps1 extension, e.g. C:\Scripts\restart_winmgmt.ps1
#   a. NOTE: The script will save a .log file with the same name as the script to this location
# 2. Open Task Scheduler by going to Start -> type in Task Scheduler
# 3. Add the script to the Task Scheduler
#   a. Create a Basic Task
#   a. Set the trigger to 'When I log on'
#   b. Set the program/script action to 'powershell'
#   c. Set the arguments to the location of the script in quotes, prefixed with -File, e.g. -File "C:\Scripts\restart_winmgmt.ps1"
#   d. Check the box to open properties dialog and hit finish
#   e. In the properties dialog under General, select "Run whether the user is logged on or not" and check the box "Run with highest privileges"
#   f. In the properties dialog under Conditions, uncheck the box "Stop if the computer switches to battery power" (for laptops)
# 4. Restart your computer
# =================================================================


# Variables
$WRITE_LOG_FILE = $true # if false, will not create a log file
$RESET_LOG_FILE = $true # reset log file each time script runs
$LOG_FILE_PATH  = "$PSScriptRoot\$(Split-Path -Leaf $MyInvocation.MyCommand.Name).log" # by default logs to same directory as the script
$RETRY_LIMIT = 30


$PROCESS_NAME = "Razer Synapse 3"
$SERVICES_TO_RESTART = [System.Collections.ArrayList]@("Winmgmt")
$DEPENDENT_SERVICES=[System.Collections.ArrayList]@()

# Main function
function Main {
    # Clear log file
    if ($RESET_LOG_FILE -and (Test-Path $LOG_FILE_PATH)) {
        Clear-Content $LOG_FILE_PATH 
    }

    WriteLog "Script starting..."

    # Check if run as administrator
    if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { 
        WriteLog "This script must be run as an administrator. Please re-run this script as an administrator."
        Exit
    }

    # Restart services
    CheckForProcess 
    StopServices 
    StartServices
    ExitScript
}

# Function to check if process is running
function CheckForProcess {
    WriteLog "Checking if $($PROCESS_NAME) is running..."
    $retryCount = 0
    do {
        $processFound = Get-Process $PROCESS_NAME -ErrorAction SilentlyContinue
        if (-not $processFound) {
            $retryCount++
            if ($retryCount -le $RETRY_LIMIT) {
                WriteLog "$($PROCESS_NAME) is not running. Waiting 10 seconds... [$($retryCount)]"
                Start-Sleep -Seconds 10
            }
            else {
                WriteLog "Did not detect $($PROCESS_NAME) after $($retryCount) attempts. Exiting."
                ExitScript
            }
        }
        else {
            WriteLog "$($PROCESS_NAME) is running. Waiting 30 seconds..."
            Start-Sleep -Seconds 30
        }
    } while (-not $processFound -and $retryCount -le $RETRY_LIMIT)
}

# Function to stop listed services along with any dependent services
function StopServices {

    $retryCount = 0
    WriteLog "Stopping services: $($SERVICES_TO_RESTART)..."

    # Iterate services to restart
    $SERVICES_TO_RESTART | ForEach-Object {
        
        $currentDependencies=[System.Collections.ArrayList]@()

        # Get dependent services 
        Get-Service $_ -DependentServices | Where-Object { $_.Status -eq 'Running' } | ForEach-Object {
            WriteLog "Dependent service: $($_.Name)..."
            [void]$currentDependencies.add($_.Name)

        }

        # First stop all dependent services, then stop the current service
        ($currentDependencies + $_) | ForEach-Object {
            
            # Retry logic
            do {
                try{
                    $SERVICE_STOPPED = $false
                    Stop-Service $_ -ErrorAction Stop
                }
                catch{
                    WriteLog $_
                }

                $serviceStatus = (Get-Service $_).Status

                if ($serviceStatus -eq 'Stopped') {
                    WriteLog "Stopped service $($_)."
                    $SERVICE_STOPPED = $true
                }

                # If service did not stop successfully, try again
                if (-not $SERVICE_STOPPED) {
                    $retryCount++
                    if ($retryCount -le $RETRY_LIMIT) {
                        WriteLog "Failed to stop $($_) service. Retrying in 5 seconds... [$($retryCount)]"
                        Start-Sleep -Seconds 5
                    }
                    else {
                        WriteLog "Failed to stop $($_) service after $($retryCount) attempts. Exiting."
                        ExitScript
                    }
                }
            } while (-not $SERVICE_STOPPED -and $retryCount -le $RETRY_LIMIT)
        }


        $currentDependencies | ForEach-Object {
            if ($DEPENDENT_SERVICES -notcontains $_) {  
                [void]$DEPENDENT_SERVICES.add($_)
            }
        }
    }
}

# Function to restart any services that were previously stopped
function StartServices {

    $retryCount = 0
    WriteLog "Starting services: $($SERVICES_TO_RESTART)..."

    ($SERVICES_TO_RESTART + $DEPENDENT_SERVICES) | ForEach-Object {

        # Retry logic
        do {
            try {
                $SERVICE_STARTED = $false
                Start-Service $_ -ErrorAction Stop
            }
            catch {
                WriteLog $_
            }

            $serviceStatus = (Get-Service $_).Status

            # Check if service was started successfully
            if ($serviceStatus -eq 'Running') {
                WriteLog "Started service $($_)."
                $SERVICE_STARTED = $true
            }

            # If service did not start successfully, try again
            if (-not $SERVICE_STARTED) {
                $retryCount++
                if ($retryCount -le $RETRY_LIMIT) {
                    WriteLog "Failed to start $($_) service. Retrying in 5 seconds... [$($retryCount)]"
                    Start-Sleep -Seconds 5
                }
                else {
                    WriteLog "Failed to start $($_) service after $($retryCount) attempts. Exiting."
                    ExitScript
                }
            }
        } while (-not $SERVICE_STARTED -and $retryCount -le $RETRY_LIMIT)
    }
}

function WriteLog($message) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - $message"
    if ($WRITE_LOG_FILE) {
        Write-Output $line >> $LOG_FILE_PATH
    }
    else {
        Write-Output $line
    }
}



function ExitScript {
    WriteLog "Script execution complete."
    Exit
}
    
Main
