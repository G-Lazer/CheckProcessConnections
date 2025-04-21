# Define output CSV file path in the current script directory
$outputCsv = Join-Path -Path $PSScriptRoot -ChildPath "network_connections.csv"

# Create an empty array to store the results
$results = @()

# Get the netstat output with process IDs
$netstatOutput = netstat -ano | Select-String "TCP|UDP"

foreach ($line in $netstatOutput) {
    # Extract relevant parts from the netstat line
    $lineParts = $line.ToString().Split(" ", [System.StringSplitOptions]::RemoveEmptyEntries)
    
    # Filter out lines with missing PID (empty PID means not associated with a process)
    if ($lineParts.Length -ge 5) {
        $protocol = $lineParts[0]       # TCP/UDP
        $localAddress = $lineParts[1]    # Local address
        $foreignAddress = $lineParts[2]  # Foreign address
        $state = $lineParts[3]           # State (for TCP connections)
        $connPID = $lineParts[4]         # PID (Process ID) renamed to avoid conflict with $PID default variable in PowerShell

        # Get the process name using the PID
        try {
            $process = Get-Process -Id $connPID -ErrorAction Stop
            $processName = $process.ProcessName
            $processPath = $process.Path
        } catch {
            # If no process found for the PID (it may have ended), continue with next line
            continue
        }

        # Create an object to store the information
        $connection = [PSCustomObject]@{
            Protocol       = $protocol
            LocalAddress   = $localAddress
            ForeignAddress = $foreignAddress
            State          = $state
            PID            = $connPID
            ProcessName    = $processName
            ProcessPath    = $processPath
        }

        # Add the result to the results array
        $results += $connection
    }
}

# Export the results to CSV in the current script directory
$results | Export-Csv -Path $outputCsv -NoTypeInformation

# Output the path to the CSV file
Write-Host "Network connections and process details have been saved to $outputCsv"