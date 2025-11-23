<#
.SYNOPSIS
Controls Hisense 100BM66D TVs over IP using ASCII protocol.

.DESCRIPTION
This script sends ASCII protocol commands to Hisense 100BM66D TVs (BM series) over IP.
Supports power on/off and screen on/off commands.
The 100BM66D uses the BM series protocol with ASCII commands on port 8088.

.PARAMETER TVName
The name of the TV to control. Use -ListTVs to see available TVs.

.PARAMETER Command
The command to send to the TV. Valid values are:
  - 'poweron' : Power on the TV
  - 'poweroff' : Power off the TV
  - 'screenon' : Turn screen on
  - 'screenoff' : Turn screen off

.PARAMETER ListTVs
List all configured TVs with their IP addresses and MAC addresses.

.PARAMETER Port
The port number for the TV control protocol. Default is 8088 (BM series standard port).

.EXAMPLE
.\Control-HisenseTV.ps1 -ListTVs
.\Control-HisenseTV.ps1 -TVName "ProjectionRoom" -Command "poweron"
.\Control-HisenseTV.ps1 -TVName "ProjectionRoom" -Command "poweroff"
.\Control-HisenseTV.ps1 -TVName "ProjectionRoom" -Command "screenon"
.\Control-HisenseTV.ps1 -TVName "ProjectionRoom" -Command "screenoff"

#>

param(
    [Parameter(Mandatory=$false)]
    [string]$TVName,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet('poweron', 'poweroff', 'screenon', 'screenoff')]
    [string]$Command,
    
    [Parameter(Mandatory=$false)]
    [switch]$ListTVs,
    
    [Parameter(Mandatory=$false)]
    [int]$Port = 8088
)

# TV Configuration List
$TVConfig = @{
    "West" = @{
        IPAddress = "192.168.60.30"
        MACAddress = "00:1A:2B:3C:4D:5E"
        Description = "West Sanctuary display (piano side)"
    }
    "East" = @{
        IPAddress = "192.168.60.31"
        MACAddress = "00:1A:2B:3C:4D:5F"
        Description = "East Sanctuary display (organ side)"
    }
    # Add more TVs here as needed
}

function Get-ConfiguredTVs {
    Write-Host "`nConfigured Hisense TVs:`n"
    Write-Host ("{0,-20} {1,-15} {2,-17}" -f "TV Name", "IP Address", "MAC Address")
    Write-Host "-------------------------------------------------------------------"
    
    foreach ($tvName in $TVConfig.Keys | Sort-Object) {
        $tv = $TVConfig[$tvName]
        Write-Host ("{0,-20} {1,-15} {2,-17}" -f $tvName, $tv.IPAddress, $tv.MACAddress)
        Write-Host ("    Description: {0}" -f $tv.Description)
    }
    Write-Host ""
}

function Send-WakeOnLan {
    param(
        [string]$MACAddress,
        [string]$BroadcastAddress = "255.255.255.255"
    )
    
    try {
        # Convert MAC address to bytes
        $macBytes = $MACAddress -split ":" | ForEach-Object { [byte]"0x$_" }
        
        # Create magic packet: FF FF FF FF FF FF (6 times) + MAC address repeated 16 times
        $magicPacket = @()
        for ($i = 0; $i -lt 6; $i++) {
            $magicPacket += 0xFF
        }
        for ($i = 0; $i -lt 16; $i++) {
            $magicPacket += $macBytes
        }
        
        # Send magic packet via UDP
        $socket = New-Object System.Net.Sockets.UdpClient
        $socket.Connect($BroadcastAddress, 9)
        $socket.Send($magicPacket, $magicPacket.Count) | Out-Null
        $socket.Close()
        
        Write-Host "Wake-on-LAN packet sent to $MACAddress"
        return $true
    }
    catch {
        Write-Host "Error sending Wake-on-LAN packet: $_"
        return $false
    }
}

function Send-TVCommand {
    param(
        [string]$TVName,
        [string]$IPAddress,
        [string]$MACAddress,
        [byte[]]$Command,
        [int]$Port
    )
    
    try {
        $socket = New-Object System.Net.Sockets.Socket([System.Net.Sockets.AddressFamily]::InterNetwork, [System.Net.Sockets.SocketType]::Stream, [System.Net.Sockets.ProtocolType]::Tcp)
        $socket.ReceiveTimeout = 5000
        $socket.SendTimeout = 5000
        
        $endpoint = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Parse($IPAddress), $Port)
        $socket.Connect($endpoint)
        
        if ($socket.Connected) {
            Write-Host "Connected to '$TVName' at $IPAddress (MAC: $MACAddress) : $Port"
            
            # Send the ASCII command (hex string as text)
            $socket.Send($Command) | Out-Null
            
            # Display the command sent
            $asciiCommand = [System.Text.Encoding]::ASCII.GetString($Command)
            Write-Host "Command sent (ASCII text): $asciiCommand"
            
            # Wait a moment for response
            Start-Sleep -Milliseconds 500
            
            # Try to receive response
            $receiveBuffer = New-Object Byte[] 1024
            try {
                $bytesReceived = $socket.Receive($receiveBuffer, 1024, [System.Net.Sockets.SocketFlags]::None)
                if ($bytesReceived -gt 0) {
                    $hexResponse = ($receiveBuffer[0..($bytesReceived-1)] | ForEach-Object { $_.ToString("X2") }) -join " "
                    Write-Host "Response received (hex): $hexResponse"
                    # Also try to display as ASCII if readable
                    $asciiResponse = [System.Text.Encoding]::ASCII.GetString($receiveBuffer, 0, $bytesReceived)
                    Write-Host "Response (ASCII): $asciiResponse"
                }
            } catch {
                # Response timeout is okay, command may still succeed
            }
            
            $socket.Close()
            return $true
        }
        else {
            Write-Host "Failed to connect to '$TVName' at $IPAddress : $Port"
            return $false
        }
    }
    catch {
        Write-Host "Error: $_"
        return $false
    }
}

# Handle list TVs command
if ($ListTVs) {
    Get-ConfiguredTVs
    exit 0
}

# Validate parameters
if (-not $TVName -or -not $Command) {
    Write-Host "Error: TVName and Command parameters are required."
    Write-Host ""
    Write-Host "Usage: .\Control-HisenseTV.ps1 -TVName <name> -Command <command>"
    Write-Host ""
    Write-Host "Use -ListTVs to see available TV names:"
    Write-Host "  .\Control-HisenseTV.ps1 -ListTVs"
    Write-Host ""
    exit 1
}

# Get TV configuration
if (-not $TVConfig.ContainsKey($TVName)) {
    Write-Host "Error: TV '$TVName' not found in configuration."
    Write-Host ""
    Write-Host "Available TVs:"
    Get-ConfiguredTVs
    exit 1
}

$tv = $TVConfig[$TVName]
$IPAddress = $tv.IPAddress
$MACAddress = $tv.MACAddress

# BM Series (100BM66D) uses ASCII protocol on port 8088
# ASCII command format (from documentation):
# Commands are sent as ASCII hex strings (not binary)
#
# Power On:  DD FF 00 08 C1 15 00 00 01 BB BB DD BB CC
# Power Off: DD FF 00 08 C1 15 00 00 01 AA AA DD BB CC
# Screen On: DD FF 00 07 C1 31 00 00 01 F7 BB CC
# Screen Off: DD FF 00 07 C1 31 00 00 00 F6 BB CC

switch ($Command.ToLower()) {
    'poweron' {
        # Try Wake-on-LAN first
        Write-Host "Attempting Wake-on-LAN for '$TVName'..."
        $wolResult = Send-WakeOnLan -MACAddress $MACAddress
        
        # Wait for TV to boot
        Write-Host "Waiting for TV to boot (10 seconds)..."
        Start-Sleep -Seconds 10
        
        # Then send the power on command via IP
        Write-Host "Sending power on command via IP..."
        $commandString = "DDFF0008C115000001BBBBDDBBCC"
        $commandBytes = [System.Text.Encoding]::ASCII.GetBytes($commandString)
    }
    'poweroff' {
        # Power off command for BM series (ASCII hex string)
        $commandString = "DDFF0008C115000001AAAADDBBCC"
        $commandBytes = [System.Text.Encoding]::ASCII.GetBytes($commandString)
    }
    'screenon' {
        # Screen on command for BM series (ASCII hex string)
        $commandString = "DDFF0007C131000001F7BBCC"
        $commandBytes = [System.Text.Encoding]::ASCII.GetBytes($commandString)
    }
    'screenoff' {
        # Screen off command for BM series (ASCII hex string)
        $commandString = "DDFF0007C131000000F6BBCC"
        $commandBytes = [System.Text.Encoding]::ASCII.GetBytes($commandString)
    }
}

Write-Host "Sending '$Command' command to Hisense TV '$TVName' at $IPAddress : $Port"
$result = Send-TVCommand -TVName $TVName -IPAddress $IPAddress -MACAddress $MACAddress -Command $commandBytes -Port $Port

if ($result) {
    Write-Host "Command executed successfully"
    exit 0
}
else {
    Write-Host "Command execution failed"
    exit 1
}
