<#
.SYNOPSIS
Controls a Hisense 100BM66D TV over IP using ASCII protocol.

.DESCRIPTION
This script sends ASCII protocol commands to a Hisense 100BM66D TV (BM series) over IP.
Supports power on/off and screen on/off commands.
The 100BM66D uses the BM series protocol with ASCII commands on port 8088.

.PARAMETER IPAddress
The IP address of the Hisense TV.

.PARAMETER Command
The command to send to the TV. Valid values are:
  - 'poweron' : Power on the TV
  - 'poweroff' : Power off the TV
  - 'screenon' : Turn screen on
  - 'screenoff' : Turn screen off

.PARAMETER Port
The port number for the TV control protocol. Default is 8088 (BM series standard port).

.EXAMPLE
.\Control-HisenseTV.ps1 -IPAddress "192.168.1.100" -Command "poweron"
.\Control-HisenseTV.ps1 -IPAddress "192.168.1.100" -Command "poweroff"
.\Control-HisenseTV.ps1 -IPAddress "192.168.1.100" -Command "screenon"
.\Control-HisenseTV.ps1 -IPAddress "192.168.1.100" -Command "screenoff"

#>

param(
    [Parameter(Mandatory=$true)]
    [string]$IPAddress,
    
    [Parameter(Mandatory=$true)]
    [ValidateSet('poweron', 'poweroff', 'screenon', 'screenoff')]
    [string]$Command,
    
    [Parameter(Mandatory=$false)]
    [int]$Port = 8088
)

function Send-TVCommand {
    param(
        [string]$IPAddress,
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
            Write-Host "Connected to TV at $IPAddress : $Port"
            
            # Send the ASCII command bytes
            $socket.Send($Command) | Out-Null
            
            Write-Host "Command sent (ASCII): $(([System.Text.Encoding]::ASCII.GetString($Command)))"
            
            # Wait a moment for response
            Start-Sleep -Milliseconds 500
            
            # Try to receive response
            $receiveBuffer = New-Object Byte[] 1024
            try {
                $bytesReceived = $socket.Receive($receiveBuffer, 1024, [System.Net.Sockets.SocketFlags]::None)
                if ($bytesReceived -gt 0) {
                    $response = [System.Text.Encoding]::ASCII.GetString($receiveBuffer, 0, $bytesReceived)
                    Write-Host "Response received: $response"
                }
            } catch {
                # Response timeout is okay, command may still succeed
            }
            
            $socket.Close()
            return $true
        }
        else {
            Write-Host "Failed to connect to TV at $IPAddress : $Port"
            return $false
        }
    }
    catch {
        Write-Host "Error: $_"
        return $false
    }
}

# BM Series (100BM66D) uses ASCII protocol on port 8088
# ASCII command format (from documentation):
#
# Power On:  DD FF 00 08 C1 15 00 00 01 BB BB DD BB CC
# Power Off: DD FF 00 08 C1 15 00 00 01 AA AA DD BB CC
# Screen On: DD FF 00 07 C1 31 00 01 01 F6 BB CC
# Screen Off: DD FF 00 07 C1 31 00 01 00 F7 BB CC

switch ($Command.ToLower()) {
    'poweron' {
        # Power on command for BM series
        # Start: DD FF, Length: 00 08, Code: C1 15, Reserved: 00 00, Data: 01 BB BB, Checksum: DD, End: BB CC
        $commandString = [char]0xDD + [char]0xFF + [char]0x00 + [char]0x08 + [char]0xC1 + [char]0x15 + [char]0x00 + [char]0x00 + [char]0x01 + [char]0xBB + [char]0xBB + [char]0xDD + [char]0xBB + [char]0xCC
        $command = [System.Text.Encoding]::ASCII.GetBytes($commandString)
    }
    'poweroff' {
        # Power off command for BM series
        # Start: DD FF, Length: 00 08, Code: C1 15, Reserved: 00 00, Data: 01 AA AA, Checksum: DD, End: BB CC
        $commandString = [char]0xDD + [char]0xFF + [char]0x00 + [char]0x08 + [char]0xC1 + [char]0x15 + [char]0x00 + [char]0x00 + [char]0x01 + [char]0xAA + [char]0xAA + [char]0xDD + [char]0xBB + [char]0xCC
        $command = [System.Text.Encoding]::ASCII.GetBytes($commandString)
    }
    'screenon' {
        # Screen on command for BM series
        # Start: DD FF, Length: 00 07, Code: C1 31, Reserved: 00 01, Data: 01, Checksum: F6, End: BB CC
        $commandString = [char]0xDD + [char]0xFF + [char]0x00 + [char]0x07 + [char]0xC1 + [char]0x31 + [char]0x00 + [char]0x01 + [char]0x01 + [char]0xF6 + [char]0xBB + [char]0xCC
        $command = [System.Text.Encoding]::ASCII.GetBytes($commandString)
    }
    'screenoff' {
        # Screen off command for BM series
        # Start: DD FF, Length: 00 07, Code: C1 31, Reserved: 00 01, Data: 00, Checksum: F7, End: BB CC
        $commandString = [char]0xDD + [char]0xFF + [char]0x00 + [char]0x07 + [char]0xC1 + [char]0x31 + [char]0x00 + [char]0x01 + [char]0x00 + [char]0xF7 + [char]0xBB + [char]0xCC
        $command = [System.Text.Encoding]::ASCII.GetBytes($commandString)
    }
}

Write-Host "Sending '$Command' command to Hisense TV at $IPAddress : $Port"
$result = Send-TVCommand -IPAddress $IPAddress -Command $command -Port $Port

if ($result) {
    Write-Host "Command executed successfully"
    exit 0
}
else {
    Write-Host "Command execution failed"
    exit 1
}
