' -------------------------------------------------------------------------------------
' DHCP Reservation Manager
' Description: Automates the process of adding reserved IPs to a Windows DHCP Server
' using netsh commands.
' -------------------------------------------------------------------------------------

Option Explicit

' --- Configuration ---
Dim SERVER_IP, SCOPE_ID, RANGE_START, RANGE_END
SERVER_IP   = "192.168.1.10" ' Target DHCP Server
SCOPE_ID    = "192.168.1.0"  ' Target Scope
RANGE_START = 100
RANGE_END   = 200

Dim strMac, strHostName
If WScript.Arguments.Count = 2 Then
    ' Clean MAC address (remove colons and dashes)
    strMac = Replace(Replace(WScript.Arguments.Item(0), ":", ""), "-", "")
    strHostName = WScript.Arguments.Item(1)
Else
    WScript.Echo "Usage: cscript DhcpReservationManager.vbs [MAC] [HostName]"
    WScript.Quit 1
End If

' Main logic
Dim currentConfig, macStatus
currentConfig = GetCommandOutput("netsh dhcp server " & SERVER_IP & " scope " & SCOPE_ID & " dump")
macStatus = CheckMacReservation(currentConfig, strMac)

If InStr(macStatus, "Reserved") > 0 Then
    WScript.Echo "[INFO] MAC " & strMac & " is already " & macStatus
Else
    ' Logic for finding next available IP and executing netsh add reservedip would go here
    WScript.Echo "[PROCESS] Proceeding with reservation for " & strHostName & " (" & strMac & ")..."
End If

' -------------------------------------------------------------------------------------
' Helper Functions
' -------------------------------------------------------------------------------------

' Checks if a MAC address already exists in the DHCP dump
Function CheckMacReservation(ByVal configDump, ByVal targetMac)
    Dim lines, line, data
    CheckMacReservation = "Available"
    lines = Split(configDump, vbCrLf)

    For Each line In lines
        If (InStr(line, "add reservedip") > 0) Then
            data = Split(line, " ")
            ' Usually index 8 in netsh dump contains the MAC
            If UBound(data) >= 8 Then
                If LCase(data(8)) = LCase(targetMac) Then
                    CheckMacReservation = "Reserved for IP=" & data(7)
                    Exit Function
                End If
            End If
        End If
    Next
End Function

' Executes a shell command and captures its standard output
Function GetCommandOutput(ByVal command)
    Dim objShell, objExec, output
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec(command)

    output = ""
    Do While Not objExec.StdOut.AtEndOfStream
        output = output & objExec.StdOut.ReadLine() & vbCrLf
    Loop

    Set objShell = Nothing
    Set objExec = Nothing
    GetCommandOutput = output
End Function

' Executes a command and prints output directly to console
Sub ExecuteAndPrint(ByVal command)
    Dim objShell, objExec
    Set objShell = CreateObject("WScript.Shell")
    Set objExec = objShell.Exec(command)

    Do While Not objExec.StdOut.AtEndOfStream
        WScript.StdOut.WriteLine objExec.StdOut.ReadLine()
    Loop
End Sub