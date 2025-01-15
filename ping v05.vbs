Option Explicit

'---------------------------------------------------------------------------------
' 1) FORCE CSCRIPT MODE (CONSOLE). If already in cscript, do nothing.
'---------------------------------------------------------------------------------
If InStr(LCase(WScript.FullName), "cscript.exe") = 0 Then
    Dim reLaunch
    reLaunch = "cscript //nologo """ & WScript.ScriptFullName & """"
    CreateObject("WScript.Shell").Run reLaunch, 1, True
    WScript.Quit
End If

Dim Ws, fso
Set Ws  = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

'---------------------------------------------------------------------------------
' 2) IP LIST
'---------------------------------------------------------------------------------
Dim IP_List
IP_List = _
"192.168.30.1,    NSS;" & _
"192.168.30.2,    Integrity_VM;" & _
"192.168.30.3,    Mosaiq;" & _
"192.168.30.4,    XVI;" & _
"192.168.30.5,    iViewGT;" & _
"192.168.30.7,    iGuide;" & _
"192.168.30.16,   UPS;" & _
"192.168.30.17,   NAS;" & _
"192.168.30.150,  TRM_Computer;" & _
"192.168.30.200,  CCP-Management;" & _
"192.168.81.2,    IntelliMax_VM;" & _
"192.168.240.244, Netgear_switch;" & _
"192.168.240.247, VM_access;" & _
"192.168.240.250, NRT_server"
'"192.168.150.1,   Integrity_VM;" & _

'---------------------------------------------------------------------------------
' 3) CREATE OUTPUT FILE NAME (CURRENT FOLDER)
'---------------------------------------------------------------------------------
Dim strSafeDate, strSafeTime, strDateTime, strOutputPath
strSafeDate = Year(Date) & Right("0" & Month(Date), 2) & Right("0" & Day(Date), 2)
strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
strDateTime = strSafeDate & "-" & strSafeTime

strOutputPath = GetCurrentFolder() & "\" & strDateTime & ".txt"

'---------------------------------------------------------------------------------
' 4) CLEAN UP IP_List: REMOVE TABS, CR, LF
'---------------------------------------------------------------------------------
IP_List = Replace(IP_List, vbTab, "")
IP_List = Replace(IP_List, vbCr,   "")
IP_List = Replace(IP_List, vbLf,   "")

'---------------------------------------------------------------------------------
' 5) SPLIT BY ";"
'---------------------------------------------------------------------------------
Dim rawLines, i, lineCount
rawLines = Split(IP_List, ";")

lineCount = 0
For i = 0 To UBound(rawLines)
    If Trim(rawLines(i)) <> "" Then lineCount = lineCount + 1
Next

WScript.Echo "========================================="
WScript.Echo " PARALLEL PING (NO FORCED TIMEOUT)"
WScript.Echo " Number of entries: " & lineCount
WScript.Echo "========================================="

If lineCount = 0 Then
    WScript.Echo "No IP entries found. Exiting."
    WScript.Quit
End If

'---------------------------------------------------------------------------------
' 6) ARRAYS FOR Exec OBJECTS, IP, NAME, STATUS, ETC.
'---------------------------------------------------------------------------------
Dim pingExec(), pingIP(), pingName(), pingDone(), pingStat()
ReDim pingExec(lineCount - 1)
ReDim pingIP  (lineCount - 1)
ReDim pingName(lineCount - 1)
ReDim pingDone(lineCount - 1)
ReDim pingStat(lineCount - 1)

Dim idx
idx = 0

'---------------------------------------------------------------------------------
' 7) START PARALLEL PING FOR EACH IP (3 PACKETS, 300ms TIMEOUT)
'---------------------------------------------------------------------------------
Dim lineStr, parts
For i = 0 To UBound(rawLines)
    lineStr = Trim(rawLines(i))
    If lineStr <> "" Then
        parts = Split(lineStr, ",")
        If UBound(parts) = 1 Then
            pingIP(idx)   = Trim(parts(0))
            pingName(idx) = Trim(parts(1))
            pingDone(idx) = False
            pingStat(idx) = "unknown"
            
            Dim cmdLine
            cmdLine = "cmd /c ping -n 3 -w 300 " & pingIP(idx)
            
            WScript.Echo "[" & idx & "] Starting ping: " & pingIP(idx) & " (" & pingName(idx) & ")"
            Set pingExec(idx) = Ws.Exec(cmdLine)
            
            idx = idx + 1
        Else
            WScript.Echo "Wrong config entry: " & lineStr
        End If
    End If
Next

'---------------------------------------------------------------------------------
' 8) WAIT FOR ALL PROCESSES TO FINISH NATURALLY (NO .Terminate)
'---------------------------------------------------------------------------------
Dim allDone
Do
    allDone = True
    
    For i = 0 To lineCount - 1
        If Not pingDone(i) Then
            If pingExec(i).Status = 0 Then
                ' Process is still running
                allDone = False
            Else
                ' Process finished => parse output
                Dim txt
                txt = LCase(pingExec(i).StdOut.ReadAll)
                
                If (InStr(txt, "unreachable") > 0) Or _
                   (InStr(txt, "timed out") > 0) Or _
                   (InStr(txt, "could not find host") > 0) Then
                    pingStat(i) = "down"
                ElseIf InStr(txt, "reply from") > 0 Then
                    pingStat(i) = "up"
                Else
                    pingStat(i) = "down"
                End If
                
                pingDone(i) = True
                WScript.Echo "[" & i & "] " & pingIP(i) & " (" & pingName(i) & ") => " & pingStat(i)
            End If
        End If
    Next
    
    If allDone Then Exit Do
    WScript.Sleep 200
Loop

WScript.Echo vbCrLf & "========================================="
WScript.Echo "  ALL PINGS FINISHED"
WScript.Echo "========================================="

'---------------------------------------------------------------------------------
' 9) WRITE RESULTS TO A TEXT FILE WITH ALIGNED COLUMNS
'---------------------------------------------------------------------------------
Dim objTextOut
Set objTextOut = fso.CreateTextFile(strOutputPath, True)

' Print column headers
' We'll use fixed-width fields: IP(18), Status(8), Device(18), DateTime(20)
objTextOut.WriteLine _
    FixedWidth("IP", 18) & _
    FixedWidth("Status", 8) & _
    FixedWidth("Device", 18) & _
    FixedWidth("Date/Time", 20)

' Write each row
Dim nowString
For i = 0 To lineCount - 1
    nowString = FormatDateTime(Now, vbShortDate) & " " & FormatDateTime(Now, vbLongTime)
    objTextOut.WriteLine _
        FixedWidth(pingIP(i), 18) & _
        FixedWidth(pingStat(i), 8) & _
        FixedWidth(pingName(i), 18) & _
        FixedWidth(nowString, 20)
Next

' Local IP info
objTextOut.WriteLine ""
objTextOut.WriteLine GetLocalIPList()
objTextOut.Close

WScript.Echo vbCrLf & "Results saved to: " & strOutputPath
Ws.Run "notepad.exe """ & strOutputPath & """"

'---------------------------------------------------------------------------------
' 10) WAIT FOR ENTER SO THE CONSOLE STAYS OPEN
'---------------------------------------------------------------------------------
'WScript.Echo vbCrLf & "Press ENTER to close this console..."
'Dim dummy
'dummy = WScript.StdIn.ReadLine

'======================== SUPPORT FUNCTIONS ===============================

Function GetCurrentFolder()
    Dim tmpFSO
    Set tmpFSO = CreateObject("Scripting.FileSystemObject")
    GetCurrentFolder = tmpFSO.GetAbsolutePathName(".")
End Function

Function GetLocalIPList()
    Dim objWMIService, IPConfigSet, strComputer, strMsg, IPConfig, k
    strComputer = "."
    strMsg = "local IP addresses list:"
    
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set IPConfigSet   = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
    
    For Each IPConfig In IPConfigSet
        If Not IsNull(IPConfig.IPAddress) Then
            For k = LBound(IPConfig.IPAddress) To UBound(IPConfig.IPAddress)
                If InStr(IPConfig.IPAddress(k), ":") = 0 Then
                    strMsg = strMsg & vbCrLf & IPConfig.IPAddress(k)
                End If
            Next
        End If
    Next
    
    GetLocalIPList = strMsg
End Function

Function FixedWidth(ByVal strValue, ByVal colWidth)
    ' Truncate or pad 'strValue' to a fixed width of 'colWidth'
    If Len(strValue) > colWidth Then
        FixedWidth = Left(strValue, colWidth)
    Else
        FixedWidth = strValue & Space(colWidth - Len(strValue))
    End If
End Function
