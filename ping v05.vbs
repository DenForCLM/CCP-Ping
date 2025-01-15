' I usually put this script in the C:\Temp folder. A new log file is created there each time. Denis Kiselev.
Option Explicit

Dim strInputPath, strOutputPath, strStatus, strSafeDate, strSafeTime, strDateTime
Dim Title, MsgTitle, MsgWaiting, IP_List
Dim objFSO, objTextOut, ReadAllFile, Lines, Line, LineIP, Ltmp, LineName
Dim Ws, Command, OpenCSVFile, Temp, StartTime, DurationTime

IP_List = _ 
"192.168.30.1,     NSS              ;" &_  
"192.168.30.2,     Integrity_VM     ;" &_
"192.168.30.3,     Mosaiq           ;" &_
"192.168.30.4,     XVI              ;" &_
"192.168.30.5,     iViewGT          ;" &_
"192.168.30.7,     iGuide           ;" &_
"192.168.30.16,    UPS              ;" &_
"192.168.30.17,    NAS              ;" &_
"192.168.30.150,   TRM_Computer     ;" &_
"192.168.30.200,   CCP-Management   ;" &_
"192.168.81.2,     IntelliMax_VM    ;" &_
"192.168.150.1,    Integrity_VM     ;" &_
"192.168.240.244,  Netgear_switch   ;" &_
"192.168.240.247,  VM_access        ;" &_
"192.168.240.250,  NRT_server       "
'"10.0.0.1,         test "


Set Ws = CreateObject("WScript.Shell")
Title = "Ping list of servers"
MsgTitle = Title
MsgWaiting = "Please wait ... the pinging is on progress ...."

' -- Show a popup window at start:
Ws.Popup MsgWaiting, 2, MsgTitle, 64  ' 2 seconds waiting

Temp = Ws.ExpandEnvironmentStrings("%Temp%")

' -- Generate date+time for the output file name:
strSafeDate = DatePart("yyyy", Date) & Right("0" & DatePart("m", Date), 2) & Right("0" & DatePart("d", Date), 2)
strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
strDateTime = strSafeDate & "-" & strSafeTime

' -- The folder to write the file to. By default, the current folder (you can specify "C:\Temp\" & strDateTime & ".txt"):
strOutputPath = GetCurrentFolder() & "\" & strDateTime & ".txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextOut = objFSO.CreateTextFile(strOutputPath)

objTextOut.WriteLine("     IP         STATUS                               DATE")

' Remove actual tab, CR, LF characters.
' If you want to remove extra spaces - use Trim or Replace, but I'm showing basically:

IP_List = Replace(IP_List, vbTab, "")
IP_List = Replace(IP_List, vbCr,   "")
IP_List = Replace(IP_List, vbLf,   "")

' -- Split into an array by ";"
Lines = Split(IP_List, ";")

StartTime = Timer

' -- Main loop
Dim i
For i = 0 To UBound(Lines)
    Line = Trim(Lines(i))       ' remove extra spaces around
    If Line <> "" Then
        Ltmp = Split(Line, ",")
        
        If UBound(Ltmp) = 1 Then
            LineIP   = Trim(Ltmp(0))
            LineName = Trim(Ltmp(1))
            
            ' Call the new OnLine function (ping via cmd with output check)
            If OnLine(LineIP) Then
                strStatus = "up"
            Else
                strStatus = "down"
            End If
            
            ' Write the line to the log.
            objTextOut.WriteLine(LineIP & ";" & Space(16 - Len(LineIP)) & _
                                 strStatus & Space(6 - Len(strStatus)) & _
                                 LineName & "; " & Space(20 - Len(LineName))  & Now)
        Else
            ' -- (No.3) detailed message:
            objTextOut.WriteLine("Wrong string in config: " & Line)
        End If
    End If
Next

objTextOut.WriteLine(vbCrLf & GetLocalIPList)

DurationTime = FormatNumber(Timer - StartTime, 0) & " seconds."
Ws.Popup "The pinging Script is finished in " & DurationTime, 3, MsgTitle, 64

' -- Open the result in Notepad
Ws.Run "notepad.exe " & DblQuote(strOutputPath)

'=================== Functions below ===========================

'  simplified ping function (cmd /c ping)
Function OnLine(strHost)
    Dim result, pingOutput, tempFile, tfile

    tempFile = Ws.ExpandEnvironmentStrings("%Temp%") & "\ping_" & Replace(strHost, ".", "_") & ".tmp"
    If objFSO.FileExists(tempFile) Then
        objFSO.DeleteFile tempFile, True
    End If

    ' -n 1 = only 1 packet, -w 300 = 300ms timeout, redirect output to a temp file
    result = Ws.Run("cmd /c ping -n 1 -w 300 " & strHost & " >" & DblQuote(tempFile) & " 2>&1", 0, True)
    
    If result = 0 Then
        If objFSO.FileExists(tempFile) Then
            Set tfile = objFSO.OpenTextFile(tempFile, 1)
            pingOutput = tfile.ReadAll
            tfile.Close
            If (InStr(LCase(pingOutput), "unreachable") > 0) Or _
               (InStr(LCase(pingOutput), "timed out") > 0) Or _
               (InStr(LCase(pingOutput), "could not find host") > 0) Then
                OnLine = False
            Else
                OnLine = True
            End If
            objFSO.DeleteFile tempFile, True
        Else
            OnLine = False
        End If
    Else
        ' if the ping command's exit code is not 0 => no successful reply
        OnLine = False
    End If
End Function


' -- Intermediate pause if needed
Sub Pause(NSeconds)
    WScript.Sleep NSeconds * 1000
End Sub
'**********************************************************************************************
Function ExcelPath()
    Dim appXL
    Set appXL = CreateObject("Excel.Application")
    ExcelPath = appXL.Path
    appXL.Quit
    Set appXL = Nothing
End Function
'**********************************************************************************************
Function DblQuote(Str)
    DblQuote = Chr(34) & Str & Chr(34)
End Function

' -- Function that returns the current folder (where the script is located)
Function GetCurrentFolder()
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetCurrentFolder = fso.GetAbsolutePathName(".")
End Function

' -- Output the local IPv4 addresses list
Function GetLocalIPList()
    Dim objWMIService, IPConfigSet, strComputer, strMsg, IPConfig, i
    strComputer = "."
    strMsg = "local IP addresses list:"

    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
    Set IPConfigSet = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")

    For Each IPConfig In IPConfigSet
        If Not IsNull(IPConfig.IPAddress) Then
            For i = LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
                ' Check that it's IPv4 (and not IPv6 with colons)
                If InStr(IPConfig.IPAddress(i), ":") = 0 Then
                    strMsg = strMsg & vbCrLf & IPConfig.IPAddress(i)
                End If
            Next
        End If
    Next

    GetLocalIPList = strMsg
End Function
'**********************************************************************************************