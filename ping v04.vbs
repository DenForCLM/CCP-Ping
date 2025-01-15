' I usually put this script in the C:\Temp folder. A new log file is created there each time. Denis Kiselev.
Option Explicit
Dim strInputPath,strOutputPath,strStatus,strSafeDate,strSafeTime,strDateTime,Title,MsgTitle,MsgWaiting, IP_List
Dim objFSO,objTextIn,objTextOut,ReadAllFile,Lines,Line, LineIP, Ltmp, LineName,Ws,Command,OpenCSVFile,oExec,Temp,StartTime,DurationTime

IP_List = _ 
"192.168.30.1,     NSS              ;" &_  
"192.168.30.2,     Integrity_VM     ;" &_
"192.168.30.3,     Mosaiq           ;" &_
"192.168.30.4,     XVI              ; "&_
"192.168.30.5,     iViewGT          ; "&_
"192.168.30.7,     iGuide           ; "&_
"192.168.30.16,    UPS              ; "&_
"192.168.30.17,    NAS              ; "&_
"192.168.30.150,   TRM_Computer     ; "&_ 
"192.168.30.200,   CCP-Management   ; "&_
"192.168.81.2,     IntelliMax_VM    ; "&_
"192.168.150.1,    Integrity_VM     ; "&_ 
"192.168.240.244,  Netgear_switch   ; "&_
"192.168.240.247,  VM_access        ; "&_
"192.168.240.250,  NRT_server        "
'"10.0.0.1,         test "


Set Ws = CreateObject("WScript.Shell")
Title = "Ping list of servers"
MsgTitle = Title
MsgWaiting = "Please wait ... the pinging is on progress ...."
Temp = ws.ExpandEnvironmentStrings("%Temp%")
strSafeDate = DatePart("yyyy",Date) & Right("0" & DatePart("m",Date), 2) & Right("0" & DatePart("d",Date), 2)
strSafeTime = Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
strDateTime = strSafeDate & "-" & strSafeTime
strOutputPath = GetCurrentFolder() & "\" & strDateTime & ".txt" '- location of output
set objFSO = CreateObject("Scripting.FileSystemObject")
set objTextOut = objFSO.CreateTextFile(strOutputPath)
objTextOut.WriteLine("     IP         STATUS                               DATE")
'ReadAllFile = objTextIn.ReadAll

IP_List = Replace(IP_List," ","",1,-1)
IP_List = Replace(IP_List,"vbTab","",1,-1)
IP_List = Replace(IP_List,"vbCr","",1,-1)
IP_List = Replace(IP_List,"vbLf","",1,-1)
Lines = Split(IP_List,";")
Call CreateProgressBar(MsgTitle,MsgWaiting) 'Create the waiting Bar
Call LaunchProgressBar()                    'Launch the progress bar
StartTime = Timer                           'Start  Timer  
For Each Line In Lines
    Ltmp = Split(Line,",")
    'Ws.Popup "Ltmp: "& Line,"5",MsgTitle,64
    If ubound(Ltmp) = 1 Then
        LineIP = Ltmp(0)
        LineName = Ltmp(1)
        If OnLine(LineIP) = True Then
            strStatus = "  up   "
        else
            strStatus = "down   "
        end if
        objTextOut.WriteLine(LineIP & ";" & Space(16 - Len(LineIP)) & strStatus & LineName & "; " & Space(20 - Len(LineName))  & Now)
    else
        objTextOut.WriteLine("Wrong string in config")
    End if
Next
objTextOut.WriteLine(vbcrlf & GetLocalIPList)  ' add local IPs
Call CloseProgressBar()                                             'Closing the progess Bar
DurationTime = FormatNumber(Timer - StartTime, 0) & " seconds."     'The duration of script execution
'Command = "cmd /c CD " & DblQuote(ExcelPath()) & " | Start Excel.exe" &" /E "& DblQuote(strOutputPath)
Ws.Popup "The pinging Script is finshed in "& DurationTime,"3",MsgTitle,64
Ws.run "notepad.exe " & DblQuote(strOutputPath)
'OpenCSVFile = ws.run(Command,0,False)

'************************************************************************************************************************************************************
Function OnLine(strHost)
    Dim objPing,z,objRetStatus,PingStatus
    Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("select * from Win32_PingStatus where address = '" & strHost & "'")
    z = 0
    Do   
        z = z + 1
        For Each objRetStatus In objPing
            If IsNull(objRetStatus.StatusCode) Or objRetStatus.StatusCode <> 0 Then
                PingStatus = False
            Else
                PingStatus = True
            End If     
        Next   
        Call Pause(0.2) '<<<<<<<<<<<<<<<<<
        If z = 3 Then Exit Do 'here you can incerase or decerase the value of z = 5
    Loop until PingStatus = True
    If PingStatus = True Then
        OnLine = True
    Else
        OnLine = False
    End If
End Function
'*********************************************************************************************
Sub Pause(NSeconds)
    Wscript.Sleep(NSeconds*1000)
End Sub
'**********************************************************************************************
Function ExcelPath()
    Dim appXL,s
    Set appXL = CreateObject("Excel.Application")
    ExcelPath = appXL.Path
    appXL.Quit
    Set appXL = Nothing
End Function
'**********************************************************************************************
Function DblQuote(Str)
    DblQuote = Chr(34) & Str & Chr(34)
End Function
'**********************************************************************************************
Sub CreateProgressBar(Title,MsgWaiting)
    Dim ws,fso,f,f2,ts,ts2,Line,i,fread,ReadAll,NbLineTotal,Temp,PathOutPutHTML,fhta,oExec
    Set ws = CreateObject("wscript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Temp = WS.ExpandEnvironmentStrings("%Temp%")
    PathOutPutHTML = Temp & "\Barre.hta"
    Set fhta = fso.OpenTextFile(PathOutPutHTML,2,True)
    fhta.WriteLine "<HTML>"
    fhta.WriteLine "<HEAD>"
    fhta.WriteLine "<Title>  " & Title & "</Title>"
    fhta.WriteLine "<HTA:APPLICATION"
    fhta.WriteLine "ICON = ""magnify.exe"" "
    fhta.WriteLine "BORDER=""THIN"" "
    fhta.WriteLine "INNERBORDER=""NO"" "
    fhta.WriteLine "MAXIMIZEBUTTON=""NO"" "
    fhta.WriteLine "MINIMIZEBUTTON=""NO"" "
    fhta.WriteLine "SCROLL=""NO"" "
    fhta.WriteLine "SYSMENU=""NO"" "
    fhta.WriteLine "SELECTION=""NO"" "
    fhta.WriteLine "SINGLEINSTANCE=""YES"">"
    fhta.WriteLine "</HEAD>"
    fhta.WriteLine "<BODY text=""10bed2""><CENTER>"
    fhta.WriteLine "<marquee DIRECTION=""LEFT"" SCROLLAMOUNT=""3"" BEHAVIOR=ALTERNATE><font face=""Comic sans MS"">" & MsgWaiting &"</font></marquee>"
    fhta.WriteLine "<br><img src=""data:image/gif;base64,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"" />"
    fhta.WriteLine "</CENTER></BODY></HTML>"
    fhta.WriteLine "<SCRIPT LANGUAGE=""VBScript""> "
    fhta.WriteLine "Set ws = CreateObject(""wscript.Shell"")"
    fhta.WriteLine "Temp = WS.ExpandEnvironmentStrings(""%Temp%"")"
    fhta.WriteLine "Sub window_onload()"
    fhta.WriteLine "    CenterWindow 350,100"
    fhta.WriteLine "    Self.document.bgColor = ""00677d"" "
    fhta.WriteLine " End Sub"
    fhta.WriteLine " Sub CenterWindow(x,y)"
    fhta.WriteLine "    Dim iLeft,itop"
    fhta.WriteLine "    window.resizeTo x,y"
    fhta.WriteLine "    iLeft = window.screen.availWidth/2 - x/2"
    fhta.WriteLine "    itop = window.screen.availHeight/2 - y/2"
    fhta.WriteLine "    window.moveTo ileft,itop"
    fhta.WriteLine "End Sub"
    fhta.WriteLine "</script>"
    fhta.close
End Sub
'**********************************************************************************************
Sub LaunchProgressBar()
    Set oExec = Ws.Exec("mshta.exe " & Temp & "\Barre.hta")
End Sub
'**********************************************************************************************
Sub CloseProgressBar()
    oExec.Terminate
End Sub
'**********************************************************************************************
Function GetCurrentFolder()
    Dim FSO
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetCurrentFolder= FSO.GetAbsolutePathName(".")
End Function
'**********************************************************************************************
Function GetLocalIPList()
    DIM objWMIService, IPConfigSet, strComputer, strMsg, IPConfig, i
    strComputer = "." 
    strMsg = "local IP addresses list:" 
    Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2") 
    Set IPConfigSet = objWMIService.ExecQuery _ 
        ("Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE") 
    For Each IPConfig in IPConfigSet 
        If Not IsNull(IPConfig.IPAddress) Then  
            For i=LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
                If Not Instr(IPConfig.IPAddress(i), ":") > 0 Then
                    strMsg = strMsg & vbcrlf & IPConfig.IPAddress(i) 
                    'WScript.Echo IPConfig.IPAddress(i) 
                End If
        Next 
        End If 
    Next 

    GetLocalIPList = strMsg
End Function
'**********************************************************************************************