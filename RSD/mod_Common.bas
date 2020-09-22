Attribute VB_Name = "mod_Common"
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetTickCount& Lib "kernel32" ()
'For List View
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LB_SETHORIZONTALEXTENT = &H194

Public FileSize As Long
Public SharedFileNameURL  As String
Public SharedFileName As String
Public Bool1 As Boolean
Public Bool2 As Boolean

Public SharedGetFileExt As String

Public SecondServer As String
Public ThirdServer As String

Public tmp1, tmp2, tmp3, tmp4, tmp5 As String
Public tmp6, tmp7, tmp8, tmp9, tmp10 As String

Dim tmp11, tmp12, tmp13 As String
Dim tmpsize1, tmpsize2 As String

Public tmpsize3 As String
Public sTimer4

Public bDone As Boolean
Public Cnt As Integer
Public Cnt1 As Integer
Public sWait15Min As String

Sub GetInfo1(Inet1 As Inet, URL As String)
    Dim Res As String
    Dim Res1 As String
    
    Res = Inet1.OpenURL(URL)
           
    Do While Inet1.StillExecuting
        DoEvents
    Loop
    
    'For File Size
    If InStr(Res, "/files/") Then
        strpos3 = InStr(Res, "/files/")
        tmpsize1 = Mid(Res, InStr(1, Res, "|") + 1)
        tmpsize2 = Left(tmpsize1, InStr(1, tmpsize1, "KB") + 1)
        tmpsize3 = Replace(tmpsize2, "KB", "")
        'Debug.Print Trim(tmpsize3)
    End If
    
    If InStr(Res, "<form action=") Then
        strpos1 = InStr(Res, "<form action=")
        tmp1 = Mid(Res, InStr(1, Res, "<form action=") + 1)
        tmp2 = Mid(tmp1, 14, Len(Res))
        tmp3 = Mid(tmp2, 1, InStr(1, tmp2, Chr(&H22)) - 1)
        
        'Second Server Name
        tmp4 = Mid(tmp3, 8, InStr(1, tmp3, "/files") - 8)
        SecondServer = tmp4
        'Debug.Print tmp4
        
        'Original file URL
        SharedFileNameURL = Trim(tmp3)
        'Debug.Print Trim(tmp3)
        
        'For Posting Value from /files....
        tmp5 = Mid(tmp3, InStr(1, tmp3, ".com") + 4)
        'Debug.Print tmp5
        
        'Zip/Rar file Name
        tmp4 = Mid(tmp3, InStr(40, tmp3, "/") + 1)
        SharedFileName = Trim(tmp4)
        'Debug.Print SharedFileName
        
        'Save File Name
        SharedGetFileExt = Left(tmp4, Len(tmp4) + 3)
        'Debug.Print SharedGetFileExt
    End If
    
    If InStr(Res, "already downloading a file") Then
        MsgBox ""
    End If
    
    Call GetInfo2(frmMain.Inet2)
End Sub

Sub GetInfo2(Inet2 As Inet)
    Dim Res1
    
    Inet2.Execute "http://" & SecondServer & tmp5, "POST", "dl.start=Free", "Content-Type: application/x-www-form-urlencoded"
    
    Do While Inet2.StillExecuting
        DoEvents
    Loop
    
    Res1 = Inet2.GetChunk(8024, icString)
    
    'For Getting Second Download Wait time
     If InStr(Res1, "try again in about") Then
        tmp0 = Mid(Res1, InStr(1, Res1, "try again in about") + 2)
        tmp1 = Mid(tmp0, 10, Len(Res1))
        tmp2 = Mid(tmp1, 1, InStr(1, tmp1, ".") - 1)
        tmp3 = Left(tmp2, Len(tmp2) - 7)
        tmp4 = Right(tmp3, Len(tmp3) - 7)
        sWait15Min = tmp4
        frmMain.LblWait2.Caption = "Limit Reached! Please wait for " & sWait15Min & " Min."
     End If
    'End Proc
    
    'For Getting Page Counter Timer Function
    If InStr(Res1, "Still") Then
        sTimer1 = Mid(Res1, InStr(1, Res1, "var c=") + 1)
        sTimer2 = Mid(sTimer1, 6, Len(Res1))
        sTimer3 = Mid(sTimer2, 1, InStr(1, sTimer2, ";") - 1)
        sTimer4 = Trim(sTimer3)
        'MsgBox sTimer4
    End If
    'End Proc
    
    'For Getting Current Load
    If InStr(Res1, "already downloading a file") Then
       frmMain.lblRapidStatus.Caption = "Already Downloading a File!!!"
    Else
    
        If InStr(Res1, "<form name=") Then
            tmp6 = Mid(Res1, InStr(1, Res1, "<form name=") + 1)
            tmp7 = Mid(tmp6, 25, Len(Res1))
            tmp8 = Mid(tmp7, 1, InStr(1, tmp7, Chr(&H22)) - 1)
            
            'Third Server Founded
            ThirdServer = tmp8
            Debug.Print ThirdServer
            
            'Form Timer For Download
            frmMain.Timer1.Enabled = True
            Cnt = sTimer4
            'End Timer Routine
        End If
    End If
End Sub

Sub DownloadCreate()
    Dim FileNumber As Integer
    Dim FileData() As Byte
    Dim FileSize  As Long
    Dim FileRemaining As Long
    
    Dim t As Long
    Dim StartT As Long
    Dim spRate As Single
    
    frmMain.Inet3.Execute ThirdServer, "POST", "mirror=on&x=44&y=34", "Content-Type: application/x-www-form-urlencoded"
        
    Do While frmMain.Inet3.StillExecuting
        DoEvents
    Loop
       
    'FileSize = tmpsize3 * 1024
    FileSize = frmMain.Inet3.GetHeader("Content-Length")
        Sz = FileSize / 1000
        frmMain.lblSize.Caption = Sz & " KB"
    FileRemaining = FileSize
    FileSize_Current = 0
   
    Debug.Print FileSize
    
    
    FileNumber = FreeFile
        
    Open App.Path & "/" & SharedGetFileExt For Binary Access Write As #FileNumber
    
    StartT = GetTickCount
    
    Do Until FileRemaining = 0
        If frmMain.Tag = "Cancel" Then
            frmMain.Inet3.Cancel
            Exit Sub
        End If
        
        If FileRemaining > 1024 Then
            FileData = frmMain.Inet3.GetChunk(1024, icByteArray)
            FileRemaining = FileRemaining - 1024
        Else
            FileData = frmMain.Inet3.GetChunk(FileRemaining, icByteArray)
            FileRemaining = 0
        End If
        
        FileSize_Current = FileSize - FileRemaining
        PBValue = CInt((100 / FileSize) * FileSize_Current)
        frmMain.lblSaved.Caption = FileSize_Current / 1000 & " KB"
        frmMain.lblRemaining.Caption = (FileSize - FileSize_Current) / 1000 & " KB"
        frmMain.lblPercentage.Caption = "% " & PBValue
        frmMain.StPanel.Panels(2).Text = PBValue & "%"
        frmMain.PB1.Value = PBValue
                
        If FileSize_Current <> 0 Then
           t = GetTickCount - StartT
           If t <> 0 Then
                spRate = (spRate + ((FileSize_Current / 1000) / (t / 1000))) / 2
                frmMain.lblSpeed.Caption = Format(spRate, "#.##") & " KBPS"
                'Time Calc Goes Here
                    EstimateDownloadTimeCalc (FileSize - FileSize_Current) / 1000, spRate, frmMain.lblTime
                    EstimateDownloadTimeCalc FileSize_Current / 1000, spRate, frmMain.lblTakeTime
                'End Time Calculation
           End If
        End If
        Put #FileNumber, , FileData
    Loop
    Close #FileNumber
    bDone = True
End Sub

Function EstimateDownloadTimeCalc(ByVal Size As String, ByVal Speed As String, ByVal EstimateLabel As Label) As String
On Error Resume Next
Dim time As Long
Dim hrs As Long
Dim mins As Long
Dim days As Long

Dim ttime As String
Dim thrs As String
Dim tmins As String
Dim tdays As String

time = Size / Speed

repeat:
If time >= 60 Then
    If time >= 86400 Then
        days = days + 1
        time = time - 86400
        GoTo repeat
    ElseIf time >= 3600 Then
        hrs = hrs + 1
        time = time - 3600
        GoTo repeat
    Else
        mins = mins + 1
        time = time - 60
        GoTo repeat
    End If
End If

If days = 0 Then
    tdays = ""
Else
    tdays = days & " Days, "
End If
If hrs = 0 Then
    thrs = ""
Else
    thrs = hrs & " Hours, "
End If
If mins = 0 Then
    tmins = ""
Else
    tmins = mins & " Minutes and "
End If
    
ttime = time & " Seconds."
EstimateLabel.Caption = tdays & thrs & tmins & ttime
End Function

Function GetStatus(st As Integer, Inet2 As Inet)
    Select Case st
        Case icError
            GetStatus = Left$(Inet2.ResponseInfo, _
            Len(Inet2.ResponseInfo) - 2)
        Case icResolvingHost, icRequesting, icRequestSent
            GetStatus = "Searching... "
        Case icHostResolved
            GetStatus = "Found." & vName
        Case icReceivingResponse, icResponseReceived
            GetStatus = "Receiving data "
        Case icResponseCompleted
            GetStatus = "Connected"
        Case icConnecting, icConnected
            GetStatus = "Connecting..."
        Case icDisconnecting
            GetStatus = "Disconnecting..."
        Case icDisconnected
            GetStatus = "Disconnected"
        Case Else
    End Select
End Function

