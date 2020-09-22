VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schedule List For Downloading..."
   ClientHeight    =   5430
   ClientLeft      =   10560
   ClientTop       =   3705
   ClientWidth     =   6495
   Icon            =   "frmList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6495
   Begin VB.CommandButton CmdClear 
      Caption         =   "&Clear List"
      Height          =   375
      Left            =   5445
      TabIndex        =   10
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   420
      Left            =   5445
      TabIndex        =   8
      Top             =   4455
      Width           =   960
   End
   Begin VB.CommandButton CmdOpen 
      Caption         =   "&Open"
      Height          =   420
      Left            =   5445
      TabIndex        =   7
      Top             =   3960
      Width           =   960
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5715
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton CmdListDown 
      Caption         =   "Down"
      Height          =   375
      Left            =   5445
      TabIndex        =   4
      Top             =   2925
      Width           =   960
   End
   Begin VB.CommandButton CmdListUP 
      Caption         =   "Up"
      Height          =   375
      Left            =   5445
      TabIndex        =   3
      Top             =   2430
      Width           =   960
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   5445
      TabIndex        =   2
      Top             =   585
      Width           =   975
   End
   Begin VB.ListBox List1 
      Height          =   4545
      Left            =   90
      TabIndex        =   6
      Top             =   450
      Width           =   5265
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   5445
      TabIndex        =   1
      Top             =   90
      Width           =   975
   End
   Begin VB.TextBox txtURL 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   630
      TabIndex        =   0
      Top             =   90
      Width           =   4725
   End
   Begin VB.Label LblTotalFiles 
      Caption         =   "Total Files"
      Height          =   240
      Left            =   135
      TabIndex        =   9
      Top             =   5085
      Width           =   2040
   End
   Begin VB.Label Label1 
      Caption         =   "URL"
      Height          =   255
      Left            =   225
      TabIndex        =   5
      Top             =   120
      Width           =   465
   End
End
Attribute VB_Name = "frmList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FileNumber As Long

Private Sub CmdAdd_Click()
    If txtURL = "" Then Exit Sub
    If InStr(txtURL.Text, "rapidshare.com") Then
        List1.AddItem txtURL.Text
            On Error Resume Next
            Open App.Path & "\downloads.txt" For Append As #1
                Print #1, Trim(txtURL.Text)
            Close #1
        txtURL.Text = ""
    Else
        MsgBox "Please Use Only Rapidshare links"
    End If
End Sub

Private Sub CmdClear_Click()
    If MsgBox("Are you Sure want to clear the hole lists", vbYesNo) = vbYes Then
        List1.Clear
    End If
End Sub

Private Sub CmdDelete_Click()
    If List1.ListIndex = -1 Then Exit Sub
    List1.RemoveItem List1.ListIndex
    Call ChangeListFile
    LblTotalFiles.Caption = "Total Download List " & List1.ListCount
End Sub


Private Sub CmdListDown_Click()
Dim strTemp As String
Dim Count As Integer

Count = List1.ListIndex

If Count > -1 Then
    strTemp = List1.List(Count)
    List1.AddItem strTemp, (Count + 2)
    List1.RemoveItem (Count)
    List1.Selected(Count + 1) = True
End If
End Sub

Private Sub CmdListUP_Click()
Dim strTemp As String
Dim Count As Integer

Count = List1.ListIndex

If Count > -1 Then
    strTemp = List1.List(Count)
    List1.AddItem strTemp, (Count - 1)
    List1.RemoveItem (Count + 1)
    List1.Selected(Count - 1) = True
End If
End Sub

Private Sub CmdOpen_Click()
    Dim strData As String
    
    CD1.Filter = "Text Files (*.txt)|*.txt|"
    CD1.ShowOpen
    CD1.Flags = &H4
    CD1.CancelError = False
    
    On Error Resume Next
    Open CD1.FileName For Input As #1
        Do While Not EOF(1)
            Line Input #1, strData
            List1.AddItem Trim(strData)
        Loop
    Close #1
    
    Dim i As Integer, intGreatestLen As Integer, lngGreatestWidth As Long
    For i = 0 To List1.ListCount - 1
        If Len(List1.List(i)) > Len(List1.List(intGreatestLen)) Then
            intGreatestLen = i
        End If
    Next i
    lngGreatestWidth = List1.Parent.TextWidth(List1.List(intGreatestLen) + Space(1))
    lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
    SendMessage List1.hwnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0

    LblTotalFiles.Caption = "Total Download List " & List1.ListCount
End Sub

Private Sub CmdSave_Click()
    Call ChangeListFile
End Sub

Sub Form_Load()
    Dim strData As String
    On Error Resume Next
    Open App.Path & "\downloads.txt" For Input As #1
        Do While Not EOF(1)
            Line Input #1, strData
            List1.AddItem Trim(strData)
        Loop
    Close #1
    
    Dim i As Integer, intGreatestLen As Integer, lngGreatestWidth As Long
    For i = 0 To List1.ListCount - 1
        If Len(List1.List(i)) > Len(List1.List(intGreatestLen)) Then
            intGreatestLen = i
        End If
    Next i
    lngGreatestWidth = List1.Parent.TextWidth(List1.List(intGreatestLen) + Space(1))
    lngGreatestWidth = lngGreatestWidth \ Screen.TwipsPerPixelX
    SendMessage List1.hwnd, LB_SETHORIZONTALEXTENT, lngGreatestWidth, 0

    LblTotalFiles.Caption = "Total Download List " & List1.ListCount
End Sub

Sub ChangeListFile()
    On Error Resume Next
    Open App.Path & "\downloads.txt" For Output As #1
        For i = 0 To List1.ListCount - 1
            Print #1, Trim(List1.List(i))
        Next
    Close #1
End Sub

Private Sub List1_Click()
    If List1.SelCount > 1 Then
        CmdListDown.Enabled = False
        CmdListUP.Enabled = False
        Exit Sub
    Else
        CmdListDown.Enabled = True
        CmdListUP.Enabled = True
    End If
    If List1.Selected(0) Then
        CmdListUP.Enabled = False
    Else
        CmdListUP.Enabled = True
    End If
    If List1.Selected(List1.ListCount - 1) Then
        CmdListDown.Enabled = False
    Else
        CmdListDown.Enabled = True
    End If
End Sub

Private Sub List1_DblClick()
    txtURL.Text = List1.Text
    frmMain.txtURL = List1.Text
End Sub
