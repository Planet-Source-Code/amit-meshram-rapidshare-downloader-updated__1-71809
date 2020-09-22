VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rapidshare Downloader..."
   ClientHeight    =   3465
   ClientLeft      =   1875
   ClientTop       =   4635
   ClientWidth     =   8175
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   8175
   Begin VB.CommandButton Command1 
      Caption         =   "Scheduler"
      Height          =   390
      Left            =   6570
      TabIndex        =   21
      Top             =   1635
      Width           =   1515
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7380
      Top             =   810
   End
   Begin InetCtlsObjects.Inet Inet2 
      Left            =   8400
      Top             =   990
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   9000
      Top             =   990
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel Download"
      Height          =   390
      Left            =   5040
      TabIndex        =   12
      Top             =   1635
      Width           =   1515
   End
   Begin VB.CommandButton CmdDownload 
      Caption         =   "Start Download"
      Height          =   390
      Left            =   3480
      TabIndex        =   11
      Top             =   1635
      Width           =   1530
   End
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   3495
      TabIndex        =   9
      Top             =   540
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StPanel 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9137
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2646
            MinWidth        =   2646
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "2:12 PM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   2025
      TabIndex        =   1
      Text            =   "http://rapidshare.com/files/180564572/7zip_www.sxforum.org.rar"
      Top             =   120
      Width           =   6060
   End
   Begin InetCtlsObjects.Inet Inet3 
      Left            =   9570
      Top             =   990
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label LblWait2 
      Height          =   255
      Left            =   5310
      TabIndex        =   23
      Top             =   2085
      Width           =   2775
   End
   Begin VB.Label lblWait1 
      Height          =   255
      Left            =   2790
      TabIndex        =   22
      Top             =   2085
      Width           =   2460
   End
   Begin VB.Label Label8 
      Caption         =   "Total Time"
      Height          =   240
      Left            =   135
      TabIndex        =   20
      Top             =   2790
      Width           =   1095
   End
   Begin VB.Label lblTakeTime 
      Caption         =   "0 Days, 0 Hours, 0 Minutes and 0 Seconds."
      Height          =   255
      Left            =   2025
      TabIndex        =   19
      Top             =   2790
      Width           =   6060
   End
   Begin VB.Label lblTime 
      Caption         =   "0 Days, 0 Hours, 0 Minutes and 0 Seconds."
      Height          =   255
      Left            =   2025
      TabIndex        =   18
      Top             =   2430
      Width           =   6060
   End
   Begin VB.Label Label5 
      Caption         =   "Time Remaining"
      Height          =   240
      Left            =   135
      TabIndex        =   17
      Top             =   2430
      Width           =   1725
   End
   Begin VB.Label lblRapidStatus 
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   3480
      TabIndex        =   16
      Top             =   1260
      Width           =   4590
   End
   Begin VB.Label Label3 
      Caption         =   "Speed"
      Height          =   240
      Left            =   150
      TabIndex        =   15
      Top             =   1710
      Width           =   1095
   End
   Begin VB.Label lblSpeed 
      Caption         =   "In bits"
      Height          =   240
      Left            =   2010
      TabIndex        =   14
      Top             =   1710
      Width           =   1365
   End
   Begin VB.Label lblWait 
      Caption         =   "Wait : "
      Height          =   240
      Left            =   150
      TabIndex        =   13
      Top             =   2100
      Width           =   2595
   End
   Begin VB.Label lblPercentage 
      Caption         =   "Persent % Completed..."
      Height          =   240
      Left            =   3510
      TabIndex        =   10
      Top             =   900
      Width           =   3285
   End
   Begin VB.Label lblRemaining 
      Caption         =   "In bits"
      Height          =   285
      Left            =   2010
      TabIndex        =   8
      Top             =   1320
      Width           =   1365
   End
   Begin VB.Label Label6 
      Caption         =   "File Remaining"
      Height          =   285
      Left            =   150
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label lblSaved 
      Caption         =   "In bits"
      Height          =   285
      Left            =   2010
      TabIndex        =   6
      Top             =   930
      Width           =   1365
   End
   Begin VB.Label Label4 
      Caption         =   "File Saved"
      Height          =   285
      Left            =   150
      TabIndex        =   5
      Top             =   930
      Width           =   1095
   End
   Begin VB.Label lblSize 
      Caption         =   "In bits"
      Height          =   285
      Left            =   2010
      TabIndex        =   3
      Top             =   540
      Width           =   1365
   End
   Begin VB.Label Label2 
      Caption         =   "File Size"
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Rapidshare URL"
      Height          =   255
      Left            =   150
      TabIndex        =   0
      Top             =   150
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub TerminateTimer()
    Timer1.Enabled = False
End Sub

Sub CmdCancel_Click()
    Inet1.Cancel
    Inet2.Cancel
    Inet3.Cancel
    
    If frmMain.Tag = "Cancel" Then
        Inet1.Cancel
        Inet2.Cancel
        Inet3.Cancel
    End If
    Timer1.Enabled = False
    lblWait.Caption = "Action : Cancelled"
    Cnt30s = 1
End Sub

Private Sub CmdDownload_Click()
    Call GetInfo1(Inet1, Trim(txtURL.Text))
    'Timer1.Enabled = True
    'Cnt = sTimer4
End Sub

Private Sub Command1_Click()
    frmList.Show
End Sub

Private Sub Form_Load()
    bDone = False
End Sub

Private Sub Timer1_Timer()
        Cnt = Cnt - 1
        
        lblWait.Caption = "Download Will Start In : " & Cnt & " Seconds"
        'lblWait1.Caption = "Please Wait For : " & sTimer4 & " Seconds"
    
        If Cnt = 0 Then 'sTimer4 Then
            lblWait.Caption = "File Downloading Started..."
            Call DownloadCreate
            Cnt = 0
            Timer1.Enabled = False
        End If
End Sub

Private Sub Inet3_StateChanged(ByVal State As Integer)
    StPanel.Panels(1).Text = GetStatus(State, Inet3)
End Sub


