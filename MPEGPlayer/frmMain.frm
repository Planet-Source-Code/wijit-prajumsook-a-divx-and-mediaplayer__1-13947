VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "MPEGPlayer"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4575
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1455
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4080
      Top             =   960
   End
   Begin MSComDlg.CommonDialog comDlg 
      Left            =   3960
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblFrame 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Total Frm: 00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   1140
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Total Time:00:00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   1470
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Caption         =   "Current Frm: 00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.Image imgNext 
      Height          =   255
      Left            =   3360
      ToolTipText     =   "::- Jump Forward -::"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image imgBack 
      Height          =   255
      Left            =   3000
      ToolTipText     =   "::- Jump Back -::"
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgClose 
      Height          =   255
      Left            =   2520
      ToolTipText     =   "::- Close -::"
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgOpen 
      Height          =   255
      Left            =   1920
      ToolTipText     =   "::- Open -::"
      Top             =   1080
      Width           =   495
   End
   Begin VB.Image imgPause 
      Height          =   255
      Left            =   1440
      ToolTipText     =   "::- Pause -::"
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgStop 
      Height          =   255
      Left            =   960
      ToolTipText     =   "::- Stop -::"
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgPlay 
      Height          =   255
      Left            =   600
      ToolTipText     =   "::- Play -::"
      Top             =   1080
      Width           =   375
   End
   Begin VB.Image imgMin 
      Height          =   255
      Left            =   4200
      ToolTipText     =   "::- Minimize -::"
      Top             =   0
      Width           =   135
   End
   Begin VB.Image imgExit 
      Height          =   255
      Left            =   4320
      ToolTipText     =   "::- Exit Program -::"
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H80000007&
      Caption         =   "lblCaption"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Image imgMain 
      Height          =   1455
      Left            =   0
      Picture         =   "frmMain.frx":030A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cur As Integer
Dim min As Integer
Dim sec As Integer

Private Sub Form_Load()
    Dim result As String
    Dim temp As String * 40
    Dim winDir As String
    
    result = GetWindowsDirectory(temp, Len(temp))
    winDir = Left$(temp, result)
    result = WritePrivateProfileString("MCI", "MPEGVideo", "mciqtz.drv", winDir & "\" & "system.ini")
    lblCaption.Caption = "[ No Media ]"
    ffSpeed = 5
End Sub

Private Sub imgBack_Click()
    Call Rewind
End Sub

Private Sub imgClose_Click()
    Call CloseMovie
    Timer.Enabled = False
    lblTime.Caption = "Total Time: 00:00"
    lblFrame.Caption = "Total Frm: 00"
    lbl.Caption = "Current Frm: 00"
End Sub

Private Sub imgExit_Click()
    Call CloseAll
    End
End Sub

Private Sub imgMain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Call DragForm(Me)
    End If
End Sub

Private Sub imgNext_Click()
    Call FForward
End Sub

Private Sub imgOpen_Click()
    On Error GoTo erropen
    If bPlaying Then
        MsgBox "MPEGPlayer is busy right now!", vbExclamation
        Exit Sub
    End If
    comDlg.Filter = "AVI(DivX) Files|*.avi|MPEG Files|*.mpeg|MPG File|*.mpg|ASF Files|*.asf|VCD DAT Files|*.dat|DVD VOB Files|*.vob|MP3 Files|*.mp3"
    comDlg.CancelError = True
    comDlg.ShowOpen
    
    If comDlg.FileName = "" Or comDlg.FileName = strFileToPlay Then
            
    Else
        strFileToPlay = comDlg.FileName
        strFileToPlay = """" & strFileToPlay & """"
        Call OpenMovie
        Call PlayMovie
        Call TotalTime
        Call TotalFrames
    
        Timer.Enabled = True
    End If
erropen:
End Sub

Private Sub imgPause_Click()
    Call PauseMovie
    Timer.Enabled = False
End Sub

Private Sub imgPlay_Click()
    Call PlayMovie
    Timer.Enabled = True
End Sub

Private Sub imgStop_Click()
    Call StopMovie
    Timer.Enabled = False
End Sub

Private Sub Timer_Timer()
    Call UpdateScreen
End Sub
