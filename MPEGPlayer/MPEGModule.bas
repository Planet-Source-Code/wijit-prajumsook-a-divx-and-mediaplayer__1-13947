Attribute VB_Name = "MPEG"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public strFileToPlay As String
Public bPlaying As Boolean
Public ffSpeed As Long
Public lTotalFrames As Long
Public lTotalTime As Long

Public Sub DragForm(frm As Form)
    Call ReleaseCapture
    Call SendMessage(frm.hwnd, &HA1, 2, 0)
End Sub

Public Sub PlayMovie()
    If strFileToPlay <> "" Then
        mciSendString "play " & strFileToPlay, 0, 0, 0
        bPlaying = True
        frmMain.lblCaption.Caption = "[ Playing ]"
        
    End If
End Sub

Public Sub StopMovie()
    If bPlaying Then
        mciSendString "stop " & strFileToPlay, 0, 0, 0
        bPlaying = False
        frmMain.lblCaption.Caption = "[ Stoped ]"
    End If
End Sub

Public Sub CloseMovie()
    If bPlaying Then
        mciSendString "close " & strFileToPlay, 0, 0, 0
        bPlaying = False
        frmMain.lblCaption.Caption = "[ No Media ]"
        UpdateScreen
    End If
End Sub

Public Sub CloseAll()
    mciSendString "close all", 0, 0, 0
End Sub

Public Sub OpenMovie()
    If strFileToPlay <> "" Then
        mciSendString "open " & strFileToPlay & " type MPEGVideo", 0, 0, 0
    End If
End Sub

Public Sub PauseMovie()
    If bPlaying Then
        mciSendString "pause " & strFileToPlay, 0, 0, 0
        bPlaying = False
        frmMain.lblCaption.Caption = "[ Paused ]"
    End If
End Sub

Public Sub FForward()
    If bPlaying Then
        Dim command As String
        Dim s As String * 40
        mciSendString "set " & strFileToPlay & " time format milliseconds", s, 128, 0&
        mciSendString "status " & strFileToPlay & " position wait", s, Len(s), 0
        command = "play " & strFileToPlay & " from " & CStr(CLng(s) + ffSpeed * 1000)
        mciSendString command, 0, 0, 0
        bPlaying = True
        mciSendString "set " & strFileToPlay & " time format frames", 0, 0, 0
    End If
End Sub

Public Sub Rewind()
    If bPlaying Then
        Dim command As String
        Dim s As String * 40
        mciSendString "set " & strFileToPlay & " time format milliseconds", s, 128, 0&
        mciSendString "status " & strFileToPlay & " position wait", s, Len(s), 0
        command = "play " & strFileToPlay & " from " & CStr(CLng(s) - ffSpeed * 1000)
        mciSendString command, 0, 0, 0
        bPlaying = True
        mciSendString "set " & strFileToPlay & " time format frames", 0, 0, 0
    End If

End Sub

Public Function GetThisTime(ByVal timein As Long) As String
    Dim conH As Integer
    Dim conM As Integer
    Dim conS As Integer
    Dim remTime As Long
    Dim strRetTime As String
    
    remTime = timein / 1000
    conH = Int(remTime / 3600)
    remTime = remTime Mod 3600
    conM = Int(remTime / 60)
    remTime = remTime Mod 60
    conS = remTime
    
    If conH > 0 Then
        strRetTime = Trim(Str(conH)) & ":"
    Else
        strRetTime = ""
    End If
    
    If conM >= 10 Then
        strRetTime = strRetTime & Trim(Str(conM))
    ElseIf conM > 0 Then
        strRetTime = strRetTime & "0" & Trim(Str(conM))
    Else
        strRetTime = strRetTime & "00"
    End If
    
    strRetTime = strRetTime & ":"
    
    If conS >= 10 Then
        strRetTime = strRetTime & Trim(Str(conS))
    ElseIf conS > 0 Then
        strRetTime = strRetTime & "0" & Trim(Str(conS))
    Else
        strRetTime = strRetTime & "00"
    End If
    
    GetThisTime = strRetTime
End Function


Public Sub TotalFrames()
    Dim TotalFrames As String * 128

    mciSendString "status " & strFileToPlay & " length", TotalFrames, 128, 0&
    lTotalFrames = Val(TotalFrames)
End Sub

Public Sub TotalTime()
    Dim TotalTime As String * 128

    mciSendString "set " & strFileToPlay & " time format ms", TotalTime, 128, 0&
    mciSendString "status " & strFileToPlay & " length", TotalTime, 128, 0&

    mciSendString "set " & strFileToPlay & " time format frames", 0&, 0&, 0&
    
    lTotalTime = Val(TotalTime)
End Sub

Public Sub UpdateScreen()
    Dim s As String * 40
    Dim t As String
    t = GetThisTime(lTotalTime)
    frmMain.lblTime.Caption = "Total Time: " & t
    frmMain.lblFrame.Caption = "Total Frm: " & lTotalFrames
    
    mciSendString "status " & strFileToPlay & " position", s, Len(s), 0
    frmMain.lbl.Caption = "Current Frm: " & s
End Sub
