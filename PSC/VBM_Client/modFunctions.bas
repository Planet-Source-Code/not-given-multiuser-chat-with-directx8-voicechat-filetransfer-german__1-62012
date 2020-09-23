Attribute VB_Name = "modFunctions"
Option Explicit
Private Const INTERNET_OPEN_TYPE_DIRECT     As Integer = 1
Private Const INTERNET_OPEN_TYPE_PROXY      As Integer = 3
Private Const INTERNET_FLAG_RELOAD          As Long = &H80000000
Private Const UserAgent                     As String = "VBM"
Private Declare Function InternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal sAgent As String, _
                                                                           ByVal lAccessType As Long, _
                                                                           ByVal sProxyName As String, _
                                                                           ByVal sProxyBypass As String, _
                                                                           ByVal lFlags As Long) As Long
Private Declare Function InternetCloseHandle Lib "wininet" (ByVal hInet As Long) As Integer
Private Declare Function InternetReadFile Lib "wininet" (ByVal hFile As Long, _
                                                         ByVal sBuffer As String, _
                                                         ByVal lNumBytesToRead As Long, _
                                                         lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, _
                                                                                 ByVal lpszUrl As String, _
                                                                                 ByVal lpszHeaders As String, _
                                                                                 ByVal dwHeadersLength As Long, _
                                                                                 ByVal dwFlags As Long, _
                                                                                 ByVal dwContext As Long) As Long
Private Declare Function PathCompactPath Lib "shlwapi" Alias "PathCompactPathA" (ByVal hDC As Long, _
                                                                                 ByVal lpszPath As String, _
                                                                                 ByVal dx As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal X As Long, _
                                                    Y, _
                                                    ByVal cx As Long, _
                                                    ByVal cy As Long, _
                                                    ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, _
                                                   ByVal bInvert As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWndLock As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long
Private Const LVM_FIRST = &H1000
Private Const LVM_SETEXTENDEDLISTVIEWSTYLE     As Long = (LVM_FIRST + 54)
Private Const LVS_EX_FULLROWSELECT             As Long = &H20
' hinten) nicht ändern
Private Declare Function GetCaretPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Function AppPath() As String

    If Right$(App.Path, 1) = "\" Then
        AppPath = Left$(App.Path, Len(App.Path) - 1)
    Else 'NOT RIGHT$(APP.PATH,...
        AppPath = App.Path
    End If

End Function

Public Function BytesToString(Bytes As Long) As String

Dim zwFileLen As Long

    zwFileLen = Bytes
    If zwFileLen <= 1024 Then
        BytesToString = zwFileLen & " Byte"
    ElseIf zwFileLen <= 1048576 Then 'NOT ZWFILELEN...
        BytesToString = Round(zwFileLen / (1024), 2) & " KB"
    Else 'NOT ZWFILELEN...
        BytesToString = Round(zwFileLen / 1048576, 2) & " MB"
    End If

End Function

Public Function CompactPath(oForm As Form, _
                            ByVal sPath As String, _
                            oControl As Control) As String

Dim nWidth As Long

    nWidth = oControl.Width / Screen.TwipsPerPixelX
    PathCompactPath oForm.hDC, sPath, nWidth
    CompactPath = sPath

End Function

Public Function DirExists(Path As String) As Boolean

    On Error Resume Next
    DirExists = CBool(GetAttr(Path) And vbDirectory)
    On Error GoTo 0

End Function

Public Function ExtractFilename(Path As String) As String

    ExtractFilename = Right$(Path, Len(Path) - InStrRev(Path, "\"))

End Function

Public Function FileExists(Path As String) As Boolean

Const NotFile As Double = vbDirectory Or vbVolume

    On Error Resume Next
    FileExists = (GetAttr(Path) And NotFile) = 0
    On Error GoTo 0

End Function

Public Sub FlashForm(hwnd As Long, _
                     Invert As Boolean)

    FlashWindow hwnd, Invert

End Sub

Public Sub FullRowSelect(LV As ListView)

Dim State As Long

    State = True
    SendMessage LV.hwnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, ByVal State

End Sub

Public Function GetDate() As String

    GetDate = Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Replace$(Date, "/", "."), "\", "."), ":", "."), "*", "."), "?", "."), Chr$(34), "."), "<", "."), ">", "."), "|", ".")

End Function

Public Function GetFileLen(Path As String) As String

Dim zwFileLen As Long

    If FileExists(Path) Then
        zwFileLen = FileLen(Path)
        If zwFileLen <= 1024 Then
            GetFileLen = zwFileLen & " Byte"
        ElseIf zwFileLen <= 1048576 Then 'NOT ZWFILELEN...
            GetFileLen = Round(zwFileLen / (1024), 2) & " KB"
        Else 'NOT ZWFILELEN...
            GetFileLen = Round(zwFileLen / 1048576, 2) & " MB"
        End If
    End If

End Function


Public Function GetNextFreeFilename(Filename As String) As String

Dim zw                       As String
Dim counter                  As Integer
Dim Extension                As String
Dim FilenameWithoutExtension As String

    zw = Filename
    counter = 1
    If FileExists(zw) Then
        If Not InStr(zw, ".") = 0 Then ' ist überhaupt ein punkt vorhanden?
            FilenameWithoutExtension = Left$(Filename, InStrRev(Filename, ".") - 1)
            Extension = Right$(zw, Len(zw) - Len(FilenameWithoutExtension))
        Else 'NOT NOT...
            FilenameWithoutExtension = zw
        End If
        Do Until Not FileExists(FilenameWithoutExtension & "(" & counter & ")" & Extension)
            counter = counter + 1
        Loop
        GetNextFreeFilename = FilenameWithoutExtension & "(" & counter & ")" & Extension
    Else 'FILEEXISTS(ZW) = FALSE/0
        GetNextFreeFilename = zw
    End If

End Function

Public Function GetTCursX() As Long

Dim pt As POINTAPI

    GetCaretPos pt
    GetTCursX = pt.X

End Function

Public Function GetTCursY() As Long

Dim pt As POINTAPI

    GetCaretPos pt
    GetTCursY = pt.Y

End Function

Public Function GetURL(URL As String) As String

'~ Internetseite auslesen

Dim l&
Dim Buffer$
Dim hOpen&
Dim hFile&
Dim Result&

    l = 50000
    Buffer = Space$(l)
    DoEvents
    hOpen = InternetOpen(UserAgent, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    hFile = InternetOpenUrl(hOpen, URL, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
    Call InternetReadFile(hFile, Buffer, l, Result)
    Call InternetCloseHandle(hFile)
    Call InternetCloseHandle(hOpen)
    Buffer = Left$(Buffer, Result)
    GetURL = Buffer

End Function

Public Function HasActiveWindow() As Boolean

Dim ret As Long

    ret = GetActiveWindow()
    HasActiveWindow = Not (ret = 0)

End Function

Public Function IsIDE() As Boolean

    On Error Resume Next
    Debug.Print 1 / 0
    IsIDE = (Err <> 0)

End Function

Public Function IsValidIP(Test As String) As Boolean

'~ Validate IP

Dim SubNets() As String
Dim i         As Integer

    If LCase$(Test) = "localhost" Then
        IsValidIP = True
        Exit Function
    End If
    If Len(Test) > 16 Then
        IsValidIP = False
        Exit Function
    End If
    SubNets = Split(Test, ".")
    If Not UBound(SubNets) = 3 Then
        IsValidIP = False
        Exit Function
    End If
    For i = 0 To 3
        If Not IsNumeric(SubNets(i)) Or SubNets(i) < 0 Or SubNets(i) > 255 Then
            IsValidIP = False
            Exit Function
        End If
    Next i
    IsValidIP = True

End Function

Public Function LoadFF(Filename As String) As String

Dim zw  As String
Dim fnr As Integer

    If Not FileExists(Filename) Then Exit Function
    fnr = FreeFile
    Open Filename For Binary As #fnr
    zw = Space$(LOF(fnr))
    Get #fnr, , zw
    Close #fnr
    LoadFF = zw

End Function

Public Function LockUpdate(hwnd As Long, _
                           LockWindow As Boolean)

    LockWindowUpdate IIf(LockWindow, hwnd, 0)

End Function

Public Sub MakeNormal(hwnd As Long)

    SetWindowPos hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS

End Sub

Public Sub MakeTopMost(hwnd As Long)

    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS

End Sub

Public Sub RefreshRTB(RTB As RichTextBox)

'ich weiß...

    RTB.Height = RTB.Height - 1
    DoEvents
    RTB.Height = RTB.Height + 1

End Sub

Public Function Rot13(ByRef Text As String) As String

'~ Verschlüsselung (einfach)

Dim i As Long

    Rot13 = Text
    For i = 1 To Len(Text)
        Select Case UCase$(Mid$(Text, i, 1))
        Case "A" To "M"
            Mid$(Rot13, i) = Chr$(Asc(Mid$(Text, i, 1)) + 13)
        Case "N" To "Z"
            Mid$(Rot13, i) = Chr$(Asc(Mid$(Text, i, 1)) - 13)
        End Select
    Next i

End Function

Public Sub ShakeForm(ByRef fForm As Form, _
                     ByVal lAmplitude As Long, _
                     ByVal lMilliSeconds As Long, _
                     Optional ByVal lFrameRefresh As Long = 10)

Dim lngOriginalLeft As Long
Dim lngOriginalTop  As Long
Dim X               As Long
Dim Y               As Long

On Error GoTo errh

    Randomize Timer * Timer
    lngOriginalLeft = fForm.Left
    lngOriginalTop = fForm.Top
    Do While lMilliSeconds >= lFrameRefresh
        fForm.Left = lngOriginalLeft
        fForm.Top = lngOriginalTop
        X = lMilliSeconds / lAmplitude
        Y = lMilliSeconds / lAmplitude
        Select Case Int((4) * Rnd + 1)
        Case 1
            fForm.Top = fForm.Top - Y
        Case 2
            fForm.Top = fForm.Top + Y
        Case 3
            fForm.Left = fForm.Left + X
        Case 4
            fForm.Left = fForm.Left - X
        End Select
        lMilliSeconds = lMilliSeconds - lFrameRefresh
        DoEvents
    Loop

errh:
End Sub



