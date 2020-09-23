Attribute VB_Name = "modSysTray"
Option Explicit
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 128
    dwState As Long
    dwStateMask As Long
    szInfo As String * 256
    uTimeout As Long
    szInfoTitle As String * 64
    dwInfoFlags As Long
End Type
Private nf_IconData As NOTIFYICONDATA
Private Const NOTIFYICON_VERSION = 3
Private Const NOTIFYICON_OLDVERSION = 0
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIM_SETFOCUS = &H3
Private Const NIM_SETVERSION = &H4
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10
Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2
Private Const NIIF_NONE = &H0
Private Const NIIF_WARNING = &H2
Private Const NIIF_ERROR = &H3
Private Const NIIF_INFO = &H1
Private Const NIIF_GUID = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_RBUTTONDBLCLK = &H206
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
                                                                                   pnid As NOTIFYICONDATA) As Boolean

Public Sub AddTray(Pic As PictureBox, _
                   traytip As String, _
                   Pict As PictureBox)

    With nf_IconData
        .cbSize = Len(nf_IconData)
        .hwnd = Pict.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Pic.Picture
        .szTip = traytip & vbNullChar 'QuickInfo des Symbols & vbNullChar
        .dwState = 0
        .dwStateMask = 0
    End With 'NF_ICONDATA
    Shell_NotifyIcon NIM_ADD, nf_IconData

End Sub

Public Sub BaloonTip(Pic As PictureBox, _
                     traytip As String, _
                     Pict As PictureBox, _
                     InfoTip As String, _
                     InfoTitle As String)

    With nf_IconData
        .cbSize = Len(nf_IconData)
        .hwnd = Pic.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Pict.Picture
        .szTip = traytip & vbNullChar 'QuickInfo des Symbols & vbNullChar
        .dwState = 0
        .dwStateMask = 0
        .szInfo = InfoTip & vbNullChar
        'Text des BallonTips. Mehrzeilige Tips sind möglich (Zeilenwechsel: vbCrLf). Am Ende muss ein Chr$(0) angehängt werden.
        .szInfoTitle = InfoTitle & vbNullChar
        'Fettformatierte Überschrift des Balloon-Tips. & Chr$(0)
        .dwInfoFlags = NIIF_INFO
        'Benutzes Icon für den Balloon (Info, Warnung oder Fehler)(NIIF_NONE, NIIF_INFO, NIIF_WARNING, NIIF_ERROR)
        .uTimeout = 1 'Zeit nachdem der Balloon spät. verschwinden soll (millisek.)
    End With 'NF_ICONDATA
    Shell_NotifyIcon NIM_MODIFY, nf_IconData

End Sub

Public Sub ModifyTray(Pic As PictureBox, _
                      traytip As String, _
                      Pict As PictureBox)

    With nf_IconData
        .cbSize = Len(nf_IconData)
        .hwnd = Pic.hwnd
        .uID = vbNull
        .uFlags = NIF_ICON Or NIF_INFO Or NIF_MESSAGE Or NIF_TIP
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Pict.Picture
        .szTip = traytip & vbNullChar 'QuickInfo des Symbols & vbNullChar
        .dwState = 0
        .dwStateMask = 0
    End With 'NF_ICONDATA
    Shell_NotifyIcon NIM_MODIFY, nf_IconData

End Sub

Public Sub RemTray()

    Shell_NotifyIcon NIM_DELETE, nf_IconData

End Sub

':)Code Fixer V3.0.9 (18.07.2005 15:59:15) 70 + 67 = 137 Lines Thanks Ulli for inspiration and lots of code.

