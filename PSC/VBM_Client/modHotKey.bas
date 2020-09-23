Attribute VB_Name = "modHotKey"
Private Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, _
    ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
    
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, _
    ByVal id As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
    (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
    (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_HOTKEY = &H312
Private Const GWL_WNDPROC = -4


'Strg
Private Const MOD_STRG = &H2
'Shift
Private Const MOD_SHIFT = &H4
'Alt
Private Const MOD_ALT = &H1
'Windows-Taste
Private Const MOD_WIN = &H8

'F1 bis F12
Private Const VK_F1 = &H70
Private Const VK_F2 = &H71
Private Const VK_F3 = &H72
Private Const VK_F4 = &H73
Private Const VK_F5 = &H74
Private Const VK_F6 = &H75
Private Const VK_F7 = &H76
Private Const VK_F8 = &H77
Private Const VK_F9 = &H78
Private Const VK_F10 = &H79
Private Const VK_F11 = &H7A
Private Const VK_F12 = &H7B

Private glWinRet As Long

Private Function CallbackMsgs(ByVal wHwnd As Long, ByVal wMsg As Long, ByVal wp_id As Long, ByVal lp_id As Long) As Long
    If wMsg = WM_HOTKEY Then
        Call frmMain.HotKeyEvent(wp_id)
        CallbackMsgs = 1
        Exit Function
    End If
    CallbackMsgs = CallWindowProc(glWinRet, wHwnd, wMsg, wp_id, lp_id)
End Function

Public Sub StartHotkey(frm As Form)
RetVal = RegisterHotKey(frm.hwnd, 0, MOD_ALT, Asc("S"))
glWinRet = SetWindowLong(frm.hwnd, GWL_WNDPROC, AddressOf CallbackMsgs)
End Sub

Public Sub EndHotkey(frm As Form)
    UnregisterHotKey frm.hwnd, 0
End Sub


