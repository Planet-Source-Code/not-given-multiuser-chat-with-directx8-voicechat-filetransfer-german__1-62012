Attribute VB_Name = "modMenu"
Option Explicit
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hwnd As Long, _
                                                                              ByVal msg As Long, _
                                                                              ByVal wParam As Long, _
                                                                              lParam As RECT) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
Private lpPrevWndProc As Long
Private gHW As Long
 
Const WM_MENUSELECT = &H11F

Public Sub HookMenu(hwnd As Long)

    If gHW Then Exit Sub
    gHW = hwnd
    lpPrevWndProc = SetWindowLong(gHW, -4, AddressOf WindowProc)

End Sub

Public Sub UnhookMenu()

    SetWindowLong gHW, -4, lpPrevWndProc
    gHW = 0

End Sub

Private Function WindowProc(ByVal hwnd As Long, _
                            ByVal uMsg As Long, _
                            ByVal wParam As Long, _
                            R As RECT) As Long




    If uMsg = WM_MENUSELECT Then  'WM_SIZING
    Debug.Print uMsg
uMsg = 0

    ElseIf uMsg = 2 Then 'WM_CLOSE'NOT UMSG...
        UnhookMenu
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, R)

End Function

