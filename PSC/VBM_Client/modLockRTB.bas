Attribute VB_Name = "modLockRTB"
Option Explicit
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hwnd As Long, _
                                                                              ByVal MSG As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
Public Const GWL_WNDPROC        As Long = (-4&)
Private Const WM_CHAR           As Long = &H102
Private Const WM_KEYDOWN        As Long = &H100
Private Const WM_KEYUP          As Long = &H101
Private Const WM_RBUTTONDOWN    As Long = &H204
Private Const WM_RBUTTONUP      As Long = &H205
Private Const WM_PAINT = &HF
Private Const WM_NCPAINT = &H85
Private Const WM_ERASEBKGND = &H14
Private Const WM_VSCROLL = &H115
Private PrevWndProc             As Long

Public Sub InitLRTB(hwnd&)

    PrevWndProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf SubWndProc)

End Sub

' hier werden keine messages abgefangen, sondern hier wird alles abgefangen.
' bis auf 4
Private Function SubWndProc(ByVal hwnd As Long, _
                            ByVal MSG As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam As Long) As Long

    If MSG = WM_PAINT Or MSG = WM_NCPAINT Or MSG = WM_ERASEBKGND Or MSG = 1246 Then
    Else 'NOT MSG...
        '  Debug.Print MSG
        Exit Function
    End If
    SubWndProc = CallWindowProc(PrevWndProc, hwnd, MSG, wParam, lParam)

End Function

Public Sub TerminateLRTB(hwnd&)

    Call SetWindowLong(hwnd, GWL_WNDPROC, PrevWndProc)

End Sub



