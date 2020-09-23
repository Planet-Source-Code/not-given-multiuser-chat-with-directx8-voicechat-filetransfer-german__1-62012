Attribute VB_Name = "modResize2"
Option Explicit
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, _
                                                                            ByVal wMsg As Long, _
                                                                            ByVal wParam As Long, _
                                                                            ByVal lParam As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hwnd As Long, _
                                                                              ByVal MSG As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory1 Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                      Source As Any, _
                                                                      ByVal Length As Long)
Private Declare Sub CopyMemory2 Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Any, _
                                                                      Source As Any, _
                                                                      ByVal Length As Long)
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type
Public Type SIZEPAR
    xMin As Long
    yMin As Long
    xMax As Long
    yMax As Long
End Type
Private Const GWL_WNDPROC As Long = -4&
Private Const WM_GETMINMAXINFO As Long = &H24&
Private WinOldProc As Long
Public spR As SIZEPAR
Private Frm As Form

Public Sub InitR(F As Form)

    Set Frm = F
    WinOldProc = SetWindowLong(Frm.hwnd, GWL_WNDPROC, AddressOf WindowProc)

End Sub

Public Sub UnHookR()

    Call SetWindowLong(Frm.hwnd, GWL_WNDPROC, WinOldProc)

End Sub

Private Function WindowProc(ByVal hwnd As Long, _
                            ByVal uMsg As Long, _
                            ByVal wParam As Long, _
                            ByVal lParam&) As Long

Dim Result As Long
Dim MM     As MINMAXINFO

    If uMsg = WM_GETMINMAXINFO And Frm.WindowState = 0 Then
        Call CopyMemory1(MM, lParam, Len(MM))
        MM.ptMaxPosition.X = 0
        MM.ptMaxPosition.Y = 0
        MM.ptMaxSize.X = Screen.Width / Screen.TwipsPerPixelX
        MM.ptMaxSize.Y = Screen.Height / Screen.TwipsPerPixelY
        MM.ptMinTrackSize.X = spR.xMin
        MM.ptMinTrackSize.Y = spR.yMin
        MM.ptMaxTrackSize.X = spR.xMax
        MM.ptMaxTrackSize.Y = spR.yMax
        Call CopyMemory2(lParam&, MM, Len(MM))
        Result = DefWindowProc(hwnd, uMsg, wParam, lParam)
    Else 'NOT UMSG...
        Result = CallWindowProc(WinOldProc, hwnd, uMsg, wParam, lParam)
    End If
    WindowProc = Result

End Function



