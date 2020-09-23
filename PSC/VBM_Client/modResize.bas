Attribute VB_Name = "modResize"
Option Explicit
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hwnd As Long, _
                                                                              ByVal MSG As Long, _
                                                                              ByVal wParam As Long, _
                                                                              lParam As RECT) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
Private lpPrevWndProc As Long
Private gHW As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Min_Width As Long     'Minimum width in pixels.
Public Min_Height As Long   'Minimum height in pixels.

Public Sub Hook(hwnd As Long)

    If gHW Then Exit Sub
    gHW = hwnd
    lpPrevWndProc = SetWindowLong(gHW, -4, AddressOf WindowProc)

End Sub

Public Sub Unhook()

    SetWindowLong gHW, -4, lpPrevWndProc
    gHW = 0

End Sub

Private Function WindowProc(ByVal hwnd As Long, _
                            ByVal uMsg As Long, _
                            ByVal wParam As Long, _
                            R As RECT) As Long

    If uMsg = 532 Then 'WM_SIZING
        If InStr(36, wParam) = 0 And R.Right - R.Left < Min_Width Then
            If InStr(147, wParam) Then
                R.Left = R.Right - Min_Width
            Else 'NOT INSTR(147,...
                R.Right = R.Left + Min_Width
            End If
        End If
        If InStr(12, wParam) = 0 And R.Bottom - R.Top < Min_Height Then
            If InStr(345, wParam) Then
                R.Top = R.Bottom - Min_Height
            Else 'NOT INSTR(345,...
                R.Bottom = R.Top + Min_Height
            End If
        End If
    ElseIf uMsg = 2 Then 'WM_CLOSE'NOT UMSG...
        Unhook
    End If
    WindowProc = CallWindowProc(lpPrevWndProc, hwnd, uMsg, wParam, R)

End Function



