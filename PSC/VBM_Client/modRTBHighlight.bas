Attribute VB_Name = "modRTBHighlight"
Option Explicit
Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    Code As Long
End Type
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
Private Type ENLINK
    hdr As NMHDR
    MSG As Long
    wParam As Long
    lParam As Long
    chrg As CHARRANGE
End Type
Private Type TEXTRANGE
    chrg As CHARRANGE
    lpstrText As String
End Type
'Used to change the window procedure which kick-starts the subclassing
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
'Used to call the default window procedure for the parent
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, _
                                                                              ByVal hwnd As Long, _
                                                                              ByVal MSG As Long, _
                                                                              ByVal wParam As Long, _
                                                                              ByVal lParam As Long) As Long
'Used to set and retrieve various information
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long
'Used to copy... memory... from pointers
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, _
                                                                     Source As Any, _
                                                                     ByVal Length As Long)
'Used to launch the URL in the user's default browser
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                           ByVal lpOperation As String, _
                                                                           ByVal lpFile As String, _
                                                                           ByVal lpParameters As String, _
                                                                           ByVal lpDirectory As String, _
                                                                           ByVal nShowCmd As Long) As Long
Private Const WM_NOTIFY = &H4E
Private Const EM_SETEVENTMASK = &H445
Private Const EM_GETEVENTMASK = &H43B
Private Const EM_GETTEXTRANGE = &H44B
Private Const EM_AUTOURLDETECT = &H45B
Private Const EN_LINK = &H70B
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Const WM_SETCURSOR = &H20
Private Const CFE_LINK = &H20
Private Const ENM_LINK = &H4000000
Private Const GWL_WNDPROC = (-4)
Private Const SW_SHOW = 5
Private lOldProc As Long    'Old windowproc
Private hWndRTB As Long     'hWnd of RTB
Private hWndParent As Long  'hWnd of parent window

Public Sub DisableURLDetect()

'Don't want to unsubclass a non-subclassed window

    If lOldProc Then
        'Stop URL detection
        SendMessage hWndRTB, EM_AUTOURLDETECT, 0, ByVal 0
        'Reset the window procedure (stop the subclassing)
        SetWindowLong hWndParent, GWL_WNDPROC, lOldProc
        'Set this to 0 so we can subclass again in future
        lOldProc = 0
    End If

End Sub

Public Sub EnableURLDetect(ByVal hWndTextbox As Long, _
                           ByVal hwndOwner As Long)

'Don't want to subclass twice!

    If lOldProc = 0 Then
        'Subclass!
        lOldProc = SetWindowLong(hwndOwner, GWL_WNDPROC, AddressOf WndProc)
        'Tell the RTB to inform us when stuff happens to URLs
        SendMessage hWndTextbox, EM_SETEVENTMASK, 0, ByVal ENM_LINK Or SendMessage(hWndTextbox, EM_GETEVENTMASK, 0, 0)
        'Tell the RTB to start automatically detecting URLs
        SendMessage hWndTextbox, EM_AUTOURLDETECT, 1, ByVal 0
        hWndParent = hwndOwner
        hWndRTB = hWndTextbox
    End If

End Sub

Public Function WndProc(ByVal hwnd As Long, _
                        ByVal uMsg As Long, _
                        ByVal wParam As Long, _
                        ByVal lParam As Long) As Long

Dim uHead As NMHDR
Dim eLink As ENLINK
Dim eText As TEXTRANGE
Dim sText As String
Dim lLen  As Long

    'Which message?
    Select Case uMsg
    Case WM_NOTIFY
        'Ooo! A notify message! Something exciting must be happening...
        'Copy the notification header into our structure from the pointer
        CopyMemory uHead, ByVal lParam, Len(uHead)
        'Peek inside the structure
        If (uHead.hwndFrom = hWndRTB) And (uHead.Code = EN_LINK) Then
            'Yay! Some kind of kinky linky message.
            'Now that we know its a link message, we can copy the whole ENLINK structure
            'into our structure
            CopyMemory eLink, ByVal lParam, Len(eLink)
            'What kind of message?
            Select Case eLink.MSG
            Case WM_LBUTTONDBLCLK
                'Other miscellaneous messages
            Case WM_LBUTTONDOWN
            Case WM_LBUTTONUP
                'clicked the link!
                'Set up out TEXTRANGE struct
                eText.chrg.cpMin = eLink.chrg.cpMin
                eText.chrg.cpMax = eLink.chrg.cpMax
                eText.lpstrText = Space$(1024)
                'Tell the RTB to fill out our TEXTRANGE with the text
                lLen = SendMessage(hWndRTB, EM_GETTEXTRANGE, 0, eText)
                'Trim the text
                sText = Left$(eText.lpstrText, lLen)
                'Launch the browser
                ShellExecute hWndParent, vbNullString, sText, vbNullString, vbNullString, SW_SHOW
            Case WM_RBUTTONDBLCLK
            Case WM_RBUTTONDOWN
            Case WM_RBUTTONUP
            Case WM_SETCURSOR
            End Select
        End If
    End Select
    'Call the stored window procedure to let it handle all the messages
    WndProc = CallWindowProc(lOldProc, hwnd, uMsg, wParam, lParam)

End Function



