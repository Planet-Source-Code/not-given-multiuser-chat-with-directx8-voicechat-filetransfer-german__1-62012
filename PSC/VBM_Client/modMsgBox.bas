Attribute VB_Name = "modMsgBox"
Option Explicit
'Option Explicit
'
'Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
'
'Const MB_ABORTRETRYIGNORE = &H2&
'Const MB_APPLMODAL = &H0&
'Const MB_COMPOSITE = &H2
'Const MB_DEFAULT_DESKTOP_ONLY = &H20000
'Const MB_DEFBUTTON1 = &H0&
'Const MB_DEFBUTTON2 = &H100&
'Const MB_DEFBUTTON3 = &H200&
'Const MB_DEFMASK = &HF00&
'Const MB_ICONASTERISK = &H40&
'Const MB_ICONEXCLAMATION = &H30&
'Const MB_ICONHAND = &H10&
'Const MB_ICONMASK = &HF0&
'Const MB_ICONQUESTION = &H20&
'Const MB_MISCMASK = &HC000&
'Const MB_MODEMASK = &H3000&
'Const MB_NOFOCUS = &H8000&
'Const MB_OK = &H0&
'Const MB_OKCANCEL = &H1&
'Const MB_PRECOMPOSED = &H1
'Const MB_RETRYCANCEL = &H5&
'Const MB_SETFOREGROUND = &H10000
'Const MB_SYSTEMMODAL = &H1000&
'Const MB_TASKMODAL = &H2000&
'Const MB_TYPEMASK = &HF&
'Const MB_USEGLYPHCHARS = &H4
'Const MB_YESNO = &H4&
'Const MB_YESNOCANCEL = &H3&
'
'
'Public Function MsgBox()
'  Dim Style&, Prompt$, Titel$, Result&
'
'    Titel = "Diese Box wurde mittels API-Aufruf generiert"
'    Prompt = "Beachten Sie, daß im aufrufenden Task jetzt im " &
''             "Gegensatz zur VB-Standard-MsgBox der Timer " &
''             "weiterläuft!"
'    Style = MB_ICONASTERISK Or MB_YESNOCANCEL Or MB_SYSTEMMODAL
'
'    Result = MessageBox(Me.hwnd, Prompt, Titel, Style)
'
'    Me.Caption = "Rückgabewert = " & Result
'End Function
Private DummyToKeepDecCommentsInDeclarations As Boolean

