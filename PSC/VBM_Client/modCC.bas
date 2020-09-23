Attribute VB_Name = "modCC"
Option Explicit
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                           ByVal nIndex As Long) As Long
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, _
                                                                                  ByVal lpfn As Long, _
                                                                                  ByVal hmod As Long, _
                                                                                  ByVal dwThreadId As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                   ByVal hWndInsertAfter As Long, _
                                                   ByVal X As Long, _
                                                   ByVal Y As Long, _
                                                   ByVal cx As Long, _
                                                   ByVal cy As Long, _
                                                   ByVal wFlags As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
Const GWL_HINSTANCE = (-6)
Const SWP_NOSIZE = &H1
Const SWP_NOZORDER = &H4
Const SWP_NOACTIVATE = &H10
Const HCBT_ACTIVATE = 5
Const WH_CBT = 5
Private hHook As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ChooseColor Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLORS) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
                                                                                    ByVal lpszShortPath As String, _
                                                                                    ByVal cchBuffer As Long) As Long
Private Declare Function ChooseFont Lib "comdlg32.dll" Alias "ChooseFontA" (pChoosefont As CHOOSEFONTS) As Long
Private Declare Function PrintDlg Lib "comdlg32.dll" Alias "PrintDlgA" (pPrintdlg As PRINTDLGS) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, _
                                                ByVal hDC As Long) As Long
Private Declare Sub CopyMemoryStr Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, _
                                                                        ByVal lpvSource As String, _
                                                                        ByVal cbCopy As Long)
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHOWHELP = &H10
Private Const OFS_MAXPATHNAME = 256
Private Const LF_FACESIZE = 32
'OFS_FILE_OPEN_FLAGS and OFS_FILE_SAVE_FLAGS below
'are mine to save long statements; they're not
'a standard Win32 type.
Private Const OFS_FILE_OPEN_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_CREATEPROMPT Or OFN_NODEREFERENCELINKS Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT
Private Const OFS_FILE_SAVE_FLAGS = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_OVERWRITEPROMPT Or OFN_HIDEREADONLY
Public Type OPENFILENAME
    nStructSize As Long
    hwndOwner As Long
    hInstance As Long
    sFilter As String
    sCustomFilter As String
    nCustFilterSize As Long
    nFilterIndex As Long
    sFile As String
    nFileSize As Long
    sFileTitle As String
    nTitleSize As Long
    sInitDir As String
    sDlgTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExt As Integer
    sDefFileExt As String
    nCustDataSize As Long
    fnHook As Long
    sTemplateName As String
End Type
Type NMHDR
    hwndFrom As Long
    idfrom As Long
    Code As Long
End Type
Type OFNOTIFY
    hdr As NMHDR
    lpOFN As OPENFILENAME
    pszFile As String        '  May be NULL
End Type
Type CHOOSECOLORS
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(LF_FACESIZE) As Byte
End Type
Public Type CHOOSEFONTS
    lStructSize As Long
    hwndOwner As Long          '  caller's window handle
    hDC As Long                '  printer DC/IC or NULL
    lpLogFont As Long          '  ptr. to a LOGFONT struct
    iPointSize As Long         '  10 * size in points of selected font
    flags As Long              '  enum. type flags
    rgbColors As Long          '  returned text color
    lCustData As Long          '  data passed to hook fn.
    lpfnHook As Long           '  ptr. to hook function
    lpTemplateName As String     '  custom template name
    hInstance As Long          '  instance handle of.EXE that
    lpszStyle As String          '  return the style field here
    nFontType As Integer          '  same value reported to the EnumFonts
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long           '  minimum pt size allowed &
    nSizeMax As Long           '  max pt size allowed if
End Type
Public Type FONT_CONST
    FontName As String
    FontSize As Long
    FontColor As Long
    FontBold As Boolean
    FontItalic As Boolean
    FontUnderline As Boolean
    FontStrikeThru As Boolean
End Type
Private Const CC_RGBINIT = &H1
Private Const CC_FULLOPEN = &H2
Private Const CC_PREVENTFULLOPEN = &H4
Private Const CC_SHOWHELP = &H8
Private Const CC_ENABLEHOOK = &H10
Private Const CC_ENABLETEMPLATE = &H20
Private Const CC_ENABLETEMPLATEHANDLE = &H40
Private Const CC_SOLIDCOLOR = &H80
Private Const CC_ANYCOLOR = &H100
Private Const COLOR_FLAGS = CC_FULLOPEN Or CC_ANYCOLOR Or CC_RGBINIT
Private Const CF_SCREENFONTS = &H1
Private Const CF_PRINTERFONTS = &H2
Private Const CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
Private Const CF_SHOWHELP = &H4&
Private Const CF_ENABLEHOOK = &H8&
Private Const CF_ENABLETEMPLATE = &H10&
Private Const CF_ENABLETEMPLATEHANDLE = &H20&
Private Const CF_INITTOLOGFONTSTRUCT = &H40&
Private Const CF_USESTYLE = &H80&
Private Const CF_EFFECTS = &H100&
Private Const CF_APPLY = &H200&
Private Const CF_ANSIONLY = &H400&
Private Const CF_SCRIPTSONLY = CF_ANSIONLY
Private Const CF_NOVECTORFONTS = &H800&
Private Const CF_NOOEMFONTS = CF_NOVECTORFONTS
Private Const CF_NOSIMULATIONS = &H1000&
Private Const CF_LIMITSIZE = &H2000&
Private Const CF_FIXEDPITCHONLY = &H4000&
Private Const CF_WYSIWYG = &H8000
'  must also have CF_SCREENFONTS CF_PRINTERFONTS
Private Const CF_FORCEFONTEXIST = &H10000
Private Const CF_SCALABLEONLY = &H20000
Private Const CF_TTONLY = &H40000
Private Const CF_NOFACESEL = &H80000
Private Const CF_NOSTYLESEL = &H100000
Private Const CF_NOSIZESEL = &H200000
Private Const CF_SELECTSCRIPT = &H400000
Private Const CF_NOSCRIPTSEL = &H800000
Private Const CF_NOVERTFONTS = &H1000000
Private Const SIMULATED_FONTTYPE = &H8000
Private Const PRINTER_FONTTYPE = &H4000
Private Const SCREEN_FONTTYPE = &H2000
Private Const BOLD_FONTTYPE = &H100
Private Const ITALIC_FONTTYPE = &H200
Private Const REGULAR_FONTTYPE = &H400
Private Const LBSELCHSTRING = "commdlg_LBSelChangedNotify"
Private Const SHAREVISTRING = "commdlg_ShareViolation"
Private Const FILEOKSTRING = "commdlg_FileNameOK"
Private Const COLOROKSTRING = "commdlg_ColorOK"
Private Const SETRGBSTRING = "commdlg_SetRGBColor"
Private Const HELPMSGSTRING = "commdlg_help"
Private Const FINDMSGSTRING = "commdlg_FindReplace"
Private Const CD_LBSELNOITEMS = -1
Private Const CD_LBSELCHANGE = 0
Private Const CD_LBSELSUB = 1
Private Const CD_LBSELADD = 2
Type PRINTDLGS
    lStructSize As Long
    hwndOwner As Long
    hDevMode As Long
    hDevNames As Long
    hDC As Long
    flags As Long
    nFromPage As Integer
    nToPage As Integer
    nMinPage As Integer
    nMaxPage As Integer
    nCopies As Integer
    hInstance As Long
    lCustData As Long
    lpfnPrintHook As Long
    lpfnSetupHook As Long
    lpPrintTemplateName As String
    lpSetupTemplateName As String
    hPrintTemplate As Long
    hSetupTemplate As Long
End Type
Private Const PD_ALLPAGES = &H0
Private Const PD_SELECTION = &H1
Private Const PD_PAGENUMS = &H2
Private Const PD_NOSELECTION = &H4
Private Const PD_NOPAGENUMS = &H8
Private Const PD_COLLATE = &H10
Private Const PD_PRINTTOFILE = &H20
Private Const PD_PRINTSETUP = &H40
Private Const PD_NOWARNING = &H80
Private Const PD_RETURNDC = &H100
Private Const PD_RETURNIC = &H200
Private Const PD_RETURNDEFAULT = &H400
Private Const PD_SHOWHELP = &H800
Private Const PD_ENABLEPRINTHOOK = &H1000
Private Const PD_ENABLESETUPHOOK = &H2000
Private Const PD_ENABLEPRINTTEMPLATE = &H4000
Private Const PD_ENABLESETUPTEMPLATE = &H8000
Private Const PD_ENABLEPRINTTEMPLATEHANDLE = &H10000
Private Const PD_ENABLESETUPTEMPLATEHANDLE = &H20000
Private Const PD_USEDEVMODECOPIES = &H40000
Private Const PD_USEDEVMODECOPIESANDCOLLATE = &H40000
Private Const PD_DISABLEPRINTTOFILE = &H80000
Private Const PD_HIDEPRINTTOFILE = &H100000
Private Const PD_NONETWORKBUTTON = &H200000
Type DEVNAMES
    wDriverOffset As Integer
    wDeviceOffset As Integer
    wOutputOffset As Integer
    wDefault As Integer
End Type
Private Const DN_DEFAULTPRN = &H1
Public Type SelectedFile
    nFilesSelected As Integer
    sFiles() As String
    sLastDirectory As String
    bCanceled As Boolean
End Type
Public Type SelectedColor
    oSelectedColor As OLE_COLOR
    bCanceled As Boolean
End Type
Public Type SelectedFont
    sSelectedFont As String
    bCanceled As Boolean
    bBold As Boolean
    bItalic As Boolean
    nSize As Integer
    bUnderline As Boolean
    bStrikeOut As Boolean
    lColor As Long
    sFaceName As String
End Type
Public FileDialog As OPENFILENAME
Private ColorDialog As CHOOSECOLORS
Public FontDialog As CHOOSEFONTS
Private PrintDialog As PRINTDLGS
Private ParenthWnd As Long

Private Function IsArrayEmpty(va As Variant) As Boolean

Dim v As Variant

    On Error Resume Next
    v = va(LBound(va))
    IsArrayEmpty = (Err <> 0)

End Function

Public Function ShowColor(ByVal hwnd As Long, _
                          Optional ByVal Color As Long = &H0&, _
                          Optional ByVal centerForm As Boolean = True) As SelectedColor

Dim customcolors() As Byte   ' dynamic (resizable) array
Dim i              As Integer
Dim ret            As Long
Dim hInst          As Long
Dim Thread         As Long

    ParenthWnd = hwnd
    ColorDialog.rgbResult = Color
    If ColorDialog.lpCustColors = "" Then
        ReDim customcolors(0 To 16 * 4 - 1) As Byte  'resize the array
        For i = LBound(customcolors) To UBound(customcolors)
            customcolors(i) = 254 ' sets all custom colors to white
        Next i
        ColorDialog.lpCustColors = StrConv(customcolors, vbUnicode)  ' convert array
    End If
    ColorDialog.hwndOwner = hwnd
    ColorDialog.lStructSize = Len(ColorDialog)
    ColorDialog.flags = COLOR_FLAGS
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else 'NOT CENTERFORM...
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    ret = ChooseColor(ColorDialog)
    If ret Then
        ShowColor.bCanceled = False
        ShowColor.oSelectedColor = ColorDialog.rgbResult
        Exit Function
    Else 'RET = FALSE/0
        ShowColor.bCanceled = True
        ShowColor.oSelectedColor = &H0&
        Exit Function
    End If

End Function

Public Function ShowFont(ByVal hwnd As Long, _
                         stFont As FONT_CONST, _
                         Optional ByVal centerForm As Boolean = True) As SelectedFont

Dim ret       As Long
Dim lfLogFont As LOGFONT
Dim hInst     As Long
Dim Thread    As Long
Dim i         As Integer
Dim fnt       As LOGFONT
Const PointsPerTwip = 1440 / 72

    ParenthWnd = hwnd
    fnt.lfHeight = -(stFont.FontSize * (PointsPerTwip / Screen.TwipsPerPixelY))
    fnt.lfWeight = IIf(stFont.FontBold, 500, 400)
    fnt.lfItalic = stFont.FontItalic
    fnt.lfUnderline = stFont.FontUnderline
    fnt.lfStrikeOut = stFont.FontStrikeThru
    StrToBytes fnt.lfFaceName, stFont.FontName
    FontDialog.nSizeMax = 0
    FontDialog.nSizeMin = 0
    FontDialog.nFontType = Screen.FontCount
    FontDialog.hwndOwner = hwnd
    FontDialog.hDC = 0
    FontDialog.lpfnHook = 0
    FontDialog.lCustData = 0
    FontDialog.lpLogFont = VarPtr(fnt)
    If FontDialog.iPointSize = 0 Then
        FontDialog.iPointSize = 10 * 10
    End If
    FontDialog.lpTemplateName = Space$(2048)
    FontDialog.rgbColors = stFont.FontColor 'RGB(0, 255, 255)
    FontDialog.lStructSize = Len(FontDialog)
    If FontDialog.flags = 0 Then
        FontDialog.flags = CF_SCREENFONTS Or CF_EFFECTS Or CF_INITTOLOGFONTSTRUCT 'Or CF_EFFECTS
    End If
    For i = 0 To Len(stFont.FontName) - 1
        lfLogFont.lfFaceName(i) = Asc(Mid$(stFont.FontName, i + 1, 1))
    Next i
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else 'NOT CENTERFORM...
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    ret = ChooseFont(FontDialog)
    If ret Then
        ShowFont.bCanceled = False
        ShowFont.bBold = IIf(fnt.lfWeight > 400, 1, 0)
        ShowFont.bItalic = fnt.lfItalic
        ShowFont.bStrikeOut = fnt.lfStrikeOut
        ShowFont.bUnderline = fnt.lfUnderline
        ShowFont.lColor = FontDialog.rgbColors
        ShowFont.nSize = FontDialog.iPointSize / 10
        For i = 0 To 31
            ShowFont.sSelectedFont = ShowFont.sSelectedFont + Chr$(fnt.lfFaceName(i))
        Next i
        ShowFont.sSelectedFont = Mid$(ShowFont.sSelectedFont, 1, InStr(1, ShowFont.sSelectedFont, vbNullChar) - 1)
        Exit Function
    Else 'RET = FALSE/0
        ShowFont.bCanceled = True
        Exit Function
    End If

End Function

Public Function ShowOpen(ByVal hwnd As Long, _
                         Optional ByVal centerForm As Boolean = True) As SelectedFile

Dim ret                 As Long
Dim Count               As Integer
Dim fileNameHolder      As String
Dim LastCharacter       As Integer
Dim NewCharacter        As Integer
Dim tempFiles(1 To 200) As String
Dim hInst               As Long
Dim Thread              As Long

    ParenthWnd = hwnd
    FileDialog.nStructSize = Len(FileDialog)
    FileDialog.hwndOwner = hwnd
    FileDialog.sFileTitle = Space$(2048)
    FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
    FileDialog.sFile = FileDialog.sFile & Space$(2047) & vbNullChar
    FileDialog.nFileSize = Len(FileDialog.sFile)
    'If FileDialog.flags = 0 Then
    FileDialog.flags = OFS_FILE_OPEN_FLAGS
    'End If
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else 'NOT CENTERFORM...
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    ret = GetOpenFileName(FileDialog)
    If ret Then
        If Trim$(FileDialog.sFileTitle) = "" Then
            LastCharacter = 0
            Count = 0
            While ShowOpen.nFilesSelected = 0
                NewCharacter = InStr(LastCharacter + 1, FileDialog.sFile, vbNullChar, vbTextCompare)
                If Count > 0 Then
                    tempFiles(Count) = Mid$(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                Else 'NOT COUNT...
                    ShowOpen.sLastDirectory = Mid$(FileDialog.sFile, LastCharacter + 1, NewCharacter - LastCharacter - 1)
                End If
                Count = Count + 1
                If InStr(NewCharacter + 1, FileDialog.sFile, vbNullChar, vbTextCompare) = InStr(NewCharacter + 1, FileDialog.sFile, vbNullChar & vbNullChar, vbTextCompare) Then
                    tempFiles(Count) = Mid$(FileDialog.sFile, NewCharacter + 1, InStr(NewCharacter + 1, FileDialog.sFile, vbNullChar & vbNullChar, vbTextCompare) - NewCharacter - 1)
                    ShowOpen.nFilesSelected = Count
                End If
                LastCharacter = NewCharacter
            Wend
            ReDim ShowOpen.sFiles(1 To ShowOpen.nFilesSelected)
            For Count = 1 To ShowOpen.nFilesSelected
                ShowOpen.sFiles(Count) = tempFiles(Count)
            Next Count
        Else 'NOT TRIM$(FILEDIALOG.SFILETITLE)...
            ReDim ShowOpen.sFiles(1 To 1)
            ShowOpen.sLastDirectory = Left$(FileDialog.sFile, FileDialog.nFileOffset)
            ShowOpen.nFilesSelected = 1
            ShowOpen.sFiles(1) = Mid$(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFile, vbNullChar, vbTextCompare) - FileDialog.nFileOffset - 1)
        End If
        ShowOpen.bCanceled = False
        Exit Function
    Else 'RET = FALSE/0
        ShowOpen.sLastDirectory = ""
        ShowOpen.nFilesSelected = 0
        ShowOpen.bCanceled = True
        Erase ShowOpen.sFiles
        Exit Function
    End If

End Function

Public Function ShowPrinter(ByVal hwnd As Long, _
                            Optional ByVal centerForm As Boolean = True) As Long

Dim hInst  As Long
Dim Thread As Long

    ParenthWnd = hwnd
    PrintDialog.hwndOwner = hwnd
    PrintDialog.lStructSize = Len(PrintDialog)
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else 'NOT CENTERFORM...
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    ShowPrinter = PrintDlg(PrintDialog)

End Function

Public Function ShowSave(ByVal hwnd As Long, _
                         Optional ByVal centerForm As Boolean = True) As SelectedFile

Dim ret    As Long
Dim hInst  As Long
Dim Thread As Long

    ParenthWnd = hwnd
    FileDialog.nStructSize = Len(FileDialog)
    FileDialog.hwndOwner = hwnd
    FileDialog.sFileTitle = Space$(2048)
    FileDialog.nTitleSize = Len(FileDialog.sFileTitle)
    FileDialog.sFile = Space$(2047) & vbNullChar
    FileDialog.nFileSize = Len(FileDialog.sFile)
    If FileDialog.flags = 0 Then
        FileDialog.flags = OFS_FILE_SAVE_FLAGS
    End If
    'Set up the CBT hook
    hInst = GetWindowLong(hwnd, GWL_HINSTANCE)
    Thread = GetCurrentThreadId()
    If centerForm = True Then
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterForm, hInst, Thread)
    Else 'NOT CENTERFORM...
        hHook = SetWindowsHookEx(WH_CBT, AddressOf WinProcCenterScreen, hInst, Thread)
    End If
    ret = GetSaveFileName(FileDialog)
    ReDim ShowSave.sFiles(1)
    If ret Then
        ShowSave.sLastDirectory = Left$(FileDialog.sFile, FileDialog.nFileOffset)
        ShowSave.nFilesSelected = 1
        ShowSave.sFiles(1) = Mid$(FileDialog.sFile, FileDialog.nFileOffset + 1, InStr(1, FileDialog.sFile, vbNullChar, vbTextCompare) - FileDialog.nFileOffset - 1)
        ShowSave.bCanceled = False
        Exit Function
    Else 'RET = FALSE/0
        ShowSave.sLastDirectory = ""
        ShowSave.nFilesSelected = 0
        ShowSave.bCanceled = True
        Erase ShowSave.sFiles
        Exit Function
    End If

End Function

Private Sub StrToBytes(ab() As Byte, _
                       s As String)

Dim cab As Long

    If IsArrayEmpty(ab) Then
        ' Assign to empty array
        ab = StrConv(s, vbFromUnicode)
    Else 'ISARRAYEMPTY(AB) = FALSE/0
        ' Copy to existing array, padding or truncating if necessary
        cab = UBound(ab) - LBound(ab) + 1
        If Len(s) < cab Then s = s & String$(cab - Len(s), 0)
        CopyMemoryStr ab(LBound(ab)), s, cab
    End If

End Sub

Private Function WinProcCenterForm(ByVal lMsg As Long, _
                                   ByVal wParam As Long, _
                                   ByVal lParam As Long) As Long

Dim rectForm As RECT, rectMsg As RECT
Dim X        As Long, Y As Long

    'On HCBT_ACTIVATE, show the MsgBox centered over Form1
    If lMsg = HCBT_ACTIVATE Then
        'Get the coordinates of the form and the message box so that
        'you can determine where the center of the form is located
        GetWindowRect ParenthWnd, rectForm
        GetWindowRect wParam, rectMsg
        X = (rectForm.Left + (rectForm.Right - rectForm.Left) / 2) - ((rectMsg.Right - rectMsg.Left) / 2)
        Y = (rectForm.Top + (rectForm.Bottom - rectForm.Top) / 2) - ((rectMsg.Bottom - rectMsg.Top) / 2)
        'Position the msgbox
        SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
    End If
    WinProcCenterForm = False

End Function

Private Function WinProcCenterScreen(ByVal lMsg As Long, _
                                     ByVal wParam As Long, _
                                     ByVal lParam As Long) As Long

Dim rectForm As RECT, rectMsg As RECT
Dim X        As Long, Y As Long

    If lMsg = HCBT_ACTIVATE Then
        'Show the MsgBox at a fixed location (0,0)
        GetWindowRect wParam, rectMsg
        X = Screen.Width / Screen.TwipsPerPixelX / 2 - (rectMsg.Right - rectMsg.Left) / 2
        Y = Screen.Height / Screen.TwipsPerPixelY / 2 - (rectMsg.Bottom - rectMsg.Top) / 2
        Debug.Print "Screen " & Screen.Height / 2
        Debug.Print "MsgBox " & (rectMsg.Right - rectMsg.Left) / 2
        SetWindowPos wParam, 0, X, Y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
        'Release the CBT hook
        UnhookWindowsHookEx hHook
    End If
    WinProcCenterScreen = False

End Function



