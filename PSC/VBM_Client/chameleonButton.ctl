VERSION 5.00
Begin VB.UserControl chameleonButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DefaultCancel   =   -1  'True
   PropertyPages   =   "chameleonButton.ctx":0000
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ToolboxBitmap   =   "chameleonButton.ctx":0035
   Begin VB.Timer OverTimer 
      Enabled         =   0   'False
      Interval        =   3
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "chameleonButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#Const isOCX = False
Private Const cbVersion             As String = "2.0.6"
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%             <<< GONCHUKI SYSTEMS >>>               %
'%                                                    %
'%                 CHAMELEON BUTTON                   %
'%         copyright ©2001-2002 by gonchuki           %
'%                                                    %
'%  this custom control will emulate the most common  %
'%      command buttons that everyone knows.          %
'%                                                    %
'%  it took me three months to develop this control   %
'% but that was a first step, now eight months after, %
'%  it turned out to be a very professional control.  %
'%                                                    %
'%     ALL THE CODE WAS WRITTEN FROM SCRATCH!!!       %
'%                                                    %
'%   ever wanted to add cool buttons to your app???   %
'%          this is the BEST solution!!!              %
'%                                                    %
'%        Copyright © 2001-2002 by gonchuki           %
'%                                                    %
'%    Commercial use of this control is FORBIDDEN     %
'%       without explicitly permission from me        %
'%    You can't either use any part of this code      %
'%              without my permission                 %
'%   You can use this code without asking for your    %
'%  personal projects or for freeware, but remember   %
'%           to give credits where its due            %
'%                                                    %
'%  If you are building an OCX version, you MUST set  %
'%      the isOCX constant to true and inlcude the    %
'%          original unmodified about form            %
'%                                                    %
'%            e-mail: gonchuki@yahoo.es               %
'%                                                    %
'%                  MADE IN URUGUAY                   %
'%                                                    %
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'######################################################
'#                    UPDTATE LOG                     #
'#  all times are GMT -03:00                          #
'#                                                    #
'# November 9  - 03:00 am                             #
'#              · first release                       #
'#                                                    #
'# November 9  - 05:00 pm                             #
'#              · added ShowFocusRect property        #
'#              · added repaint before triggering the #
'#                click event                         #
'#                                                    #
'# November 9  - 07:20 pm                             #
'#              · fixed the color shifting so it will #
'#                display the correct color and not a #
'#                weird one.                          #
'#              · improved Java button drawing        #
'#              · added custom colors capability      #
'#                now it looks better than ever COOL! #
'#              · improved Flat button drawing        #
'#                                                    #
'# November 13 - 03:40 pm                             #
'#              · fixed the WinXP button colors and   #
'#                styles. Note that as the colors are #
'#                relative to a base, and for this    #
'#                button i made a color work-around,  #
'#                some colors will be un-reachable    #
'#              · added MouseMove event as requested  #
'#                                                    #
'# November 18 - 10:40 am                             #
'#              · translated all the line methods to  #
'#                API calls. It's now faster than     #
'#                ever. It will also decrease the     #
'#                extra size of your exe!!!           #
'#              · improved Win32 button drawing       #
'#              · moved the direct calls to SetPixel  #
'#                to use less inline .hDC calls       #
'#              · fixed KeyDown/KeyUp events so they  #
'#                now act as they should              #
'#                                                    #
'# November 23 - 3:55 pm  (not updating on PSC...)    #
'#              · upgraded version to 1.1             #
'#              · added FontBold, and other similar   #
'#                properties as requested             #
'#              · greatly improved drawing speed by   #
'#                replacing lots of duplicated code   #
'#                with the new-brand function made by #
'#                me: "DrawFrame"                     #
'#              · fixed MouseDown/MouseUp events so   #
'#                they now act as they should         #
'#              · added MousePointer property         #
'#                                                    #
'# December 1  - 10:10 pm                             #
'#              · replaced the RECT types assignment  #
'#                in the resize event with API calls  #
'#                that take 3/4 the time of raw vb    #
'#              · added "use container" to the color  #
'#                schemes                             #
'#              · button now initializes with it's    #
'#                caption set as it's name            #
'#                                                    #
'# December 23 - 2:00 pm                              #
'#              · finally got all the code in API by  #
'#                replacing the Usercontrol.ForeColor #
'#                calls with CreatePen API            #
'#              · added support for wrapping captions #
'#              · changed a bit the XP button gradient#
'#                thanks to Ghuran Kartal for this    #
'#              · added refresh sub to force a button #
'#                redraw.                             #
'#              · MouseIcon property added            #
'#              · MouseOver/MouseOut events added and #
'#                also a ForeOver property is provided#
'#                to change font color on mouse over. #
'#                this also fixed the WinXP button,   #
'#                which design is now perfect.        #
'#              · added FlatHover button style that is#
'#                the real toolbar button.            #
'#                                                    #
'# January 1  - 11:15 am                 year 2002!!! #
'#              · some minor fixes                    #
'#              · new release!!!                      #
'#                                                    #
'# January 5  - 10:15 am                              #
'#              · fixed the memory leaks (only 1% of  #
'#                gdi is lost per 15-20 runs of demo) #
'#              · the font assignment has changed     #
'#              · fixed a very rare and random bug in #
'#                the XP-button. Problem was in the   #
'#                DrawLine sub. Thanks goes to Dennis #
'#                Vanderspek                          #
'#              · changed Mid and LCase to the faster #
'#                Mid$ and LCase$ way                 #
'#                                                    #
'# January 22  - 11:55 pm                             #
'#              · fixed the "not redrawing" bug under #
'#                Win 2K/NT/ME.                       #
'#              · fixed a bug that prevented hot keys #
'#                to work properly                    #
'#              · fixed the font alignment problem    #
'#                many many thanks to Carles P.V.     #
'#                                                    #
'# February 6  - 4:15 pm                              #
'#              · fixed property assignment problems  #
'#              · fixed "Use Container" color scheme  #
'#              · optimized a bit the code            #
'#              · fixed problem with system colors    #
'#              · added SoftBevel prop to allow the   #
'#                buton to be "flatter"               #
'#                                                    #
'# February 8  - 10:15 pm                             #
'#              · fixed click event when user double  #
'#                clicks on the button                #
'#                                                    #
'# February 10 - 2:35 pm                              #
'#              · added Office XP button style        #
'#              · added "DrawCaption" sub for easier  #
'#                caption management                  #
'#              · changed focus rects for flat buttons#
'#              · added "DisableRefresh" sub to allow #
'#                property changes without repainting #
'#                until needed to do so.              #
'#              · added BackOver property             #
'#                                                    #
'# February 11 - 1:15 am                              #
'#              · added primitive support for pictures#
'#              · fixed colors when mouse re-enters   #
'#                button area while holding the mouse #
'#                button.                             #
'#                                                    #
'# February 12 - 4:30 pm                              #
'#              · finished with the picture property! #
'#              · Java focus rect fixed               #
'#              · Office XP style fixed               #
'#              · Changed "ConvertFromSystemColor" sub#
'#                                                    #
'# February 14 - 6:20 pm                              #
'#              · replaced the transparent blitting   #
'#                function with one 10 times better   #
'#              · joined bitmaps & icons drawing      #
'#              · added "UseGreyscale" option         #
'#                                                    #
'# February 18 - 4:30 pm                              #
'#              · added embossed/engraved/shadowed fx #
'#              · added category for each property    #
'#              · added standard property pages       #
'#                                                    #
'# March 3 - 9:10 pm                                  #
'#              · fixed effects for XP styles         #
'#              · added mouseover detection function  #
'#              · some minor adjustments              #
'#                                                    #
'# March 31 - 2:55 am                                 #
'#              · upgraded to version 2.0             #
'#              · added transparent, 3D Hover and     #
'#                oval button types                   #
'#                                                    #
'# April 1 - 9:45 pm                                  #
'#              · fixed transparent button drawing    #
'#                                                    #
'# April 19 - 6:00 pm                                 #
'#              · fixed Ofice XP button colors        #
'#              · added built-in hand cursor          #
'#                                                    #
'# May 11 - 12:40 pm                                  #
'#              · added KDE 2 button style!           #
'#              · slightly optimized Mac button code  #
'#                                                    #
'# May 16 - 7:00 pm                                   #
'#              · added version property              #
'#              · added complilation options for lite #
'#                version (evaluation purpose only)   #
'#              · some optimizations for drawing fx   #
'#                                                    #
'# May 22 - 5:20 pm                                   #
'#              · added some code to make more robust #
'#                the lite version                    #
'#              · added background picture option     #
'#                                                    #
'# June 29 - 4:00 pm                                  #
'#              · added CheckBoxBehaviour option to   #
'#                allow the button behave as one of em#
'#                                                    #
'# July 25 - 11:55 pm                                 #
'#              · slightly optimized code, specially  #
'#                by removing the slow IIf's          #
'#              · corrected default state for KDE2    #
'#                                                    #
'# August 1 - 12:30 pm                                #
'#              · NEW PUBLIC RELEASE!!!    (ver 2.04) #
'#            2:40 pm                           2.05  #
'#              · button was not updating when "value"#
'#                prop was changed by the code. Thanks#
'#                to Steve and uZiGuLa.               #
'#              · fixed drawing for Win32 button while#
'#                being CheckBox and Value = True     #
'#                                                    #
'# August 2 - 11:30 pm                                #
'#              · fixed (i hope) the problem with the #
'#                WinXP disabled picture              #
'#              · fixed the "not redrawing" problem   #
'#                                                    #
'######################################################
Private Const COLOR_HIGHLIGHT       As Integer = 13
Private Const COLOR_BTNFACE         As Integer = 15
Private Const COLOR_BTNSHADOW       As Integer = 16
Private Const COLOR_BTNTEXT         As Integer = 18
Private Const COLOR_BTNHIGHLIGHT    As Integer = 20
Private Const COLOR_BTNDKSHADOW     As Integer = 21
Private Const COLOR_BTNLIGHT        As Integer = 22
Private Const DT_CALCRECT           As Long = &H400
Private Const DT_WORDBREAK          As Long = &H10
Private Const DT_CENTER             As Long = &H1 Or DT_WORDBREAK Or &H4
Private Const PS_SOLID              As Integer = 0
Private Const RGN_DIFF              As Integer = 4
Private Type RECT
    Left                                As Long
    Top                                 As Long
    Right                               As Long
    Bottom                              As Long
End Type
Private Type POINTAPI
    X                                   As Long
    Y                                   As Long
End Type
Private Type BITMAPINFOHEADER
    biSize                              As Long
    biWidth                             As Long
    biHeight                            As Long
    biPlanes                            As Integer
    biBitCount                          As Integer
    biCompression                       As Long
    biSizeImage                         As Long
    biXPelsPerMeter                     As Long
    biYPelsPerMeter                     As Long
    biClrUsed                           As Long
    biClrImportant                      As Long
End Type
Private Type RGBTRIPLE
    rgbBlue                             As Byte
    rgbGreen                            As Byte
    rgbRed                              As Byte
End Type
Private Type BITMAPINFO
    bmiHeader                           As BITMAPINFOHEADER
    bmiColors                           As RGBTRIPLE
End Type
Public Enum ButtonTypes
    [Windows 16-bit] = 1    'the old-fashioned Win16 button
    [Windows 32-bit] = 2    'the classic windows button
    [Windows XP] = 3        'the new brand XP button totally owner-drawn
    [Mac] = 4
    'i suppose it looks exactly as a Mac button... i took the style from a GetRight skin!!!
    [Java metal] = 5        'there are also other styles but not so different from windows one
    [Netscape 6] = 6
    'this is the button displayed in web-pages, it also appears in some java apps
    [Simple Flat] = 7       'the standard flat button seen on toolbars
    [Flat Highlight] = 8
    'again the flat button but this one has no border until the mouse is over it
    [Office XP] = 9         'the new Office XP button
    '[MacOS-X] = 10         'this is a plan for the future...
    [Transparent] = 11      'suggested from a user...
    [3D Hover] = 12         'took this one from "Noteworthy Composer" toolbal
    [Oval Flat] = 13        'a simple Oval Button
    [KDE 2] = 14            'the great standard KDE2 button!
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Mac, Transparent
#End If
Public Enum ColorTypes
    [Use Windows] = 1
    [Custom] = 2
    [Force Standard] = 3
    [Use Container] = 4
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Custom
#End If
Public Enum PicPositions
    cbLeft = 0
    cbRight = 1
    cbTop = 2
    cbBottom = 3
    cbBackground = 4
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private cbLeft, cbRight, cbTop, cbBottom, cbBackground
#End If
Public Enum fx
    cbNone = 0
    cbEmbossed = 1
    cbEngraved = 2
    cbShadowed = 3
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private cbNone, cbEmbossed, cbEngraved, cbShadowed
#End If
Private Const FXDEPTH               As Long = &H28
'events
Public Event Click()
Attribute Click.VB_MemberFlags = "200"
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseOver()
Public Event MouseOut()
'variables
Private MyButtonType                As ButtonTypes
Private MyColorType                 As ColorTypes
Private PicPosition                 As PicPositions
Private SFX                         As fx      'font and picture effects
Private He                          As Long    'the height of the button
Private Wi                          As Long    'the width of the button
Private BackC                       As Long    'back color
Private BackO                       As Long    'back color when mouse is over
Private ForeC                       As Long    'fore color
Private ForeO                       As Long    'fore color when mouse is over
Private MaskC                       As Long    'mask color
Private OXPb                        As Long
Private OXPf                        As Long
Private useMask                     As Boolean
Private useGrey                     As Boolean
Private useHand                     As Boolean
Private picNormal                   As StdPicture
Private picHover                    As StdPicture
Private pDC                         As Long    'used for the treansparent button
Private pBM                         As Long
Private oBM                         As Long
Private elTex                       As String 'current text
Private rc                          As RECT    'text and focus rect locations
Private rc2                         As RECT
Private rc3                         As RECT
Private fc                          As POINTAPI
Private picPT                       As POINTAPI 'picture Position & Size
Private picSZ                       As POINTAPI
Private rgnNorm                     As Long
Private LastButton                  As Byte
Private LastKeyDown                 As Byte
Private isEnabled                   As Boolean
Private isSoft                      As Boolean
Private HasFocus                    As Boolean
Private showFocusR                  As Boolean
Private cFace                       As Long
Private cLight                      As Long
Private cHighLight                  As Long
Private cShadow                     As Long
Private cDarkShadow                 As Long
Private cText                       As Long
Private cTextO                      As Long
Private cFaceO                      As Long
Private cMask                       As Long
Private XPFace                      As Long
Private lastStat                    As Byte    'used to avoid unnecessary repaints
Private TE                          As String
Private isShown                     As Boolean
Private isOver                      As Boolean
Private inLoop                      As Boolean
Private Locked                      As Boolean
Private captOpt                     As Long
Private isCheckbox                  As Boolean
Private cValue                      As Boolean
Private Declare Function SetPixel Lib "gdi32" Alias "SetPixelV" (ByVal hDC As Long, _
                                                                 ByVal X As Long, _
                                                                 ByVal Y As Long, _
                                                                 ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, _
                                                               ByVal lHPalette As Long, _
                                                               lColorRef As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, _
                                                   ByVal crColor As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, _
                                                                  ByVal lpStr As String, _
                                                                  ByVal nCount As Long, _
                                                                  lpRect As RECT, _
                                                                  ByVal wFormat As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, _
                                                lpRect As RECT, _
                                                ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hDC As Long, _
                                                 lpRect As RECT, _
                                                 ByVal hBrush As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, _
                                                     lpRect As RECT) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, _
                                              ByVal x1 As Long, _
                                              ByVal y1 As Long, _
                                              ByVal x2 As Long, _
                                              ByVal y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, _
                                                   ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, _
                                               ByVal X As Long, _
                                               ByVal Y As Long, _
                                               lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                                ByVal nWidth As Long, _
                                                ByVal crColor As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, _
                                                    ByVal y1 As Long, _
                                                    ByVal x2 As Long, _
                                                    ByVal y2 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, _
                                                        ByVal y1 As Long, _
                                                        ByVal x2 As Long, _
                                                        ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, _
                                                 ByVal hSrcRgn1 As Long, _
                                                 ByVal hSrcRgn2 As Long, _
                                                 ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hRgn As Long, _
                                                    ByVal bRedraw As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, _
                                                     lpRect As RECT) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, _
                                                   ByVal X As Long, _
                                                   ByVal Y As Long) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, _
                                                  ByVal X As Long, _
                                                  ByVal Y As Long) As Long
Private Declare Function CopyRect Lib "user32" (lpDestRect As RECT, _
                                                lpSourceRect As RECT) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, _
                                                       ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, _
                                                ByVal hBitmap As Long, _
                                                ByVal nStartScan As Long, _
                                                ByVal nNumScans As Long, _
                                                lpBits As Any, _
                                                lpBI As BITMAPINFO, _
                                                ByVal wUsage As Long) As Long
Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDC As Long, _
                                                        ByVal X As Long, _
                                                        ByVal Y As Long, _
                                                        ByVal dx As Long, _
                                                        ByVal dy As Long, _
                                                        ByVal SrcX As Long, _
                                                        ByVal SrcY As Long, _
                                                        ByVal Scan As Long, _
                                                        ByVal NumScans As Long, _
                                                        Bits As Any, _
                                                        BitsInfo As BITMAPINFO, _
                                                        ByVal wUsage As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hSrcDC As Long, _
                                             ByVal xSrc As Long, _
                                             ByVal ySrc As Long, _
                                             ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, _
                                                  ByVal xLeft As Long, _
                                                  ByVal yTop As Long, _
                                                  ByVal hIcon As Long, _
                                                  ByVal cxWidth As Long, _
                                                  ByVal cyWidth As Long, _
                                                  ByVal istepIfAniCur As Long, _
                                                  ByVal hbrFlickerFreeDraw As Long, _
                                                  ByVal diFlags As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, _
                                                   ByVal nHeight As Long, _
                                                   ByVal nPlanes As Long, _
                                                   ByVal nBitCount As Long, _
                                                   lpBits As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, _
                                                    ByVal nIndex As Long) As Long

'########## BUTTON PROPERTIES ##########
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501

    BackColor = BackC

End Property

Public Property Let BackColor(ByVal theCol As OLE_COLOR)

    BackC = theCol
    If Not Ambient.UserMode Then
        BackO = theCol
    End If
    Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "BCOL"

End Property

Public Property Get BackOver() As OLE_COLOR

    BackOver = BackO

End Property

Public Property Let BackOver(ByVal theCol As OLE_COLOR)

    BackO = theCol
    Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "BCOLO"

End Property

Public Property Get ButtonType() As ButtonTypes

    ButtonType = MyButtonType

End Property

Public Property Let ButtonType(ByVal newValue As ButtonTypes)

    MyButtonType = newValue
    If MyButtonType = [Java metal] And Not Ambient.UserMode Then
        UserControl.FontBold = True
    ElseIf MyButtonType = 11 And isShown Then 'NOT MYBUTTONTYPE...
        Call GetParentPic
    End If
    Call UserControl_Resize
    PropertyChanged "BTYPE"

End Property

Private Sub CalcPicPos()

'exit if there's no picture

    If picNormal Is Nothing And picHover Is Nothing Then
        Exit Sub
    End If
    If (Trim$(elTex) <> "") And (PicPosition <> 4) Then
        'if there is no caption, or we have the picture as background, then we put the picture at the center of the button
        Select Case PicPosition
        Case 0 'left
            picPT.X = rc.Left - picSZ.X - 4
            picPT.Y = (He - picSZ.Y) \ 2
        Case 1 'right
            picPT.X = rc.Right + 4
            picPT.Y = (He - picSZ.Y) \ 2
        Case 2 'top
            picPT.X = (Wi - picSZ.X) \ 2
            picPT.Y = rc.Top - picSZ.Y - 2
        Case 3 'bottom
            picPT.X = (Wi - picSZ.X) \ 2
            picPT.Y = rc.Bottom + 2
        End Select
    Else 'center the picture'NOT (TRIM$(ELTEX)...
        picPT.X = (Wi - picSZ.X) \ 2
        picPT.Y = (He - picSZ.Y) \ 2
    End If

End Sub

Private Sub CalcPicSize()

    If Not picNormal Is Nothing Then
        picSZ.X = UserControl.ScaleX(picNormal.Width, 8, UserControl.ScaleMode)
        picSZ.Y = UserControl.ScaleY(picNormal.Height, 8, UserControl.ScaleMode)
    Else 'NOT NOT...
        picSZ.X = 0
        picSZ.Y = 0
    End If

End Sub

Private Sub CalcTextRects()

'this sub will calculate the rects required to draw the text

    Select Case PicPosition
    Case 0
        With rc2
            .Left = 1 + picSZ.X
            .Right = Wi - 2
            .Top = 1
            .Bottom = He - 2
        End With 'rc2
    Case 1
        With rc2
            .Left = 1
            .Right = Wi - 2 - picSZ.X
            .Top = 1
            .Bottom = He - 2
        End With 'rc2
    Case 2
        With rc2
            .Left = 1
            .Right = Wi - 2
            .Top = 1 + picSZ.Y
            .Bottom = He - 2
        End With 'rc2
    Case 3
        With rc2
            .Left = 1
            .Right = Wi - 2
            .Top = 1
            .Bottom = He - 2 - picSZ.Y
        End With 'rc2
    Case 4
        With rc2
            .Left = 1
            .Right = Wi - 2
            .Top = 1
            .Bottom = He - 2
        End With 'rc2
    End Select
    DrawText UserControl.hDC, elTex, Len(elTex), rc2, DT_CALCRECT Or DT_WORDBREAK
    CopyRect rc, rc2
    fc.X = rc.Right - rc.Left
    fc.Y = rc.Bottom - rc.Top
    Select Case PicPosition
    Case 0, 2
        OffsetRect rc, (Wi - rc.Right) \ 2, (He - rc.Bottom) \ 2
    Case 1
        OffsetRect rc, (Wi - rc.Right - picSZ.X - 4) \ 2, (He - rc.Bottom) \ 2
    Case 3
        OffsetRect rc, (Wi - rc.Right) \ 2, (He - rc.Bottom - picSZ.Y - 4) \ 2
    Case 4
        OffsetRect rc, (Wi - rc.Right) \ 2, (He - rc.Bottom) \ 2
    End Select
    CopyRect rc2, rc
    OffsetRect rc2, 1, 1
    Call CalcPicPos 'once we have the text position we are able to calculate the pic position

End Sub

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = 0

    Caption = elTex

End Property

Public Property Let Caption(ByVal newValue As String)

    elTex = newValue
    Call SetAccessKeys
    Call CalcTextRects
    Call Redraw(0, True)
    PropertyChanged "TX"

End Property

Public Property Get CheckBoxBehaviour() As Boolean

    CheckBoxBehaviour = isCheckbox

End Property

Public Property Let CheckBoxBehaviour(ByVal newValue As Boolean)

    isCheckbox = newValue
    Call Redraw(lastStat, True)
    PropertyChanged "CHECK"

End Property

'it is very common that a windows user uses custom color
'schemes to view his/her desktop, and is also very
'common that this color scheme has weird colors that
'would alter the nice look of my buttons.
'So if you want to force the button to use the windows
'standard colors you may change this property to "Force Standard"
Public Property Get ColorScheme() As ColorTypes

    ColorScheme = MyColorType

End Property

Public Property Let ColorScheme(ByVal newValue As ColorTypes)

    MyColorType = newValue
    Call SetColors
    Call Redraw(0, True)
    PropertyChanged "COLTYPE"

End Property

Private Function ConvertFromSystemColor(ByVal theColor As Long) As Long

    Call OleTranslateColor(theColor, 0, ConvertFromSystemColor)

End Function

Public Sub DisableRefresh()

'this is for fast button editing, once you disable the refresh,
' you can change every prop without triggering the drawing methods.
' once you are done, you call Refresh.

    isShown = False

End Sub

Private Sub DoFX(ByVal offset As Long, _
                 ByVal thePic As StdPicture)

Dim curFace As Long

    If SFX > cbNone Then
        If MyButtonType = [Windows XP] Then
            curFace = XPFace
        Else 'NOT MYBUTTONTYPE...
            If offset = -1 And MyColorType <> Custom Then
                curFace = OXPf
            Else 'NOT OFFSET...
                curFace = cFace
            End If
        End If
        TransBlt UserControl.hDC, picPT.X + 1 + offset, picPT.Y + 1 + offset, picSZ.X, picSZ.Y, thePic, cMask, ShiftColor(curFace, Abs(SFX = cbEngraved) * FXDEPTH + (SFX <> cbEngraved) * FXDEPTH)
        If SFX < cbShadowed Then
            TransBlt UserControl.hDC, picPT.X - 1 + offset, picPT.Y - 1 + offset, picSZ.X, picSZ.Y, thePic, cMask, ShiftColor(curFace, Abs(SFX <> cbEngraved) * FXDEPTH + (SFX = cbEngraved) * FXDEPTH)
        End If
    End If

End Sub

Private Sub DrawCaption(ByVal State As Byte)

'this code is commonly shared through all the buttons so
' i took it and put it toghether here for easier readability
' of the code, and to cut-down disk size.

    captOpt = State
    With UserControl
        Select Case State
            'in this select case, we only change the text color and draw only text that needs rc2, at the end, text that uses rc will be drawn
        Case 0 'normal caption
            txtFX rc
            SetTextColor .hDC, cText
        Case 1 'hover caption
            txtFX rc
            SetTextColor .hDC, cTextO
        Case 2 'down caption
            txtFX rc2
            If MyButtonType = Mac Then
                SetTextColor .hDC, cLight
            Else 'NOT MYBUTTONTYPE...
                SetTextColor .hDC, cTextO
            End If
            DrawText .hDC, elTex, Len(elTex), rc2, DT_CENTER
        Case 3 'disabled embossed caption
            SetTextColor .hDC, cHighLight
            DrawText .hDC, elTex, Len(elTex), rc2, DT_CENTER
            SetTextColor .hDC, cShadow
        Case 4 'disabled grey caption
            SetTextColor .hDC, cShadow
        Case 5 'WinXP disabled caption
            SetTextColor .hDC, ShiftColor(XPFace, -&H68, True)
        Case 6 'KDE 2 disabled
            SetTextColor .hDC, cHighLight
            DrawText .hDC, elTex, Len(elTex), rc2, DT_CENTER
            SetTextColor .hDC, cFace
        Case 7 'KDE 2 down
            SetTextColor .hDC, ShiftColor(cShadow, -&H32)
            DrawText .hDC, elTex, Len(elTex), rc2, DT_CENTER
            SetTextColor .hDC, cHighLight
        End Select
        'we now draw the text that is common in all the captions
        If State <> 2 Then
            DrawText .hDC, elTex, Len(elTex), rc, DT_CENTER
        End If
    End With 'USERCONTROL

End Sub

Private Sub DrawEllipse(ByVal X As Long, _
                        ByVal Y As Long, _
                        ByVal Width As Long, _
                        ByVal Height As Long, _
                        ByVal BorderColor As Long, _
                        ByVal FillColor As Long)

Dim pBrush As Long
Dim pPen   As Long

    pBrush = SelectObject(hDC, CreateSolidBrush(FillColor))
    pPen = SelectObject(hDC, CreatePen(PS_SOLID, 2, BorderColor))
    Call Ellipse(hDC, X, Y, X + Width, Y + Height)
    Call DeleteObject(SelectObject(hDC, pBrush))
    Call DeleteObject(SelectObject(hDC, pPen))

End Sub

Private Sub DrawFocusR()

    If showFocusR And HasFocus Then
        SetTextColor UserControl.hDC, cText
        DrawFocusRect UserControl.hDC, rc3
    End If

End Sub

Private Sub DrawFrame(ByVal ColHigh As Long, _
                      ByVal ColDark As Long, _
                      ByVal ColLight As Long, _
                      ByVal ColShadow As Long, _
                      ByVal ExtraOffset As Boolean, _
                      Optional ByVal Flat As Boolean = False)

'a very fast way to draw windows-like frames

Dim pt     As POINTAPI
Dim frHe   As Long
Dim frWi   As Long
Dim frXtra As Long

    frHe = He - 1 + ExtraOffset
    frWi = Wi - 1 + ExtraOffset
    frXtra = Abs(ExtraOffset)
    With UserControl
        Call DeleteObject(SelectObject(.hDC, CreatePen(PS_SOLID, 1, ColHigh)))
        '=============================
        MoveToEx .hDC, frXtra, frHe, pt
        LineTo .hDC, frXtra, frXtra
        LineTo .hDC, frWi, frXtra
        '=============================
        Call DeleteObject(SelectObject(.hDC, CreatePen(PS_SOLID, 1, ColDark)))
        '=============================
        LineTo .hDC, frWi, frHe
        LineTo .hDC, frXtra - 1, frHe
        MoveToEx .hDC, frXtra + 1, frHe - 1, pt
        If Flat Then
            Exit Sub
        End If
        '=============================
        Call DeleteObject(SelectObject(.hDC, CreatePen(PS_SOLID, 1, ColLight)))
        '=============================
        LineTo .hDC, frXtra + 1, frXtra + 1
        LineTo .hDC, frWi - 1, frXtra + 1
        '=============================
        Call DeleteObject(SelectObject(.hDC, CreatePen(PS_SOLID, 1, ColShadow)))
        '=============================
        LineTo .hDC, frWi - 1, frHe - 1
        LineTo .hDC, frXtra, frHe - 1
    End With 'USERCONTROL

End Sub

Private Sub DrawLine(ByVal x1 As Long, _
                     ByVal y1 As Long, _
                     ByVal x2 As Long, _
                     ByVal y2 As Long, _
                     ByVal Color As Long)

'a fast way to draw lines

Dim pt     As POINTAPI
Dim oldPen As Long
Dim hPen   As Long

    With UserControl
        hPen = CreatePen(PS_SOLID, 1, Color)
        oldPen = SelectObject(.hDC, hPen)
        MoveToEx .hDC, x1, y1, pt
        LineTo .hDC, x2, y2
        SelectObject .hDC, oldPen
        DeleteObject hPen
    End With 'USERCONTROL

End Sub

Private Sub DrawPictures(ByVal State As Byte)

'check if there is a main picture, if not then exit

    If picNormal Is Nothing Then
        Exit Sub
    End If
    With UserControl
        Select Case State
        Case 0 'normal & hover
            If Not isOver Then
                Call DoFX(0, picNormal)
                TransBlt .hDC, picPT.X, picPT.Y, picSZ.X, picSZ.Y, picNormal, cMask, , , useGrey, (MyButtonType = [Office XP])
            Else 'NOT NOT...
                If MyButtonType = [Office XP] Then
                    Call DoFX(-1, picNormal)
                    TransBlt .hDC, picPT.X + 1, picPT.Y + 1, picSZ.X, picSZ.Y, picNormal, cMask, cShadow
                    TransBlt .hDC, picPT.X - 1, picPT.Y - 1, picSZ.X, picSZ.Y, picNormal, cMask
                Else 'NOT MYBUTTONTYPE...
                    If Not picHover Is Nothing Then
                        Call DoFX(0, picHover)
                        TransBlt .hDC, picPT.X, picPT.Y, picSZ.X, picSZ.Y, picHover, cMask
                    Else 'NOT NOT...
                        Call DoFX(0, picNormal)
                        TransBlt .hDC, picPT.X, picPT.Y, picSZ.X, picSZ.Y, picNormal, cMask
                    End If
                End If
            End If
        Case 1 'down
            If picHover Is Nothing Or MyButtonType = [Office XP] Then
                Select Case MyButtonType
                Case 5, 9
                    Call DoFX(0, picNormal)
                    TransBlt .hDC, picPT.X, picPT.Y, picSZ.X, picSZ.Y, picNormal, cMask
                Case Else
                    Call DoFX(1, picNormal)
                    TransBlt .hDC, picPT.X + 1, picPT.Y + 1, picSZ.X, picSZ.Y, picNormal, cMask
                End Select
            Else 'NOT PICHOVER...
                TransBlt .hDC, picPT.X + Abs(MyButtonType <> [Java metal]), picPT.Y + Abs(MyButtonType <> [Java metal]), picSZ.X, picSZ.Y, picHover, cMask
            End If
        Case 2 'disabled
            Select Case MyButtonType
            Case 5, 6, 9    'draw flat grey pictures
                TransBlt .hDC, picPT.X, picPT.Y, picSZ.X, picSZ.Y, picNormal, cMask, Abs(MyButtonType = [Office XP]) * ShiftColor(cShadow, &HD) + Abs(MyButtonType <> [Office XP]) * cShadow, True
            Case 3          'for WinXP draw a greyscaled image
                TransBlt .hDC, picPT.X + 1, picPT.Y + 1, picSZ.X, picSZ.Y, picNormal, cMask, , , True
            Case Else       'draw classic embossed pictures
                TransBlt .hDC, picPT.X + 1, picPT.Y + 1, picSZ.X, picSZ.Y, picNormal, cMask, cHighLight, True
                TransBlt .hDC, picPT.X, picPT.Y, picSZ.X, picSZ.Y, picNormal, cMask, cShadow, True
            End Select
        End Select
    End With 'USERCONTROL
    If PicPosition = cbBackground Then
        Call DrawCaption(captOpt)
    End If

End Sub

Private Sub DrawRectangle(ByVal X As Long, _
                          ByVal Y As Long, _
                          ByVal Width As Long, _
                          ByVal Height As Long, _
                          ByVal Color As Long, _
                          Optional OnlyBorder As Boolean = False)

'this is my custom function to draw rectangles and frames
'it's faster and smoother than using the line method

Dim bRECT  As RECT
Dim hBrush As Long

    With bRECT
        .Left = X
        .Top = Y
        .Right = X + Width
        .Bottom = Y + Height
    End With 'bRECT
    hBrush = CreateSolidBrush(Color)
    If OnlyBorder Then
        FrameRect UserControl.hDC, bRECT, hBrush
    Else 'ONLYBORDER = FALSE/0
        FillRect UserControl.hDC, bRECT, hBrush
    End If
    DeleteObject hBrush

End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514

    Enabled = isEnabled

End Property

Public Property Let Enabled(ByVal newValue As Boolean)

    isEnabled = newValue
    Call Redraw(0, True)
    UserControl.Enabled = isEnabled
    PropertyChanged "ENAB"

End Property

Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512

    Set Font = UserControl.Font

End Property

Public Property Set Font(ByRef newFont As Font)

    Set UserControl.Font = newFont
    Call CalcTextRects
    Call Redraw(0, True)
    PropertyChanged "FONT"

End Property

Public Property Get FontBold() As Boolean

    FontBold = UserControl.FontBold

End Property

Public Property Let FontBold(ByVal newValue As Boolean)

    UserControl.FontBold = newValue
    Call CalcTextRects
    Call Redraw(0, True)

End Property

Public Property Get FontItalic() As Boolean

    FontItalic = UserControl.FontItalic

End Property

Public Property Let FontItalic(ByVal newValue As Boolean)

    UserControl.FontItalic = newValue
    Call CalcTextRects
    Call Redraw(0, True)

End Property

Public Property Get FontName() As String

    FontName = UserControl.FontName

End Property

Public Property Let FontName(ByVal newValue As String)

    UserControl.FontName = newValue
    Call CalcTextRects
    Call Redraw(0, True)

End Property

Public Property Get FontSize() As Integer

    FontSize = UserControl.FontSize

End Property

Public Property Let FontSize(ByVal newValue As Integer)

    UserControl.FontSize = newValue
    Call CalcTextRects
    Call Redraw(0, True)

End Property

Public Property Get FontUnderline() As Boolean

    FontUnderline = UserControl.FontUnderline

End Property

Public Property Let FontUnderline(ByVal newValue As Boolean)

    UserControl.FontUnderline = newValue
    Call CalcTextRects
    Call Redraw(0, True)

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513

    ForeColor = ForeC

End Property

Public Property Let ForeColor(ByVal theCol As OLE_COLOR)

    ForeC = theCol
    If Not Ambient.UserMode Then
        ForeO = theCol
    End If
    Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "FCOL"

End Property

Public Property Get ForeOver() As OLE_COLOR

    ForeOver = ForeO

End Property

Public Property Let ForeOver(ByVal theCol As OLE_COLOR)

    ForeO = theCol
    Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "FCOLO"

End Property

Private Sub GetParentPic()

    On Local Error Resume Next
    inLoop = True
    UserControl.Height = 0
    DoEvents
    BitBlt pDC, 0, 0, Wi, He, GetDC(GetParent(hwnd)), Extender.Left, Extender.Top, vbSrcCopy
    UserControl.Height = ScaleY(He, vbPixels, vbTwips)
    inLoop = False

End Sub

Public Property Get HandPointer() As Boolean

    HandPointer = useHand

End Property

Public Property Let HandPointer(ByVal newVal As Boolean)

    useHand = newVal
    If useHand Then
        Set UserControl.MouseIcon = LoadResPicture(101, 2)
        UserControl.MousePointer = 99
    Else 'USEHAND = FALSE/0
        Set UserControl.MouseIcon = Nothing
        UserControl.MousePointer = 1
    End If
    PropertyChanged "HAND"

End Property

Public Property Get hwnd() As Long
Attribute hwnd.VB_UserMemId = -515

    hwnd = UserControl.hwnd

End Property

Private Function isMouseOver() As Boolean

Dim pt As POINTAPI

    GetCursorPos pt
    isMouseOver = (WindowFromPoint(pt.X, pt.Y) = hwnd)

End Function

Private Sub MakeRegion()

'this function creates the regions to "cut" the UserControl
'so it will be transparent in certain areas

Dim rgn1 As Long
Dim rgn2 As Long

    DeleteObject rgnNorm
    rgnNorm = CreateRectRgn(0, 0, Wi, He)
    rgn2 = CreateRectRgn(0, 0, 0, 0)
    Select Case MyButtonType
    Case 1, 5, 14 'Windows 16-bit, Java & KDE 2
        rgn1 = CreateRectRgn(0, He, 1, He - 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 0, Wi - 1, 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        If MyButtonType <> 5 Then  'the above was common code
            rgn1 = CreateRectRgn(0, 0, 1, 1)
            CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
            DeleteObject rgn1
            rgn1 = CreateRectRgn(Wi, He, Wi - 1, He - 1)
            CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
            DeleteObject rgn1
        End If
    Case 3, 4 'Windows XP and Mac
        rgn1 = CreateRectRgn(0, 0, 2, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, He, 2, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 0, Wi - 2, 1)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, He, Wi - 2, He - 1)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, 1, 1, 2)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(0, He - 1, 1, He - 2)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, 1, Wi - 1, 2)
        CombineRgn rgn2, rgnNorm, rgn1, RGN_DIFF
        DeleteObject rgn1
        rgn1 = CreateRectRgn(Wi, He - 1, Wi - 1, He - 2)
        CombineRgn rgnNorm, rgn2, rgn1, RGN_DIFF
        DeleteObject rgn1
    Case 13
        DeleteObject rgnNorm
        rgnNorm = CreateEllipticRgn(0, 0, Wi, He)
    End Select
    DeleteObject rgn2

End Sub

Public Property Get MaskColor() As OLE_COLOR

    MaskColor = MaskC

End Property

Public Property Let MaskColor(ByVal theCol As OLE_COLOR)

    MaskC = theCol
    Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "MCOL"

End Property

Public Property Get MouseIcon() As StdPicture

    Set MouseIcon = UserControl.MouseIcon

End Property

Public Property Set MouseIcon(ByVal newIcon As StdPicture)

    On Local Error Resume Next
    Set UserControl.MouseIcon = newIcon
    PropertyChanged "MICON"

End Property

Public Property Get MousePointer() As MousePointerConstants

    MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(ByVal newPointer As MousePointerConstants)

    UserControl.MousePointer = newPointer
    PropertyChanged "MPTR"

End Property

Private Sub mSetPixel(ByVal X As Long, _
                      ByVal Y As Long, _
                      ByVal Color As Long)

    Call SetPixel(UserControl.hDC, X, Y, Color)

End Sub

Private Sub OverTimer_Timer()

    If Not isMouseOver Then
        OverTimer.Enabled = False
        isOver = False
        Call Redraw(0, True)
        RaiseEvent MouseOut
    End If

End Sub

Public Property Get PictureNormal() As StdPicture

    Set PictureNormal = picNormal

End Property

Public Property Set PictureNormal(ByVal newPic As StdPicture)

    Set picNormal = newPic
    Call CalcPicSize
    Call CalcTextRects
    Call Redraw(lastStat, True)
    PropertyChanged "PICN"

End Property

Public Property Get PictureOver() As StdPicture

    Set PictureOver = picHover

End Property

Public Property Set PictureOver(ByVal newPic As StdPicture)

    Set picHover = newPic
    'only redraw i we need to see this picture immediately
    If isOver Then
        Call Redraw(lastStat, True)
    End If
    PropertyChanged "PICO"

End Property

Public Property Get PicturePosition() As PicPositions

    PicturePosition = PicPosition

End Property

Public Property Let PicturePosition(ByVal newPicPos As PicPositions)

    PicPosition = newPicPos
    PropertyChanged "PICPOS"
    Call CalcTextRects
    Call Redraw(lastStat, True)

End Property

Private Sub Redraw(ByVal curStat As Byte, _
                   ByVal Force As Boolean)

'here is the CORE of the button, everything is drawn here
'it's not well commented but i think that everything is
'pretty self explanatory...

Dim i        As Long
Dim stepXP1  As Single
Dim XPFace2  As Long
Dim tempCol  As Long
Dim prevBold As Boolean

    If isCheckbox And cValue Then
        curStat = 2
    End If
    If Not Force Then  'check drawing redundancy
        If (curStat = lastStat) And (TE = elTex) Then
            Exit Sub
        End If
    End If
    'we don't want errors
    If He = 0 Or Not isShown Then
        Exit Sub
    End If
    lastStat = curStat
    TE = elTex
    With UserControl
        .Cls
        If isOver And MyColorType = Custom Then
            tempCol = BackC
            BackC = BackO
            SetColors
        End If
        DrawRectangle 0, 0, Wi, He, cFace
        If isEnabled Then
            If curStat = 0 Then
                '#@#@#@#@#@# BUTTON NORMAL STATE #@#@#@#@#@#
                Select Case MyButtonType
                Case 1 'Windows 16-bit
                    Call DrawCaption(Abs(isOver))
                    DrawFrame cHighLight, cShadow, cHighLight, cShadow, True
                    DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                    Call DrawFocusR
                Case 2 'Windows 32-bit
                    Call DrawCaption(Abs(isOver))
                    If Ambient.DisplayAsDefault And showFocusR Then
                        DrawFrame cHighLight, cDarkShadow, cLight, cShadow, True
                        Call DrawFocusR
                        DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                    Else 'NOT AMBIENT.DISPLAYASDEFAULT...
                        DrawFrame cHighLight, cDarkShadow, cLight, cShadow, False
                    End If
                Case 3 'Windows XP
                    stepXP1 = 25 / He
                    For i = 1 To He
                        DrawLine 0, i, Wi, i, ShiftColor(XPFace, -stepXP1 * i, True)
                    Next i
                    Call DrawCaption(Abs(isOver))
                    DrawRectangle 0, 0, Wi, He, &H733C00, True
                    mSetPixel 1, 1, &H7B4D10
                    mSetPixel 1, He - 2, &H7B4D10
                    mSetPixel Wi - 2, 1, &H7B4D10
                    mSetPixel Wi - 2, He - 2, &H7B4D10
                    If isOver Then
                        DrawRectangle 1, 2, Wi - 2, He - 4, &H31B2FF, True
                        DrawLine 2, He - 2, Wi - 2, He - 2, &H96E7&
                        DrawLine 2, 1, Wi - 2, 1, &HCEF3FF
                        DrawLine 1, 2, Wi - 1, 2, &H8CDBFF
                        DrawLine 2, 3, 2, He - 3, &H6BCBFF
                        DrawLine Wi - 3, 3, Wi - 3, He - 3, &H6BCBFF
                    ElseIf ((HasFocus Or Ambient.DisplayAsDefault) And showFocusR) Then 'ISOVER = FALSE/0
                        DrawRectangle 1, 2, Wi - 2, He - 4, &HE7AE8C, True
                        DrawLine 2, He - 2, Wi - 2, He - 2, &HEF826B
                        DrawLine 2, 1, Wi - 2, 1, &HFFE7CE
                        DrawLine 1, 2, Wi - 1, 2, &HF7D7BD
                        DrawLine 2, 3, 2, He - 3, &HF0D1B5
                        DrawLine Wi - 3, 3, Wi - 3, He - 3, &HF0D1B5
                    Else 'NOT ((HASFOCUS...
                        'we do not draw the bevel always because the above code would repaint over it'NOT ((HASFOCUS...
                        DrawLine 2, He - 2, Wi - 2, He - 2, ShiftColor(XPFace, -&H30, True)
                        DrawLine 1, He - 3, Wi - 2, He - 3, ShiftColor(XPFace, -&H20, True)
                        DrawLine Wi - 2, 2, Wi - 2, He - 2, ShiftColor(XPFace, -&H24, True)
                        DrawLine Wi - 3, 3, Wi - 3, He - 3, ShiftColor(XPFace, -&H18, True)
                        DrawLine 2, 1, Wi - 2, 1, ShiftColor(XPFace, &H10, True)
                        DrawLine 1, 2, Wi - 2, 2, ShiftColor(XPFace, &HA, True)
                        DrawLine 1, 2, 1, He - 2, ShiftColor(XPFace, -&H5, True)
                        DrawLine 2, 3, 2, He - 3, ShiftColor(XPFace, -&HA, True)
                    End If
                Case 4 'Mac
                    DrawRectangle 1, 1, Wi - 2, He - 2, cLight
                    Call DrawCaption(Abs(isOver))
                    DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                    mSetPixel 1, 1, cDarkShadow
                    mSetPixel 1, He - 2, cDarkShadow
                    mSetPixel Wi - 2, 1, cDarkShadow
                    mSetPixel Wi - 2, He - 2, cDarkShadow
                    DrawLine 1, 2, 2, 0, cFace
                    DrawLine 3, 2, Wi - 3, 2, cHighLight
                    DrawLine 2, 2, 2, He - 3, cHighLight
                    mSetPixel 3, 3, cHighLight
                    DrawLine Wi - 3, 1, Wi - 3, He - 3, cFace
                    DrawLine 1, He - 3, Wi - 3, He - 3, cFace
                    mSetPixel Wi - 4, He - 4, cFace
                    DrawLine Wi - 2, 2, Wi - 2, He - 2, cShadow
                    DrawLine 2, He - 2, Wi - 2, He - 2, cShadow
                    mSetPixel Wi - 3, He - 3, cShadow
                Case 5 'Java
                    DrawRectangle 1, 1, Wi - 1, He - 1, ShiftColor(cFace, &HC)
                    Call DrawCaption(Abs(isOver))
                    DrawRectangle 1, 1, Wi - 1, He - 1, cHighLight, True
                    DrawRectangle 0, 0, Wi - 1, He - 1, ShiftColor(cShadow, -&H1A), True
                    mSetPixel 1, He - 2, ShiftColor(cShadow, &H1A)
                    mSetPixel Wi - 2, 1, ShiftColor(cShadow, &H1A)
                    If HasFocus And showFocusR Then
                        DrawRectangle rc.Left - 2, rc.Top - 1, fc.X + 4, fc.Y + 2, &HCC9999, True
                    End If
                Case 6 'Netscape
                    Call DrawCaption(Abs(isOver))
                    DrawFrame ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), cShadow, False
                    Call DrawFocusR
                Case 7, 8, 12 'Flat buttons
                    Call DrawCaption(Abs(isOver))
                    If (MyButtonType = [Simple Flat]) Then
                        DrawFrame cHighLight, cShadow, 0, 0, False, True
                    ElseIf isOver Then 'NOT (MYBUTTONTYPE...
                        If MyButtonType = [Flat Highlight] Then
                            DrawFrame cHighLight, cShadow, 0, 0, False, True
                        Else 'NOT MYBUTTONTYPE...
                            DrawFrame cHighLight, cDarkShadow, cLight, cShadow, False, False
                        End If
                    End If
                    Call DrawFocusR
                Case 9 'Office XP
                    If isOver Then
                        DrawRectangle 1, 1, Wi, He, OXPf
                    End If
                    Call DrawCaption(Abs(isOver))
                    If isOver Then
                        DrawRectangle 0, 0, Wi, He, OXPb, True
                    End If
                    Call DrawFocusR
                Case 11 'transparent
                    BitBlt hDC, 0, 0, Wi, He, pDC, 0, 0, vbSrcCopy
                    Call DrawCaption(Abs(isOver))
                    Call DrawFocusR
                Case 13 'Oval
                    DrawEllipse 0, 0, Wi, He, Abs(isOver) * cShadow + Abs(Not isOver) * cFace, cFace
                    Call DrawCaption(Abs(isOver))
                Case 14 'KDE 2
                    If Not isOver Then
                        stepXP1 = 58 / He
                        For i = 1 To He
                            DrawLine 0, i, Wi, i, ShiftColor(cHighLight, -stepXP1 * i)
                        Next i
                    Else 'NOT NOT...
                        DrawRectangle 0, 0, Wi, He, cLight
                    End If
                    If Ambient.DisplayAsDefault Then
                        isShown = False
                        prevBold = Me.FontBold
                        Me.FontBold = True
                    End If
                    Call DrawCaption(Abs(isOver))
                    If Ambient.DisplayAsDefault Then
                        Me.FontBold = prevBold
                        isShown = True
                    End If
                    DrawRectangle 0, 0, Wi, He, ShiftColor(cShadow, -&H32), True
                    DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cFace, -&H9), True
                    DrawRectangle 2, 2, Wi - 4, 2, cHighLight
                    DrawRectangle 2, 4, 2, He - 6, cHighLight
                    Call DrawFocusR
                End Select
                Call DrawPictures(0)
            ElseIf curStat = 2 Then 'NOT CURSTAT...
                '#@#@#@#@#@# BUTTON IS DOWN #@#@#@#@#@#
                Select Case MyButtonType
                Case 1 'Windows 16-bit
                    Call DrawCaption(2)
                    DrawFrame cShadow, cHighLight, cShadow, cHighLight, True
                    DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                    Call DrawFocusR
                Case 2 'Windows 32-bit
                    Call DrawCaption(2)
                    If showFocusR And Ambient.DisplayAsDefault Then
                        DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                        DrawRectangle 1, 1, Wi - 2, He - 2, cShadow, True
                        Call DrawFocusR
                    Else 'NOT SHOWFOCUSR...
                        DrawFrame cDarkShadow, cHighLight, cShadow, cLight, False
                    End If
                Case 3 'Windows XP
                    stepXP1 = 25 / He
                    XPFace2 = ShiftColor(XPFace, -32, True)
                    For i = 1 To He
                        DrawLine 0, He - i, Wi, He - i, ShiftColor(XPFace2, -stepXP1 * i, True)
                    Next i
                    Call DrawCaption(2)
                    DrawRectangle 0, 0, Wi, He, &H733C00, True
                    mSetPixel 1, 1, &H7B4D10
                    mSetPixel 1, He - 2, &H7B4D10
                    mSetPixel Wi - 2, 1, &H7B4D10
                    mSetPixel Wi - 2, He - 2, &H7B4D10
                    DrawLine 2, He - 2, Wi - 2, He - 2, ShiftColor(XPFace2, &H10, True)
                    DrawLine 1, He - 3, Wi - 2, He - 3, ShiftColor(XPFace2, &HA, True)
                    DrawLine Wi - 2, 2, Wi - 2, He - 2, ShiftColor(XPFace2, &H5, True)
                    DrawLine Wi - 3, 3, Wi - 3, He - 3, XPFace
                    DrawLine 2, 1, Wi - 2, 1, ShiftColor(XPFace2, -&H20, True)
                    DrawLine 1, 2, Wi - 2, 2, ShiftColor(XPFace2, -&H18, True)
                    DrawLine 1, 2, 1, He - 2, ShiftColor(XPFace2, -&H20, True)
                    DrawLine 2, 2, 2, He - 2, ShiftColor(XPFace2, -&H16, True)
                Case 4 'Mac
                    DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                    XPFace = ShiftColor(cShadow, -&H10)
                    Call DrawCaption(2)
                    XPFace = ShiftColor(cFace, &H30)
                    DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                    DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, -&H40), True
                    DrawRectangle 2, 2, Wi - 4, He - 4, ShiftColor(cShadow, -&H20), True
                    mSetPixel 2, 2, ShiftColor(cShadow, -&H40)
                    mSetPixel 3, 3, ShiftColor(cShadow, -&H20)
                    mSetPixel 1, 1, cDarkShadow
                    mSetPixel 1, He - 2, cDarkShadow
                    mSetPixel Wi - 2, 1, cDarkShadow
                    mSetPixel Wi - 2, He - 2, cDarkShadow
                    DrawLine Wi - 3, 1, Wi - 3, He - 3, cShadow
                    DrawLine 1, He - 3, Wi - 2, He - 3, cShadow
                    mSetPixel Wi - 4, He - 4, cShadow
                    DrawLine Wi - 2, 3, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                    DrawLine 3, He - 2, Wi - 2, He - 2, ShiftColor(cShadow, -&H10)
                    DrawLine Wi - 2, He - 3, Wi - 4, He - 1, ShiftColor(cShadow, -&H20)
                    mSetPixel 2, He - 2, ShiftColor(cShadow, -&H20)
                    mSetPixel Wi - 2, 2, ShiftColor(cShadow, -&H20)
                Case 5 'Java
                    DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cShadow, &H10), False
                    DrawRectangle 0, 0, Wi - 1, He - 1, ShiftColor(cShadow, -&H1A), True
                    DrawLine Wi - 1, 1, Wi - 1, He, cHighLight
                    DrawLine 1, He - 1, Wi - 1, He - 1, cHighLight
                    SetTextColor .hDC, cTextO
                    DrawText .hDC, elTex, Len(elTex), rc, DT_CENTER
                    If HasFocus And showFocusR Then
                        DrawRectangle rc.Left - 2, rc.Top - 1, fc.X + 4, fc.Y + 2, &HCC9999, True
                    End If
                Case 6 'Netscape
                    Call DrawCaption(2)
                    DrawFrame cShadow, ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), False
                    Call DrawFocusR
                Case 7, 8, 12 'Flat buttons
                    Call DrawCaption(2)
                    If MyButtonType = [3D Hover] Then
                        DrawFrame cDarkShadow, cHighLight, cShadow, cLight, False, False
                    Else 'NOT MYBUTTONTYPE...
                        DrawFrame cShadow, cHighLight, 0, 0, False, True
                    End If
                    Call DrawFocusR
                Case 9 'Office XP
                    If isOver Then
                        DrawRectangle 0, 0, Wi, He, Abs(MyColorType = 2) * ShiftColor(OXPf, -&H20) + Abs(MyColorType <> 2) * ShiftColorOXP(OXPb, &H80)
                    End If
                    Call DrawCaption(2)
                    DrawRectangle 0, 0, Wi, He, OXPb, True
                    Call DrawFocusR
                Case 11 'transparent
                    BitBlt hDC, 0, 0, Wi, He, pDC, 0, 0, vbSrcCopy
                    Call DrawCaption(2)
                    Call DrawFocusR
                Case 13 'Oval
                    DrawEllipse 0, 0, Wi, He, cDarkShadow, ShiftColor(cFace, -&H20)
                    Call DrawCaption(2)
                Case 14 'KDE 2
                    DrawRectangle 1, 1, Wi, He, ShiftColor(cFace, -&H9)
                    DrawRectangle 0, 0, Wi, He, ShiftColor(cShadow, -&H30), True
                    DrawLine 2, He - 2, Wi - 2, He - 2, cHighLight
                    DrawLine Wi - 2, 2, Wi - 2, He - 1, cHighLight
                    Call DrawCaption(7)
                    Call DrawFocusR
                End Select
                Call DrawPictures(1)
            End If
        Else 'ISENABLED = FALSE/0
            '#~#~#~#~#~# DISABLED STATUS #~#~#~#~#~#
            Select Case MyButtonType
            Case 1 'Windows 16-bit
                Call DrawCaption(3)
                DrawFrame cHighLight, cShadow, cHighLight, cShadow, True
                DrawRectangle 0, 0, Wi, He, cDarkShadow, True
            Case 2 'Windows 32-bit
                Call DrawCaption(3)
                DrawFrame cHighLight, cDarkShadow, cLight, cShadow, False
            Case 3 'Windows XP
                DrawRectangle 0, 0, Wi, He, ShiftColor(XPFace, -&H18, True)
                Call DrawCaption(5)
                DrawRectangle 0, 0, Wi, He, ShiftColor(XPFace, -&H54, True), True
                mSetPixel 1, 1, ShiftColor(XPFace, -&H48, True)
                mSetPixel 1, He - 2, ShiftColor(XPFace, -&H48, True)
                mSetPixel Wi - 2, 1, ShiftColor(XPFace, -&H48, True)
                mSetPixel Wi - 2, He - 2, ShiftColor(XPFace, -&H48, True)
            Case 4 'Mac
                DrawRectangle 1, 1, Wi - 2, He - 2, cLight
                Call DrawCaption(3)
                DrawRectangle 0, 0, Wi, He, cDarkShadow, True
                mSetPixel 1, 1, cDarkShadow
                mSetPixel 1, He - 2, cDarkShadow
                mSetPixel Wi - 2, 1, cDarkShadow
                mSetPixel Wi - 2, He - 2, cDarkShadow
                DrawLine 1, 2, 2, 0, cFace
                DrawLine 3, 2, Wi - 3, 2, cHighLight
                DrawLine 2, 2, 2, He - 3, cHighLight
                mSetPixel 3, 3, cHighLight
                DrawLine Wi - 3, 1, Wi - 3, He - 3, cFace
                DrawLine 1, He - 3, Wi - 3, He - 3, cFace
                mSetPixel Wi - 4, He - 4, cFace
                DrawLine Wi - 2, 2, Wi - 2, He - 2, cShadow
                DrawLine 2, He - 2, Wi - 2, He - 2, cShadow
                mSetPixel Wi - 3, He - 3, cShadow
            Case 5 'Java
                Call DrawCaption(4)
                DrawRectangle 0, 0, Wi, He, cShadow, True
            Case 6 'Netscape
                Call DrawCaption(4)
                DrawFrame ShiftColor(cLight, &H8), cShadow, ShiftColor(cLight, &H8), cShadow, False
            Case 7, 8, 12, 13 'Flat buttons
                Call DrawCaption(3)
                If MyButtonType = [Simple Flat] Then
                    DrawFrame cHighLight, cShadow, 0, 0, False, True
                End If
            Case 9 'Office XP
                Call DrawCaption(4)
            Case 11 'transparent
                BitBlt hDC, 0, 0, Wi, He, pDC, 0, 0, vbSrcCopy
                Call DrawCaption(3)
            Case 14 'KDE 2
                stepXP1 = 58 / He
                For i = 1 To He
                    DrawLine 0, i, Wi, i, ShiftColor(cHighLight, -stepXP1 * i)
                Next i
                DrawRectangle 0, 0, Wi, He, ShiftColor(cShadow, -&H32), True
                DrawRectangle 1, 1, Wi - 2, He - 2, ShiftColor(cFace, -&H9), True
                DrawRectangle 2, 2, Wi - 4, 2, cHighLight
                DrawRectangle 2, 4, 2, He - 6, cHighLight
                Call DrawCaption(6)
            End Select
            Call DrawPictures(1)
        End If
    End With 'USERCONTROL
    If isOver And MyColorType = Custom Then
        BackC = tempCol
        SetColors
    End If

End Sub

Public Sub Refresh()

    If MyButtonType = 11 Then
        Call GetParentPic
    End If
    Call SetColors
    Call CalcTextRects
    isShown = True
    Call Redraw(lastStat, True)

End Sub

Private Sub SetAccessKeys()

'this is a TRUE access keys parser
'the basic rule is that if an ampersand is followed by another,
'  a single ampersand is drawn and this is not the access key.
'  So we continue searching for another possible access key.
'   I only do a second pass because no one writes text like "Me & them & everyone"
'   so the caption prop should be "Me && them && &everyone", this is rubbish and a
'   search like this would only waste time

Dim ampersandPos As Long

    'we first clear the AccessKeys property, and will be filled if one is found
    UserControl.AccessKeys = vbNullString
    If Len(elTex) > 1 Then
        ampersandPos = InStr(1, elTex, "&", vbTextCompare)
        If (ampersandPos < Len(elTex)) And (ampersandPos > 0) Then
            If Mid$(elTex, ampersandPos + 1, 1) <> "&" Then
                'if text is sonething like && then no access key should be assigned, so continue searching
                UserControl.AccessKeys = LCase$(Mid$(elTex, ampersandPos + 1, 1))
            Else 'do only a second pass to find another ampersand character'NOT MID$(ELTEX,...
                ampersandPos = InStr(ampersandPos + 2, elTex, "&", vbTextCompare)
                If Mid$(elTex, ampersandPos + 1, 1) <> "&" Then
                    UserControl.AccessKeys = LCase$(Mid$(elTex, ampersandPos + 1, 1))
                End If
            End If
        End If
    End If

End Sub

Private Sub SetColors()

'this function sets the colors taken as a base to build
'all the other colors and styles.

    If MyColorType = Custom Then
        cFace = ConvertFromSystemColor(BackC)
        cFaceO = ConvertFromSystemColor(BackO)
        cText = ConvertFromSystemColor(ForeC)
        cTextO = ConvertFromSystemColor(ForeO)
        cShadow = ShiftColor(cFace, -&H40)
        cLight = ShiftColor(cFace, &H1F)
        cHighLight = ShiftColor(cFace, &H2F) 'it should be 3F but it looks too lighter
        cDarkShadow = ShiftColor(cFace, -&HC0)
        OXPb = ShiftColor(cFace, -&H80)
        OXPf = cFace
    ElseIf MyColorType = [Force Standard] Then 'NOT MYCOLORTYPE...
        cFace = &HC0C0C0
        cFaceO = cFace
        cShadow = &H808080
        cLight = &HDFDFDF
        cDarkShadow = &H0
        cHighLight = &HFFFFFF
        cText = &H0
        cTextO = cText
        OXPb = &H800000
        OXPf = &HD1ADAD
    ElseIf MyColorType = [Use Container] Then 'NOT MYCOLORTYPE...
        cFace = GetBkColor(GetDC(GetParent(hwnd)))
        cFaceO = cFace
        cText = GetTextColor(GetDC(GetParent(hwnd)))
        cTextO = cText
        cShadow = ShiftColor(cFace, -&H40)
        cLight = ShiftColor(cFace, &H1F)
        cHighLight = ShiftColor(cFace, &H2F)
        cDarkShadow = ShiftColor(cFace, -&HC0)
        OXPb = GetSysColor(COLOR_HIGHLIGHT)
        OXPf = ShiftColorOXP(OXPb)
    Else 'NOT MYCOLORTYPE...
        'if MyColorType is 1 or has not been set then use windows colors
        cFace = GetSysColor(COLOR_BTNFACE)
        cFaceO = cFace
        cShadow = GetSysColor(COLOR_BTNSHADOW)
        cLight = GetSysColor(COLOR_BTNLIGHT)
        cDarkShadow = GetSysColor(COLOR_BTNDKSHADOW)
        cHighLight = GetSysColor(COLOR_BTNHIGHLIGHT)
        cText = GetSysColor(COLOR_BTNTEXT)
        cTextO = cText
        OXPb = GetSysColor(COLOR_HIGHLIGHT)
        OXPf = ShiftColorOXP(OXPb)
    End If
    cMask = ConvertFromSystemColor(MaskC)
    XPFace = ShiftColor(cFace, &H30, MyButtonType = [Windows XP])

End Sub

Private Function ShiftColor(ByVal Color As Long, _
                            ByVal Value As Long, _
                            Optional isXP As Boolean = False) As Long

'this function will add or remove a certain color
'quantity and return the result

Dim Red   As Long
Dim Blue  As Long
Dim Green As Long

    'this is just a tricky way to do it and will result in weird colors for WinXP and KDE2
    If isSoft Then
        Value = Value \ 2
    End If
    If Not isXP Then 'for XP button i use a work-aroud that works fine
        Blue = ((Color \ &H10000) Mod &H100) + Value
    Else 'NOT NOT...
        Blue = ((Color \ &H10000) Mod &H100)
        Blue = Blue + ((Blue * Value) \ &HC0)
    End If
    Green = ((Color \ &H100) Mod &H100) + Value
    Red = (Color And &HFF) + Value
    'a bit of optimization done here, values will overflow a
    ' byte only in one direction... eg: if we added 32 to our
    ' color, then only a > 255 overflow can occurr.
    If Value > 0 Then
        If Red > 255 Then
            Red = 255
        End If
        If Green > 255 Then
            Green = 255
        End If
        If Blue > 255 Then
            Blue = 255
        End If
    ElseIf Value < 0 Then 'NOT VALUE...
        If Red < 0 Then
            Red = 0
        End If
        If Green < 0 Then
            Green = 0
        End If
        If Blue < 0 Then
            Blue = 0
        End If
    End If
    'more optimization by replacing the RGB function by its correspondent calculation
    ShiftColor = Red + 256& * Green + 65536 * Blue

End Function

Private Function ShiftColorOXP(ByVal theColor As Long, _
                               Optional ByVal Base As Long = &HB0) As Long

Dim Red   As Long
Dim Blue  As Long
Dim Green As Long
Dim Delta As Long

    Blue = ((theColor \ &H10000) Mod &H100)
    Green = ((theColor \ &H100) Mod &H100)
    Red = (theColor And &HFF)
    Delta = &HFF - Base
    Blue = Base + Blue * Delta \ &HFF
    Green = Base + Green * Delta \ &HFF
    Red = Base + Red * Delta \ &HFF
    If Red > 255 Then
        Red = 255
    End If
    If Green > 255 Then
        Green = 255
    End If
    If Blue > 255 Then
        Blue = 255
    End If
    ShiftColorOXP = Red + 256& * Green + 65536 * Blue

End Function

Public Property Get ShowFocusRect() As Boolean

    ShowFocusRect = showFocusR

End Property

Public Property Let ShowFocusRect(ByVal newValue As Boolean)

    showFocusR = newValue
    Call Redraw(lastStat, True)
    PropertyChanged "FOCUSR"

End Property

Public Property Get SoftBevel() As Boolean

    SoftBevel = isSoft

End Property

Public Property Let SoftBevel(ByVal newValue As Boolean)

    isSoft = newValue
    Call SetColors
    Call Redraw(lastStat, True)
    PropertyChanged "SOFT"

End Property

Public Property Get SpecialEffect() As fx

    SpecialEffect = SFX

End Property

Public Property Let SpecialEffect(ByVal newValue As fx)

    SFX = newValue
    Call Redraw(lastStat, True)
    PropertyChanged "FX"

End Property

Private Sub TransBlt(ByVal DstDC As Long, _
                     ByVal DstX As Long, _
                     ByVal DstY As Long, _
                     ByVal DstW As Long, _
                     ByVal DstH As Long, _
                     ByVal SrcPic As StdPicture, _
                     Optional ByVal TransColor As Long = -1, _
                     Optional ByVal BrushColor As Long = -1, _
                     Optional ByVal MonoMask As Boolean = False, _
                     Optional ByVal isGreyscale As Boolean = False, _
                     Optional ByVal XPBlend As Boolean = False)

Dim B        As Long
Dim H        As Long
Dim F        As Long
Dim i        As Long
Dim newW     As Long
Dim TmpDC    As Long
Dim TmpBmp   As Long
Dim TmpObj   As Long
Dim Sr2DC    As Long
Dim Sr2Bmp   As Long
Dim Sr2Obj   As Long
Dim Data1()  As RGBTRIPLE
Dim Data2()  As RGBTRIPLE
Dim Info     As BITMAPINFO
Dim BrushRGB As RGBTRIPLE
Dim gCol     As Long
Dim SrcDC    As Long
Dim tObj     As Long
Dim ttt      As Long
Dim br       As RECT
Dim hBrush   As Long

    If DstW = 0 Or DstH = 0 Then
        Exit Sub
    End If
    SrcDC = CreateCompatibleDC(hDC)
    If DstW < 0 Then
        DstW = UserControl.ScaleX(SrcPic.Width, 8, UserControl.ScaleMode)
    End If
    If DstH < 0 Then
        DstH = UserControl.ScaleY(SrcPic.Height, 8, UserControl.ScaleMode)
    End If
    If SrcPic.Type = 1 Then 'check if it's an icon or a bitmap
        tObj = SelectObject(SrcDC, SrcPic)
    Else 'NOT SRCPIC.TYPE...
        br.Right = DstW
        br.Bottom = DstH
        ttt = CreateCompatibleBitmap(DstDC, DstW, DstH)
        tObj = SelectObject(SrcDC, ttt)
        hBrush = CreateSolidBrush(MaskColor)
        FillRect SrcDC, br, hBrush
        DeleteObject hBrush
        DrawIconEx SrcDC, 0, 0, SrcPic.Handle, 0, 0, 0, 0, &H1 Or &H2
    End If
    TmpDC = CreateCompatibleDC(SrcDC)
    Sr2DC = CreateCompatibleDC(SrcDC)
    TmpBmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    Sr2Bmp = CreateCompatibleBitmap(DstDC, DstW, DstH)
    TmpObj = SelectObject(TmpDC, TmpBmp)
    Sr2Obj = SelectObject(Sr2DC, Sr2Bmp)
    ReDim Data1(DstW * DstH * 3 - 1)
    ReDim Data2(UBound(Data1))
    With Info.bmiHeader
        .biSize = Len(Info.bmiHeader)
        .biWidth = DstW
        .biHeight = DstH
        .biPlanes = 1
        .biBitCount = 24
    End With 'INFO.BMIHEADER
    BitBlt TmpDC, 0, 0, DstW, DstH, DstDC, DstX, DstY, vbSrcCopy
    BitBlt Sr2DC, 0, 0, DstW, DstH, SrcDC, 0, 0, vbSrcCopy
    GetDIBits TmpDC, TmpBmp, 0, DstH, Data1(0), Info, 0
    GetDIBits Sr2DC, Sr2Bmp, 0, DstH, Data2(0), Info, 0
    If BrushColor > 0 Then
        With BrushRGB
            .rgbBlue = (BrushColor \ &H10000) Mod &H100
            .rgbGreen = (BrushColor \ &H100) Mod &H100
            .rgbRed = BrushColor And &HFF
        End With 'BrushRGB
    End If
    If Not useMask Then
        TransColor = -1
    End If
    newW = DstW - 1
    For H = 0 To DstH - 1
        F = H * DstW
        For B = 0 To newW
            i = F + B
            If (CLng(Data2(i).rgbRed) + 256& * Data2(i).rgbGreen + 65536 * Data2(i).rgbBlue) <> TransColor Then
                With Data1(i)
                    If BrushColor > -1 Then
                        If MonoMask Then
                            If (CLng(Data2(i).rgbRed) + Data2(i).rgbGreen + Data2(i).rgbBlue) <= 384 Then
                                Data1(i) = BrushRGB
                            End If
                        Else 'MONOMASK = FALSE/0
                            Data1(i) = BrushRGB
                        End If
                    Else 'NOT BRUSHCOLOR...
                        If isGreyscale Then
                            gCol = CLng(Data2(i).rgbRed * 0.3) + Data2(i).rgbGreen * 0.59 + Data2(i).rgbBlue * 0.11
                            .rgbRed = gCol
                            .rgbGreen = gCol
                            .rgbBlue = gCol
                        Else 'ISGREYSCALE = FALSE/0
                            If XPBlend Then
                                .rgbRed = (CLng(.rgbRed) + Data2(i).rgbRed * 2) \ 3
                                .rgbGreen = (CLng(.rgbGreen) + Data2(i).rgbGreen * 2) \ 3
                                .rgbBlue = (CLng(.rgbBlue) + Data2(i).rgbBlue * 2) \ 3
                            Else 'XPBLEND = FALSE/0
                                Data1(i) = Data2(i)
                            End If
                        End If
                    End If
                End With 'DATA1(I)
            End If
        Next B
    Next H
    SetDIBitsToDevice DstDC, DstX, DstY, DstW, DstH, 0, 0, 0, DstH, Data1(0), Info, 0
    Erase Data1, Data2
    DeleteObject SelectObject(TmpDC, TmpObj)
    DeleteObject SelectObject(Sr2DC, Sr2Obj)
    If SrcPic.Type = 3 Then
        DeleteObject SelectObject(SrcDC, tObj)
    End If
    DeleteDC TmpDC
    DeleteDC Sr2DC
    DeleteObject tObj
    DeleteObject ttt
    DeleteDC SrcDC

End Sub

Private Sub txtFX(ByRef theRect As RECT)

Dim curFace As Long
Dim tempR   As RECT

    If SFX > cbNone Then
        With UserControl
            CopyRect tempR, theRect
            OffsetRect tempR, 1, 1
            Select Case MyButtonType
            Case 3, 4, 14
                curFace = XPFace
            Case Else
                If lastStat = 0 And isOver And MyColorType <> Custom And MyButtonType = [Office XP] Then
                    curFace = OXPf
                Else 'NOT LASTSTAT...
                    curFace = cFace
                End If
            End Select
            SetTextColor .hDC, ShiftColor(curFace, Abs(SFX = cbEngraved) * FXDEPTH + (SFX <> cbEngraved) * FXDEPTH)
            DrawText .hDC, elTex, Len(elTex), tempR, DT_CENTER
            If SFX < cbShadowed Then
                OffsetRect tempR, -2, -2
                SetTextColor .hDC, ShiftColor(curFace, Abs(SFX <> cbEngraved) * FXDEPTH + (SFX = cbEngraved) * FXDEPTH)
                DrawText .hDC, elTex, Len(elTex), tempR, DT_CENTER
            End If
        End With 'USERCONTROL
    End If

End Sub

Public Property Get UseGreyscale() As Boolean

    UseGreyscale = useGrey

End Property

Public Property Let UseGreyscale(ByVal newValue As Boolean)

    useGrey = newValue
    If Not picNormal Is Nothing Then
        Call Redraw(lastStat, True)
    End If
    PropertyChanged "NGREY"

End Property

Public Property Get UseMaskColor() As Boolean

    UseMaskColor = useMask

End Property

Public Property Let UseMaskColor(ByVal newValue As Boolean)

    useMask = newValue
    If Not picNormal Is Nothing Then
        Call Redraw(lastStat, True)
    End If
    PropertyChanged "UMCOL"

End Property

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)

    LastButton = 1
    Call UserControl_Click

End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)

    If Not MyColorType = [Custom] Then
        Call SetColors
        Call Redraw(lastStat, True)
    End If

End Sub

Private Sub UserControl_Click()

    If LastButton = 1 And isEnabled Then
        If isCheckbox Then
            cValue = Not cValue
        End If
        Call Redraw(0, True) 'be sure that the normal status is drawn
        UserControl.Refresh
        RaiseEvent Click
    End If

End Sub

Private Sub UserControl_DblClick()

    If LastButton = 1 Then
        Call UserControl_MouseDown(1, 0, 0, 0)
        SetCapture hwnd
    End If

End Sub

Private Sub UserControl_GotFocus()

    HasFocus = True
    Call Redraw(lastStat, True)

End Sub

Private Sub UserControl_Hide()

    isShown = False

End Sub

Private Sub UserControl_Initialize()

'this makes the control to be slow, remark this line if the "not redrawing" problem is not important for you: ie, you intercept the Load_Event (with breakpoint or messageBox) and the button does not repaint...

    isShown = True

End Sub

Private Sub UserControl_InitProperties()

    isEnabled = True
    showFocusR = True
    useMask = True
    elTex = Ambient.DisplayName
    Set UserControl.Font = Ambient.Font
    MyButtonType = [Windows 32-bit]
    MyColorType = [Use Windows]
    Call SetColors
    BackC = cFace
    BackO = BackC
    ForeC = cText
    ForeO = ForeC
    MaskC = &HC0C0C0
    Call CalcTextRects

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, _
                                Shift As Integer)

    RaiseEvent KeyDown(KeyCode, Shift)
    LastKeyDown = KeyCode
    Select Case KeyCode
    Case 32 'spacebar pressed
        Call Redraw(2, False)
    Case 39, 40 'right and down arrows
        SendKeys "{Tab}"
    Case 37, 38 'left and up arrows
        SendKeys "+{Tab}"
    End Select

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, _
                              Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)
    If (KeyCode = 32) And (LastKeyDown = 32) Then 'spacebar pressed, and not cancelled by the user
        If isCheckbox Then
            cValue = Not cValue
        End If
        Call Redraw(0, False)
        UserControl.Refresh
        RaiseEvent Click
    End If

End Sub

Private Sub UserControl_LostFocus()

    HasFocus = False
    Call Redraw(lastStat, True)

End Sub

Private Sub UserControl_MouseDown(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)
    LastButton = Button
    If Button <> 2 Then
        Call Redraw(2, False)
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, _
                                  Shift As Integer, _
                                  X As Single, _
                                  Y As Single)

    RaiseEvent MouseMove(Button, Shift, X, Y)
    If Button < 2 Then
        If Not isMouseOver Then
            'we are outside the button
            Call Redraw(0, False)
        Else 'NOT NOT...
            'we are inside the button
            If Button = 0 And Not isOver Then
                OverTimer.Enabled = True
                isOver = True
                Call Redraw(0, True)
                RaiseEvent MouseOver
            ElseIf Button = 1 Then 'NOT BUTTON...
                isOver = True
                Call Redraw(2, False)
                isOver = False
            End If
        End If
    End If

End Sub

Private Sub UserControl_MouseUp(Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

    RaiseEvent MouseUp(Button, Shift, X, Y)
    If Button <> 2 Then
        Call Redraw(0, False)
    End If

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        MyButtonType = .ReadProperty("BTYPE", 2)
        elTex = .ReadProperty("TX", vbNullString)
        isEnabled = .ReadProperty("ENAB", True)
        Set UserControl.Font = .ReadProperty("FONT", UserControl.Font)
        MyColorType = .ReadProperty("COLTYPE", 1)
        showFocusR = .ReadProperty("FOCUSR", True)
        BackC = .ReadProperty("BCOL", GetSysColor(COLOR_BTNFACE))
        BackO = .ReadProperty("BCOLO", BackC)
        ForeC = .ReadProperty("FCOL", GetSysColor(COLOR_BTNTEXT))
        ForeO = .ReadProperty("FCOLO", ForeC)
        MaskC = .ReadProperty("MCOL", &HC0C0C0)
        UserControl.MousePointer = .ReadProperty("MPTR", 0)
        Set UserControl.MouseIcon = .ReadProperty("MICON", Nothing)
        Set picNormal = .ReadProperty("PICN", Nothing)
        Set picHover = .ReadProperty("PICH", Nothing)
        useMask = .ReadProperty("UMCOL", True)
        isSoft = .ReadProperty("SOFT", False)
        PicPosition = .ReadProperty("PICPOS", 0)
        useGrey = .ReadProperty("NGREY", False)
        SFX = .ReadProperty("FX", 0)
        Me.HandPointer = .ReadProperty("HAND", False)
        isCheckbox = .ReadProperty("CHECK", False)
        cValue = .ReadProperty("VALUE", False)
    End With 'PROPBAG
    UserControl.Enabled = isEnabled
    Call CalcPicSize
    Call CalcTextRects
    Call SetAccessKeys

End Sub

'########## END OF PROPERTIES ##########
Private Sub UserControl_Resize()

    If inLoop Then
        Exit Sub
    End If
    'get button size
    GetClientRect UserControl.hwnd, rc3
    'assign these values to He and Wi
    He = rc3.Bottom
    Wi = rc3.Right
    'build the FocusRect size and position depending on the button type
    If MyButtonType >= [Simple Flat] And MyButtonType <= [Oval Flat] Then
        InflateRect rc3, -3, -3
    ElseIf MyButtonType = [KDE 2] Then 'NOT MYBUTTONTYPE...
        InflateRect rc3, -5, -5
        OffsetRect rc3, 1, 1
    Else 'NOT MYBUTTONTYPE...
        InflateRect rc3, -4, -4
    End If
    Call CalcTextRects
    If rgnNorm Then
        DeleteObject rgnNorm
    End If
    Call MakeRegion
    SetWindowRgn UserControl.hwnd, rgnNorm, True
    If He Then
        Call Redraw(0, True)
    End If

End Sub

Private Sub UserControl_Show()

    If MyButtonType = 11 Then
        If pDC = 0 Then
            pDC = CreateCompatibleDC(UserControl.hDC)
            pBM = CreateBitmap(Wi, He, 1, GetDeviceCaps(hDC, 12), ByVal 0&)
            oBM = SelectObject(pDC, pBM)
        End If
        Call GetParentPic
    End If
    isShown = True
    Call SetColors
    Call Redraw(0, True)

End Sub

Private Sub UserControl_Terminate()

    isShown = False
    DeleteObject rgnNorm
    If pDC Then
        DeleteObject SelectObject(pDC, oBM)
        DeleteDC pDC
    End If

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("BTYPE", MyButtonType)
        Call .WriteProperty("TX", elTex)
        Call .WriteProperty("ENAB", isEnabled)
        Call .WriteProperty("FONT", UserControl.Font)
        Call .WriteProperty("COLTYPE", MyColorType)
        Call .WriteProperty("FOCUSR", showFocusR)
        Call .WriteProperty("BCOL", BackC)
        Call .WriteProperty("BCOLO", BackO)
        Call .WriteProperty("FCOL", ForeC)
        Call .WriteProperty("FCOLO", ForeO)
        Call .WriteProperty("MCOL", MaskC)
        Call .WriteProperty("MPTR", UserControl.MousePointer)
        Call .WriteProperty("MICON", UserControl.MouseIcon)
        Call .WriteProperty("PICN", picNormal)
        Call .WriteProperty("PICH", picHover)
        Call .WriteProperty("UMCOL", useMask)
        Call .WriteProperty("SOFT", isSoft)
        Call .WriteProperty("PICPOS", PicPosition)
        Call .WriteProperty("NGREY", useGrey)
        Call .WriteProperty("FX", SFX)
        Call .WriteProperty("HAND", useHand)
        Call .WriteProperty("CHECK", isCheckbox)
        Call .WriteProperty("VALUE", cValue)
    End With 'PROPBAG

End Sub

Public Property Get Value() As Boolean

    Value = cValue

End Property

Public Property Let Value(ByVal newValue As Boolean)

    cValue = newValue
    If isCheckbox Then
        Call Redraw(0, True)
    End If
    PropertyChanged "VALUE"

End Property

Public Property Get Version() As String

    Version = cbVersion

End Property



