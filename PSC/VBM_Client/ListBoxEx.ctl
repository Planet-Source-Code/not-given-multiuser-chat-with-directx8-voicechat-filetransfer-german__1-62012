VERSION 5.00
Begin VB.UserControl ListBoxEX 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2325
   EditAtDesignTime=   -1  'True
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   OLEDropMode     =   1  'Manuell
   ScaleHeight     =   163
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   155
   ToolboxBitmap   =   "ListBoxEx.ctx":0000
   Begin VB.VScrollBar VScroll 
      Height          =   2415
      LargeChange     =   10
      Left            =   2040
      Max             =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox picDraw 
      Align           =   3  'Links ausrichten
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      DrawStyle       =   2  'Punkt
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Ausgefüllt
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2445
      Left            =   0
      OLEDropMode     =   1  'Manuell
      ScaleHeight     =   163
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "ListBoxEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^¶¶^^^^^^¶¶^^^^^^^^^^^^^¶¶¶¶¶¶¶^^^^^^^^^^^^^^^^¶¶¶¶¶¶¶^^^^^^^^^^^^^¶¶^^^^¶¶^^^¶¶^^^^^$
'$^^^^¶¶^^^^^^^^^^^^^^^^^¶¶^^¶¶^^^^¶¶^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^¶¶^^^^¶¶^^¶¶¶^^^^^$
'$^^^^¶¶^^^^^^^^^^^^^^^^^¶¶^^¶¶^^^^¶¶^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^¶¶^^¶¶^^^^¶¶^^^^^$
'$^^^^¶¶^^^^^^¶¶^^^¶¶¶¶^^¶¶¶^¶¶^^^^¶¶^^¶¶¶¶^^¶¶^^¶¶^¶¶^^^^^^^¶¶^^¶¶^^^^^^¶¶^^¶¶^^^^¶¶^^^^^$
'$^^^^¶¶^^^^^^¶¶^^¶¶^^¶¶^¶¶^^¶¶¶¶¶¶¶^^¶¶^^¶¶^^¶¶¶¶^^¶¶¶¶¶¶^^^^¶¶¶¶^^^^^^^¶¶^^¶¶^^^^¶¶^^^^^$
'$^^^^¶¶^^^^^^¶¶^^^¶¶¶^^^¶¶^^¶¶^^^^¶¶^¶¶^^¶¶^^^¶¶^^^¶¶^^^^^^^^^¶¶^^^^^^^^^¶¶¶¶^^^^^¶¶^^^^^$
'$^^^^¶¶^^^^^^¶¶^^^^^¶¶^^¶¶^^¶¶^^^^¶¶^¶¶^^¶¶^^^¶¶^^^¶¶^^^^^^^^^¶¶^^^^^^^^^¶¶¶¶^^^^^¶¶^^^^^$
'$^^^^¶¶^^^^^^¶¶^^¶¶^^¶¶^¶¶^^¶¶^^^^¶¶^¶¶^^¶¶^^¶¶¶¶^^¶¶^^^^^^^^¶¶¶¶^^^^^^^^^¶¶^^^^^^¶¶^^^^^$
'$^^^^¶¶¶¶¶¶¶^¶¶^^^¶¶¶¶^^^¶¶^¶¶¶¶¶¶¶^^^¶¶¶¶^^¶¶^^¶¶^¶¶¶¶¶¶¶^^¶¶^^¶¶^^^^^^^^¶¶^^^^^^¶¶^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^¶¶¶¶¶¶^^^^^^^^^^^^^^^¶^¶^^^^^^^^^^^^^^^^^^^¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^¶^^^^^¶^^^^^^^^^^^^^^¶^^^^^^^^^^^^^^^^^^^^^¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^¶^^^^^¶^^^^^^^^^^^^^^¶^^^^^^^^^^^^^^^^^^^^^¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^¶^^^^^¶^¶^^^¶^^^^^^^^¶^¶^^¶¶¶¶^¶¶¶^^^^^^^^^¶^^¶¶¶^^^¶¶¶^^^¶¶¶^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^¶¶¶¶¶¶^^¶^^^¶^^^^^^^^¶^¶^^¶^^^¶^^^¶^^^^^^^^¶^¶^^^¶^¶^^^¶^¶^^^¶^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^¶^^^^^¶^¶^^^¶^^^^^^^^¶^¶^^¶^^^¶^^^¶^^^^^^^^¶^¶^^^¶^^¶¶^^^¶¶¶¶¶^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^¶^^^^^¶^¶^^^¶^^^^¶^^^¶^¶^^¶^^^¶^^^¶^^^^¶^^^¶^¶^^^¶^^^^¶^^¶^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^¶^^^^^¶^^¶^¶^^^^^¶^^^¶^¶^^¶^^^¶^^^¶^^^^¶^^^¶^¶^^^¶^¶^^^¶^¶^^^¶^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^¶¶¶¶¶¶^^^^¶^^^^^^^¶¶¶^^¶^^¶^^^¶^^^¶^^^^^¶¶¶^^^¶¶¶^^^¶¶¶^^^¶¶¶^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'------------------------------------------------------------------------------------------
' SourceCode : ListBoxEX V1
' Auther     : Jim Jose
' Email      : jimjosev33@yahoo.com
' Date       : 3-6-2005
' Purpose    : An upgraded version of VBListBox with Icons and many more
' Comment    : This is the first version of this control.
'            : This version aimed for a clear and simple code.
'            : Use your imaginations to visualize more features.
'            : Please send me your better ideas and additional features you need.
' CopyRight  : JimJose © Gtech Creations - 2005
'------------------------------------------------------------------------------------------
Option Explicit
'[APIs]
'[Types]
Private Type RECT
    Left                                   As Long
    Top                                    As Long
    Right                                  As Long
    Bottom                                 As Long
End Type
'[Enums]
Public Enum AppearanceEnum
    [Flat] = 0
    [3D] = 1
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Flat
#End If
Public Enum BorderEnum
    [None] = 0
    [Fixed Single] = 1
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private None
#End If
Public Enum SortOrderEnum
    [Ascending] = -1
    [Desending] = 1
End Enum
#If False Then 'Trick preserves Case of Enums when typing in IDE
Private Ascending, Desending
#End If
'[Local Variables]
Private m_SelItem                      As Long
Private m_iHeight                      As Double
Private m_iCount                       As Long
Private m_iTop                         As Long
Private m_hMode                        As Long
Private m_KeyControl                   As Boolean
'[Data Storage]
Private m_Items                        As New Collection
'[Property Variables]
Private m_Picture                      As New StdPicture
Private m_ListIcon                     As New StdPicture
Private m_BackColor                    As OLE_COLOR
Private m_ForeColor                    As OLE_COLOR
Private m_Font                         As New StdFont
Private m_SelColor                     As OLE_COLOR
Private m_FullRowSel                   As Boolean
Private m_SortOrder                    As SortOrderEnum
Private m_SortItems                    As Boolean
Private m_SelForeColor                 As OLE_COLOR
Private m_StrechIcon                   As Boolean
Private m_IconFocus                    As Boolean
Private m_TextAllineMent               As AlignmentConstants
'[Default Property Values]
Private Const m_def_BackColor          As Long = &HFFFFFF
Private Const m_def_ForeColor          As Long = &H80000012
Private Const m_def_SelColor           As Long = &HFF8C1A
Private Const m_def_SelForeColor       As Long = &HFFFFFF
Private Const m_def_StrechIcon         As Boolean = False
Private Const m_def_Appearance         As Integer = 1
Private Const m_def_BorderStyle        As Integer = 1
Private Const m_def_FullRowSel         As Boolean = False
Private Const m_def_SortOrder          As Long = Ascending
Private Const m_def_SortItems          As Boolean = False
Private Const m_def_IconFocus          As Boolean = True
Private Const m_def_TextAllignMent     As Long = vbLeftJustify
'[Events]
Public Event FileDragDrop(Filename As String)
Public Event MouseClick()
Public Event SelChange()
Public Event DbClick()
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, _
                                                                      ByVal lpStr As String, _
                                                                      ByVal nCount As Long, _
                                                                      ByRef lpRect As RECT, _
                                                                      ByVal wFormat As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hDC As Long, _
                                                         ByRef lpRect As RECT) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, _
                                                    ByVal x1 As Long, _
                                                    ByVal y1 As Long, _
                                                    ByVal x2 As Long, _
                                                    ByVal y2 As Long) As Long

'-------------------------------------------------------------------------
' Procedure  : AddItem
' Auther     : Jim Jose
' Input      : New item
' OutPut     : None
' Purpose    : To add an item to listBox
'-------------------------------------------------------------------------
Public Sub AddItem(vText As String, _
                   Optional vIndex As Long = -1)

    If vIndex = -1 Then
        ' Index not specified , add to last
        m_Items.Add vText
    Else 'NOT VINDEX...
        ' add to specified index
        m_Items.Add vText, , vIndex
    End If
    ' Sort items iff needed
    If m_SortItems Then
        SortList
    End If
    Me.Refresh

End Sub

'-------------------------------------------------------------------------
' Procedure  : Appearance
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property Appearance
'-------------------------------------------------------------------------
Public Property Get Appearance() As AppearanceEnum

    Appearance = UserControl.Appearance

End Property

Public Property Let Appearance(ByVal vNewAppearance As AppearanceEnum)

    UserControl.Appearance = vNewAppearance
    PropertyChanged "Appearance"

End Property

'-------------------------------------------------------------------------
' Procedure  : BackColor
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property BackColor
'-------------------------------------------------------------------------
Public Property Get BackColor() As OLE_COLOR

    BackColor = m_BackColor

End Property

Public Property Let BackColor(ByVal vNewCol As OLE_COLOR)

    m_BackColor = vNewCol
    PropertyChanged "BackColor"
    ReDrawList

End Property

'-------------------------------------------------------------------------
' Procedure  : BorderStyle
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property BorderStyle
'-------------------------------------------------------------------------
Public Property Get BorderStyle() As BorderEnum

    BorderStyle = UserControl.BorderStyle

End Property

Public Property Let BorderStyle(ByVal vNewBorder As BorderEnum)

    UserControl.BorderStyle = vNewBorder
    PropertyChanged "BorderStyle"

End Property

'-------------------------------------------------------------------------
' Procedure  : CheckSelected
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To check if the selected item is in permissible range
'              and reset the scroll bars
'-------------------------------------------------------------------------
Private Sub CheckSelected()

    If m_SelItem > m_Items.Count Then
        m_SelItem = m_Items.Count
    End If
    If m_SelItem < 1 Then
        m_SelItem = 1
    End If
    If m_KeyControl = False Then
        Exit Sub
    End If
    If m_SelItem < m_iTop + 1 Then
        VScroll.Value = m_SelItem - 1
    End If
    If m_SelItem > m_iTop + m_iCount Then
        VScroll.Value = m_SelItem - m_iCount
    End If

End Sub

'-------------------------------------------------------------------------
' Procedure  : Clear
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Clear List
'-------------------------------------------------------------------------
Public Sub Clear()

Dim X As Long

    ' Remove each Item
    For X = 1 To m_Items.Count
        m_Items.Remove (1)
    Next X
    ' Redraw
    picDraw.Cls
    Me.Refresh

End Sub

'-------------------------------------------------------------------------
' Procedure  : Font
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property Font
'-------------------------------------------------------------------------
Public Property Get Font() As Font

    Set Font = m_Font

End Property

Public Property Set Font(ByVal vNewFont As Font)

    Set m_Font = vNewFont
    PropertyChanged "Font"
    Me.Refresh

End Property

'-------------------------------------------------------------------------
' Procedure  : ForeColor
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property ForeColor
'-------------------------------------------------------------------------
Public Property Get ForeColor() As OLE_COLOR

    ForeColor = m_ForeColor

End Property

Public Property Let ForeColor(ByVal vNewCol As OLE_COLOR)

    m_ForeColor = vNewCol
    PropertyChanged "ForeColor"
    ReDrawList

End Property

'-------------------------------------------------------------------------
' Procedure  : FullRowSelect
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Get/Let Property FullRowSelect
'-------------------------------------------------------------------------
Public Property Get FullRowSelect() As Boolean

    FullRowSelect = m_FullRowSel

End Property

Public Property Let FullRowSelect(ByVal vNewValue As Boolean)

    m_FullRowSel = vNewValue
    PropertyChanged "FullRowSelect"
    ReDrawList

End Property

'-------------------------------------------------------------------------
' Procedure  : SortItems
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Get/Let Property SortItems
'-------------------------------------------------------------------------
Public Property Get IconFocus() As Boolean

    IconFocus = m_IconFocus

End Property

Public Property Let IconFocus(ByVal vNewValue As Boolean)

    m_IconFocus = vNewValue
    PropertyChanged "IconFocus"
    ReDrawList

End Property

'-------------------------------------------------------------------------
' Procedure  : IsThere
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To check if the Picture is loaded
'-------------------------------------------------------------------------
Private Function IsThere(vPicture As StdPicture) As Boolean

    On Error GoTo Handle
    If Not vPicture.Handle = 0 And Not vPicture.Height = 0 And Not vPicture.Width = 0 Then
        IsThere = True
    Else 'NOT NOT...
        IsThere = False
    End If

Exit Function

Handle:
    IsThere = False

End Function

'-------------------------------------------------------------------------
' Procedure  : ListCount
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : to get ListCount
'-------------------------------------------------------------------------
Public Function ListCount() As Long

    On Error GoTo Handle
    ListCount = m_Items.Count

Exit Function

Handle:
    ListCount = 0

End Function

'-------------------------------------------------------------------------
' Procedure  : ListIcon
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property ListIcon
'-------------------------------------------------------------------------
Public Property Get ListIcon() As Picture

    Set ListIcon = m_ListIcon

End Property

Public Property Set ListIcon(ByVal vNewPicture As Picture)

    Set m_ListIcon = vNewPicture
    PropertyChanged "ListIcon"
    ReDrawList

End Property

'-------------------------------------------------------------------------
' Procedure  : ListItems
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Get/Let Property ListItems
'-------------------------------------------------------------------------
Public Property Get ListItems(ByVal vIndex As Long) As String

    On Error Resume Next
    ListItems = m_Items(vIndex)

End Property

Public Property Let ListItems(ByVal vIndex As Long, _
                              ByVal vNewValue As String)

    m_Items(vIndex) = vNewValue

End Property

'-------------------------------------------------------------------------
' Procedure  : picDraw_Click
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : RaiseEvent MouseClick
'-------------------------------------------------------------------------
Private Sub picDraw_Click()

    RaiseEvent MouseClick

End Sub

'-------------------------------------------------------------------------
' Procedure  : picDraw_DblClick
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : RaiseEvent DbClick
'-------------------------------------------------------------------------
Private Sub picDraw_DblClick()

    RaiseEvent DbClick

End Sub

'-------------------------------------------------------------------------
' Procedure  : picDraw_KeyDown
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Move Selection by keyboard
'-------------------------------------------------------------------------
Private Sub picDraw_KeyDown(KeyCode As Integer, _
                            Shift As Integer)

' Select each Key

    Select Case KeyCode
    Case vbKeyUp
        m_SelItem = m_SelItem - 1
    Case vbKeyDown
        m_SelItem = m_SelItem + 1
    Case vbKeyEnd
        m_SelItem = ListCount
    Case vbKeyHome
        m_SelItem = 0
    Case vbKeyPageDown
        m_SelItem = m_SelItem + m_iCount
    Case vbKeyPageUp
        m_SelItem = m_SelItem - m_iCount
    End Select
    ' Refrech Control
    Me.Refresh
    RaiseEvent SelChange

End Sub

'-------------------------------------------------------------------------
' Procedure  : picDraw_MouseDown
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To calculate selection by mouse
'-------------------------------------------------------------------------
Private Sub picDraw_MouseDown(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

' Calculate row from mouse 'Y'

    m_SelItem = m_iTop + Int(Y / m_iHeight) + 1
    ReDrawList
    RaiseEvent SelChange
    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

'-------------------------------------------------------------------------
' Procedure  : picDraw_MouseMove
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To RaiseEvent MouseMove
'-------------------------------------------------------------------------
Private Sub picDraw_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

'-------------------------------------------------------------------------
' Procedure  : picDraw_MouseUp
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To RaiseEvent MouseUp
'-------------------------------------------------------------------------
Private Sub picDraw_MouseUp(Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)

    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Private Sub picDraw_OLEDragDrop(Data As DataObject, _
                                Effect As Long, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

Dim i As Integer

    If Data.GetFormat(vbCFFiles) Then
        SelectedItem = m_iTop + Fix(Y / m_iHeight)
        For i = 1 To Data.Files.Count
            RaiseEvent FileDragDrop(Data.Files(i))
        Next i
    End If

End Sub

Private Sub picDraw_OLEDragOver(Data As DataObject, _
                                Effect As Long, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single, _
                                State As Integer)

    SelectedItem = m_iTop + Fix(Y / m_iHeight)

End Sub

'-------------------------------------------------------------------------
' Procedure  : Picture
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property Picture
'-------------------------------------------------------------------------
Public Property Get Picture() As Picture

    Set Picture = m_Picture

End Property

Public Property Set Picture(ByVal vNewPicture As Picture)

    Set m_Picture = vNewPicture
    PropertyChanged "Picture"
    ReDrawList

End Property

'-------------------------------------------------------------------------
' Procedure  : ReDrawList
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To draw the entire region
'-------------------------------------------------------------------------
Private Sub ReDrawList()

Dim X      As Long
Dim Rct    As RECT
Dim vText  As String
Dim vCount As Long
Dim vTop   As Long
Dim vIcon  As Boolean
Dim iLeft  As Long
Dim iTop   As Long

    On Error GoTo Handle
    ' Some initial preperation
    CheckSelected
    picDraw.Cls
    picDraw.BackColor = m_BackColor
    vCount = m_iTop + m_iCount
    picDraw.ForeColor = m_ForeColor
    Set picDraw.Font = m_Font
    If IsThere(m_Picture) Then
        ' kleine Änderung für den VBM
        ' -> pic wird unten gezeichnet
        'picDraw.PaintPicture m_Picture, 0, picDraw.ScaleHeight - m_Picture.Height / Screen.TwipsPerPixelY
    End If
    If vCount > m_Items.Count Then
        vCount = m_Items.Count
    End If
    vIcon = IsThere(m_ListIcon)
    ' Define space for Listicon\Rect
    If vIcon Then
        Rct.Left = m_iHeight + 3
        If m_StrechIcon Then
            iLeft = 1
            iTop = 1
        Else 'M_STRECHICON = FALSE/0
            iLeft = m_iHeight / 2 - ScaleX(m_ListIcon.Width) / 2
            iTop = m_iHeight / 2 - ScaleY(m_ListIcon.Height) / 2
        End If
    Else 'VICON = FALSE/0
        Rct.Left = 3
    End If
    Rct.Right = picDraw.Width
    ' Draw each item
    For X = m_iTop To vCount - 1
        ' Downward shift
        Rct.Top = vTop
        Rct.Bottom = Rct.Top + m_iHeight
        ' Get the item text
        vText = " " & m_Items(X + 1) & " "
        DrawText picDraw.hDC, vText, Len(vText), Rct, m_TextAllineMent
        ' Draw Icons
        If m_StrechIcon Then
            If vIcon Then
                picDraw.PaintPicture m_ListIcon, iLeft, Rct.Top + iTop, m_iHeight - 1, m_iHeight - 1
            End If
        Else 'M_STRECHICON = FALSE/0
            If vIcon Then
                picDraw.PaintPicture m_ListIcon, iLeft, Rct.Top + iTop
            End If
        End If
        ' Downward shift
        vTop = Rct.Bottom
    Next X
    ' Prepare to draw selection
    X = Rct.Left
    If m_FullRowSel Then
        Rct.Left = 0
    End If
    With picDraw
        .DrawStyle = vbSolid
        .FillStyle = vbSolid
        .FillColor = m_SelColor
    End With 'picDraw
    With Rct
        .Top = (m_SelItem - m_iTop - 1) * m_iHeight
        .Bottom = .Top + m_iHeight
        ' Draw the sel back
        Rectangle picDraw.hDC, .Left, .Top, .Right, .Bottom
        ' Draw Focus
    End With 'Rct
    DrawFocusRect picDraw.hDC, Rct
    ' Draw iCon on selection
    If m_StrechIcon Then
        If vIcon Then
            picDraw.PaintPicture m_ListIcon, iLeft, Rct.Top + iTop, m_iHeight - 1, m_iHeight - 1
        End If
    Else 'M_STRECHICON = FALSE/0
        If vIcon Then
            picDraw.PaintPicture m_ListIcon, iLeft, Rct.Top + iTop
        End If
    End If
    ' Draw selected text
    vText = " " & m_Items(m_SelItem) & " "
    picDraw.ForeColor = m_SelForeColor
    Rct.Left = X
    DrawText picDraw.hDC, vText, Len(vText), Rct, m_TextAllineMent
    ' Draw Icon Focus
    If vIcon And m_IconFocus Then
        picDraw.ForeColor = m_ForeColor
        Rct.Left = 1
        Rct.Right = Rct.Left + m_iHeight
        Rct.Bottom = Rct.Top + m_iHeight
        DrawFocusRect picDraw.hDC, Rct
    End If
    ' Refresh Box
    picDraw.Refresh
Handle:

End Sub

'-------------------------------------------------------------------------
' Procedure  : Refresh
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Arrage control and calculate local variables
'-------------------------------------------------------------------------
Public Sub Refresh()

    On Error Resume Next
    ' Determine item height & item cound per Screen
    Set picDraw.Font = m_Font
    m_iHeight = picDraw.TextHeight("A")
    m_iCount = Int(ScaleHeight / m_iHeight)
    ' Arrange\Set controls
    If m_Items.Count > m_iCount Then
        With VScroll
            .Visible = True
            .Move ScaleWidth - .Width, 0, .Width, ScaleHeight
            .Max = m_Items.Count - m_iCount
            picDraw.Move 0, 0, ScaleWidth - .Width, ScaleHeight
        End With 'VScroll
    Else 'NOT M_ITEMS.COUNT...
        VScroll.Value = 0
        VScroll.Visible = False
        picDraw.Move 0, 0, ScaleWidth, ScaleHeight
    End If
    ' Redraw the list
    ReDrawList

End Sub

'-------------------------------------------------------------------------
' Procedure  : Remove
' Auther     : Jim Jose
' Input      : Index
' OutPut     : None
' Purpose    : To remove an item from List
'-------------------------------------------------------------------------
Public Sub Remove(Optional ByVal vIndex As Long = -1)

    If vIndex = -1 Then
        ' Index not specifid, remove selected item
        m_Items.Remove m_SelItem
    Else 'NOT VINDEX...
        ' Remove specified item
        m_Items.Remove vIndex
    End If
    ' Sort If needed
    If m_SortItems Then
        SortList
    End If
    Me.Refresh

End Sub

'-------------------------------------------------------------------------
' Procedure  : SelColor
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property SelColor
'-------------------------------------------------------------------------
Public Property Get SelColor() As OLE_COLOR

    SelColor = m_SelColor

End Property

Public Property Let SelColor(ByVal vNewCol As OLE_COLOR)

    m_SelColor = vNewCol
    PropertyChanged "SelColor"
    ReDrawList

End Property

'-------------------------------------------------------------------------
' Procedure  : SelectedItem
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property SelectedItem
'-------------------------------------------------------------------------
Public Property Get SelectedItem() As Long

    SelectedItem = m_SelItem

End Property

Public Property Let SelectedItem(ByVal vNewValue As Long)

    m_SelItem = vNewValue
    PropertyChanged "SelectedItem"
    CheckSelected
    ReDrawList

End Property

'-------------------------------------------------------------------------
' Procedure  : SelectedText
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To get selected text
'-------------------------------------------------------------------------
Public Property Get SelectedText() As String

    If ListCount = 0 Then
        Exit Property
    End If
    SelectedText = m_Items(m_SelItem)

End Property

'-------------------------------------------------------------------------
' Procedure  : SelForeColor
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property SelForeColor
'-------------------------------------------------------------------------
Public Property Get SelForeColor() As OLE_COLOR

    SelForeColor = m_SelForeColor

End Property

Public Property Let SelForeColor(ByVal vNewCol As OLE_COLOR)

    m_SelForeColor = vNewCol
    PropertyChanged "SelForeColor"
    ReDrawList

End Property

'-------------------------------------------------------------------------
' Procedure  : SortItems
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Get/Let Property SortItems
'-------------------------------------------------------------------------
Public Property Get SortItems() As Boolean

    SortItems = m_SortItems

End Property

Public Property Let SortItems(ByVal vNewValue As Boolean)

    m_SortItems = vNewValue
    PropertyChanged "SortItems"
    If m_SortItems Then
        SortList
    End If
    ReDrawList

End Property

'-------------------------------------------------------------------------
' Procedure  : SortList
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To sort the Data-Collection Ascending/Descending
'-------------------------------------------------------------------------
Private Sub SortList()

Dim X         As Long
Dim vPos      As Long
Dim vRtn      As Long
Dim vCount    As Long
Dim vStart    As Long
Dim vNewCount As Long
Dim vNew      As New Collection

    ' Get current Count
    vStart = 1
    vCount = m_Items.Count
    ' Loop through Current collection
    For X = vStart To vCount
        ' Get new collection count
        vNewCount = vNew.Count
        ' Loop through new collection
        For vPos = 1 To vNewCount
            ' Compair each item in new collection
            vRtn = StrComp(m_Items(X), vNew(vPos), vbTextCompare)
            ' Escape with purpose
            If vRtn = m_SortOrder Then
                Exit For
            End If
        Next vPos
        If X = vStart Or vPos = vNewCount + 1 Then
            ' New item at last
            vNew.Add m_Items(X), "K " & X
        Else 'NOT X...
            ' New item somewhere b/w
            vNew.Add m_Items(X), "K " & X, vPos
        End If
    Next X
    ' Return Sorted Collection
    Set m_Items = vNew

End Sub

'-------------------------------------------------------------------------
' Procedure  : SortOrder
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Get/Let Property SortOrder
'-------------------------------------------------------------------------
Public Property Get SortOrder() As SortOrderEnum

    SortOrder = m_SortOrder

End Property

Public Property Let SortOrder(ByVal vNewValue As SortOrderEnum)

    m_SortOrder = vNewValue
    PropertyChanged "SortOrder"
    If m_SortItems Then
        SortList
    End If
    ReDrawList

End Property

'-------------------------------------------------------------------------
' Procedure  : StrechIcon
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Let/Get property StrechIcon
'-------------------------------------------------------------------------
Public Property Get StrechIcon() As Boolean

    StrechIcon = m_StrechIcon

End Property

Public Property Let StrechIcon(ByVal vNewValue As Boolean)

    m_StrechIcon = vNewValue
    PropertyChanged "StrechIcon"
    ReDrawList

End Property

'-------------------------------------------------------------------------
' Procedure  : TextAlignment
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To Get/Let Property TextAlignment
'-------------------------------------------------------------------------
Public Property Get TextAlignment() As AlignmentConstants

    TextAlignment = m_TextAllineMent

End Property

Public Property Let TextAlignment(ByVal vNewValue As AlignmentConstants)

    m_TextAllineMent = vNewValue
    PropertyChanged "TextAlignment"
    ReDrawList

End Property

'-------------------------------------------------------------------------
' Procedure  : UserControl_Initialize
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Initialise control
'-------------------------------------------------------------------------
Private Sub UserControl_Initialize()

' Used to prevent crashes on XP

    m_hMode = LoadLibrary("shell32.dll")
    m_KeyControl = True

End Sub

'-------------------------------------------------------------------------
' Procedure  : UserControl_InitProperties
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Initialise default property values
'-------------------------------------------------------------------------
Private Sub UserControl_InitProperties()

    m_BackColor = m_def_BackColor
    m_ForeColor = m_def_ForeColor
    m_SelColor = m_def_SelColor
    m_SelForeColor = m_def_SelForeColor
    Set m_Picture = Nothing
    Set m_ListIcon = Nothing
    Set m_Font = Ambient.Font
    m_StrechIcon = m_def_StrechIcon
    m_iHeight = TextHeight("A")
    m_FullRowSel = m_def_FullRowSel
    m_SortOrder = m_def_SortOrder
    m_SortItems = m_def_SortItems
    m_IconFocus = m_def_IconFocus
    m_TextAllineMent = m_def_TextAllignMent

End Sub

'-------------------------------------------------------------------------
' Procedure  : UserControl_ReadProperties
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Read design time propery changes
'-------------------------------------------------------------------------
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    With PropBag
        Set m_Picture = .ReadProperty("Picture", Nothing)
        Set m_ListIcon = .ReadProperty("ListIcon", Nothing)
        Set m_Font = .ReadProperty("Font", Ambient.Font)
        m_BackColor = .ReadProperty("BackColor", m_def_BackColor)
        m_ForeColor = .ReadProperty("ForeColor", m_def_ForeColor)
        m_SelColor = .ReadProperty("SelColor", m_def_SelColor)
        m_SelForeColor = .ReadProperty("SelForeColor", m_def_SelForeColor)
        m_StrechIcon = .ReadProperty("StrechIcon", m_def_StrechIcon)
        Me.Appearance = .ReadProperty("Appearance", m_def_Appearance)
        Me.BorderStyle = .ReadProperty("BorderStyle", m_def_BorderStyle)
        m_FullRowSel = .ReadProperty("FullRowSelect", m_def_FullRowSel)
        m_SortItems = .ReadProperty("SortItems", m_def_SortItems)
        m_SortOrder = .ReadProperty("SortOrder", m_def_SortOrder)
        m_IconFocus = .ReadProperty("IconFocus", m_def_IconFocus)
        m_TextAllineMent = .ReadProperty("TextAlignment", m_def_TextAllignMent)
    End With 'PropBag
    ReDrawList

End Sub

'-------------------------------------------------------------------------
' Procedure  : UserControl_Resize
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Resize Controls
'-------------------------------------------------------------------------
Private Sub UserControl_Resize()

    Me.Refresh

End Sub

'-------------------------------------------------------------------------
' Procedure  : UserControl_Terminate
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Free the memory
'-------------------------------------------------------------------------
Private Sub UserControl_Terminate()

    FreeLibrary m_hMode
    Me.Clear
    Set m_Items = Nothing

End Sub

'-------------------------------------------------------------------------
' Procedure  : UserControl_WriteProperties
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Write design time propery changes
'-------------------------------------------------------------------------
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    With PropBag
        Call .WriteProperty("ListIcon", m_ListIcon, Nothing)
        Call .WriteProperty("Picture", m_Picture, Nothing)
        Call .WriteProperty("Font", m_Font, Ambient.Font)
        Call .WriteProperty("BackColor", m_BackColor, m_def_BackColor)
        Call .WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
        Call .WriteProperty("SelColor", m_SelColor, m_def_SelColor)
        Call .WriteProperty("SelForeColor", m_SelForeColor, m_def_SelForeColor)
        Call .WriteProperty("StrechIcon", m_StrechIcon, m_def_StrechIcon)
        Call .WriteProperty("Appearance", UserControl.Appearance, m_def_Appearance)
        Call .WriteProperty("BorderStyle", UserControl.BorderStyle, m_def_BorderStyle)
        Call .WriteProperty("FullRowSelect", m_FullRowSel, m_def_FullRowSel)
        Call .WriteProperty("SortItems", m_SortItems, m_def_SortItems)
        Call .WriteProperty("SortOrder", m_SortOrder, m_def_SortOrder)
        Call .WriteProperty("IconFocus", m_IconFocus, m_def_IconFocus)
        Call .WriteProperty("TextAlignment", m_TextAllineMent, m_def_TextAllignMent)
    End With 'PropBag

End Sub

'-------------------------------------------------------------------------
' Procedure  : VScroll_Change
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : Scroll List
'-------------------------------------------------------------------------
Private Sub VScroll_Change()

    On Error Resume Next
    m_iTop = VScroll.Value
    ReDrawList

End Sub

'-------------------------------------------------------------------------
' Procedure  : VScroll_GotFocus/VScroll_LostFocus
' Auther     : Jim Jose
' Input      : None
' OutPut     : None
' Purpose    : To determine control using keyboard
'-------------------------------------------------------------------------
Private Sub VScroll_GotFocus()

    m_KeyControl = False

End Sub

Private Sub VScroll_LostFocus()

    m_KeyControl = True

End Sub



