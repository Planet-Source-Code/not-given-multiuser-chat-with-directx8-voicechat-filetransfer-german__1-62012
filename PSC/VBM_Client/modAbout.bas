Attribute VB_Name = "modAbout"
'*********************************
'(Adout / Credits) Form - Animated
'*********************************

'*********************************
'By     Jim Jose
'email  jimjosev33@yahoo.com
'*********************************

'*********************************
'This code is only in initialstage.
'It can only draw in a perticular font
'and no text allining is added. But
'it is sufficient to my needs.
'
'You can choose any Start/Mid/End
'fading Color. I set those accoring to
' my background picture. You can
'also choose the shading region.
'cRegion=1/6 means the 1/6th height
'of  the picturebox will be shaded at
'top  and bottom.
'
'If you like this I am ready to
'complete it.Yourown upgrades are also
'welcomed.( If so please inform me too)
'Kindly give me your Comments/suggessions.
'****************************
Option Explicit

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public vTop         As Long     'Stores the Text Top pos
Public CrdLines()   As String   'Stores the text lines

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, _
                                        ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long


'Returns the shaded color in the specified color depth
Public Function GetShade(ByVal StartCol As Long, ByVal EndCol As Long, ByVal ColDepth As Double) As Long
On Error Resume Next
Dim sRate As Double
Dim cBlue As Long, cGreen As Long, cRed As Long   'Determines the pixel color
Dim sBlue As Long, sGreen As Long, sRed As Long   'Determines the SHADING color
    sRate = ColDepth
    GetRGB EndCol, sRed, sGreen, sBlue
    GetRGB StartCol, cRed, cGreen, cBlue
    cRed = cRed + (sRed - cRed) * sRate
    cGreen = cGreen + (sGreen - cGreen) * sRate
    cBlue = cBlue + (sBlue - cBlue) * sRate
    If cRed < 0 Then cRed = -cRed
    If cGreen < 0 Then cGreen = -cGreen
    If cBlue < 0 Then cBlue = -cBlue
GetShade = RGB(cRed, cGreen, cBlue)
End Function

'Returns the RGB values
Private Sub GetRGB(ByVal LngCol As Long, R As Long, G As Long, B As Long)
  R = LngCol Mod 256
  G = (LngCol And vbGreen) / 256 'Green
  B = (LngCol And vbBlue) / 65536 'Blue
End Sub

'Drawing the Text
Public Function SendCredits(PicBox As PictureBox, Txt As String, _
                        ByVal X As Integer, ByVal Y As Integer, _
                        Optional StartCol As Long = 0, Optional MidCol As Long = 111111, _
                        Optional EndCol As Long = 0, Optional ByVal cRegion As Double)
Dim hLength   As Integer 'Region over which the text fades
Dim DrawCol   As Long    'The current faded color
Dim rctDraw   As RECT
    hLength = PicBox.Height * cRegion   'Determines the fade region
    If Y <= hLength And Y >= -50 Then   'Some Calculations
        DrawCol = GetShade(MidCol, EndCol, (hLength - Y) / (hLength + 20))   'Getting the shaded color
    ElseIf Y <= PicBox.Height And Y >= PicBox.Height * (1 - cRegion) Then
        DrawCol = GetShade(StartCol, MidCol, (PicBox.Height - Y) / hLength)  'Getting the shaded color
    Else
        DrawCol = MidCol
    End If
    With rctDraw
        .Left = X
        .Top = Y
        .Right = PicBox.Width
        .Bottom = PicBox.Height
    End With
    PicBox.ForeColor = DrawCol  'Setting the DrawColor
    DrawText PicBox.hdc, Txt, -1, rctDraw, &H800    'Drawing the text
End Function


