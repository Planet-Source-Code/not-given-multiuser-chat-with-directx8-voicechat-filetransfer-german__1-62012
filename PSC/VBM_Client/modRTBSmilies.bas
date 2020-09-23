Attribute VB_Name = "modRTBSmilies"
Option Explicit

Public Function CheckRichTextForSmilies(RichtText As String) As String

Dim i  As Integer
Dim zw As String

    zw = RichtText
    For i = 0 To UBound(SmilieArray)
        zw = modFastReplace.Replace(zw, SmilieArray(i).CharCode, SmilieArray(i).rtfCode)
    Next i
    CheckRichTextForSmilies = zw

End Function

Public Sub InitSmilies()

Dim i As Integer

    ReDim SmilieArray(38)
    SmilieArray(0).CharCode = ":)"
    SmilieArray(1).CharCode = ":D"
    SmilieArray(2).CharCode = ":-O"
    SmilieArray(3).CharCode = ";)"
    SmilieArray(4).CharCode = "(H)"
    SmilieArray(5).CharCode = ":P"
    SmilieArray(6).CharCode = ":S"
    SmilieArray(7).CharCode = ":@"
    SmilieArray(8).CharCode = ":("
    SmilieArray(9).CharCode = ":$"
    SmilieArray(10).CharCode = ":I"
    SmilieArray(11).CharCode = ":'("
    SmilieArray(12).CharCode = "8oI"
    SmilieArray(13).CharCode = "(A)"
    SmilieArray(14).CharCode = "+o("
    SmilieArray(15).CharCode = "8-I"
    SmilieArray(16).CharCode = "<:o)"
    SmilieArray(17).CharCode = "I-)"
    SmilieArray(18).CharCode = "*-)"
    SmilieArray(19).CharCode = ":-#"
    SmilieArray(20).CharCode = ":-*"
    SmilieArray(21).CharCode = "^o)"
    SmilieArray(22).CharCode = "8-)"
    SmilieArray(23).CharCode = "(L)"
    SmilieArray(24).CharCode = "(U)"
    SmilieArray(25).CharCode = "(M)"
    SmilieArray(26).CharCode = "(@)"
    SmilieArray(27).CharCode = "(&)"
    SmilieArray(28).CharCode = "(SN)"
    SmilieArray(29).CharCode = "(SH)"
    SmilieArray(30).CharCode = "(S)"
    SmilieArray(31).CharCode = "(*)"
    SmilieArray(32).CharCode = "(#)"
    SmilieArray(33).CharCode = "(R)"
    SmilieArray(34).CharCode = "(h-)"
    SmilieArray(35).CharCode = "(-h)"
    SmilieArray(36).CharCode = "(K)"
    SmilieArray(37).CharCode = "(F)"
    SmilieArray(38).CharCode = "(O)"
    For i = 0 To UBound(SmilieArray)
        SmilieArray(i).rtfCode = modFunctions.LoadFF(AppPath & "\Smilies\Smilie" & CStr(i) & ".rtf")
    Next i

End Sub


