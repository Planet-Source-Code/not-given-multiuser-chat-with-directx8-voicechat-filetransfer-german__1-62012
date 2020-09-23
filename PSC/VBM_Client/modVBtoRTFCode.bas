Attribute VB_Name = "modVBtoRTFCode"
Option Explicit

Private Const CODE_CONST As Long = 53
Public Codes(1 To 53) As String
Dim Initialized As Boolean

Private Sub InitCode()
    Codes(1) = "If"
    Codes(2) = "Then"
    Codes(3) = "Else"
    Codes(4) = "End"
    Codes(5) = "Sub"
    Codes(6) = "Function"
    Codes(7) = "Private"
    Codes(8) = "Public"
    Codes(9) = "Private"
    Codes(10) = "Select"
    Codes(11) = "Case"
    Codes(12) = "Exit"
    Codes(13) = "For"
    Codes(14) = "To"
    Codes(15) = "Next"
    Codes(16) = "Do"
    Codes(17) = "While"
    Codes(18) = "Until"
    Codes(19) = "Call"
    Codes(20) = "Loop"
    Codes(21) = "CInt"
    Codes(22) = "CDbl"
    Codes(23) = "CBool"
    Codes(24) = "CLng"
    Codes(25) = "CStr"
    Codes(26) = "CSng"
    Codes(27) = "+"
    Codes(28) = "-"
    Codes(29) = "*"
    Codes(30) = "/"
    Codes(31) = "^"
    Codes(32) = "SQR"
    Codes(33) = "CByte"
    Codes(34) = "CDate"
    Codes(35) = "CCur"
    Codes(36) = "CDec"
    Codes(37) = "Const"
    Codes(38) = "Dim"
    Codes(39) = "ElseIf"
    Codes(40) = "Private"
    Codes(41) = "Error"
    Codes(42) = "GoTo"
    Codes(43) = "On"
    Codes(44) = "Option"
    Codes(45) = "Explicit"
    Codes(46) = "False"
    Codes(47) = "True"
    Codes(48) = "Set"
    Codes(49) = "Nothing"
    Codes(50) = "Wend"
    Codes(51) = "Type"
    Codes(52) = "Randomize"
    Codes(53) = "Rnd"
    Initialized = True
End Sub

Public Function CreateColoredString(ByVal txMessage As String) As String
    Dim TempAr$(), I&, pos&, pos1&, pos2&, pos3&, tmp$
    If Not Initialized Then InitCode
    TempAr = Split(txMessage, vbCrLf)
txMessage = "{\rtf1\ansi\ansicpg1252\deff0\deflang1031{\fonttbl{\f0\fmodern\fprq1\fcharset0 Courier New;}}{\colortbl ;\red0\green0\blue0;\red0\green128\blue0;\red0\green0\blue128;\red128\green128\blue128;}" & vbCrLf & _
                "\viewkind4\uc1\pard\cf1\f0\fs20" & vbCrLf
    
    For I = 0 To UBound(TempAr)
        pos = 1
        Do While pos > 0
            pos1 = InStr(pos, TempAr(I), "'")
            pos2 = InStr(pos, TempAr(I), """")
            pos3 = FindNextWord(TempAr(I), pos)
            If (pos1 > 0) Or (pos2 > 0) Or (pos3 > 0) Then
                If IsMax(pos1, pos2, pos3) Then
                    TempAr(I) = Left$(TempAr(I), pos1 - 1) & "\cf2 " & Right$(TempAr(I), Len(TempAr(I)) - pos1 + 1)
                    pos = 0
                ElseIf IsMax(pos2, pos1, pos3) Then
                    pos1 = InStr(pos2 + 1, TempAr(I), """")
                    If pos1 = 0 Then pos1 = Len(TempAr(I))
                    pos3 = Len(TempAr(I))
                    tmp = TempAr(I)
                    TempAr(I) = Left$(TempAr(I), pos2 - 1) & "\cf4" & Mid$(TempAr(I), pos2, pos1 - pos2 + 1) & "\cf0"
                    If pos1 < pos3 Then TempAr(I) = TempAr(I) & Right$(tmp, pos3 - pos1)
                    pos = pos1 + 5
                    tmp = ""
                Else
                    TempAr(I) = ColorNextWord(TempAr(I), pos3)
                    pos = pos3 + 5
                End If
            Else
                pos = 0
            End If
'            DoEvents
        Loop
        TempAr(I) = TempAr(I) & "\cf0\par"
'        DoEvents
    Next
    CreateColoredString = txMessage & Join(TempAr, vbCrLf) & "}"
End Function

Private Function IsMax(ByVal Item1 As Long, ByVal Item2 As Long, ByVal Item3 As Long) As Boolean
    IsMax = ((Item1 < Item2) Or (Item2 = 0)) And ((Item1 < Item3) Or (Item3 = 0)) And (Not (Item1 = 0))
End Function

Private Function FindNextWord(ByVal txMessage As String, ByVal startPos As Long) As Long
    Dim I&, pos&, pos1&, tmp$
    pos = Len(txMessage)
    For I = 1 To CODE_CONST
        pos1 = 0
        pos1 = InStr(startPos, txMessage, Codes(I), vbTextCompare)
        If (pos1 < pos) And (pos1 > 0) Then
            If pos1 - 1 = 0 Then
                If Len(txMessage) = pos1 + Len(Codes(I)) Then
                    pos = pos1
                Else
                    tmp = Mid$(txMessage, pos1 + Len(Codes(I)), 1)
                    If tmp = " " Or tmp = "(" Or tmp = """" Then
                        pos = pos1
                    End If
                End If
            Else
                If (Mid$(txMessage, pos1 - 1, 1) = " ") Then
                    If Len(txMessage) <= pos1 + Len(Codes(I)) Then
                        pos = pos1
                    Else
                        tmp = Mid$(txMessage, pos1 + Len(Codes(I)), 1)
                        If tmp = " " Or tmp = "(" Or tmp = """" Then
                            pos = pos1
                        End If
                    End If
                End If
            End If
        End If
    Next
    If pos = Len(txMessage) Then
        FindNextWord = 0
    Else
        FindNextWord = pos
    End If
End Function

Private Function ColorNextWord(ByVal txMessage, ByVal startPos As Long) As String
    Dim I&, pos&, pos1&, counter&, tmp$
    pos = Len(txMessage)
    For I = 1 To CODE_CONST
        pos1 = InStr(startPos, txMessage, Codes(I), vbTextCompare)
        If (pos1 < pos) And (pos1 > 0) Then
            If pos1 - 1 = 0 Then
                If Len(txMessage) = pos1 + Len(Codes(I)) Then
                    pos = pos1
                    counter = I
                Else
                    tmp = Mid$(txMessage, pos1 + Len(Codes(I)), 1)
                    If tmp = " " Or tmp = "(" Or tmp = """" Then
                        pos = pos1
                        counter = I
                    End If
                End If
            Else
                If (Mid$(txMessage, pos1 - 1, 1) = " ") Then
                    If Len(txMessage) <= pos1 + Len(Codes(I)) Then
                        pos = pos1
                        counter = I
                    Else
                        tmp = Mid$(txMessage, pos1 + Len(Codes(I)), 1)
                        If tmp = " " Or tmp = "(" Or tmp = """" Then
                            pos = pos1
                            counter = I
                        End If
                    End If
                End If
            End If
        End If
    Next
    ColorNextWord = Left$(txMessage, pos - 1) & "\cf3" & Codes(counter) & "\cf0 " & Right$(txMessage, Len(txMessage) - pos - Len(Codes(counter)) + 1)
End Function

