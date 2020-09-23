VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "VBMessenger 9 - About"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3435
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   229
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer tmrScroll 
      Interval        =   20
      Left            =   5160
      Top             =   2880
   End
   Begin RichTextLib.RichTextBox rtbCredits 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   4048
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"frmAbout.frx":628A
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINTL
    X As Long
    Y As Long
End Type
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, _
                                                                        ByVal wMsg As Long, _
                                                                        ByVal wParam As Long, _
                                                                        lParam As Any) As Long
Private Const WM_USER = &H400
Private Const EM_GETSCROLLPOS = (WM_USER + 221)
Private Const EM_SETSCROLLPOS = (WM_USER + 222)
Private P As POINTL

Private Sub Form_Load()

Dim check As New clsCRC

    If Not FileExists(AppPath & "\Credits.rtf") Then
        MsgBox "VBMessenger9 by Thorben Linneweber (2005)"
        Unload Me
    Else 'NOT NOT...
        If Not 542002045 = check.CalculateFile(AppPath & "\Credits.rtf") Then
            MsgBox "VBMessenger9 by Thorben Linneweber (2005)"
            Unload Me
        Else 'NOT NOT...
            Call modFunctions.MakeTopMost(hwnd)
            rtbCredits.LoadFile AppPath & "\Credits.rtf"
            modLockRTB.InitLRTB rtbCredits.hwnd
        End If
    End If
Debug.Print check.CalculateFile(AppPath & "\Credits.rtf")
End Sub

Private Sub Form_Unload(Cancel As Integer)

    modLockRTB.TerminateLRTB rtbCredits.hwnd
    Set frmAbout = Nothing

End Sub

Private Sub tmrScroll_Timer()

    SendMessage rtbCredits.hwnd, EM_SETSCROLLPOS, 0, P 'make the other RichTextBox match
    P.Y = P.Y + 1
    If P.Y = 1250 Then
        tmrScroll.Enabled = False
        Unload Me
    End If

End Sub



