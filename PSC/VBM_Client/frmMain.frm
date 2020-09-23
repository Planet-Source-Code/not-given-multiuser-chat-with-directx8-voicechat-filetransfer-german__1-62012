VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00EBF5F4&
   Caption         =   "VBMessenger"
   ClientHeight    =   6105
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9165
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   407
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   611
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox lstIntelliSense 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1080
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox picTray 
      BorderStyle     =   0  'Kein
      Height          =   240
      Left            =   120
      Picture         =   "frmMain.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Frame frameButtons 
      BackColor       =   &H00EBF5F4&
      BorderStyle     =   0  'Kein
      Height          =   855
      Left            =   6960
      TabIndex        =   7
      Top             =   4905
      Width           =   2175
      Begin VBMessenger9.chameleonButton cmdSendFile 
         Height          =   375
         Left            =   480
         TabIndex        =   8
         ToolTipText     =   "Eine Datei senden"
         Top             =   0
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   ""
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   15463924
         BCOLO           =   15463924
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":0B14
         PICN            =   "frmMain.frx":0B30
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VBMessenger9.chameleonButton cmdPrivateMessage 
         Height          =   375
         Left            =   0
         TabIndex        =   9
         ToolTipText     =   "Private Nachricht senden"
         Top             =   0
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   15463924
         BCOLO           =   15463924
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":1266
         PICN            =   "frmMain.frx":1282
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VBMessenger9.chameleonButton cmdNudge 
         Height          =   375
         Left            =   1080
         TabIndex        =   10
         ToolTipText     =   "Einen 'Nudge' senden"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   ""
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   15463924
         BCOLO           =   15463924
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":1C94
         PICN            =   "frmMain.frx":1CB0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VBMessenger9.isButton cmdSend 
         Default         =   -1  'True
         Height          =   345
         Left            =   70
         TabIndex        =   11
         Top             =   470
         Width           =   2000
         _ExtentX        =   3519
         _ExtentY        =   609
         Style           =   7
         Caption         =   "Senden"
         iNonThemeStyle  =   0
         USeCustomColors =   -1  'True
         BackColor       =   14737632
         HighlightColor  =   14737632
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VBMessenger9.chameleonButton cmdFormatCode 
         Height          =   375
         Left            =   1680
         TabIndex        =   14
         ToolTipText     =   "Code formatieren"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   15463924
         BCOLO           =   15463924
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMain.frx":21B2
         PICN            =   "frmMain.frx":21CE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin VB.Timer tmrGetData 
      Interval        =   25
      Left            =   600
      Top             =   600
   End
   Begin VB.Timer tmrType 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   120
      Top             =   600
   End
   Begin VB.PictureBox picSmiliesContainer 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2580
      Left            =   2550
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   3975
      Begin VBMessenger9.isButton isButton1 
         Height          =   360
         Left            =   3480
         TabIndex        =   4
         Top             =   2055
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         Icon            =   "frmMain.frx":24F0
         Style           =   7
         Caption         =   "<"
         iNonThemeStyle  =   4
         BackColor       =   16777215
         Tooltiptitle    =   ""
         ToolTipIcon     =   0
         ToolTipType     =   1
         ttForeColor     =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaskColor       =   0
         RoundedBordersByTheme=   0   'False
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   38
         Left            =   3000
         Picture         =   "frmMain.frx":250C
         Top             =   2040
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   37
         Left            =   2520
         Picture         =   "frmMain.frx":2CBA
         Top             =   2040
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   36
         Left            =   2040
         Picture         =   "frmMain.frx":3468
         Top             =   2040
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   35
         Left            =   1560
         Picture         =   "frmMain.frx":3C16
         Top             =   2040
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   34
         Left            =   1080
         Picture         =   "frmMain.frx":43C4
         Top             =   2040
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   33
         Left            =   600
         Picture         =   "frmMain.frx":4B72
         Top             =   2040
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   32
         Left            =   120
         Picture         =   "frmMain.frx":5320
         Top             =   2040
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   31
         Left            =   3480
         Picture         =   "frmMain.frx":5ACE
         Top             =   1560
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   30
         Left            =   3000
         Picture         =   "frmMain.frx":627C
         Top             =   1560
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   29
         Left            =   2520
         Picture         =   "frmMain.frx":6A2A
         Top             =   1560
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   28
         Left            =   2040
         Picture         =   "frmMain.frx":71D8
         Top             =   1560
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   27
         Left            =   1560
         Picture         =   "frmMain.frx":7986
         Top             =   1560
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   26
         Left            =   1080
         Picture         =   "frmMain.frx":8134
         Top             =   1560
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   25
         Left            =   600
         Picture         =   "frmMain.frx":88E2
         Top             =   1560
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   24
         Left            =   120
         Picture         =   "frmMain.frx":9090
         Top             =   1560
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   23
         Left            =   3480
         Picture         =   "frmMain.frx":983E
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   22
         Left            =   3000
         Picture         =   "frmMain.frx":9FEC
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   21
         Left            =   2520
         Picture         =   "frmMain.frx":A79A
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   20
         Left            =   2040
         Picture         =   "frmMain.frx":AF48
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   19
         Left            =   1560
         Picture         =   "frmMain.frx":B6F6
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   18
         Left            =   1080
         Picture         =   "frmMain.frx":BEA4
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   17
         Left            =   600
         Picture         =   "frmMain.frx":C652
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   16
         Left            =   120
         Picture         =   "frmMain.frx":CE00
         Top             =   1080
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   15
         Left            =   3480
         Picture         =   "frmMain.frx":D5AE
         Top             =   600
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   14
         Left            =   3000
         Picture         =   "frmMain.frx":DD5C
         Top             =   600
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   13
         Left            =   2520
         Picture         =   "frmMain.frx":E50A
         Top             =   600
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   12
         Left            =   2040
         Picture         =   "frmMain.frx":ECB8
         Top             =   600
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   11
         Left            =   1560
         Picture         =   "frmMain.frx":F466
         Top             =   600
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   10
         Left            =   1080
         Picture         =   "frmMain.frx":FC14
         Top             =   600
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   9
         Left            =   600
         Picture         =   "frmMain.frx":103C2
         Top             =   600
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   8
         Left            =   120
         Picture         =   "frmMain.frx":10B70
         Top             =   600
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   7
         Left            =   3480
         Picture         =   "frmMain.frx":1131E
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   6
         Left            =   3000
         Picture         =   "frmMain.frx":11ACC
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   5
         Left            =   2520
         Picture         =   "frmMain.frx":1227A
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   4
         Left            =   2040
         Picture         =   "frmMain.frx":12A28
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   3
         Left            =   1560
         Picture         =   "frmMain.frx":131D6
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   2
         Left            =   1080
         Picture         =   "frmMain.frx":13984
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   0
         Left            =   120
         Picture         =   "frmMain.frx":14132
         Top             =   120
         Width           =   375
      End
      Begin VB.Image imgSmilie 
         Appearance      =   0  '2D
         Height          =   375
         Index           =   1
         Left            =   600
         Picture         =   "frmMain.frx":148E0
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Timer tmrConnectionCheck 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1080
      Top             =   120
   End
   Begin VB.Timer tmrLogin 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   600
      Top             =   120
   End
   Begin VB.Timer tmrConnect 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin MSWinsockLib.Winsock wsc 
      Left            =   1080
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Unten ausrichten
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   5805
      Width           =   9165
      _ExtentX        =   16166
      _ExtentY        =   529
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VBMessenger9.ListBoxEX lstBuddys 
      Height          =   4770
      Left            =   6990
      TabIndex        =   1
      Top             =   45
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   8414
      ListIcon        =   "frmMain.frx":1508E
      Picture         =   "frmMain.frx":152EA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   11034163
      SelColor        =   11034163
      Appearance      =   0
      BorderStyle     =   0
      TextAlignment   =   2
   End
   Begin VB.PictureBox picSmilieTmp 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   2760
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   23
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   345
   End
   Begin RichTextLib.RichTextBox RTBChat 
      Height          =   4755
      Left            =   75
      TabIndex        =   0
      Top             =   60
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   8387
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":15306
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RTBMessage 
      Height          =   780
      Left            =   75
      TabIndex        =   6
      Top             =   4935
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   1376
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":1537D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Shape shpChat 
      BorderColor     =   &H00C0C0C0&
      Height          =   4815
      Left            =   60
      Top             =   30
      Width           =   6855
   End
   Begin VB.Shape shpBuddy 
      BorderColor     =   &H00808080&
      Height          =   4815
      Left            =   6975
      Top             =   30
      Width           =   2130
   End
   Begin VB.Shape shpMessage 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   60
      Top             =   4920
      Width           =   6855
   End
   Begin VB.Menu mnuVBMessenger 
      Caption         =   "VBMessenger"
      Begin VB.Menu mnuFont 
         Caption         =   "Schriftart ändern "
      End
      Begin VB.Menu mnuFontColor 
         Caption         =   "Schriftfarbe ändern"
      End
      Begin VB.Menu mnuSendMessage 
         Caption         =   "Nachricht senden"
      End
      Begin VB.Menu TRZ1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLogOut 
         Caption         =   "Ausloggen"
      End
      Begin VB.Menu mnuCloseMessenger 
         Caption         =   "Messenger Beenden"
      End
   End
   Begin VB.Menu mnuOtherUser 
      Caption         =   "Andere Benutzer"
      Begin VB.Menu mnuNudge 
         Caption         =   "'Nudge' senden"
      End
      Begin VB.Menu mnuprivateMessage 
         Caption         =   "Private Nachricht senden"
      End
      Begin VB.Menu mnuFileTransfer 
         Caption         =   "Dateitransfer"
      End
   End
   Begin VB.Menu mnuSmilies 
      Caption         =   "Smilies"
   End
   Begin VB.Menu mnuVoiceChat 
      Caption         =   "VoiceChat beitreten"
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Optionen"
      Begin VB.Menu mnuAudioA 
         Caption         =   "Audio Assistenten starten"
      End
      Begin VB.Menu mnuMOptions 
         Caption         =   "Messenger Optionen"
      End
   End
   Begin VB.Menu mnuQuestionMark 
      Caption         =   "?"
      Begin VB.Menu mnuHelp 
         Caption         =   "Hilfe"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --------------------------------------------------
'
'             VBMessenger9 Server & Client
'
' (C) by Thorben Linneweber (2005)
'
' OpenSource
'
' Darf unverändert weitergegeben werden.
' Darf nur für den privaten Gebrauch verändert werden.
' Kommerzielle Nutzung ist untersagt!
'
' Verbesserungsvorschläge und Änderungen bitte an

' thorben_linneweber@hotmail.com
'
' ----------------------------------------------------

' --------------------------------------------------
'
' VBMessenger 9 Client
'
' Serverbefehle
' ---------------------
'
'
'    login
'    [Username][Passwort]
'    -> LoginAnfrage
'
'    message
'    [Nachricht]
'    -> [Nachricht] an Alle
'
'    privatemessage
'    [Username]
'    -> Nachricht an [Username]
'
'    nudge
'    [Username]
'    -> Nudge an [Username]
'
'    file
'    [Username][Filename][Filegröße]
'    -> sendet an [Username] eine Fileanfrage für [Filename][Filegröße]
'
'
'    acceptfile
'    [Username][AcceptString(true\false\alreadyreceiving)]
'    -> sendet an [Username] den [AcceptString(true\false\alreadyreceiving)] für eine FileTransferanfrage
'
'    typing
'    [TypeString(true\false]
'    -> sendet an alle Benutzer den [TypeString(true\false]
'
'
'    kick
'    [Username]
'    -> kickt [Username] (geht nur, wenn der Sender Adminrechte hat)
'
'    ban
'    [Username]
'    -> banned [Username] (geht nur, wenn der Sender Adminrechte hat)
'
'    makeadmin
'    [Username]
'    -> macht [Username] zum Admin (geht nur, wenn der Sender Adminrechte hat)
'
'    giveupadmin
'    []
'    -> gibt die Adminrechte ab
'
'    kickvoice
'    [Username]
'    -> schmeißt [Username] aus dem VoiceChat
'
'    loginadmin
'    []
'    -> stellt die festen Adminrechte wieder her
'
'    askvoice
'    []
'    -> Anfrage für VoiceChat
'
'
' Rückgabewerte vom Server (Events) (aus Sicht des Clienten)
' ------------------------
'
'    login
'    [AcceptString(accept\deny)]
'    -> LoginAnfrage-Antwort
'
'    message
'    [Benutzername][Nachricht][Admin(true\false)]
'    -> [Nachricht] an Alle von [Benutzername]
'
'    privatemessage
'    [Benutzername][Nachricht][Admin(true\false)]
'    -> private [Nachricht] von [Benutzername]
'
'    privatemessageb
'    [Benutzername][Nachricht][Admin(true\false)]
'    -> Bestätigung von einer privaten [Nachricht]
'
'    nudge
'    [Benutzername][Admin(true\false)]
'    -> Nudge von [Benutzername]
'
'    nudgeb
'    [Benutzername][Admin(true\false)]
'    -> Nudge-Sende-Bestätigung
'
'    fileb
'    [Benutzername][Filename][Filesize][Admin(true\false)]
'    -> FileAnfragesendebestätigung
'
'    acceptfile
'    [Benutzername][AcceptString(true\false\alreadyreceiving)][Admin(true\false)]
'    -> Filetransfer Anfragebestätigung von [Benutzername], [AcceptString(true\false\alreadyreceiving)]
'
'    acceptfileb
'    [Benutzername][AcceptString(true\false\alreadyreceiving)][Admin(true\false)]
'    -> FiletransferAnfrageAntwortBestätigung
'
'    typing
'    [Username][TypeString(true\false)][Admin(true\false)]
'    -> [Username] tippt \ tippt nicht ; [TypeString(true\false)]
'
'    voice
'    [Username][Username][VoiceStr(enabled\disabled]
'    -> [Username] tippt \ tippt nicht ; [TypeString(true\false)]
'
'    askvoice
'    [AllowStr(true\false)]
'    -> Antwort auf Voiceanfrage
'
'    kick
'    [Username][Username2]
'    -> [Username] hat [Username2] gekickt
'
'    ban
'    [Username][Username2]
'    -> [Username] hat [Username2] gebannt
'
'    makeadmin
'    [Username][Username2]
'    -> [Username] hat [Username2] zum Admin gemacht
'
'    giveupadmin
'    [Username]
'    -> [Username] hat seine Adminrechte abgelegt
'
'    kickvoice
'    [Username][Username2]
'    -> [Username] kickt [Username2] aus dem VoiceChat
'
Option Explicit

Private IntelliVisible    As Boolean
Private TimeOutC          As Integer
Private Typing()          As tTyping
Private ReceiveBuffer     As String
Public oMagneticWnd       As New cMagneticWnd

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                           ByVal lpOperation As String, _
                                                                           ByVal lpFile As String, _
                                                                           ByVal lpParameters As String, _
                                                                           ByVal lpDirectory As String, _
                                                                           ByVal nShowCmd As Long) As Long

Private Sub AcceptFile(Username As String, _
                       AcceptStr As String, _
                       RemoteIP As String, _
                       Admin As String)

' ~ Antwort des anderen auf die file-anfrage

    If RemoteIP = "127.0.0.1" Then
        ' Client, der eine Datei senden möchte, läuft auf gleichem computer wie server
        ' sonderfall (auf drängen von Herr Bierwirth noch eingebaut ;) )
        RemoteIP = modDeclaration.ServerIP
    End If
    If AcceptStr = "true" Then
        AddTextToRTB Username & IIf(Admin = "true", " (Administrator)", "") & "(" & RemoteIP & ") hat die Dateitransferanfrage angenommen", EventColor, True, True
        modDeclaration.RemoteIP = RemoteIP
        modDeclaration.SendFile = True
        frmTransfer.Show
    ElseIf AcceptStr = "false" Then 'NOT ACCEPTSTR...
        AddTextToRTB Username & IIf(Admin = "true", " (Administrator)", "") & "(" & RemoteIP & ") hat die Dateitransferanfrage abgelehnt", EventColor, True, True
        modDeclaration.SendingOrReceivingFile = False
    ElseIf AcceptStr = "alreadyreceiving" Then 'NOT ACCEPTSTR...
        AddTextToRTB Username & IIf(Admin = "true", " (Administrator)", "") & "(" & RemoteIP & ") kann momentan keine Dateien empfangen, da er/sie bereits Dateien sendet/empfängt", EventColor, True, True
        modDeclaration.SendingOrReceivingFile = False
    End If
    FlashAndSound

End Sub

Private Sub AcceptFileB(Username As String, _
                        AcceptStr As String, _
                        Admin As String)

' ~ Wenn man der FileAnfrage zugestimmt hat

    If AcceptStr = "true" Then
        AddTextToRTB "Sie haben dem Dateitransfer mit " & Username & " zugestimmt", EventColor, True, True
        frmTransfer.Show
    ElseIf AcceptStr = "false" Then 'NOT ACCEPTSTR...
        AddTextToRTB "Sie haben den Dateintransfer mit " & Username & " abgelehnt", EventColor, True, True
    ElseIf AcceptStr = "alreadyreceiving" Then 'NOT ACCEPTSTR...
        AddTextToRTB "Ein Dateitransfer wurde automatisch abgelehnt, da Sie bereits eine Datei empfangen", EventColor, True, True
    End If
    FlashAndSound

End Sub

Private Sub AddMessageTextToRTB(Username As String, _
                                Message As String, _
                                PrivateMessage As Boolean, _
                                Admin As Boolean)

' ~ Diese wunderschöne Funktion added
'   - Nachrichten
'   - private Nachrichten

    Message = modRTBSmilies.CheckRichTextForSmilies(Message) ' Smilie-Codes ersetzen
    AddTextToRTB Username & IIf(Admin, " (Administrator)", "") & IIf(PrivateMessage, " (private Nachricht)", "") & " sagt:", 5526612, , PrivateMessage
    AddTextToRTB Message, , , PrivateMessage

End Sub

Private Sub AddTextToRTB(Text As String, _
                         Optional Color As ColorConstants = vbBlack, _
                         Optional Bold As Boolean = False, _
                         Optional Italic As Boolean = False)

' ~ Vereinfachen des "RTB-Text-Addens"

    With RTBChat
        .SelStart = Len(.Text)
        .SelBold = Bold
        .SelColor = Color
        .SelItalic = Italic
        .SelText = Text & vbNewLine
    End With 'RTBChat

End Sub

Private Sub AskVoice(Allow As String)

' der server hat unsere Anfrage angenommen, einen VoiceChat führen zu dürfen
' -> jetzt können wir eine Verbindung zum DX8-Voice-Server aufnehmen (sonst würden wir
' nicht zum Voice-Server connecten können)

    If Not bLoadedfrmVoice Then
        If Allow = "true" Then
            AddTextToRTB "Der Server hat der Anfrage für einen VoiceChat zugestimmt", EventColor, True, True
            frmVoice.Left = frmMain.Left + frmMain.Width
            frmVoice.Top = frmMain.Top
            frmVoice.Show vbModeless
        Else 'NOT ALLOW...
            AddTextToRTB "Der Server hat die Anfrage für einen VoiceChat abgelehnt", EventColor, True, True
        End If
    End If

End Sub

Private Sub Ban(AdminName As String, _
                Username As String)

    AddTextToRTB AdminName & " hat " & Username & " gebannt!", EventColor, True, True

End Sub

Private Sub cmdFormatCode_Click()
On Error Resume Next
RTBMessage.TextRTF = modVBtoRTFCode.CreateColoredString(RTBMessage.Text)

End Sub

Private Sub cmdNudge_Click()

Dim ToUser As String

    ToUser = lstBuddys.ListItems(lstBuddys.SelectedItem)
    ' ~ Sendet einen Nudge an den gewählten Benutzer
    If LenB(ToUser) = 0 Then Exit Sub
    If ToUser = modDeclaration.Username Then
        MsgBox "Sie können keinen 'Nudge' an sich selbst senden", vbInformation
    Else 'NOT TOUSER...
        SendData "nudge|" & lstBuddys.ListItems(lstBuddys.SelectedItem)
    End If

End Sub

Private Sub cmdPrivateMessage_Click()

Dim MessageToSend As String
Dim ToUser        As String

    MessageToSend = modFastReplace.Replace(RTBMessage.TextRTF, "|", "")
    ToUser = lstBuddys.ListItems(lstBuddys.SelectedItem)
    ' ~ Sendet eine private Nachricht, die nur der andere (und man selber) sehen kann
    If LenB(modFastReplace.Replace(RTBMessage.TextRTF, "|", "")) = 0 Then
        RTBMessage.Text = vbNullString
        Exit Sub
    End If
    If LenB(ToUser) = 0 Then Exit Sub
    If Len(MessageToSend) > (204800) Then
        MsgBox "Die Nachricht darf maximal 200kb groß sein!"
    Else 'NOT LEN(RTBMESSAGE.TEXTRTF)...'NOT LEN(MESSAGETOSEND)...
        If Not ToUser = modDeclaration.Username Then
            SendData "privatemessage|" & lstBuddys.ListItems(lstBuddys.SelectedItem) & "|" & MessageToSend
            RTBMessage.Text = vbNullString
            On Error Resume Next
            RTBMessage.SetFocus
            On Error GoTo 0
        Else 'NOT NOT...
            MsgBox "Sie können keine privaten Nachrichten an sich selbst senden!", vbInformation
        End If
    End If

End Sub

Private Sub cmdSend_Click()

Dim MessageToSend As String

    MessageToSend = modFastReplace.Replace(RTBMessage.TextRTF, "|", "")
    '~ sendet eine Nachricht an alle
    If Not LenB(modFastReplace.Replace(RTBMessage.Text, "|", "")) = 0 Then
        tmrType.Enabled = False
        If Len(MessageToSend) > (204800) Then
            MsgBox "Die Nachricht darf maximal 200kb groß sein!"
        Else 'NOT LEN(RTBMESSAGE.TEXTRTF)...'NOT LEN(MESSAGETOSEND)...
            SendData "typing|false"
            If RTBMessage.Text = "\\kick" Then
                SendData "kick|" & lstBuddys.ListItems(lstBuddys.SelectedItem)
            ElseIf RTBMessage.Text = "\\ban" Then 'NOT RTBMESSAGE.TEXT...
                SendData "ban|" & lstBuddys.ListItems(lstBuddys.SelectedItem)
            ElseIf RTBMessage.Text = "\\makeadmin" Then 'NOT RTBMESSAGE.TEXT...
                SendData "makeadmin|" & lstBuddys.ListItems(lstBuddys.SelectedItem)
            ElseIf RTBMessage.Text = "\\giveupadmin" Then 'NOT RTBMESSAGE.TEXT...
                SendData "giveupadmin"
            ElseIf RTBMessage.Text = "\\loginadmin" Then 'NOT RTBMESSAGE.TEXT...
                SendData "loginadmin|" & modDeclaration.Username
            ElseIf RTBMessage.Text = "\\kickvoice" Then 'NOT RTBMESSAGE.TEXT...
                SendData "kickvoice|" & lstBuddys.ListItems(lstBuddys.SelectedItem)
            Else 'NOT RTBMESSAGE.TEXT...
                SendData "message|" & MessageToSend
            End If
            RTBMessage.Text = vbNullString
            On Error Resume Next
            RTBMessage.SetFocus
            On Error GoTo 0
        End If
    Else 'NOT NOT...
        RTBMessage.Text = vbNullString
    End If

End Sub

Private Sub cmdSendFile_Click()

Dim ToUser   As String
Dim Filename As String

    ToUser = lstBuddys.ListItems(lstBuddys.SelectedItem)
    If LenB(ToUser) = 0 Then Exit Sub
    If modDeclaration.SendingOrReceivingFile Then
        MsgBox "Sie senden oder empfangen bereits eine Datei."
    Else 'MODDECLARATION.SENDINGORRECEIVINGFILE = FALSE/0
        If Not ToUser = modDeclaration.Username Then
            modDeclaration.SendingOrReceivingFile = True
            modCC.ShowOpen hwnd, True
            Filename = modCC.FileDialog.sFile
            Filename = Left$(Filename, InStr(Filename, vbNullChar) - 1) 'chr(0) abschneiden
            modDeclaration.PathOfFileToSendOrReceive = Filename
            If Not FileExists(Filename) Then
                modDeclaration.SendingOrReceivingFile = False
                MsgBox "Die angebene Datei existiert nicht!", vbInformation
            Else 'NOT NOT...
                modDeclaration.ReceiverOrSender = ToUser
                SendData "file|" & ToUser & "|" & modFunctions.ExtractFilename(Filename) & "|" & modFunctions.GetFileLen(Filename)
            End If
        Else 'NOT NOT...
            MsgBox "Sie können keine Dateien an sich selbst senden!", vbInformation
        End If
    End If

End Sub

Private Sub Connected()

' ~ Sollte aufgerufen werden, wenn Verbindung aufgebaut wurde & Login

    modSysTray.ModifyTray picTray, "Messenger 9 - Online", picTray
    RTBMessage.Enabled = True
    cmdSend.Enabled = True
    cmdNudge.Enabled = True
    cmdSendFile.Enabled = True
    cmdPrivateMessage.Enabled = True
    On Error Resume Next
    RTBMessage.SetFocus
    On Error GoTo 0
    DoEvents

End Sub

Private Sub Disconnected()

'~ Sollte aufgerufen werden, wenn keine Vebindung besteht (auch im Form_Load!) -> alles wird disabled/zurückgesetzt

    wsc.Close
    modSysTray.ModifyTray picTray, "Messenger 9 - nicht verbunden", picTray
    tmrConnect.Enabled = False
    tmrConnectionCheck.Enabled = False
    tmrLogin.Enabled = False
    TimeOutC = 0
    lstBuddys.Clear
    RTBMessage.Enabled = False
    cmdSend.Enabled = False
    cmdNudge.Enabled = False
    cmdSendFile.Enabled = False
    cmdPrivateMessage.Enabled = False
    If bLoadedfrmTransfer Then frmTransfer.UnloadfrmTransfer
    modDeclaration.SendingOrReceivingFile = False
    DoEvents

End Sub

Private Sub File(Username As String, _
                 Filename As String, _
                 FileSize As String, _
                 RemoteIP As String, _
                 Admin As String)

    If RemoteIP = "127.0.0.1" Then
        ' Client, der eine Datei senden möchte, läuft auf gleichem computer wie server
        ' sonderfall (auf drängen von Herr Bierwirth noch eingebaut ;) )
        RemoteIP = modDeclaration.ServerIP
    End If
    modDeclaration.ReceiverOrSender = Username
    '~ man bekommt ne Anfrage eine Datei anzunehmen
    '-> wird automatisch abgelehnt, wenn bereits eine Datei gesendet/empfangen wird
    If modDeclaration.SendingOrReceivingFile Then
        SendData "acceptfile|" & Username & "|alreadyreceiving"
    Else 'MODDECLARATION.SENDINGORRECEIVINGFILE = FALSE/0
        modDeclaration.SendingOrReceivingFile = True
        cmdSendFile.Value = True
        PlaySound FileOrVoiceRequest
        AddTextToRTB Username & IIf(Admin = "true", " (Administrator)", "") & "(" & RemoteIP & ") wartet darauf, dass Sie die Datei '" & Filename & "' (" & FileSize & ") annehmen.", EventColor, True, True
        With frmAcceptFile
            .imgIcon.Picture = modExtractIcon.LoadIcon(modExtractIcon.Large, Right$(Filename, Len(Filename) - InStrRev(Filename, ".")))
            .AcceptString = "acceptfile|" & Username & "|true"
            .DenyString = "acceptfile|" & Username & "|false"
            .lblAsk.Caption = "Möchten Sie die Datei '" & Filename & "' (" & FileSize & ") von " & Username & IIf(Admin = "true", " (Administrator)", "") & " (" & RemoteIP & ") annehmen?"
        End With 'frmAcceptFile
        modDeclaration.RemoteIP = RemoteIP
        modDeclaration.SendFile = False
        frmAcceptFile.Show
    End If
    FlashAndSound

End Sub

Private Sub FileB(Username As String, _
                  Filename As String, _
                  FileSize As String, _
                  Admin As String)

' ~ Eine Bestätigung, dass man eine Anfrage gesendet hat

    AddTextToRTB "Sie haben eine Anfrage an den Benutzer " & Username & " gesendet, die Datei '" & Filename & "' (" & FileSize & ") anzunehmen", EventColor, True, True
    FlashAndSound

End Sub

Private Sub FillIntelliSenseWithAdminRights()

    lstIntelliSense.AddItem "kick"
    lstIntelliSense.AddItem "giveupadmin"
    lstIntelliSense.AddItem "ban"
    lstIntelliSense.AddItem "makeadmin"
    lstIntelliSense.AddItem "loginadmin"
    lstIntelliSense.AddItem "kickvoice"

End Sub


Private Sub FixedDataArrival(strData As String)

' ~ hier kommen die Daten schön voneinander getrennt -so wie sie abgeschickt wurden- wieder an
' -> geniales system ;) <-

Dim v As Variant

    If Not Len(strData) = 0 Then
        v = Split(strData, "|")
        Select Case CStr(v(0))
            '~ Empfang von Login-Daten
        Case "login"
            Login CStr(v(1))
            '~ Erhält 'normale' Nachricht
        Case "message"
            Message CStr(v(1)), CStr(v(2)), CStr(v(3))
            '~ Wer ist Online?
        Case "userlst"
            UpdateLst CStr(v(1)), CStr(v(2))
            '~ Empfangen einer private Nachricht / SendeBestätigung
        Case "privatemessage"
            PrivateMessage CStr(v(1)), CStr(v(2)), CStr(v(3))
        Case "privatemessageb"
            PrivateMessageB CStr(v(1)), CStr(v(2)), CStr(v(3))
            '~ Empfangen eines 'Nudges' / SendeBestätigung
        Case "nudge"
            Nudge CStr(v(1)), CStr(v(2))
        Case "nudgeb"
            NudgeB CStr(v(1)), CStr(v(2))
            '~ FileTransferANFRAGE empfangen / SendeBestätigung
        Case "file" ' -> bestätigung, dass man versucht was zu senden
            File CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5))
        Case "fileb" ' -> man bekommt ne anfrage
            FileB CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4))
            '~ FileTransferBESTÄTIGUNG empfangen / SendeBestätigung der FileTransferBESTÄTIGUNG
        Case "acceptfile" ' -> bestätigung, dass man annimmt/nicht annimmt
            AcceptFile CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4))
        Case "acceptfileb" ' -> man bekommt ne anfrage
            AcceptFileB CStr(v(1)), CStr(v(2)), CStr(v(3))
        Case "typing"
            UserTyping CStr(v(1)), CStr(v(2)), CStr(v(3))
        Case "voice"
            Voice CStr(v(1)), CStr(v(2)), CStr(v(3))
        Case "askvoice"
            AskVoice CStr(v(1))
    
            '
            '
            ' ~ der admin hat was gemacht :)
        Case "kick"
            Kick CStr(v(1)), CStr(v(2))
        Case "ban"
            Ban CStr(v(1)), CStr(v(2))
        Case "makeadmin"
            MakeAdmin CStr(v(1)), CStr(v(2))
        Case "giveupadmin"
            GiveUpAdmin CStr(v(1))
        Case "loginadmin"
            LoginAdmin CStr(v(1))
        Case "kickvoice"
            KickVoice CStr(v(1)), CStr(v(2))
        End Select
    End If
    ' If Not HasActiveWindow Then
    '  modFunctions.FlashForm frmMain.hwnd, False ' ~ es ist was passiert, also form flashen
    ' End If

End Sub

Private Sub Form_Initialize()

    Call InitCommonControls

End Sub

Private Sub Form_Load()

    ReDim Typing(0)
    modFunctions2.LoadOptions
    With RTBMessage
        .SelFontName = modDeclaration.SavedOptions.Font.FontName
        .SelFontSize = modDeclaration.SavedOptions.Font.FontSize
        .SelColor = modDeclaration.SavedOptions.Font.FontColor
        .SelBold = modDeclaration.SavedOptions.Font.FontBold
        .SelItalic = modDeclaration.SavedOptions.Font.FontItalic
        .SelUnderline = modDeclaration.SavedOptions.Font.FontUnderline
        .SelStrikeThru = modDeclaration.SavedOptions.Font.FontStrikeThru
    End With 'RTBMESSAGE
    FillIntelliSenseWithAdminRights
    InitSmilies
    On Error Resume Next
    If modDeclaration.SavedOptions.SaveHistory Then
        RTBChat.LoadFile AppPath & "\history\" & GetDate & ".rtf"
        RTBChat.SelStart = Len(RTBChat.Text)
    End If
    On Error GoTo 0
    If Not IsIDE Then
        Call oMagneticWnd.AddWindow(Me.hwnd)
        EnableURLDetect RTBChat.hwnd, Me.hwnd
        modResize.Min_Width = 200
        modResize.Min_Height = 250
        modResize.Hook hwnd
    End If
    modSysTray.AddTray picTray, "", picTray
    Disconnected
    Me.Show
    DoEvents
    SignIn

End Sub

Private Sub Form_Resize()

Const Seperator As String = 6

    If Me.WindowState = vbMinimized Then
        Me.Hide
    End If
    On Error Resume Next
    '~ resize controls
    ' X
    RTBChat.Width = frmMain.ScaleWidth - RTBChat.Left - lstBuddys.Width - Seperator * 2
    lstBuddys.Left = frmMain.ScaleWidth - lstBuddys.Width - Seperator
    RTBMessage.Width = frmMain.ScaleWidth - RTBMessage.Left - lstBuddys.Width - Seperator * 2
    frameButtons.Left = lstBuddys.Left
    'Y
    RTBChat.Height = Me.ScaleHeight - RTBMessage.Height - StatusBar.Height - Seperator * 3
    lstBuddys.Height = RTBChat.Height
    RTBMessage.Top = Me.ScaleHeight - RTBMessage.Height - StatusBar.Height - Seperator
    frameButtons.Top = RTBMessage.Top
    ' Place-Shapes
    shpChat.Move RTBChat.Left - 1, RTBChat.Top - 1, RTBChat.Width + 2, RTBChat.Height + 2
    shpBuddy.Move lstBuddys.Left - 1, lstBuddys.Top - 1, lstBuddys.Width + 2, lstBuddys.Height + 2
    shpMessage.Move RTBMessage.Left - 1, RTBMessage.Top - 1, RTBMessage.Width + 2, RTBMessage.Height + 2
    RTBMessage_KeyUp 0, 0

End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim Frm As Form

    If modDeclaration.bLoadedfrmTransfer Then
        frmTransfer.tmrEnd.Enabled = True
    End If
    If modDeclaration.bLoadedfrmVoice Then
        Unload frmVoice
    End If
    For Each Frm In VB.Forms
        Frm.Hide
    Next Frm
    DoEvents
    For Each Frm In VB.Forms
        Unload Frm
        Set Frm = Nothing
    Next Frm
    modFunctions2.SaveOptions
    modSysTray.RemTray
    modRTBHighlight.DisableURLDetect
    modResize.Unhook
    oMagneticWnd.RemoveWindow hwnd
    Set oMagneticWnd = Nothing
    On Error Resume Next
    If modDeclaration.SavedOptions.SaveHistory Then
        MkDir App.Path & "\history"
        RTBChat.SaveFile AppPath & "\History\" & GetDate & ".rtf"
    End If
    On Error GoTo 0
    wsc.Close
    '  End
    ' This is the End
    ' my only friend the end
    ' lalalalalalalalalalallalalla

End Sub

Private Sub GiveUpAdmin(Username As String)

    AddTextToRTB Username & " hat seine Administrator-Privilegien aufgegeben!", EventColor, True, True

End Sub

Private Sub imgSmilie_Click(Index As Integer)

    RTBMessage.SelText = modDeclaration.SmilieArray(Index).CharCode
    picSmiliesContainer.Visible = False

End Sub

Private Sub imgSmilie_MouseMove(Index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

Dim i As Integer

    For i = 0 To imgSmilie.Count - 1
        imgSmilie(i).BorderStyle = IIf((i = Index), 1, 0)
    Next i

End Sub

Private Sub isButton1_Click()

    picSmiliesContainer.Visible = False

End Sub

Private Sub Kick(AdminName As String, _
                 Username As String)

    AddTextToRTB AdminName & " hat " & Username & " gekicked!", EventColor, True, True

End Sub

Private Sub KickVoice(AdminName As String, _
                      Username As String)

    AddTextToRTB AdminName & " hat " & Username & " aus dem VoiceChat gekicked!", EventColor, True, True

End Sub

Private Sub Login(LoginString As String)

' ~ hier schickt der Server zurück, ob er Username/Passwort akzeptiert hat

    If LoginString = "accept" Then
        tmrLogin.Enabled = False
        StatusBar.SimpleText = "Der Server hat die Login Anforderung akzeptiert..."
        ' so... anmeldevorgang erfolgreich - ab hier wird getestet,ob man online ist
        tmrConnectionCheck.Enabled = True
        Connected
    Else 'NOT LOGINSTRING...
        If LoginString = "deny" Then
            StatusBar.SimpleText = "Der Server hat die Login Anforderung nicht akzeptiert... Falsches Passwort?"
            Disconnected
            SignIn
        End If
    End If

End Sub

Private Sub LoginAdmin(Username As String)

    AddTextToRTB Username & " hat seine Administrator-Privilegien wiedererlangt!", EventColor, True, True

End Sub

Private Sub lstBuddys_FileDragDrop(Filename As String)

Dim ToUser   As String

    ToUser = lstBuddys.ListItems(lstBuddys.SelectedItem)
    If LenB(ToUser) = 0 Then Exit Sub
    If modDeclaration.SendingOrReceivingFile Then
        MsgBox "Sie senden oder empfangen bereits eine Datei."
    Else 'MODDECLARATION.SENDINGORRECEIVINGFILE = FALSE/0
        If Not ToUser = modDeclaration.Username Then
            modDeclaration.SendingOrReceivingFile = True
            modDeclaration.PathOfFileToSendOrReceive = Filename
            If Not FileExists(Filename) Then
                modDeclaration.SendingOrReceivingFile = False
                MsgBox "Die angebene Datei existiert nicht!", vbInformation
            Else 'NOT NOT...
                modDeclaration.ReceiverOrSender = ToUser
                SendData "file|" & ToUser & "|" & modFunctions.ExtractFilename(Filename) & "|" & modFunctions.GetFileLen(Filename)
            End If
        Else 'NOT NOT...
            MsgBox "Sie können keine Dateien an sich selbst senden!", vbInformation
        End If
    End If

End Sub

Private Sub lstIntelliSense_DblClick()

    RTBMessage.Text = "\\" & lstIntelliSense.Text
    RTBMessage.SelStart = Len(RTBMessage.Text)
    lstIntelliSense.Visible = False
    IntelliVisible = False
    RTBMessage.SetFocus

End Sub

Private Sub MakeAdmin(AdminName As String, _
                      Username As String)

    AddTextToRTB AdminName & " hat " & Username & " Administrator-Privilegien zugesprochen!", EventColor, True, True

End Sub

Private Sub Message(Username As String, _
                    Message As String, _
                    Admin As String)

' ~ eine gaaanz normale Message kommt an

    AddMessageTextToRTB Username, Message, False, (Admin = "true")
    FlashAndSound

End Sub

Private Sub mnuAbout_Click()

    On Error Resume Next
    picSmiliesContainer.Visible = False
    frmAbout.Show

End Sub

Private Sub mnuAudioA_Click()

    RunAudioAssistant Me.hwnd

End Sub

Private Sub mnuCloseMessenger_Click()

    Unload Me

End Sub

Private Sub mnuFileTransfer_Click()

    cmdSendFile_Click

End Sub

Private Sub mnuFont_Click()

Dim sFont As SelectedFont
Dim aFont As FONT_CONST

    On Error GoTo e_Trap
    FontDialog.iPointSize = 12 * 10
    With RTBMessage
        aFont.FontBold = .SelBold
        aFont.FontColor = .SelColor
        aFont.FontItalic = .SelItalic
        aFont.FontName = .SelFontName
        aFont.FontSize = .SelFontSize
        aFont.FontStrikeThru = .SelStrikeThru
        aFont.FontUnderline = .SelUnderline
        sFont = ShowFont(hwnd, aFont)
        If sFont.bCanceled = True Then Exit Sub
        .SelFontName = sFont.sSelectedFont
        .SelFontSize = sFont.nSize
        '.SelColor = sFont.lColor
        .SelBold = sFont.bBold
        .SelItalic = sFont.bItalic
        .SelUnderline = sFont.bUnderline
        .SelStrikeThru = sFont.bStrikeOut
    End With 'RTBMESSAGE
    ' gleich in die speciher-struct schreiben
    With modDeclaration.SavedOptions.Font
        .FontBold = RTBMessage.SelBold
        '.FontColor = RTBMessage.SelColor
        .FontItalic = RTBMessage.SelItalic
        .FontName = RTBMessage.SelFontName
        .FontSize = RTBMessage.SelFontSize
        .FontStrikeThru = RTBMessage.SelStrikeThru
        .FontUnderline = RTBMessage.SelUnderline
    End With 'MODDECLARATION.SAVEDOPTIONS.FONT
    RTBMessage.SetFocus
e_Trap:

End Sub

Private Sub mnuFontColor_Click()

Dim sColor As SelectedColor

    On Error GoTo e_Trap
    sColor.oSelectedColor = RTBMessage.SelColor
    sColor = ShowColor(Me.hwnd, RTBMessage.SelColor)
    If sColor.bCanceled = True Then Exit Sub
    RTBMessage.SelColor = sColor.oSelectedColor
    RTBMessage.SetFocus
    modDeclaration.SavedOptions.Font.FontColor = sColor.oSelectedColor
e_Trap:

End Sub



Private Sub mnuHelp_Click()
If FileExists(AppPath & "\help.txt") Then
ShellExecute hwnd, vbNullString, AppPath & "\help.txt", vbNullString, vbNullString, 5
Else
MsgBox "Konnte 'help.txt' nicht finden!", vbCritical
End If
End Sub

Private Sub mnuLogOut_Click()

    Disconnected
    SignIn

End Sub

Private Sub mnuMOptions_Click()

    frmOptions.Show

End Sub

Private Sub mnuNudge_Click()

    cmdNudge_Click

End Sub

Private Sub mnuOptions_Click()

    picSmiliesContainer.Visible = False

End Sub

Private Sub mnuOtherUser_Click()

    picSmiliesContainer.Visible = False

End Sub

Private Sub mnuprivateMessage_Click()

    cmdPrivateMessage_Click

End Sub

Private Sub mnuSendMessage_Click()

    cmdSend_Click

End Sub

Private Sub mnuSmilies_Click()

Dim i As Integer

    On Error Resume Next
    picSmiliesContainer.Visible = Not picSmiliesContainer.Visible
    For i = 0 To imgSmilie.Count - 1
        imgSmilie(i).BorderStyle = 0
    Next i
    picSmiliesContainer.SetFocus

End Sub

Private Sub mnuVBMessenger_Click()

    picSmiliesContainer.Visible = False

End Sub

Private Sub mnuVoiceChat_Click()

    picSmiliesContainer.Visible = False
    If Not bLoadedfrmVoice Then
        SendData "askvoice"
    End If

End Sub

Private Sub Nudge(Username As String, _
                  Admin As String)

' ~ ein Nudge kommt an

    AddTextToRTB Username & IIf(Admin = "true", " (Administrator)", "") & " hat Ihnen einen 'Nudge' gesendet", EventColor, True, True
    PlaySound NudgeSendOrReceived
    modFunctions.ShakeForm frmMain, 5, 1000

End Sub

Private Sub NudgeB(Username As String, _
                   Admin As String)

' ~ bestätigung, dass man einen Nudge an jmd. gesendet hat, damit man selber auch
' ein bischen leiden muss, wir der nudge auch hier ausgelöst

    PlaySound NudgeSendOrReceived
    AddTextToRTB "Sie haben einen 'Nudge' an " & Username & " gesendet", EventColor, True, True
    modFunctions.ShakeForm frmMain, 5, 1000

End Sub

Private Sub picSmiliesContainer_LostFocus()

    picSmiliesContainer.Visible = False

End Sub

Private Sub picTray_MouseMove(Button As Integer, _
                              Shift As Integer, _
                              X As Single, _
                              Y As Single)

Static lngMsg As Long

    lngMsg = X / Screen.TwipsPerPixelX
    Select Case lngMsg
    Case WM_LBUTTONDBLCLK
        On Error Resume Next
        frmMain.WindowState = vbNormal
        frmMain.Show
        On Error GoTo 0
    End Select

End Sub


Private Sub PrivateMessage(Username As String, _
                           Message As String, _
                           Admin As String)

' ~ eine private Nachricht kommt an

    AddMessageTextToRTB Username, Message, True, (Admin = "true")
    FlashAndSound

End Sub

Private Sub PrivateMessageB(Username As String, _
                            Message As String, _
                            Admin As String)

' ~ Bestätigung, dass man eine private Nachricht bekommen hat

    AddMessageTextToRTB modDeclaration.Username, Message, True, (Admin = "true")
    'AddTextToRTB "[" & modDeclaration.Username & "][Sie haben eine private Nachricht an " & Username & " gesendet] " & Message, , (Admin = "true"), True
    FlashAndSound

End Sub

Private Sub RTBChat_Change()

    modFunctions.RefreshRTB RTBChat

End Sub

Private Sub rtbmessage_Change()

    cmdSend.Enabled = Not LenB(RTBMessage.Text) = 0
    If LenB(RTBMessage.Text) = 0 Then
        With RTBMessage
            .SelFontName = modDeclaration.SavedOptions.Font.FontName
            .SelFontSize = modDeclaration.SavedOptions.Font.FontSize
            .SelColor = modDeclaration.SavedOptions.Font.FontColor
            .SelBold = modDeclaration.SavedOptions.Font.FontBold
            .SelItalic = modDeclaration.SavedOptions.Font.FontItalic
            .SelUnderline = modDeclaration.SavedOptions.Font.FontUnderline
            .SelStrikeThru = modDeclaration.SavedOptions.Font.FontStrikeThru
        End With 'RTBMESSAGE
    End If

End Sub

Private Sub RTBMessage_KeyDown(KeyCode As Integer, _
                               Shift As Integer)

    If (KeyCode = 32 Or KeyCode = 13) And IntelliVisible Then
        RTBMessage.Text = "\\" & lstIntelliSense.Text
        RTBMessage.SelStart = Len(RTBMessage.Text)
        lstIntelliSense.Visible = False
        IntelliVisible = False
        If KeyCode = 32 Then KeyCode = 0
    End If

End Sub

Private Sub RTBMessage_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        cmdSend_Click
        KeyAscii = 0
        Exit Sub
    End If
    If Not tmrType.Enabled Then
        SendData "typing|true"
    End If
    tmrType.Enabled = False
    tmrType.Enabled = True

End Sub

Private Sub RTBMessage_KeyUp(KeyCode As Integer, _
                             Shift As Integer)

Dim TextEntered As String
Dim Found       As Integer
Dim i           As Integer

    If Len(RTBMessage.Text) > 2 Then
        TextEntered = Mid$(RTBMessage.Text, 3, Len(RTBMessage.Text))
    End If
    If Left$(RTBMessage.Text, 2) = "\\" Then
        IntelliVisible = True                       ' generell anzeigen ( \\ am anfang)
        Found = -1
        For i = 0 To lstIntelliSense.ListCount - 1
            If Left$(lstIntelliSense.List(i), Len(TextEntered)) = TextEntered Then
                If Not Len(lstIntelliSense.List(i)) = Len(TextEntered) Then
                    Found = i
                End If
                Exit For
            End If
        Next i
        If Found = -1 Then
            IntelliVisible = False
                                 ' wenns nicht in der box gefunden werden kann, dann nicht
        Else 'NOT FOUND...
            lstIntelliSense.Selected(Found) = True
        End If
    Else 'NOT LEFT$(RTBMESSAGE.TEXT,...
        IntelliVisible = False
    End If
    If IntelliVisible Then
        lstIntelliSense.Top = RTBMessage.Top + modFunctions.GetTCursY + 2
        lstIntelliSense.Left = RTBMessage.Left + modFunctions.GetTCursX + 4
        lstIntelliSense.Visible = True
    Else 'INTELLIVISIBLE = FALSE/0
        lstIntelliSense.Visible = False
    End If

End Sub

Public Sub SendData(Text As String)

    If wsc.State = 7 Then
        wsc.SendData Text & Seperator
    End If

End Sub

Private Sub SignIn()

    modDeclaration.SignIn = False
    frmConnect.Show vbModal
    If Not modDeclaration.SignIn Then
        Unload Me
    Else 'NOT NOT...
        Caption = "VBMessenger - Angemeldet als " & modDeclaration.Username
        wsc.Connect modDeclaration.ServerIP, 81
    End If

End Sub

Private Sub tmrConnect_Timer()

' ~ es wird 10sec lang versucht zu connecten

    TimeOutC = TimeOutC + 1
    If wsc.State = 7 Then
        TimeOutC = 0
        StatusBar.SimpleText = "Verbunden... Sende Login Anforderung..."
        tmrConnect.Enabled = False
        tmrLogin.Enabled = True
        SendData "login|" & modDeclaration.Username & "|" & modDeclaration.UserPass
        Exit Sub
    End If
    If TimeOutC = 10 Then
        StatusBar.SimpleText = "Es konnte keine Verbindung hergestellt werden..."
        Disconnected
        SignIn
    End If

End Sub

Private Sub tmrConnectionCheck_Timer()

' ~ sobald die Verbindung steht wird dieser Timer aktiv und prüft, ob man noch
' mit dem Server verbunden ist.

    If wsc.State = 0 Or wsc.State = 8 Then
        StatusBar.SimpleText = "Sie wurden vom Server getrennt..."
        Disconnected
        SignIn
    End If

End Sub

Private Sub tmrGetData_Timer()

' Das prüfen kommt in einen Timer?
' WIESO?!
' Ganz einfach....
' angenommen, es kommt eine Nachricht an, und die wird gerade in dieser Schleife verarbeitei,
' allerdings im Winsock Data_Arrival. Dann gibt es große Probleme, wenn wärend dieser Bearbeitung ein
' weiteres Packets arrived! (hört sich komisch an, iss aber so)

Dim temp As Long

    Do While InStr(1, ReceiveBuffer, Seperator) > 0
        temp = InStr(1, ReceiveBuffer, Seperator)
        If temp > 1 Then
            FixedDataArrival Left$(ReceiveBuffer, temp - 1)
        End If
        ReceiveBuffer = Mid$(ReceiveBuffer, temp + Len(Seperator))
    Loop

End Sub

Private Sub tmrLogin_Timer()

' ~ ein TimeOut-Timer... wird disabled sobald der Login vom Server akzeptiert wurden
' und enabled sobald man verbunden ist

    tmrLogin.Enabled = False
    StatusBar.SimpleText = "Der Server antwortet nicht auf die Login Anforderung..."
    Disconnected
    SignIn

End Sub

Private Sub tmrType_Timer()

    tmrType.Enabled = False
    SendData "typing|false"

End Sub

Private Sub UpdateLst(Username As String, _
                      Online As String)

' ~ Der Server sendet die Benutzerliste
' dies tut er nur, wenn ein neuer Benutzer sich angemeldet hat

Dim InLst        As Boolean
Dim i            As Integer
Dim E            As Integer
Dim SelectedUser As String

    InLst = False ' man weiß ja nie...
    For i = 1 To lstBuddys.ListCount
        If lstBuddys.ListItems(i) = Username Then
            InLst = True
            E = i
            Exit For
        End If
    Next i
    If Online = "true" And Not InLst Then
        lstBuddys.AddItem Username
    ElseIf Not Online = "true" And InLst Then 'NOT ONLINE...
        lstBuddys.Remove (E)
    End If
    If Online = "true" Then PlaySound UserOnline

End Sub

Private Sub UserTyping(Username As String, _
                       IsTyping As String, Admin As String)

Dim i           As Integer
Dim E           As Integer
Dim InTypeArray As Boolean
Dim Strtext     As String

    For i = 1 To UBound(Typing)
        If Username = Typing(i).Username Then
            InTypeArray = True
            E = i
            Exit For
        End If
    Next i
    If Not InTypeArray Then
        ReDim Preserve Typing(UBound(Typing) + 1)
        E = UBound(Typing)
        Typing(E).Username = Username
    End If
    Typing(E).Typing = (IsTyping = "true")
    ' TODO
    E = 0
    For i = 1 To UBound(Typing)
        If Typing(i).Typing Then
            Strtext = Strtext & Typing(i).Username & ","
            E = E + 1
        End If
    Next i
    If Not E = 0 Then
        Strtext = Left$(Strtext, Len(Strtext) - 1)
        If E = 1 Then
            Strtext = Strtext & " macht eine Eingabe"
        Else 'NOT E...
            Strtext = Strtext & " machen eine Eingabe"
        End If
    End If
    StatusBar.SimpleText = Strtext
    ' -> es gibt ein array, indem der user nur einmal vorkommt
    ' -> das e gibt nun den Aufenthalt im array an

End Sub



Private Sub Voice(Username As String, VoiceEnabled As String, Admin As String)
If VoiceEnabled = "enabled" Then
    AddTextToRTB Username & " hat den VoiceChat betreten", EventColor, True, True
Else
    AddTextToRTB Username & " hat den VoiceChat verlassen", EventColor, True, True
End If
End Sub


Private Sub wsc_DataArrival(ByVal bytesTotal As Long)

Dim strData As String

    wsc.GetData strData
    ReceiveBuffer = ReceiveBuffer & strData

End Sub

Private Sub wsc_Error(ByVal Number As Integer, _
                      Description As String, _
                      ByVal Scode As Long, _
                      ByVal Source As String, _
                      ByVal HelpFile As String, _
                      ByVal HelpContext As Long, _
                      CancelDisplay As Boolean)

    StatusBar.SimpleText = "Es konnte keine Verbindung hergestellt werden..."
    Disconnected
    SignIn

End Sub


