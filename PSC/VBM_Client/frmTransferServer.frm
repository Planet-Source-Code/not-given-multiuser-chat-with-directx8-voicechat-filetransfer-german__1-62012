VERSION 5.00
Begin VB.Form frmTransferServer 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
End
Attribute VB_Name = "frmTransferServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type da
    FileToSend As String
    FileName As String
    RemoteIP As String
    FileSize As Double
    SaveAs As String
    Pstatus As Double
    LastAmount As Double
End Type
Private Info As da
Function getfilename(ByVal filepath As String)
On Error Resume Next
Dim ta() As String
ta = Split(filepath, "\")
getfilename = ta(UBound(ta))
End Function
Private Sub Command1_Click()
On Error Resume Next
w1.Close
Close #1
Close #2
End
End Sub

Private Sub Command2_Click()
On Error GoTo Y
C1.Filter = "All Files(*.*)|*.*"
C1.ShowOpen
Open C1.FileName For Append As #1
If LOF(1) = 0 Then
    MsgBox "The File Is Empty"
    Close #1
    Exit Sub
End If
Info.FileToSend = C1.FileName
Info.FileSize = LOF(1)
p1.Max = LOF(1) \ 2
Close #1
w1.SendData "sendrequest|" & getfilename(Info.FileToSend) & "|" & Info.FileSize & "|"
DoEvents
DoEvents
DoEvents
Y:
End Sub

Private Sub Command3_Click()
On Error Resume Next
MsgBox "File Sender ( Server ) By George Papadopoulos"
End Sub

Private Sub Form_Load()
On Error Resume Next
RefreshSock
Me.Show
End Sub

Private Sub spider_Timer()
On Error Resume Next
e = Info.LastAmount
e = e \ 1024
spid.Caption = e
Info.LastAmount = 0
DoEvents
End Sub

Private Sub unfreezer_Timer()
DoEvents
End Sub

Private Sub w1_Close()
status 1
End Sub

Private Sub w1_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
'Connection Requested
If w1.State <> sckClosed Then w1.Close
w1.Accept requestID
'Connection Accepted
status 2
End Sub

Private Sub w1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim dat As String
w1.GetData dat, vbString
If LCase(Mid(dat, 1, 11)) = "sendrequest" Then
    Dim temparray() As String
    Dim fname As String
    Dim fsize As Double
    temparray = Split(dat, "|")
    fname = temparray(1)
    fsize = temparray(2)
    Form2.fsize = fsize
    Form2.fname = fname
    pa = App.Path
    If Len(pa) = 3 Then pa = Mid(pa, 1, 2)
    pa = pa & "\"
    p1.Max = fsize \ 2
    Form2.fpath = pa & fname
    Form2.Show 1
    Exit Sub
End If

If LCase(Mid(dat, 1, 2)) = "ok" Then
    Dim temparray2() As String
    Dim fname2 As String
    Dim fsize2 As Double
    temparray2 = Split(dat, "|")
    fname2 = temparray2(1)
    fsize2 = temparray2(2)
    If fname2 <> getfilename(Info.FileToSend) Or fsize2 <> Info.FileSize Then Exit Sub
    Open Info.FileToSend For Binary Access Read As #1
        If LOF(1) = 0 Then Exit Sub
        Dim SendBuffer As String
        SendBuffer = Space$(LOF(1))
        Get #1, , SendBuffer
    Close #1
    DoEvents
    w1.SendData SendBuffer & "/\/\ENDOFFILE/\/\"
    e = 2
    r = Timer
    Do Until Timer > r + 2  'leave the pc to send the file
    DoEvents
    Loop
    Exit Sub
End If

If LCase(Mid(dat, 1, 5)) = "notok" Then
    MsgBox "The Client Does Not Accept The File Tranfer Request"
    Exit Sub
End If

If Right(dat, 17) = "/\/\ENDOFFILE/\/\" Then
    Dim aaa As String
    aaa = Mid(dat, 1, Len(dat) - 17)
    Put #2, , aaa
    Close #2
    MsgBox "File Transfer Completed"
    p1.Value = 0
    Command2.Enabled = True
    Exit Sub
End If

adt = Len(dat) \ 2
If adt + p1.Value > p1.Max Then p1.Value = p1.Max Else p1.Value = p1.Value + adt
Info.LastAmount = Info.LastAmount + Len(dat)
Put #2, , dat

End Sub

Private Sub w1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
MsgBox "Error : " & Description
RefreshSock
End Sub

Sub status(ByVal st As Integer)
On Error Resume Next
Select Case st
Case 1
    'disconnected
    Label3.Caption = "Disconnected"
    Me.Caption = "Server [Disconnected] - George Papadopoulos"
    Command2.Enabled = False
Case 2
    'connected
    cli.Caption = w1.RemoteHostIP
    Label3.Caption = "Connected"
    Me.Caption = "Server [Connected] - George Papadopoulos"
    Command2.Enabled = True
End Select
End Sub

Sub RefreshSock()
On Error Resume Next
status 1
w1.Close
w1.LocalPort = 1357
w1.Listen
DoEvents
End Sub

Private Sub w1_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
On Error Resume Next
Dim adt As Double
adt = bytesSent \ 2
If adt + p1.Value > p1.Max Then p1.Value = p1.Max Else p1.Value = p1.Value + adt
Info.LastAmount = Info.LastAmount + bytesSent
End Sub
