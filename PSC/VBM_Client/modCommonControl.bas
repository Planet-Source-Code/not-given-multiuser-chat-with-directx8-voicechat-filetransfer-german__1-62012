Attribute VB_Name = "modCommonControl"
Option Explicit
Private Type BROWSEINFO
    hwndOwner                              As Long
    pIDLRoot                               As Long
    pszDisplayName                         As Long
    lpszTitle                              As String
    ulFlags                                As Long
    lpfnCallback                           As Long
    lParam                                 As Long
    iImage                                 As Long
End Type
Private Const MAX_PATH                 As Long = 260
Private Const BIF_RETURNONLYFSDIRS     As Long = &H1
Private Declare Function SHBrowseForFolder Lib "shell32" (lpBI As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, _
                                                            ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32" (ByVal hMem As Long)

Public Function BrowseForFolder(Optional Parent As Variant, _
                                Optional Title As Variant) As String

Dim tBI         As BROWSEINFO
Dim lhWndParent As Long
Dim lngPIDL     As Long
Dim strPath     As String

    If IsMissing(Title) Then
        Title = "Wählen Sie einen Ordner aus"
    End If
    If IsMissing(Parent) = False Then
        lhWndParent = Parent.hwnd
    End If
    With tBI
        .hwndOwner = lhWndParent
        .lpszTitle = Title
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With 'TBI
    lngPIDL = SHBrowseForFolder(tBI)
    If lngPIDL <> 0 Then
        '// Pfad aus Item ID List ermitteln:
        strPath = Space$(MAX_PATH)
        SHGetPathFromIDList lngPIDL, strPath
        strPath = Left$(strPath, InStr(strPath, vbNullChar) - 1)
        '// PIDL freigeben:
        CoTaskMemFree lngPIDL
    End If
    BrowseForFolder = strPath

End Function



