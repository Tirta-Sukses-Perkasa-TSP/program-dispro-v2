Attribute VB_Name = "ModuleBF"
Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As Long
    lpszTitle      As Long
    ulFlags        As Long
    lpfnCallback   As Long
    lparam         As Long
    iImage         As Long
End Type

Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000
Private Const BIF_BROWSEINCLUDEFILES = &H4000
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_EDITBOX = &H10
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_USENEWUI = (BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)


Private Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" _
    (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" _
    Alias "lstrcatA" (ByVal lpString1 As String, _
    ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" _
    (lpBI As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" _
    (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Const CSIDL_DESKTOP = &H0
Const CSIDL_PROGRAMS = &H2
Const CSIDL_CONTROLS = &H3
Const CSIDL_PRINTERS = &H4
Const CSIDL_PERSONAL = &H5
Const CSIDL_FAVORITES = &H6
Const CSIDL_STARTUP = &H7
Const CSIDL_RECENT = &H8
Const CSIDL_SENDTO = &H9
Const CSIDL_BITBUCKET = &HA
Const CSIDL_STARTMENU = &HB
Const CSIDL_DESKTOPDIRECTORY = &H10
Const CSIDL_DRIVES = &H11
Const CSIDL_NETWORK = &H12
Const CSIDL_NETHOOD = &H13
Const CSIDL_FONTS = &H14
Const CSIDL_TEMPLATES = &H15
Const CSIDL_COMMON_STARTMENU = &H16
Const CSIDL_COMMON_PROGRAMS = &H17
Const CSIDL_COMMON_STARTUP = &H18
Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
Const CSIDL_APPDATA = &H1A
Const CSIDL_PRINTHOOD = &H1B

Private Type SHITEMID
    Cb   As Long
    AbID As Byte
End Type

Private Type ITEMIDLIST
    Mkid As SHITEMID
End Type

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function OleInitialize Lib "ole32.dll" (lp As Any) As Long
Private Declare Sub OleUninitialize Lib "ole32" ()


Private Function fGetSpecialFolder(CSIDL As Long, IDL As ITEMIDLIST) As String
Dim sPath As String
If SHGetSpecialFolderLocation(hWnd, CSIDL, IDL) = 0 Then

    sPath = Space$(MAX_PATH)
    If SHGetPathFromIDList(ByVal IDL.Mkid.Cb, ByVal sPath) Then
        fGetSpecialFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1) & "\"
    End If
End If
End Function


Public Function fBrowseForFolder(hwndOwner As Long, sPrompt As String) As String
Dim iNull    As Long
Dim lpIDList As Long
Dim lResult  As Long
Dim sPath    As String
Dim sPath1   As String
Dim udtBI    As BrowseInfo
Dim IDL      As ITEMIDLIST
'
sPath1 = fGetSpecialFolder(CSIDL_DESKTOP, IDL)
Call OleInitialize(ByVal 0&)

With udtBI
    .pIDLRoot = IDL.Mkid.Cb

    .hwndOwner = hwndOwner
    .lpszTitle = lstrcat(sPrompt, vbNullString)
    .ulFlags = BIF_RETURNONLYFSDIRS + BIF_USENEWUI
End With

lpIDList = SHBrowseForFolder(udtBI)

If lpIDList Then
    sPath = String$(MAX_PATH, 0)
    lResult = SHGetPathFromIDList(lpIDList, sPath)
    Call CoTaskMemFree(lpIDList)
    
    iNull = InStr(sPath, vbNullChar)
    If iNull Then sPath = Left$(sPath, iNull - 1)
End If

Call OleUninitialize
fBrowseForFolder = sPath
End Function



