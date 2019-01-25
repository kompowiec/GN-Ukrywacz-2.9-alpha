Attribute VB_Name = "BrowseForFolder"
Option Explicit

Private Type BROWSEINFO
   hOwner As Long
   pidlRoot As Long
   pszDisplayName As String
   lpszTitle As String
   ulFlags As Long
   lpfn As Long
   lParam As Long
   iImage As Long
End Type

Private Declare Function SHGetPathFromIDList Lib "SHELL32.DLL" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "SHELL32.DLL" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000

Public Function BrowseFolder(f As Form, szDialogTitle As String) As String
   Dim x As Long, BI As BROWSEINFO, dwIList As Long, szPath As String, wPos As Integer
   
   BI.hOwner = f.hwnd
   BI.lpszTitle = szDialogTitle
   BI.ulFlags = BIF_RETURNONLYFSDIRS
   dwIList = SHBrowseForFolder(BI)
   szPath = Space$(512)
   x = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)
   If x Then
       wPos = InStr(szPath, Chr(0))
       BrowseFolder = Left$(szPath, wPos - 1)
   Else
       BrowseFolder = ""
   End If
End Function
