VERSION 5.00
Begin VB.Form frmFolder 
   Caption         =   "Wybierz folder"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   Icon            =   "GN UkrywaczFolder.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstFoldery 
      Height          =   4170
      IntegralHeight  =   0   'False
      Left            =   3480
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdAnuluj 
      Cancel          =   -1  'True
      Caption         =   "&Anuluj"
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1080
      TabIndex        =   2
      Top             =   4560
      Width           =   1575
   End
   Begin VB.DriveListBox drvNapedy 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   3135
   End
   Begin VB.DirListBox dirFoldery 
      Height          =   3690
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim blnBezHasla As Boolean
Dim blnAuto As Boolean
Dim blnDostepny As Boolean
Dim blnChroniony As Boolean
Dim blnIstnieje As Boolean
Dim strFolder As String
Dim strHaslo As String

Dim intLicznik As Integer
Const strKodKlasy As String = ".{21EC2020-3AEA-1069-A2DD-08002B30309D}"
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
  dwFileAttributes As Long
  ftCreationTime As FILETIME
  ftLastAccessTime As FILETIME
  ftLastWriteTime As FILETIME
  nFileSizeHigh As Long
  nFileSizeLow As Long
  dwReserved0 As Long
  dwReserved1 As Long
  cFileName As String * MAX_PATH
  cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Sub SearchDir(ByVal strPath As String, ByVal strMask As String)
  Dim lHandle As Long, lRet As Long, strTPath As String
  Dim fdFile As WIN32_FIND_DATA
  Dim strFile As String

  If Right$(strPath, 1) <> "\" Then strPath = strPath & "\"
  strTPath = strPath & strMask

  fdFile.dwFileAttributes = 1
  lRet = 1
  lHandle = FindFirstFile(strTPath, fdFile)
  Do While lHandle <> INVALID_HANDLE_VALUE And lRet <> 0
    strFile = strip(fdFile.cFileName)
    If strFile = "." Or strFile = ".." Then
      'Omijamy
    ElseIf (fdFile.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = _
      FILE_ATTRIBUTE_DIRECTORY Then
      'Katalog
      'Debug.Print "dir " & strPath & strFile
      If Right(strFile, Len(".{21EC2020-3AEA-1069-A2DD-08002B30309D}")) = ".{21EC2020-3AEA-1069-A2DD-08002B30309D}" Then strFile = Left(strFile, Len(strFile) - Len(".{21EC2020-3AEA-1069-A2DD-08002B30309D}")) & " (ukryty)"
      lstFoldery.AddItem strFile
    Else
      'Plik
    End If
    lRet = FindNextFile(lHandle, fdFile)
  Loop
  FindClose lHandle
End Sub

Private Function strip(ByVal str As String) As String
  Dim i As Long
  i = InStr(str, Chr$(0))
  If i > 0 Then
    strip = Left$(str, i - 1)
  Else
    strip = str
  End If
End Function

Private Sub cmdAnuluj_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
blnBezHasla = frmForm1.blnBezHasla
blnAuto = frmForm1.blnAuto
blnDostepny = frmForm1.blnDostepny
blnChroniony = frmForm1.blnChroniony
blnIstnieje = frmForm1.blnIstnieje
strFolder = frmForm1.strFolder
strHaslo = frmForm1.strHaslo
   For intLicznik = 0 To lstFoldery.ListCount - 1
      If lstFoldery.Selected(intLicznik) = True Then
         If Len(dirFoldery.Path) > 3 Then
            strFolder = dirFoldery.Path & "\" & lstFoldery.List(intLicznik)
         Else
            strFolder = dirFoldery.Path & lstFoldery.List(intLicznik)
         End If
         'sprawdzenie, czy wybrany folder nie jest folderem windozy
         If UCase(strFolder) = UCase(GetWindowsDir) Then MsgBox "Wybrany folder jest folderem Windowsa. Ukrycie go spowodowa³oby zawiesznie siê komputera i problemy z uruchomieniem Windowsa.", vbExclamation + vbOKOnly, "B³¹d wybierania foldera": Exit Sub
         'sprawdzenie, czy ktoœ nie wybra³ program files. Jest to sprawdzane trochê lamersko, ale w 99,9 % przypadków wystarczy
         If UCase(strFolder) = UCase("c:\Program Files") Then MsgBox "Ukrycie tego foldera spowoduje niemo¿liwoœæ uruchomienia wielu programów.", vbExclamation + vbOKOnly, "GN Ukrywacz - bl¹d wybierania foldera": Exit Sub
         'sprawdznie, czy ktoœ przypadkiem nie chce ukryæ foldera, w którym jest GN Ukrywacz
         If InStr(UCase(App.Path) & "\", UCase(strFolder) & "\") = 1 Then MsgBox "Nie mo¿na ukryæ foldera, w którym znajduje siê GN Ukrywacz.", vbExclamation + vbOKOnly, "GN Ukrywacz - b³¹d wybierania foldera": Exit Sub
         If Right(strFolder, Len(" (ukryty)")) = " (ukryty)" Then strFolder = Left(strFolder, Len(strFolder) - Len(" (ukryty)")) & strKodKlasy
         Properties strFolder, frmForm1.cmdZmien, frmForm1.lblStan, frmForm1.txtHaslo, strHaslo, blnDostepny, blnIstnieje, blnChroniony, blnAuto, blnBezHasla, frmForm1.blnAdmin
         frmForm1.strFolder = strFolder
         frmForm1.strHaslo = strHaslo
         frmForm1.blnAuto = blnAuto
         frmForm1.blnBezHasla = blnBezHasla
         frmForm1.blnChroniony = blnChroniony
         frmForm1.blnDostepny = blnDostepny
         frmForm1.blnIstnieje = blnIstnieje
         frmForm1.lblFolder.Caption = strFolder
         Exit For
      End If
   Next intLicznik
   Unload Me
End Sub

Private Sub dirFoldery_Change()
   lstFoldery.Clear
   SearchDir dirFoldery.Path, "*.*"
   cmdOK.Enabled = False
End Sub

Private Sub drvNapedy_Change()
   On Error Resume Next
   dirFoldery.Path = UCase(drvNapedy.Drive)
   dirFoldery.Path = Mid(dirFoldery.Path, 1, 3)
   If Err.Number <> 0 Then MsgBox "Wybrany napêd jest niedostêpny.", vbExclamation + vbSystemModal + vbOKOnly, "B³¹d wejœcia - wyjœcia"
End Sub

Private Sub Form_Load()
   On Error Resume Next
   dirFoldery.Path = "C:\"
   drvNapedy.Drive = Mid(dirFoldery.Path, 1, 3)
   Me.Icon = frmForm1.Icon
End Sub

Private Sub lstFoldery_Click()
   cmdOK.Enabled = True
End Sub

Private Sub lstFoldery_DblClick()
   cmdOK_Click
End Sub
