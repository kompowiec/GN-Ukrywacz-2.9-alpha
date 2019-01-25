Attribute VB_Name = "KodujDekodujFileExistsWindirPropertiesPass"
Option Explicit
Dim lngDlugosc As Long
Dim intPozycja As Integer
Dim strHaselko As String
Dim lngLicznik As Long

Const strKodKlasy As String = ".{21EC2020-3AEA-1069-A2DD-08002B30309D}"
Private Const MAX_PATH = 260
Private Const INVALID_HANDLE_VALUE = -1
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

Private Declare Function FindFirstFileA Lib "kernel32" _
(ByVal lpFileName As String, _
lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" _
(ByVal hFindFile As Long) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias _
   "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
 
Public Function GetWindowsDir() As String
   Dim strBuffer As String, lRet As Long
   
   strBuffer = String$(MAX_PATH, Chr(0))
   lRet = GetWindowsDirectory(strBuffer, Len(strBuffer))
   GetWindowsDir = Left$(strBuffer, lRet)
End Function

Public Sub KodujPlik()
   'Procedura koduj¹ca plik
   strHaselko = "herloom" 'Mo¿e byæ oczywiœcie inne ale takie samo jak w procedurze dekoduj¹cej
   'Przepisywanie pliku do tablicy.
   Open "dirdat.sys" For Binary As #1
   lngDlugosc = LOF(1)
   If lngDlugosc > 0 Then
      ReDim bytTablica(lngDlugosc - 1) As Byte
      ReDim intTablica(lngDlugosc - 1) As Integer
      Get #1, , bytTablica
      Close
      'Szyfrowanie bajtów z tablicy
      'dodawanie bajtów has³a do bajtów pliku, jes³i suma>255, wtedy odjêcie 256
      lngLicznik = 0
      Do
         For intPozycja = 1 To Len(strHaselko)
            intTablica(lngLicznik) = bytTablica(lngLicznik) + Asc(Mid(strHaselko, intPozycja, 1))
            If intTablica(lngLicznik) > 255 Then intTablica(lngLicznik) = intTablica(lngLicznik) - 256
            bytTablica(lngLicznik) = intTablica(lngLicznik)
            lngLicznik = lngLicznik + 1
            If lngLicznik > (lngDlugosc - 1) Then Exit For
         Next intPozycja
      Loop Until lngLicznik > (lngDlugosc - 1)
      'Przepisywanie tablicy do pliku.
      Open "dirdat.tmp" For Binary As #1
      Put #1, , bytTablica
      Close
      Kill "dirdat.sys"  'zamiana plików
      Name "dirdat.tmp" As "dirdat.sys"
   End If
   Close
End Sub

Public Sub DekodujPlik()
   'procedura dekoduj¹ca
   strHaselko = "herloom"
   'Przepisywanie pliku do tablicy.
   Open "dirdat.sys" For Binary As #1
   lngDlugosc = LOF(1)
   If lngDlugosc > 0 Then
      ReDim bytTablica(lngDlugosc - 1) As Byte
      ReDim intTablica(lngDlugosc - 1) As Integer
      Get #1, , bytTablica
      Close
      'Deszyfrowanie bajtów z tablicy, tak jak szyfrowanie, ale bajty s¹ odejmowane zamiast dodawane
      lngLicznik = 0
      Do
         For intPozycja = 1 To Len(strHaselko)
            intTablica(lngLicznik) = bytTablica(lngLicznik) - Asc(Mid(strHaselko, intPozycja, 1))
            If intTablica(lngLicznik) < 0 Then intTablica(lngLicznik) = intTablica(lngLicznik) + 256
            bytTablica(lngLicznik) = intTablica(lngLicznik)
            lngLicznik = lngLicznik + 1
            If lngLicznik > (lngDlugosc - 1) Then Exit For
         Next intPozycja
      Loop Until lngLicznik > (lngDlugosc - 1)
      'Przepisywanie tablicy do pliku.
      Open "dirdat.tmp" For Binary As #1
      Put #1, , bytTablica
      Close
      Kill "dirdat.sys"
      Name "dirdat.tmp" As "dirdat.sys"
   End If
   Close
End Sub

Public Function FileExists(ByVal sFile As String) As Boolean
   '-------------------------------------------------------------'
   '   Okreœla, czy podana œcie¿ka lub plik ju¿ istniej¹.        '
   '-------------------------------------------------------------'
   Dim r As Long
   Dim uFIND_DATA As WIN32_FIND_DATA
       
   r = FindFirstFileA(sFile, uFIND_DATA)
   If r = INVALID_HANDLE_VALUE Then
      FileExists = False
   Else
      FileExists = True
      Call FindClose(r)
   End If
End Function

Public Sub Szyfruj(ByVal strPlik As String)
   Open strPlik For Binary As #5
   ReDim bytTablica(1 To LOF(5)) As Byte
   Get #5, , bytTablica
   For intPozycja = 1 To LOF(5)
      bytTablica(intPozycja) = bytTablica(intPozycja) Xor 55
   Next intPozycja
   Close #5
   Kill strPlik
   Open strPlik For Binary As #5
   Put #5, , bytTablica
   Close #5
End Sub

Public Sub Deszyfruj(ByVal strPlik As String)
   Szyfruj (strPlik)
End Sub

Public Sub Properties(ByRef strFolder As String, ByRef cmdZmien As CommandButton, ByRef lblStan As Label, ByRef txtHaslo As TextBox, ByRef strHaslo As String, ByRef blnDostepny As Boolean, ByRef blnIstnieje As Boolean, ByRef blnChroniony As Boolean, ByRef blnAuto As Boolean, ByRef blnBezHasla As Boolean, ByVal blnAdmin)
   Dim strTekst As String
   blnChroniony = False
   blnIstnieje = False
   If strFolder = "" Then Exit Sub
   If FileExists(strFolder) = False Then 'jeœli folder nie istnieje, sprawdzamy czy istnieje jego ukryta wersja
      If FileExists(strFolder & strKodKlasy) = True Then
         blnIstnieje = True
         strFolder = strFolder & strKodKlasy
      Else
         lblStan.Caption = "Folder nie istnieje."
         blnIstnieje = False
         cmdZmien.Caption = ""
         cmdZmien.Enabled = False
         Exit Sub
      End If
   Else
      If Right(strFolder, Len(strKodKlasy)) = strKodKlasy Then 'jeœli wybrany zosta³ folder ukryty
         If FileExists(Left(strFolder, Len(strFolder) - Len(strKodKlasy))) Then 'sprawdamy, czy istnieje te¿ udostêpniony folder
            cmdZmien.Enabled = False
            lblStan.Caption = "B³¹d"
            strHaslo = ""
            MsgBox "Istniej¹ dwa foldery o tej samej nazwie(nie licz¹c rozszerzenia), z których jeden jest ukryty, a drugi udostêpniony. Zanim bêdzie mo¿liwe wykonanie jakiejkolwiek czynnoœci z którymœ z tych folderów za pomoc¹ programu GN Ukrywacz, nale¿y jednemu z nich zmieniæ nazwê.", vbExclamation + vbOKOnly, "GN Ukrywacz - dwa foldery o tej samej nazwie"
            Exit Sub
         End If
      Else 'jeœli wybrany zosta³ udostêpniony folder
         If FileExists(strFolder & strKodKlasy) Then 'sprawdzamy, czy istnieje te¿ ukryty folder
            cmdZmien.Enabled = False
            lblStan.Caption = "B³¹d"
            strHaslo = ""
            MsgBox "Istniej¹ dwa foldery o tej samej nazwie(nie licz¹c rozszerzenia), z których jeden jest ukryty, a drugi udostêpniony. Zanim bêdzie mo¿liwe wykonanie jakiejkolwiek czynnoœci z którymœ z tych folderów za pomoc¹ programu GN Ukrywacz, nale¿y jednemu z nich zmieniæ nazwê.", vbExclamation + vbOKOnly, "GN Ukrywacz - dwa foldery o tej samej nazwie"
            Exit Sub
         End If
      End If
      blnIstnieje = True
   End If
   If (GetAttr(strFolder) And vbDirectory) = 0 Then 'sprawdzamy, czy wybrany zosta³ folder a nie plik
      cmdZmien.Caption = ""
      cmdZmien.Enabled = False
      lblStan.Caption = "B³¹d"
      MsgBox "Wskazany zosta³ plik a nie folder.", vbExclamation + vbOKOnly, "GN Ukrywacz - z³y folder"
      blnIstnieje = False
      blnChroniony = False
      Exit Sub
   End If
   If Right(strFolder, Len(strKodKlasy)) = strKodKlasy Then 'ukryty
      blnDostepny = False
      If FileExists(strFolder & "\gnukr2_0_etc_shadow.pass") = True Then
         lblStan.Caption = "Folder jest ukryty."
         cmdZmien.Caption = "&Udostêpnij"
         blnChroniony = True
      Else
         lblStan.Caption = "Folder jest ukryty, ale nie jest chroniony has³em."
         cmdZmien.Enabled = True
         cmdZmien.Caption = "&Udostêpnij"
         blnChroniony = False
      End If
   Else 'udostêpniony
      blnDostepny = True
      If FileExists(strFolder & "\gnukr2_0_etc_shadow.pass") = True Then
         lblStan.Caption = "Folder jest udostêpniony."
         cmdZmien.Caption = "&Ukryj"
         blnChroniony = True
      Else
         lblStan.Caption = "Folder nie jest chroniony przez GN Ukrywacz."
         cmdZmien.Caption = "&Chroñ"
         cmdZmien.Enabled = True
         blnChroniony = False
      End If
   End If
   If blnChroniony = True Then
      blnAuto = False
      Open "auto.dat" For Input As #1
      Do Until EOF(1)
         Input #1, strTekst
         If blnDostepny = True Then
            If UCase(strTekst) = UCase(strFolder) Then blnAuto = True
         Else
            If UCase(strTekst) & strKodKlasy = UCase(strFolder) Then blnAuto = True
         End If
      Loop
      Close
      SetAttr strFolder & "\gnukr2_0_etc_shadow.pass", vbNormal
      Deszyfruj strFolder & "\gnukr2_0_etc_shadow.pass"
      Open strFolder & "\gnukr2_0_etc_shadow.pass" For Input As #1
      Input #1, strHaslo
      Input #1, blnBezHasla
      Close
      Szyfruj strFolder & "\gnukr2_0_etc_shadow.pass"
      PasswordCheck strHaslo, txtHaslo, cmdZmien, blnChroniony, blnBezHasla, frmForm1.mnuUstawienia, blnDostepny, blnAdmin
   End If
End Sub

Public Sub PasswordCheck(ByRef strHaslo As String, ByRef txtHaslo As TextBox, ByRef cmdZmien As CommandButton, ByRef blnChroniony As Boolean, ByRef blnBezHasla As Boolean, ByRef mnuUstawienia As Menu, ByRef blnDostepny As Boolean, ByVal blnAdmin As Boolean)
   If blnChroniony = True Then
      If (UCase(txtHaslo.Text) = UCase(strHaslo)) And (strHaslo <> "") Or blnAdmin Then
         cmdZmien.Enabled = True
         mnuUstawienia.Enabled = True
      Else
         cmdZmien.Enabled = False
         mnuUstawienia.Enabled = False
         If blnBezHasla = True And blnDostepny = True Then cmdZmien.Enabled = True
      End If
   Else
      mnuUstawienia.Enabled = False
   End If
End Sub

Public Function EncryptAdminPasswd(ByVal strText As String) As String
   Dim intLicznik As Integer
   For intLicznik = 1 To Len(strText)
      Mid(strText, intLicznik, 5) = Chr(Asc(Mid(strText, intLicznik, 1)) + 5)
   Next intLicznik
   EncryptAdminPasswd = strText
End Function

Public Function DecryptAdminPasswd(ByVal strText As String) As String
   Dim intLicznik As Integer
   For intLicznik = 1 To Len(strText)
      Mid(strText, intLicznik, 5) = Chr(Asc(Mid(strText, intLicznik, 1)) - 5)
   Next intLicznik
   DecryptAdminPasswd = strText
End Function

