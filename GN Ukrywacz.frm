VERSION 5.00
Begin VB.Form frmForm1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GN Ukrywacz"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6135
   Icon            =   "GN Ukrywacz.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdWyczysc 
      Caption         =   "Wyczy��"
      Height          =   495
      Left            =   5160
      TabIndex        =   10
      Top             =   1920
      Width           =   855
   End
   Begin VB.CommandButton cmdUsun 
      Caption         =   "Usu�"
      Height          =   495
      Left            =   5160
      TabIndex        =   9
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdDodaj 
      Caption         =   "Dodaj..."
      Height          =   495
      Left            =   5160
      TabIndex        =   8
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdPrzegladaj 
      Caption         =   "Przegl�daj..."
      Height          =   375
      Left            =   4200
      TabIndex        =   4
      Top             =   3480
      Width           =   1695
   End
   Begin VB.CommandButton cmdZmien 
      Enabled         =   0   'False
      Height          =   735
      Left            =   2040
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4800
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.TextBox txtHaslo 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   4080
      Width           =   5055
   End
   Begin VB.ListBox lstFoldery 
      Height          =   2205
      ItemData        =   "GN Ukrywacz.frx":0442
      Left            =   120
      List            =   "GN Ukrywacz.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4860
   End
   Begin VB.Frame Frame1 
      Caption         =   "Folder"
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   5895
      Begin VB.Label lblStan 
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   "Has�o:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblFolder 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5655
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Ulubione:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu mnuPlik 
      Caption         =   "&Plik"
      Begin VB.Menu mnuPlikKoniec 
         Caption         =   "&Koniec"
      End
   End
   Begin VB.Menu mnuUstawienia 
      Caption         =   "&Ustawienia"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "&Administrator"
      Begin VB.Menu mnuAdminLogon 
         Caption         =   "&Zaloguj..."
      End
      Begin VB.Menu mnuAdminPasswd 
         Caption         =   "Zmie� &has�o..."
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuPomoc 
      Caption         =   "Pomo&c"
      Begin VB.Menu mnuPomocOchrona 
         Caption         =   "Na czym polega ochrona foldera"
      End
      Begin VB.Menu mnuPomocFolder 
         Caption         =   "Wybieranie foldera"
      End
      Begin VB.Menu mnuPomocUkrBezH 
         Caption         =   "Ukrywanie bez has�a"
      End
      Begin VB.Menu mnuPomocAuto 
         Caption         =   "Automatyczne ukrywanie"
      End
      Begin VB.Menu mnuPomocHaslo 
         Caption         =   "Zmiana has�a"
      End
      Begin VB.Menu mnuKreska 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPomocOProgramie 
         Caption         =   "O programie..."
      End
   End
End
Attribute VB_Name = "frmForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Grzegorz Niemirowski
'grzegorz@grzegorz.net
'www.grzegorz.net


Option Explicit

Dim intZmienna As Integer
Dim intLicznik As Integer
Dim intIloscFolderow As Integer
Public blnBezHasla As Boolean
Public blnAuto As Boolean
Public blnDostepny As Boolean
Public blnChroniony As Boolean
Public blnIstnieje As Boolean
Public strFolder As String
Dim strTekst As String
Dim strSciezka As String
Public strHaslo As String
Public strUlubiony As String
Public strAdminPasswd As String
Public blnAdmin As Boolean
Const strKodKlasy As String = ".{21EC2020-3AEA-1069-A2DD-08002B30309D}"

Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long

Private Sub cmdDodaj_Click()
   frmUlubiony.Show vbModal, Me
   If Len(strUlubiony) < 4 Then Exit Sub
   lstFoldery.AddItem (strUlubiony)
   Open "ulubione.fav" For Output As #1
   For intLicznik = 0 To lstFoldery.ListCount - 1
      Write #1, lstFoldery.List(intLicznik)
   Next intLicznik
   Close
End Sub

Private Sub cmdPrzegladaj_Click()
   frmFolder.Show vbModal, Me
End Sub

Private Sub cmdUsun_Click()
   For intLicznik = 0 To lstFoldery.ListCount - 1
      If lstFoldery.Selected(intLicznik) = True Then
         lstFoldery.RemoveItem intLicznik
         Exit For
      End If
   Next intLicznik
   ChDir App.Path
   Open "ulubione.fav" For Output As #1
   For intLicznik = 0 To lstFoldery.ListCount - 1
      Write #1, lstFoldery.List(intLicznik)
   Next intLicznik
   Close
End Sub

Private Sub cmdWyczysc_Click()
   intZmienna = MsgBox("Wyczy�ci� list� Ulubionych?", vbQuestion + vbYesNo, "GN Ukrywacz - czyszczenie ulubionych")
   If intZmienna = vbYes Then
      lstFoldery.Clear
      Open "ulubione.fav" For Output As #1
      Close
   End If
End Sub

Private Sub cmdZmien_Click()
   'procedura ukrywa lub udost�pnia folder
   If strFolder = "" Then Exit Sub
   If blnChroniony = False Then
      If blnDostepny = True Then
         On Error Resume Next
         frmHaslo.Show vbModal, Me
         If strHaslo = "" Then Exit Sub
         intZmienna = MsgBox("Czy ma by� mo�liwo�� ukrywania foldera bez podawania has�a?", vbQuestion + vbYesNo, "GN Ukrywacz - ukrywanie bez has�a")
         If intZmienna = vbYes Then blnBezHasla = True Else blnBezHasla = False
         intZmienna = MsgBox("Czy folder ma by� ukrywany przy uruchamianiu komputera?", vbQuestion + vbYesNo, "GN Ukrywacz - automatyczne ukrywanie")
         If intZmienna = vbYes Then blnAuto = True Else blnAuto = False
         'utworzenie pliku z has�em w chronionym folderze
         Open strFolder & "\gnukr2_0_etc_shadow.pass" For Output As #1
         If Err.Number <> 0 Then
            MsgBox "Nie mo�na ustawi� has�a dla foldera. Znajduje si� on na dysku zabezpieczonym przed zapisem lub brak dysku w nap�dzie.", vbExclamation + vbOKOnly, "GN Ukrywacz - nie mo�na ustawi� has�a"
            Err.Clear
            Exit Sub
         End If
         Write #1, strHaslo
         Write #1, blnBezHasla
         Close
         Szyfruj strFolder & "\gnukr2_0_etc_shadow.pass"
         'je�li ma by� automatycznie ukrywany, to dopisanie go do listy
         If blnAuto = True Then
            Open App.Path & "\auto.dat" For Append As #1
            Write #1, strFolder
            Close
         End If
         cmdZmien.Caption = "&Ukryj"
         blnChroniony = True
         blnIstnieje = True
         blnDostepny = True
         txtHaslo.Text = strHaslo
         lblStan.Caption = "Folder jest udost�pniony."
      Else
         Name strFolder As Left(strFolder, Len(strFolder) - Len(strKodKlasy))
         strFolder = Left(strFolder, Len(strFolder) - Len(strKodKlasy))
         SetAttr strFolder, (GetAttr(strFolder) And 239) And (255 - vbHidden)
         blnDostepny = True
         txtHaslo = ""
         Properties strFolder, cmdZmien, lblStan, txtHaslo, strHaslo, blnDostepny, blnIstnieje, blnChroniony, blnAuto, blnBezHasla, blnAdmin
      End If
   Else
      On Error Resume Next
      If blnDostepny = True Then 'ukrycie foldera
         SetAttr strFolder, (GetAttr(strFolder) And 239) Or vbHidden  'Ustawia atrybut ukryty
         If Err.Number <> 0 Then 'Sprawdzenie czy nie wystapi� b��d zapisu
            Err.Clear
            intZmienna = MsgBox("Nie mo�na ukry� tego foldera. Znajduje si� on na dysku zabezpieczonym przed zapisem lub brak dysku z tym folderem w nap�dzie.", vbOKOnly + vbExclamation + vbSystemModal, "B��d wej�cia-wyj�cia")
         Else
            Name strFolder As strFolder & strKodKlasy
            strFolder = strFolder & strKodKlasy
            blnDostepny = False
            lblStan.Caption = "Folder jest ukryty."
            cmdZmien.Caption = "&Udost�pnij"
            PasswordCheck strHaslo, txtHaslo, cmdZmien, blnChroniony, blnBezHasla, mnuUstawienia, blnDostepny, blnAdmin
            Properties strFolder, cmdZmien, lblStan, txtHaslo, strHaslo, blnDostepny, blnIstnieje, blnChroniony, blnAuto, blnBezHasla, blnAdmin
         End If
      Else  'Udost�pnienie foldera, czynno�ci j.w. ale w odwrotnym kierunku
         SetAttr strFolder, (GetAttr(strFolder) And 239) And (255 - vbHidden)
         If Err.Number <> 0 Then
            Err.Clear
            intZmienna = MsgBox("Nie mo�na udost�pni� tego foldera. Znajduje si� on na dysku zabezpieczonym przed zapisem lub brak dysku z tym folderem w nap�dzie.", vbOKOnly + vbExclamation + vbSystemModal, "B��d wej�cia-wyj�cia")
         Else
            strFolder = Left(strFolder, Len(strFolder) - Len(strKodKlasy))
            Name strFolder & strKodKlasy As strFolder
            lblStan.Caption = "Folder jest udost�pniony."
            cmdZmien.Caption = "&Ukryj"
            Properties strFolder, cmdZmien, lblStan, txtHaslo, strHaslo, blnDostepny, blnIstnieje, blnChroniony, blnAuto, blnBezHasla, blnAdmin
         End If
      End If
   End If
End Sub

Private Sub Form_GotFocus()
   Properties strFolder, cmdZmien, lblStan, txtHaslo, strHaslo, blnDostepny, blnIstnieje, blnChroniony, blnAuto, blnBezHasla, blnAdmin
End Sub

Private Sub Form_Initialize()
   InitCommonControls
End Sub

Private Sub Form_Load()
   If App.PrevInstance = True Then MsgBox "   Nie mo�na uruchomi� wi�cej ni� jednej kopii programu w danej chwili.", vbExclamation + vbOKOnly, "GN Ukrywacz - problem z uruchomieniem": End
   cmdZmien.Picture = frmForm1.Icon
   ChDir App.Path
   'sprawdzenie, czy istnieje plik "dirdat.sys"
   If FileExists(App.Path & "\dirdat.sys") = True Then
      'import danych ze starej wersji
      DekodujPlik
      Open "dirdat.sys" For Input As #1
      Open "auto.dat" For Append As #3
      Do Until EOF(1)
         Input #1, strFolder
         Input #1, strHaslo
         Input #1, blnBezHasla
         Input #1, blnAuto
         Input #1, blnDostepny
         If blnDostepny = True Then Open strFolder & "\gnukr2_0_etc_shadow.pass" For Output As #2 Else Open strFolder & ".{21EC2020-3AEA-1069-A2DD-08002B30309D}\gnukr2_0_etc_shadow.pass" For Output As #2
         Write #2, strHaslo
         Write #2, blnBezHasla
         Close #2
         Szyfruj strFolder & "\gnukr2_0_etc_shadow.pass"
         If blnAuto = True Then Write #3, strFolder
      Loop
      Close
      Kill "dirdat.sys"
   End If
   'wpis do rejestru, uruchamiaj�cy ukrywanie przy starcie
   SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "GN Ukrywacz", Chr(34) & App.Path & "\" & App.EXEName & Chr(34) & " /auto"
   strAdminPasswd = DecryptAdminPasswd(GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions", "lhag"))
   If strAdminPasswd = "" Then frmAdminPasswd.Show vbModal, Me
   'je�li brakuje plik�w z ulubionymi i z autoukrywaniem, to s� tworzone
   If FileExists("auto.dat") = False Then
      Open "auto.dat" For Output As #1
      Close
   End If
   If FileExists("ulubione.fav") = False Then
      Open "ulubione.fav" For Output As #1
      Close
   End If
   'Sprawdzenie parametru. Program rozpoznaje, czy ma by� uruchomiony normalnie, czy w trybie automatycznego ukrywania.
   If Command$ = "/auto" Then
      'Otwarcie pliku
      Open "auto.dat" For Input As #1
      Do Until EOF(1)
         Input #1, strFolder
         If FileExists(strFolder) Then
            Name strFolder As strFolder & strKodKlasy
            SetAttr strFolder & strKodKlasy, (GetAttr(strFolder & strKodKlasy) And 239) Or vbHidden
         End If
      Loop
      Close
      End
   Else
      'wyszarzenie odpowiednich kontrolek itp.
      lstFoldery.Clear
      mnuUstawienia.Enabled = False
      'Wprowadzanie element�w do listy
      ChDir App.Path
      Open "ulubione.fav" For Input As #1
      Do Until EOF(1)
         Input #1, strFolder
         lstFoldery.AddItem (strFolder)
      Loop
      Close
      txtHaslo.Text = ""
   End If
   'Modyfikacja log�w instalatora ST6UNST.LOG �eby przy deinstalacji deinstalator usuwa� wpis w rejestrze i pliki utworzone przez program
   If FileExists(App.Path & "\ST6UNST.LOG") Then
      ChDir App.Path
      Open "ST6UNST.LOG" For Input As #1
      Input #1, strTekst
      Close
      If strTekst <> "%% Modified by GN Ukrywacz %%" Then 'na pocz�tku tej lini w pliku jest spacja, ale w zmiennnej strTekst jej nie ma, nie wiem czemu, wi�c przy por�wnaniu te� nie ma tej spacji
         Open "ST6UNST.LOG" For Binary As #1
         ReDim bytTablica(LOF(1) - 1) As Byte
         Get #1, , bytTablica
         Close
         Kill "ST6UNST.LOG"
         Open "ST6UNST.LOG" For Binary As #1
         Put #1, , " %% Modified by GN Ukrywacz %%"
         Put #1, , vbCrLf
         Put #1, , bytTablica
         Close
         Open "ST6UNST.LOG" For Append As #1
         Print #1, "ACTION: RegValue: " & Chr(34) & "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Run" & Chr(34) & ", " & Chr(34) & "GN Ukrywacz" & Chr(34)
         Print #1,
         Print #1, "ACTION: PrivateFile: " & Chr(34) & App.Path & "\auto.dat" & Chr(34)
         Print #1,
         Print #1, "ACTION: PrivateFile: " & Chr(34) & App.Path & "\ulubione.fav" & Chr(34)
         Print #1,
         Close
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub

Private Sub lblFolder_Change()
   If Right(strFolder, Len(strKodKlasy)) = strKodKlasy And Right(lblFolder.Caption, Len(strKodKlasy)) = strKodKlasy Then lblFolder.Caption = Left(strFolder, Len(strFolder) - Len(strKodKlasy))
   If Left(lblFolder.Caption, 1) <> "'" Then lblFolder.Caption = "'" & lblFolder.Caption & "'"
End Sub

Private Sub lstFoldery_Click()
   'Program przemiata plik danych. Je�li znajdzie nazw� foldera, tak� jaka
   'zosta�a klikni�ta w li�cie, wprowadza potrzebne zmiany
   'do element�w steruj�cych i zmiennych.
   txtHaslo.Text = ""
   intIloscFolderow = lstFoldery.ListCount
   If intIloscFolderow > 0 Then
      For intZmienna = 0 To intIloscFolderow - 1
         If lstFoldery.Selected(intZmienna) = True Then
            strFolder = lstFoldery.List(intZmienna)
            lblFolder.Caption = strFolder
            Properties strFolder, cmdZmien, lblStan, txtHaslo, strHaslo, blnDostepny, blnIstnieje, blnChroniony, blnAuto, blnBezHasla, blnAdmin
         End If
      Next intZmienna
   End If
End Sub

Private Sub mnuAdminLogon_Click()
   If blnAdmin = False Then
      frmAdminLogon.Show vbModal, Me
   Else
      blnAdmin = False
      mnuAdminLogon.Caption = "&Zaloguj"
      mnuAdminPasswd.Enabled = False
   End If
   Properties strFolder, cmdZmien, lblStan, txtHaslo, strHaslo, blnDostepny, blnIstnieje, blnChroniony, blnAuto, blnBezHasla, blnAdmin
End Sub

Private Sub mnuAdminPasswd_Click()
   frmAdminPasswd.Show vbModal, Me
End Sub

Private Sub mnuPlikKoniec_Click()
   intZmienna = MsgBox("Czy na pewno chcesz zako�czy� program?", vbQuestion + vbYesNo, "Zako�czenie")
   If intZmienna = 6 Then Unload Me
End Sub

Private Sub mnuPomocAuto_Click()
   MsgBox "   Opcja ta znajduje si� w Ustawieniach. Umo�liwia ukrywanie foldera przy starcie komputera. Pomocna jest ona wtedy, gdy zapominamy zabezpieczy� folder. Je�li nie ukryjemy foldera, zostanie on ukryty przy starcie komputera. Je�li opcja automatycznego ukrywania jest zaznaczona a folder jest ukryty, to przy starcie nie zostan� do niego wprowadzone �adne zmiany.", vbInformation, "Automatyczne ukrywanie"
End Sub

Private Sub mnuPomocFolder_Click()
   MsgBox "   Gdy chcemy ukry�/udost�pni� folder lub zmieni� jego ustawienia, musimy najpierw go wskaza�. Mo�na go wybra� z listy Ulubionych lub klikn�� na Przegl�daj i wybra� folder z listy z prawej strony. Lista z lewej strony s�u�y do wybrania foldera, w kt�rym jest szukany przez nas folder.", vbInformation + vbOKOnly, "GN Ukrywacz - wyb�r foldera"
End Sub

Private Sub mnuPomocHaslo_Click()
   MsgBox "   Aby zmieni� has�o, trzeba wybra� folder, wpisa� stare has�o aby uaktywni�o si� menu Ustawienia i w okienku Ustawienia wpisa� nowe has�o. Zatwierdzamy je klikaj�c na OK.", vbInformation + vbOKOnly, "GN Ukrywacz - zmiana has�a"
End Sub

Private Sub mnuPomocOchrona_Click()
   MsgBox "   Ochrona foldera polega na dodaniu do jego nazwy specjalnego rozszerzenia. Ponadto w folderze zostaje umieszczony plik zawieraj�cy has�o i ustawienie okre�laj�ce, czy b�dzie mo�na ukry� folder nie podaj�c has�a. Plik z has�em jest szyfrowany. Dzi�ki umieszczeniu pliku z has�em w chronionym folderze, mo�liwe jest jego dowolne przenoszenie, a ponadto reinstalacja programu nie pozbawia nas dost�pu do ukrytych folder�w.", vbInformation + vbOKOnly, "GN Ukrywacz - ochrona folder�w"
End Sub

Private Sub mnuPomocOProgramie_Click()
   frmAbout.Show vbModal, Me
End Sub

Private Sub mnuPomocUkrBezH_Click()
   MsgBox "   Zaznaczenie tej opcji w Ustawieniach umo�liwia ukrycie zawarto��i foldera bez konieczno�ci podawania has�a. Jej zaznaczenie wymaga wpisania has�a.", vbInformation, "Ukrywanie bez has�a"
End Sub

Private Sub mnuUstawienia_Click()
   frmUstawienia.Show vbModal
End Sub

Private Sub txtHaslo_Change()
   PasswordCheck strHaslo, txtHaslo, cmdZmien, blnChroniony, blnBezHasla, mnuUstawienia, blnDostepny, blnAdmin
End Sub
