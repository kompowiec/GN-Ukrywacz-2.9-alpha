VERSION 5.00
Begin VB.Form frmUlubiony 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GN Ukrywacz - dodawanie foldera do Ulubionych"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6540
   Icon            =   "GN UkrywaczUlubiony.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAnuluj 
      Cancel          =   -1  'True
      Caption         =   "Anuluj"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrzegladaj 
      Caption         =   "Przegl¹daj..."
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   160
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "frmUlubiony"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strTemp As String
Dim intLicznik As Integer
Const strKodKlasy As String = ".{21EC2020-3AEA-1069-A2DD-08002B30309D}"

Private Sub cmdAnuluj_Click()
   frmForm1.strUlubiony = ""
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If FileExists(Text1.Text) = True Or FileExists(Text1.Text & ".{21EC2020-3AEA-1069-A2DD-08002B30309D}") = True Then
      'sprawdzenie, czy wybrany folder nie jest folderem windozy
      If UCase(Text1.Text) = UCase(GetWindowsDir) Then MsgBox "Wybrany folder jest folderem Windowsa. Ukrycie go spowodowa³oby zawiesznie siê komputera i problemy z uruchomieniem Windowsa.", vbExclamation + vbOKOnly, "B³¹d wybierania foldera": Exit Sub
      'sprawdzenie, czy ktoœ nie wybra³ program files. Jest to sprawdzane trochê lamersko, ale w 99,9 % przypadków wystarczy
      If UCase(Text1.Text) = UCase("c:\Program Files") Then MsgBox "Ukrycie tego foldera spowoduje niemo¿liwoœæ uruchomienia wielu programów.", vbExclamation + vbOKOnly, "GN Ukrywacz - bl¹d wybierania foldera": Exit Sub
      'sprawdznie, czy ktoœ przypadkiem nie chce ukryæ foldera, w którym jest GN Ukrywacz
      If InStr(UCase(App.Path) & "\", UCase(Text1.Text) & "\") = 1 Then MsgBox "Nie mo¿na ukryæ foldera, w którym znajduje siê GN Ukrywacz.", vbExclamation + vbOKOnly, "GN Ukrywacz - b³¹d wybierania foldera": Exit Sub
      For intLicznik = 0 To frmForm1.lstFoldery.ListCount - 1
         If UCase(frmForm1.lstFoldery.List(intLicznik)) = UCase(Text1.Text) Then
            MsgBox "Ten folder ju¿ jest na liœcie ulubionych.", vbExclamation + vbOKOnly, "GN Ukrywacz - powtórzony folder"
            Exit Sub
         End If
      Next intLicznik
      frmForm1.strUlubiony = Text1.Text
      Unload Me
   Else
      MsgBox "Wpisany folder jest nieprawid³owy.", vbExclamation + vbOKOnly, "GN Ukrywacz - z³y folder"
   End If
End Sub

Private Sub cmdPrzegladaj_Click()
   strTemp = BrowseFolder(Me, "Wybierz folder:")
   If strTemp <> "" Then Text1.Text = strTemp
End Sub

Private Sub Form_Load()
   Text1.Text = frmForm1.strFolder
   If Right(Text1.Text, Len(strKodKlasy)) = strKodKlasy Then Text1.Text = Left(Text1.Text, Len(Text1.Text) - Len(strKodKlasy))
   frmUlubiony.Icon = frmForm1.Icon
End Sub
