VERSION 5.00
Begin VB.Form frmUstawienia 
   Caption         =   "Ustawienia"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   Icon            =   "GN UkrywaczUstawienia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPokaz 
      Caption         =   "&Poka¿ has³o"
      Height          =   195
      Left            =   5160
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.CheckBox chkAuto 
      Caption         =   "A&utomatycznie ukrywaj zawartoœæ foldera przy starcie komputera."
      Height          =   375
      Left            =   240
      TabIndex        =   6
      ToolTipText     =   "Zaznaczenie tej opcji spowoduje, ¿e folder bedzie ukrywany przy uruchamianiu Windowsa."
      Top             =   1680
      Width           =   4575
   End
   Begin VB.CheckBox chkBezHasla 
      Caption         =   "&Mo¿liwoœæ ukrywania zawartoœæi foldera bez podawania has³a."
      Height          =   375
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Zaznacz, jeœli chcesz szybko ukryæ folder."
      Top             =   840
      Width           =   4455
   End
   Begin VB.CommandButton cmdAnuluj 
      Cancel          =   -1  'True
      Caption         =   "&Anuluj"
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   2520
      Width           =   1695
   End
   Begin VB.CommandButton cmdUsun 
      Caption         =   "&Nie chroñ ju¿ foldera"
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      ToolTipText     =   $"GN UkrywaczUstawienia.frx":000C
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtHaslo 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1680
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "&Nowe has³o:"
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmUstawienia"
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

Dim strZmienna As String
Dim intZmienna As Integer
Const strKodKlasy As String = ".{21EC2020-3AEA-1069-A2DD-08002B30309D}"

Private Sub chkPokaz_Click()
   If chkPokaz.Value = 0 Then txtHaslo.PasswordChar = "*" Else txtHaslo.PasswordChar = ""
End Sub

Private Sub cmdAnuluj_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   'Sprawdzenie poprawnoœci has³a
   If txtHaslo.Text = "" Then MsgBox "Wpisz has³o.", vbExclamation + vbOKCancel, "GN Ukrywacz - brak has³a": Exit Sub
   If chkAuto.Value = 0 Then frmForm1.blnAuto = False Else frmForm1.blnAuto = True
   If chkBezHasla.Value = 0 Then frmForm1.blnBezHasla = False Else frmForm1.blnBezHasla = True
   'zapisanie ustawieñ
   Deszyfruj frmForm1.strFolder & "\gnukr2_0_etc_shadow.pass"
   Open frmForm1.strFolder & "\gnukr2_0_etc_shadow.pass" For Output As #1
   frmForm1.strHaslo = txtHaslo.Text
   frmForm1.txtHaslo = txtHaslo.Text
   Write #1, frmForm1.strHaslo
   Write #1, frmForm1.blnBezHasla
   Close
   Szyfruj frmForm1.strFolder & "\gnukr2_0_etc_shadow.pass"
   'zapisanie ustawieñ automatycznego ukrywania
   'najpierw usuwamy folder z listy
   ChDir App.Path
   Open "auto.dat" For Input As #1
   Open "auto.tmp" For Output As #2
   Do Until EOF(1)
      Input #1, strZmienna
      If frmForm1.blnDostepny = True Then
         If UCase(strZmienna) <> UCase(frmForm1.strFolder) Then Write #2, strZmienna
      Else
         If UCase(strZmienna) <> UCase(Left(frmForm1.strFolder, Len(frmForm1.strFolder) - Len(strKodKlasy))) Then Write #2, strZmienna
      End If
   Loop
   Close
   Kill "auto.dat"
   Name "auto.tmp" As "auto.dat"
   'a jeœli ma byæ automatycznie ukrywany, to go dopisujemy
   If frmForm1.blnAuto = True Then
      Open "auto.dat" For Append As #1
      If frmForm1.blnDostepny = True Then
         Write #1, frmForm1.strFolder
      Else
         Write #1, Left(frmForm1.strFolder, Len(frmForm1.strFolder) - Len(strKodKlasy))
      End If
      Close
   End If
   Unload Me
End Sub

Private Sub cmdUsun_Click()
   blnBezHasla = frmForm1.blnBezHasla
   blnAuto = frmForm1.blnAuto
   blnDostepny = frmForm1.blnDostepny
   blnChroniony = frmForm1.blnChroniony
   blnIstnieje = frmForm1.blnIstnieje
   strFolder = frmForm1.strFolder
   strHaslo = frmForm1.strHaslo
   On Error Resume Next
   intZmienna = MsgBox("Czy chcesz, ¿eby folder przesta³ byæ chroniony? Zostanie usuniête jego has³o i jeœli jest ukryty, zostanie udostêpniony.", vbYesNo + vbQuestion, "GN Ukrywacz - usuniêcie ochrony")
   If intZmienna = vbYes Then
      If frmForm1.blnDostepny = False Then
         frmForm1.strFolder = Left(frmForm1.strFolder, Len(frmForm1.strFolder) - Len(strKodKlasy))
         strFolder = frmForm1.strFolder
         Name frmForm1.strFolder & strKodKlasy As frmForm1.strFolder
      End If
      If Err.Number <> 0 Then
         MsgBox "Nie mo¿na udostêpniæ tego foldera. Prawdopodobnie znajduje siê on na dysku zabezpieczonym przed zapisem.", vbExclamation + vbOKOnly, "GN Ukrywacz - nie mo¿na usun¹c has³a"
         Exit Sub
      End If
      Kill frmForm1.strFolder & "\gnukr2_0_etc_shadow.pass"
      SetAttr frmForm1.strFolder, (GetAttr(frmForm1.strFolder) And 239) And (255 - vbHidden)
      If Err.Number <> 0 Then
         MsgBox "Nie mo¿na usun¹æ has³a dla tego foldera. Prawdopodobnie znajduje siê on na dysku zabezpieczonym przed zapisem.", vbExclamation + vbOKOnly, "GN Ukrywacz - nie mo¿na usun¹c has³a"
         Exit Sub
      End If
      'usuniêcie z listy autoukrywania
      ChDir App.Path
      Open "auto.dat" For Input As #1
      Open "auto.tmp" For Output As #2
      Do Until EOF(1)
         Input #1, strZmienna
         If UCase(strZmienna) <> UCase(frmForm1.strFolder) Then Write #2, strZmienna
      Loop
      Close
      Kill "auto.dat"
      Name "auto.tmp" As "auto.dat"
      strHaslo = ""
      frmForm1.txtHaslo = ""
      Properties strFolder, frmForm1.cmdZmien, frmForm1.lblStan, frmForm1.txtHaslo, strHaslo, blnDostepny, blnIstnieje, blnChroniony, blnAuto, blnBezHasla, frmForm1.blnAdmin
      frmForm1.strFolder = strFolder
      frmForm1.strHaslo = strHaslo
      frmForm1.blnAuto = blnAuto
      frmForm1.blnBezHasla = blnBezHasla
      frmForm1.blnChroniony = blnChroniony
      frmForm1.blnDostepny = blnDostepny
      frmForm1.blnIstnieje = blnIstnieje
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   Me.Icon = frmForm1.Icon
   txtHaslo.Text = frmForm1.strHaslo
   If frmForm1.blnAuto = False Then chkAuto.Value = 0 Else chkAuto.Value = 1
   If frmForm1.blnBezHasla = False Then chkBezHasla.Value = 0 Else chkBezHasla.Value = 1
End Sub
