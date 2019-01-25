VERSION 5.00
Begin VB.Form frmHaslo 
   Caption         =   "Has≥o"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6990
   Icon            =   "GN UkrywaczHaslo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6990
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtHaslo2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   720
      Width           =   4575
   End
   Begin VB.CommandButton cmdAnuluj 
      Cancel          =   -1  'True
      Caption         =   "&Anuluj"
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1440
      Width           =   1455
   End
   Begin VB.TextBox txtHaslo 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   210
      Width           =   4575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Potwierdü has≥o:"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Podaj has≥o dla foldera:"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
End
Attribute VB_Name = "frmHaslo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnuluj_Click()
   frmForm1.strHaslo = ""
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If UCase(txtHaslo.Text) <> UCase(txtHaslo2.Text) Then
      MsgBox "B≥πd w potwierdzniu has≥a! Sprawdü, czy w obu polach wpisane jest jednakowe has≥o.", vbExclamation + vbOKOnly, "GN Ukrywacz - z≥e potwierdzenie"
      txtHaslo.SetFocus
      Exit Sub
   End If
   If txtHaslo.Text <> "" Then
      frmForm1.strHaslo = txtHaslo.Text
      Unload Me
   Else
      txtHaslo.SetFocus
   End If
End Sub

Private Sub Form_Load()
   Me.Icon = frmForm1.Icon
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Cancel <> 0 Then frmForm1.strHaslo = ""
End Sub
