VERSION 5.00
Begin VB.Form frmAdminLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GN Ukrywacz - logowanie administratora"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "GN UkrywaczAdminLogon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAnuluj 
      Cancel          =   -1  'True
      Caption         =   "&Anuluj"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txtHaslo 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Has³o:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmAdminLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnuluj_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If frmForm1.strAdminPasswd = txtHaslo Then
      frmForm1.blnAdmin = True
      frmForm1.mnuAdminLogon.Caption = "&Wyloguj"
      frmForm1.mnuAdminPasswd.Enabled = True
      Unload Me
   Else
      MsgBox "Wpisane has³o jest niepoprawne.", vbExclamation + vbOKOnly, "GN Ukrywacz - b³¹d logowania"
      txtHaslo.SetFocus
      txtHaslo.SelStart = 0
      txtHaslo.SelLength = Len(txtHaslo.Text)
   End If
End Sub

Private Sub Form_Load()
   Me.Icon = frmForm1.Icon
End Sub
