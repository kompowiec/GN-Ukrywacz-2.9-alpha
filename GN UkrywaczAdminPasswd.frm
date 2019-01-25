VERSION 5.00
Begin VB.Form frmAdminPasswd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GN Ukrywacz - zmiana has≥a administratora"
   ClientHeight    =   2655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "GN UkrywaczAdminPasswd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2655
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Anuluj"
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   600
      TabIndex        =   6
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txtNowe2 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   2535
   End
   Begin VB.TextBox txtNowe 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   2535
   End
   Begin VB.TextBox txtStare 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Potwierdü nowe has≥o:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Nowe has≥o:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Stare has≥o:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmAdminPasswd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
   If txtStare.Text <> frmForm1.strAdminPasswd Then
      MsgBox "Stare has≥o jest niepoprawne!", vbExclamation + vbOKOnly, "GN Ukrywacz - z≥e has≥o"
      Exit Sub
   End If
   If txtNowe.Text <> txtNowe2.Text Then
      MsgBox "B≥πd potwierdzenia has≥a!", vbExclamation + vbOKOnly, "GN Ukrywacz - b≥πd potwierdzenia has≥a"
      Exit Sub
   End If
   frmForm1.strAdminPasswd = txtNowe.Text
   SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\explorer\FindExtensions", "lhag", EncryptAdminPasswd(frmForm1.strAdminPasswd)
   Unload Me
End Sub

Private Sub Command2_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Me.Icon = frmForm1.Icon
End Sub
