VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "O programie"
   ClientHeight    =   7110
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "GN UkrywaczOP.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4907.449
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   2130
      TabIndex        =   0
      Top             =   6240
      Width           =   1380
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   360
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "Najnowsze wersje programu, informacje i nie tylko:"
      Height          =   255
      Left            =   720
      TabIndex        =   9
      Top             =   3840
      Width           =   4215
   End
   Begin VB.Label Label4 
      Caption         =   "Pytania, komentarze, sugestie, pomoc:"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   2760
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "www.grzegorz.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF000D&
      Height          =   240
      Left            =   720
      MouseIcon       =   "GN UkrywaczOP.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4200
      Width           =   1395
   End
   Begin VB.Label Label2 
      Caption         =   "grzegorz@grzegorz.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   720
      MouseIcon       =   "GN UkrywaczOP.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3120
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Copyright (C) Grzegorz Niemirowski 1999-2002"
      Height          =   240
      Left            =   720
      TabIndex        =   5
      Top             =   2025
      Width           =   3855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   98.6
      X2              =   5299.069
      Y1              =   3380.687
      Y2              =   3380.687
   End
   Begin VB.Label lblDescription 
      Caption         =   "Program do ukrywania zawartoœci folderów"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   720
      TabIndex        =   1
      Top             =   1680
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "GN Ukrywacz"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   600
      Left            =   1560
      TabIndex        =   3
      Top             =   240
      Width           =   3765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5299.069
      Y1              =   3395.872
      Y2              =   3396.562
   End
   Begin VB.Label lblVersion 
      Caption         =   "Wersja 2.9"
      Height          =   225
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Alignment       =   2  'Center
      Caption         =   "GN Ukrywacz jest programem typu FREEWARE i nie mo¿e byæ u¿ywany w celach komercyjnych bez kontaktu z autorem."
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   240
      TabIndex        =   2
      Top             =   5280
      Width           =   5145
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "SHELL32.DLL" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
   Image1.Picture = frmForm1.Icon
   Me.Icon = frmForm1.Icon
End Sub

Private Sub Label2_Click()
ShellExecute Me.hwnd, "open", "mailto:grzegorz@grzegorz.net", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

Private Sub Label3_Click()
ShellExecute Me.hwnd, "open", "http://www.grzegorz.net", vbNullString, vbNullString, SW_SHOWNORMAL
End Sub

