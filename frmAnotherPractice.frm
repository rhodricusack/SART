VERSION 5.00
Begin VB.Form frmAnotherPractice 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   3975
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "Run test"
      Height          =   735
      Left            =   2160
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1800
      Width           =   1575
   End
   Begin VB.CommandButton cmdRepeat 
      Caption         =   "Repeat practice"
      Height          =   735
      Left            =   2160
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   960
      Width           =   1575
   End
End
Attribute VB_Name = "frmAnotherPractice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intResponse As Integer


Private Sub cmdRepeat_Click()
intResponse = 1
End Sub

Private Sub cmdRun_Click()
intResponse = 2
End Sub
