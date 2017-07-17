VERSION 5.00
Begin VB.Form frmEndOfTest 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "Done"
      Height          =   615
      Left            =   3360
      TabIndex        =   1
      Top             =   2400
      Width           =   1095
   End
   Begin VB.CommandButton cmdResultsNow 
      Caption         =   "Look at results now"
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "End of test - well done!"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   1695
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   3975
   End
End
Attribute VB_Name = "frmEndOfTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intResponse As Integer

Private Sub cmdOK_Click()
intResponse = 2
End Sub

Private Sub cmdResultsNow_Click()
intResponse = 1
End Sub
