VERSION 5.00
Begin VB.Form frmInstructions 
   BorderStyle     =   0  'None
   Caption         =   "Instructions"
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5715
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Press enter to start."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   2280
      TabIndex        =   2
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Instructions"
      ForeColor       =   &H00000040&
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   360
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   600
      TabIndex        =   0
      Top             =   1320
      Width           =   6375
   End
End
Attribute VB_Name = "frmInstructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
If (KeyAscii = 27) Then
    cleanup
    End
End If
If (KeyAscii = 13) Then
    Me.Visible = False
    frmExpt.lblReadyMsg = "Practice - ready?"
    frmExpt.SetUp
    frmExpt.Visible = True
End If
End Sub


