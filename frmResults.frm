VERSION 5.00
Begin VB.Form frmResults 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Results summary"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   735
      Left            =   4200
      TabIndex        =   17
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtSDRT 
      Height          =   285
      Left            =   2880
      TabIndex        =   16
      Text            =   "Text2"
      Top             =   2640
      Width           =   855
   End
   Begin VB.TextBox txtMeanRT 
      Height          =   285
      Left            =   2880
      TabIndex        =   15
      Top             =   2280
      Width           =   855
   End
   Begin VB.TextBox txtAnticip 
      Height          =   285
      Left            =   960
      TabIndex        =   14
      Top             =   1920
      Width           =   375
   End
   Begin VB.TextBox txtOmission 
      Height          =   285
      Left            =   3720
      TabIndex        =   13
      Top             =   1560
      Width           =   495
   End
   Begin VB.TextBox txtCommission 
      Height          =   285
      Left            =   3720
      TabIndex        =   12
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtTest 
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtDOB 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label5 
      Caption         =   "Test"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Standard deviation of reaction time"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label9 
      Caption         =   "Mean correct reaction time (ms)"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "were anticipatory responses"
      Height          =   855
      Left            =   1440
      TabIndex        =   5
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "- of which"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "Date of birth"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "Name"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Errors of omission (missed presses for non-3s)"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "Errors of commission (number of 3s pressed for)"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   3495
   End
End
Attribute VB_Name = "frmResults"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Me.Visible = False

End Sub
