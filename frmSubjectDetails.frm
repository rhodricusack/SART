VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSubjectDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Participant's details"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   615
      Left            =   3360
      TabIndex        =   11
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtNotes 
      Height          =   615
      Left            =   1320
      TabIndex        =   9
      Top             =   1680
      Width           =   2775
   End
   Begin VB.OptionButton optFemale 
      Caption         =   "Option2"
      Height          =   255
      Left            =   1680
      TabIndex        =   5
      Top             =   1200
      Width           =   255
   End
   Begin VB.OptionButton optMale 
      Caption         =   "Option1"
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox txtDOB 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Frame frmSex 
      Caption         =   "Sex"
      Height          =   615
      Left            =   720
      TabIndex        =   6
      Top             =   960
      Width           =   1935
      Begin VB.Label Label4 
         Caption         =   "Female"
         Height          =   255
         Left            =   1200
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Male"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Any notes"
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Date of birth"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmSubjectDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strSex As String

Private Sub cmdOK_Click()
vntFilename = GetFileName(frmSubjectDetails.txtName)
booPractice = True
frmSubjectDetails.Visible = False
End Sub


Function GetFileName(txtSubjName As String)
' Set CancelError is True
CommonDialog1.CancelError = True
On Error GoTo errhandler
' Set flags
CommonDialog1.flags = cdlOFNHideReadOnly
' Set filters
CommonDialog1.Filter = "All Files (*.*)|*.*|Text Files" & _
"(*.txt)|*.txt"
CommonDialog1.DialogTitle = "Save results file as"
CommonDialog1.FileName = txtSubjName & "_sart_results.txt"
' Specify default filter
CommonDialog1.FilterIndex = 2
' Display the Open dialog box
CommonDialog1.ShowSave
GetFileName = CommonDialog1.FileName

Exit Function

errhandler:
GetFileName = Null

End Function

Private Sub optFemale_Click()
strSex = "Female"
End Sub

Private Sub optMale_Click()
strSex = "Male"
End Sub
