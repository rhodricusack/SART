VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00404000&
   BorderStyle     =   0  'None
   Caption         =   "Sustained Attention to Response Task"
   ClientHeight    =   4935
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   4935
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   4200
      Width           =   615
   End
   Begin VB.CommandButton cmdFSS_YRL 
      Caption         =   "Fixed sequence SART with response locking"
      Height          =   975
      Left            =   1800
      TabIndex        =   3
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CommandButton cmdTSS_NRL 
      Caption         =   "Random sequence SART without response locking"
      Height          =   975
      Left            =   2880
      TabIndex        =   2
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   4200
      Width           =   735
   End
   Begin VB.CommandButton cmdRSS_YRL 
      Caption         =   "Random sequence SART with response locking"
      Height          =   975
      Left            =   720
      TabIndex        =   0
      Top             =   1680
      Width           =   1695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAbout_Click()
frmAbout.Visible = True

End Sub

Private Sub cmdExit_Click()
End
End Sub

Sub RunAll(strInstr As String)
frmSubjectDetails.Visible = True
Do While frmSubjectDetails.Visible = True
    DoEvents
Loop

If (Not (IsNull(vntFilename) Or IsEmpty(vntFilename))) Then
    Dim ff As Integer
    ff = FreeFile
    Open vntFilename For Append As #ff
    Print #ff, "** SART experiment conducted on:" & vbTab & Now()
    Print #ff, "Participant name: " & vbTab & frmSubjectDetails.txtName
    Print #ff, "Participant DOB: " & vbTab & frmSubjectDetails.txtDOB
    Print #ff, "Participant Sex: " & vbTab & frmSubjectDetails.strSex
    Print #ff, "Notes:" & vbTab & frmSubjectDetails.txtNotes
    Print #ff, "Response locking:" & vbTab & booResponseLocking
    Print #ff, "Random order:" & vbTab & booRandom
    Close #ff
    frmBlank.Visible = True
    frmInstructions.Label1 = strInstr
    frmInstructions.Visible = True
End If

End Sub

Private Sub cmdFSS_YRL_Click()
Dim strInstr As String
strInstr = "In this test you will see numbers between 1 and 9 appear on the screen in a repeating, regular sequence (1 2 3 4 5 6 7 8 9 1 2 3 and so on). "
strInstr = strInstr & "Part way through each trial the cross will brighten briefly. Please press the space bar in time with this brightening after each number except 3. "
strInstr = strInstr & "If you see a 3 don't press the button - just wait for the next number to appear. Try to keep your responding in time with the brightening of the cross and try to avoid pressing for the 3."
booResponseLocking = True
booRandom = False

RunAll strInstr
End Sub

Private Sub cmdRSS_YRL_Click()
Dim strInstr As String
strInstr = "In this test you will see numbers between 1 and 9 appear on the screen in a random sequence, followed by a cross. "
strInstr = strInstr & "Part way through each trial the cross will brighten briefly. Please press the space bar in time with this brightening after each number except 3. "
strInstr = strInstr & "If you see a 3 don't press the button - just wait for the next number to appear. Try to keep your responding in time with the brightening of the cross and try to avoid pressing for the 3."
booResponseLocking = True
booRandom = True

RunAll strInstr
End Sub

Private Sub cmdTSS_NRL_Click()
Dim strInstr As String
strInstr = "In this test you will see numbers between 1 and 9 appear on the screen in a random sequence, followed by a cross. "
strInstr = strinster & "Please press the space bar as quickly as possible for each number except 3."
strInstr = strInstr & "If you see a 3 don't press the button - just wait for the next number to appear. Try and respond as quickly as possible to each number while maintaining your accuracy (not pressing for 3s)."
booResponseLocking = False
booRandom = True

RunAll strInstr
End Sub

Private Sub Form_Load()
Randomize
StartTimer
End Sub
