VERSION 5.00
Begin VB.Form frmExpt 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   ScaleHeight     =   633
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   716
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
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
      Left            =   4440
      TabIndex        =   1
      Top             =   8280
      Width           =   3015
   End
   Begin VB.Label lblReadyMsg 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   735
      Left            =   3720
      TabIndex        =   0
      Top             =   3840
      Width           =   4455
   End
End
Attribute VB_Name = "frmExpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const CX = 400
Const CY = 300
Const FIXSIZE = 30
Const ISI = 1150

Dim lngBackColor As Long
Dim booExptStarted As Boolean

Dim dblNextTrial As Double

Dim ddsd2 As DDSURFACEDESC2

Dim booHadResponseThisTrial As Boolean
Dim booHadResponseNextTrial As Boolean

Dim dblRTthistrial As Double
Dim dblRTnexttrial As Double

Dim intWasTarget() As Integer
Dim dblRT() As Double
Dim intWasResp() As Boolean
Dim intTrialNum As Integer

Private Sub Form_KeyPress(KeyAscii As Integer)
Debug.Print "Key press " & KeyAscii

If (KeyAscii = 32 And booExptStarted) Then
    dblKeyTime = GetTimer
    If (intWasResp(intTrialNum) And dblRT(intTrialNum) > 0) Then
        If (Not intWasResp(intTrialNum + 1)) Then
            dblRT(intTrialNum + 1) = dblKeyTime - dblNextTrial
            If (dblRT(intTrialNum + 1) > 0) Then dblRT(intTrialNum + 1) = dblRT(intTrialNum + 1) - ISI
            intWasResp(intTrialNum + 1) = 1
        End If
    Else
        dblRT(intTrialNum) = dblKeyTime - dblNextTrial
        If (dblRT(intTrialNum) < 0) Then dblRT(intTrialNum) = dblRT(intTrialNum) + ISI
        intWasResp(intTrialNum) = 1
    End If
End If

If (KeyAscii = 27) Then
    booExptStarted = False
    cleanup
    End
End If

If (KeyAscii = 13 And Not booExptStarted) Then
    booExptStarted = True
    RunExpt
    
    SyncAndBlank
    WaitForTime (GetTimer + 1500#)
    
    If (booPractice) Then
        Me.Visible = False
        frmAnotherPractice.Visible = True
        frmAnotherPractice.Refresh
        DoEvents
        frmAnotherPractice.intResponse = 0
        ShowCursor 1
        Do
            DoEvents
        Loop Until frmAnotherPractice.intResponse <> 0
        frmAnotherPractice.Visible = False
        ShowCursor 0
        
        If (frmAnotherPractice.intResponse = 2) Then
             booPractice = False
            Me.lblReadyMsg = "Test... ready?"
        End If
        Me.Visible = True
        Me.Refresh
   End If
    booExptStarted = False
End If
End Sub

Sub SetUp()
ShowCursor 0

Set dd7 = dx7.DirectDrawCreate("")
dd7.SetCooperativeLevel Me.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
dd7.SetDisplayMode 800, 600, 32, 0, DDSDM_DEFAULT

ddsd2.lFlags = DDSD_CAPS
ddsd2.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE

Set ds7Screen = dd7.CreateSurface(ddsd2)
Debug.Print Err.Number
    
If (booPractice) Then
    Me.lblReadyMsg = "Practice... ready?"
Else
    Me.lblReadyMsg = "Test... ready?"
End If

End Sub

Sub RunExpt()

' Set up graphics
Dim area As RECT
area.Left = 0
area.Right = 800
area.Top = 0
area.Bottom = 599
Dim myfont As New StdFont
myfont.Name = "Arial"
myfont.Size = 48

lngBackColor = RGB(192, 192, 192)
ds7Screen.SetFont myfont
ds7Screen.BltColorFill area, lngBackColor



' Determine number of blocks
If (booPractice) Then
    intNumBlocks = 2
Else
    If (booRandom) Then
        intNumBlocks = 25
    Else
        intNumBlocks = 42
    End If
End If

intNumBlocks = 1

Debug.Print "Number of blocks: " & intNumBlocks

' Trial records
Dim intTotalTrials As Integer
intTotalTrials = 9 * intNumBlocks
ReDim intWasResp(0 To intTotalTrials + 1) '+1 for final anticipatory
ReDim dblRT(0 To intTotalTrials + 1)
ReDim intWasTarget(1 To intTotalTrials)
intTrialNum = 0

' Now timing

StartTimer


dblNextTrial = 500#

Dim intDigits(1 To 9) As Integer
Dim sngRand(1 To 9) As Single
Dim i, j

booHadResponseNextTrial = False
booHadResponseThisTrial = False

For i = 1 To intNumBlocks
    Debug.Print "Block " & i
    For j = 1 To 9
        intDigits(j) = j
    Next
    Debug.Print "Sorting"
    If (booRandom) Then
        For j = 1 To 9
            sngRand(j) = Rnd()
        Next
        Dim booSwap As Boolean
        booSwap = True
        Do While booSwap
            booSwap = False
            For j = 1 To 8
                If (sngRand(j) > sngRand(j + 1)) Then
                    booSwap = True
                    tmp = sngRand(j)
                    sngRand(j) = sngRand(j + 1)
                    sngRand(j + 1) = tmp
                    tmp = intDigits(j)
                    intDigits(j) = intDigits(j + 1)
                    intDigits(j + 1) = tmp
                End If
            Next
        Loop
    End If
    For j = 1 To 9
        Debug.Print "Block " & i & " Trial " & j
        PresentTrial booResponseLocking, intDigits(j), dblNextTrial
        If (intDigits(j) = 3) Then intWasTarget(intTrialNum) = 1 Else intWasTarget(intTrialNum) = 0
        dblNextTrial = dblNextTrial + ISI
    Next
Next

Do While GetTimer < dblNextTrial
    DoEvents
Loop


    
    
If (Not booPractice) Then
        
    Dim ff As Integer
    Dim intErrOmiss As Integer
    Dim intErrComiss As Integer
    Dim intAnticip As Integer
    Dim dblRTtot As Double
    Dim dblRTsqtot As Double
    Dim intNuminMean As Integer
    
    ff = FreeFile
    Open vntFilename For Append As #ff
    For i = 1 To intTotalTrials
        If (intWasTarget(i) And intWasResp(i)) Then intErrComiss = intErrComiss + 1
        If (dblRT(i) < 0) Then
            intAnticip = intAnticip + 1
        Else
            If (intWasTarget(i) = 0 And (Not intWasResp(i))) Then
                Debug.Print intWasTarget(i), intWasResp(i)
                intErrOmiss = intErrOmiss + 1
            Else
                If (intWasResp(i)) Then
                    dblRTtot = dblRTtot + dblRT(i)
                    dblRTsqtot = dblRTsqtot + dblRT(i) * dblRT(i)
                    intNuminMean = intNuminMean + 1
                End If
            End If
        End If
    Next
    dblRTMean = dblRTtot / intNuminMean
    dblRTsd = Sqr((intNuminMean * dblRTsqtot - dblRTtot ^ 2) / (intNuminMean * (intNuminMean - 1#)))
    
    Print #ff,
    Print #ff, "* RESULTS SUMMARY"
    Print #ff, "Errors of commission (number of 3s pressed for)" & vbTab & intErrComiss
    Print #ff, "Errors of omission (missed presses for non-3s)" & vbTab & intErrOmiss
    Print #ff, "  - of which " & intAnticip & " were anticipatory responses."
    Print #ff, "Mean correct RT " & dblRTMean
    Print #ff, "Standard deviation of RT "; dblRTsd
    
    Print #ff,
    Print #ff, "* RAW DATA"
    Print #ff, "Trial"; vbTab; "Target?"; vbTab; "Response?"; vbTab; "RT (ms)"
    For i = 1 To intTotalTrials
        Print #ff, i; vbTab; intWasTarget(i); vbTab; intWasResp(i); vbTab; Round(dblRT(i), 1)
    Next
    Close #ff
    
    
    SyncAndBlank
    WaitForTime (GetTimer + 1500#)
    
    Me.Visible = False
    frmEndOfTest.Visible = True
    frmEndOfTest.Refresh
    DoEvents
    frmEndOfTest.intResponse = 0
    ShowCursor 1
    Do
        DoEvents
    Loop Until frmEndOfTest.intResponse <> 0
    frmEndOfTest.Visible = False
    ShowCursor 0

    cleanup

    Me.Visible = False
    frmBlank.Visible = False
    If (frmEndOfTest.intResponse = 1) Then
        frmResults.Visible = True
        frmResults.txtName = frmSubjectDetails.txtName
        frmResults.txtDOB = frmSubjectDetails.txtDOB
        frmResults.txtCommission = intErrComiss
        frmResults.txtOmission = intErrOmiss
        frmResults.txtAnticip = intAnticip
        frmResults.txtMeanRT = Round(dblRTMean, 1)
        frmResults.txtSDRT = Round(dblRTsd, 1)
        If (booRandom) Then
            frmResults.txtTest = "Random, "
        Else
            frmResults.txtTest = "Fixed, "
        End If
        If (booResponseLocking) Then
            frmResults.txtTest = frmResults.txtTest & "response locked"
        Else
            frmResults.txtTest = frmResults.txtTest & "not response locked"
        End If
    End If

End If

End Sub

Sub SyncAndBlank()
Debug.Print "Sync and blank at " & GetTimer

If (dd7 Is Nothing) Then Exit Sub
dd7.WaitForVerticalBlank DDWAITVB_BLOCKBEGIN, 0
Dim area As RECT
area.Left = CX - FIXSIZE - 20
area.Top = CY - FIXSIZE - 20
area.Right = CX + FIXSIZE + 20
area.Bottom = CY + FIXSIZE + 20
ds7Screen.BltColorFill area, lngBackColor

End Sub
Sub DrawFixation(booHeavy As Boolean)

Debug.Print "Drawing fixation at " & GetTimer

SyncAndBlank
ds7Screen.SetForeColor 0
ds7Screen.setDrawWidth 1
ds7Screen.DrawCircle CX, CY, FIXSIZE
If (booHeavy) Then ds7Screen.setDrawWidth 4
ds7Screen.DrawLine CX - FIXSIZE, CY - FIXSIZE, CX + FIXSIZE, CY + FIXSIZE
ds7Screen.DrawLine CX - FIXSIZE, CY + FIXSIZE, CX + FIXSIZE, CY - FIXSIZE

End Sub

Sub PresentTrial(booResponseLocking As Boolean, intDigit As Integer, lngStartTime As Double)
Dim i, j

Do While (GetTimer < lngStartTime)
    DoEvents
Loop

SyncAndBlank

intTrialNum = intTrialNum + 1
ds7Screen.DrawText CX - 20, CY - 30, intDigit, False

Do While (GetTimer < (lngStartTime + 250))
    DoEvents
Loop
DrawFixation False

If (booResponseLocking) Then
    Do While (GetTimer < (lngStartTime + 350))
        DoEvents
    Loop
    DrawFixation True
    
    Do While (GetTimer < (lngStartTime + 400))
        DoEvents
    Loop
    DrawFixation False
End If


End Sub



