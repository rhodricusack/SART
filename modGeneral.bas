Attribute VB_Name = "modGeneral"
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

' DirectX globals
Public dx7 As New DirectX7
Public dd7 As DirectDraw7
Public ds7Screen As DirectDrawSurface7

' Experiment control globals
Public booPractice As Boolean
Public booResponseLocking As Boolean
Public booRandom As Boolean
Public vntFilename As Variant

Sub cleanup()
ShowCursor 1
If Not (IsNull(dd7)) Then
    dd7.RestoreDisplayMode
End If
Set ds7Screen = Nothing
Set dd7 = Nothing


End Sub
