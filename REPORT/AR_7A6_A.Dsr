VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_7A6_A 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19606
   SectionData     =   "AR_7A6_A.dsx":0000
End
Attribute VB_Name = "AR_7A6_A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Me.Hide
End If
End Sub


Private Sub Detail_BeforePrint()
Static i As Long

i = i + 1

fldno = i & "."

If fldP_disp <> fldP_disp1 Then
    fldP_disp.BackStyle = ddBKNormal
    fldP_disp.BackColor = vbYellow
    
    fldP_disp1.BackStyle = ddBKNormal
    fldP_disp1.BackColor = vbYellow
Else
    fldP_disp.BackStyle = ddBKTransparent
    fldP_disp1.BackStyle = ddBKTransparent
End If


If fldP_SHW <> fldP_SHW1 Then
    fldP_SHW.BackStyle = ddBKNormal
    fldP_SHW.BackColor = vbYellow
    
    fldP_SHW1.BackStyle = ddBKNormal
    fldP_SHW1.BackColor = vbYellow
Else
    fldP_SHW.BackStyle = ddBKTransparent
    fldP_SHW1.BackStyle = ddBKTransparent
End If


If fldP_RG <> fldp_rg1 Then
    fldP_RG.BackStyle = ddBKNormal
    fldP_RG.BackColor = vbYellow
    
    fldp_rg1.BackStyle = ddBKNormal
    fldp_rg1.BackColor = vbYellow
Else
    fldP_RG.BackStyle = ddBKTransparent
    fldp_rg1.BackStyle = ddBKTransparent
End If


If fldS_disp <> fldS_disp1 Then
    fldS_disp.BackStyle = ddBKNormal
    fldS_disp.BackColor = vbYellow
    
    fldS_disp1.BackStyle = ddBKNormal
    fldS_disp1.BackColor = vbYellow
Else
    fldS_disp.BackStyle = ddBKTransparent
    fldS_disp1.BackStyle = ddBKTransparent
End If


End Sub



