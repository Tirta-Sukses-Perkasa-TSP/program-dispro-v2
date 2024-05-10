VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_SPP_LAMPIRAN 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   9045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   24421
   _ExtentY        =   15954
   SectionData     =   "AR_SPP_LAMPIRAN.dsx":0000
End
Attribute VB_Name = "AR_SPP_LAMPIRAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Me.Hide
End If
End Sub

Private Sub GroupFooter1_BeforePrint()
If Month(lbltglSPP) = 1 Then
lbltglSPP = Format(lbltglSPP, "dd ") & "Januari " & Format(lbltglSPP, "yyyy")
ElseIf Month(lbltglSPP) = 2 Then
lbltglSPP = Format(lbltglSPP, "dd ") & "Februari " & Format(lbltglSPP, "yyyy")
ElseIf Month(lbltglSPP) = 3 Then
lbltglSPP = Format(lbltglSPP, "dd ") & "Maret " & Format(lbltglSPP, "yyyy")
ElseIf Month(lbltglSPP) = 4 Then
lbltglSPP = Format(lbltglSPP, "dd ") & "April " & Format(lbltglSPP, "yyyy")
ElseIf Month(lbltglSPP) = 5 Then
lbltglSPP = Format(lbltglSPP, "dd ") & "Mei " & Format(lbltglSPP, "yyyy")
ElseIf Month(lbltglSPP) = 6 Then
lbltglSPP = Format(lbltglSPP, "dd ") & "Juni " & Format(lbltglSPP, "yyyy")
ElseIf Month(lbltglSPP) = 7 Then
lbltglSPP = Format(lbltglSPP, "dd ") & "Juli " & Format(lbltglSPP, "yyyy")
ElseIf Month(lbltglSPP) = 8 Then
lbltglSPP = Format(lbltglSPP, "dd ") & "Agustus " & Format(lbltglSPP, "yyyy")
ElseIf Month(lbltglSPP) = 9 Then
lbltglSPP = Format(lbltglSPP, "dd ") & "September " & Format(lbltglSPP, "yyyy")
ElseIf Month(lbltglSPP) = 10 Then
lbltglSPP = Format(lbltglSPP, "dd ") & "Oktober " & Format(lbltglSPP, "yyyy")
ElseIf Month(lbltglSPP) = 11 Then
lbltglSPP = Format(lbltglSPP, "dd ") & "November " & Format(lbltglSPP, "yyyy")
ElseIf Month(lbltglSPP) = 12 Then
lbltglSPP = Format(lbltglSPP, "dd ") & "Desember " & Format(lbltglSPP, "yyyy")
End If

End Sub

