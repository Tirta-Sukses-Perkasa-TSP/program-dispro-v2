VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_SPP 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   8760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   24765
   _ExtentY        =   15452
   SectionData     =   "AR_SPP.dsx":0000
End
Attribute VB_Name = "AR_SPP"
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
On Error Resume Next

ImageTT.Picture = LoadPicture(App.Path & "\gambar\TT_SPP.gif")
End Sub

Private Sub Detail_Format()
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
