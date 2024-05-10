VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} X_AR_rptA1 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   8760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   27173
   _ExtentY        =   15452
   SectionData     =   "X_AR_rptA1.dsx":0000
End
Attribute VB_Name = "X_AR_rptA1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Me.Hide
End If
End Sub


Private Sub Detail_Format()
If fldurut = "1" And fldx = "1" Then
fldcabang.Font.Bold = False
fld120.Font.Bold = False
fld150.Font.Bold = False
fld220.Font.Bold = False
fld240.Font.Bold = False
fld250.Font.Bold = False
fld330.Font.Bold = False
fld600.Font.Bold = False
fld1500.Font.Bold = False
fld19.Font.Bold = False
fldsps.Font.Bold = False
fldtotal.Font.Bold = False
fldgln.Font.Bold = False

fldcabang.Visible = True
fldshipdt.Visible = True
fldperCD.Visible = True
fldnmsp.Visible = True
fldcustCD.Visible = True
fldcustnm.Visible = True
fldaddr.Visible = True

fldnmsp.WordWrap = True
fldcustCD.WordWrap = True
fldcustnm.WordWrap = True
fldaddr.WordWrap = True

Frame1.BackColor = vbWhite

ElseIf fldurut = "1" And fldx <> "1" Then

fldcabang.Font.Bold = False
fld120.Font.Bold = False
fld150.Font.Bold = False
fld220.Font.Bold = False
fld240.Font.Bold = False
fld250.Font.Bold = False
fld330.Font.Bold = False
fld600.Font.Bold = False
fld1500.Font.Bold = False
fld19.Font.Bold = False
fldsps.Font.Bold = False
fldtotal.Font.Bold = False
fldgln.Font.Bold = False

fldcabang.Visible = False
fldshipdt.Visible = True
fldperCD.Visible = True
fldnmsp.Visible = False
fldcustCD.Visible = False
fldcustnm.Visible = False
fldaddr.Visible = False

fldnmsp.WordWrap = False
fldcustCD.WordWrap = False
fldcustnm.WordWrap = False
fldaddr.WordWrap = False

Frame1.BackColor = vbWhite

ElseIf fldurut = "2" Or fldurut = "3" Then

fldcabang.Font.Bold = True
fld120.Font.Bold = True
fld150.Font.Bold = True
fld220.Font.Bold = True
fld240.Font.Bold = True
fld250.Font.Bold = True
fld330.Font.Bold = True
fld600.Font.Bold = True
fld1500.Font.Bold = True
fld19.Font.Bold = True
fldsps.Font.Bold = True
fldtotal.Font.Bold = True
fldgln.Font.Bold = True


fldcabang.Visible = True
fldshipdt.Visible = False
fldperCD.Visible = False
fldnmsp.Visible = False
fldcustCD.Visible = False
fldcustnm.Visible = False
fldaddr.Visible = False

fldnmsp.WordWrap = False
fldcustCD.WordWrap = False
fldcustnm.WordWrap = False
fldaddr.WordWrap = False

Frame1.BackColor = vbYellow

End If
End Sub
