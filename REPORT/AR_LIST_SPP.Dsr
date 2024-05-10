VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_LIST_SPP 
   Caption         =   "ActiveReport1"
   ClientHeight    =   11115
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35930
   _ExtentY        =   19606
   SectionData     =   "AR_LIST_SPP.dsx":0000
End
Attribute VB_Name = "AR_LIST_SPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject

Private Sub Detail_BeforePrint()
Static i As Long

i = i + 1

fldno = i & "."

If fso.FileExists(fldfileSPP) = False Then
fldno.ForeColor = vbRed
fldkdcustomer.ForeColor = vbRed
fldnmcustomer.ForeColor = vbRed
fldalamat.ForeColor = vbRed
fldnospp.ForeColor = vbRed
fldtglSPP.ForeColor = vbRed
fldtgllampiran.ForeColor = vbRed
fldketerangan.ForeColor = vbRed
fldscan.ForeColor = vbRed
fldscan = "BELUM"
Else
fldno.ForeColor = vbBlack
fldkdcustomer.ForeColor = vbBlack
fldnmcustomer.ForeColor = vbBlack
fldalamat.ForeColor = vbBlack
fldnospp.ForeColor = vbBlack
fldtglSPP.ForeColor = vbBlack
fldtgllampiran.ForeColor = vbBlack
fldketerangan.ForeColor = vbBlack
fldscan.ForeColor = vbBlack
fldscan = "SUDAH"
End If
End Sub

