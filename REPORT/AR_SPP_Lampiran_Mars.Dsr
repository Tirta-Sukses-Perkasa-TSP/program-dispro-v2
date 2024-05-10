VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_SPP_Lampiran_Mars 
   BorderStyle     =   0  'None
   Caption         =   "ActiveReport1"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   23865
   _ExtentY        =   19606
   SectionData     =   "AR_SPP_Lampiran_Mars.dsx":0000
End
Attribute VB_Name = "AR_SPP_Lampiran_Mars"
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
' Detail barang
Set SR1.object = New AR_SPP_Lampiran1_Mars

sqlA = "exec sp_lampiran_spp_mars @tgl1='" & Format(Print_Form_MARS.txttglSPP1, "yyyy/MM/dd") & "', @kdcustomer='" & fldkdcustomer & "'"

With SR1.object.DC1
.ConnectionString = koneksi
.Source = sqlA
End With


With SR1.object
.fldno.DataField = "x"
.fldkdbarang_mars.DataField = "kdbarang_mars"
.fldjnsbrg_mars.DataField = "jnsbrg_MARS"
.fldstatus.DataField = "status_kepemilikan"
.fldhrgsewa.DataField = "hrgsewa"

End With

'----------------------------------

End Sub
