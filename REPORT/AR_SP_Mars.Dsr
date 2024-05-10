VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_SP_Mars 
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
   SectionData     =   "AR_SP_Mars.dsx":0000
End
Attribute VB_Name = "AR_SP_Mars"
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
Set SR1.object = New AR_SP_MARS_D

sqlA = "select row_number() OVER (partition BY kdx ORDER BY tglinput) as x,* from V_SP_MARS_D where kdx='" & fldkdx & "'"

With SR1.object.DC1
.ConnectionString = koneksi
.Source = sqlA
End With


With SR1.object
.fldno.DataField = "x"
.fldkdbarang_mars.DataField = "kdbarang_mars"
.fldjnsbrg.DataField = "jnsbrg_MARS"
.fldstatus_kepemilikan.DataField = "status_kepemilikan"
.fldkondisi.DataField = "kondisi"

End With

'----------------------------------


' Detail TOTAL
Set SR2.object = New AR_SP_MARS_Total

sqlB = "select  'Total ' + jnsbrg_mars as jnsbrg_mars, sum(unit)as jml from V_SP_MARS_D where kdx='" & fldkdx & "' group by jnsbrg_mars"

With SR2.object.DC1
.ConnectionString = koneksi
.Source = sqlB
End With


With SR2.object
.fldjnsbrg.DataField = "jnsbrg_MARS"
.fldjml.DataField = "jml"

End With

'----------------------------------

End Sub
