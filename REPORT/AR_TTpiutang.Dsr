VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} AR_TTpiutang 
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
   SectionData     =   "AR_TTpiutang.dsx":0000
End
Attribute VB_Name = "AR_TTpiutang"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rst As ADODB.Recordset

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Me.Hide
End If
End Sub

Private Sub Detail_BeforePrint()
If fldUP <> "" Then
fldUP = "UP : " & fldUP
fldFrom = "FROM : " & fldFrom
End If
End Sub

Private Sub Detail_Format()
Set SR1.object = New AR_TTpiutang_d
sqlA1 = "select a.nomer,a.kdcustomer,isnull(b.kdpiutang,'') as kdpiutang,c.nmcustomer,case when alamat_X=1 then c.Alamat_tgh else c.alamat End as Alamat,isnull(b.jmlpiutang,0) as jmlpiutang from list_cetak_TT a left join" & vbCrLf & _
       "PiutangSewa b on a.kdcustomer=b.kdcustomer left join Customer c on a.kdcustomer=c.kdcustomer where b.bln=" & Cetak_TTpiutang.CMBbulan.ListIndex + 1 & " and b.tahun = " & Cetak_TTpiutang.txttahun & " and a.nomer=" & FLDNOMER & ""
       
sqlA = "Select row_number() over (partition by x.nomer order by x.kdcustomer) as Urut,x.* from (" & sqlA1 & ") x where jmlpiutang<>0"



With SR1.object.DC1
.ConnectionString = koneksi
.Source = sqlA
End With


With SR1.object
.fldno.DataField = "urut"
.fldkdpiutang.DataField = "kdpiutang"
.fldnmcustomer.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.fldnominal.DataField = "jmlpiutang"


End With




Set SR2.object = New AR_TTPiutang_D1
sqlT = "select nomer,sum(jmlPiutang) as jmlpiutang from (" & sqlA1 & ") x group by nomer"

With SR2.object.DC1
.ConnectionString = koneksi
.Source = sqlT
End With


With SR2.object
.fldtotal.DataField = "jmlpiutang"

End With

End Sub
