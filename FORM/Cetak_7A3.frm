VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Cetak_7A3 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   10920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18795
   LinkTopic       =   "Form2"
   ScaleHeight     =   10920
   ScaleWidth      =   18795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Chk1 
      Caption         =   "Isi"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9990
      TabIndex        =   5
      Top             =   1800
      Width           =   555
   End
   Begin VB.Timer Timerxls 
      Left            =   14310
      Top             =   2295
   End
   Begin VB.Timer TimerRtf 
      Left            =   13860
      Top             =   2295
   End
   Begin VB.Timer TimerPdf 
      Left            =   14805
      Top             =   2295
   End
   Begin VB.TextBox txttgl1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1620
      TabIndex        =   0
      Top             =   1215
      Width           =   1590
   End
   Begin Threed.SSCommand cmdfs 
      Height          =   300
      Left            =   15975
      TabIndex        =   6
      Top             =   1800
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
      _Version        =   262144
      ForeColor       =   255
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Cetak_7A3.frx":0000
      Caption         =   "&Full Screen"
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   315
      TabIndex        =   7
      Top             =   810
      Width           =   17205
      _Version        =   524288
      _ExtentX        =   30348
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdGO 
      Height          =   780
      Left            =   17685
      TabIndex        =   1
      ToolTipText     =   "Simpan"
      Top             =   1170
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1376
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Cetak_7A3.frx":6862
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdPdf 
      Height          =   780
      Left            =   17730
      TabIndex        =   4
      ToolTipText     =   "Simpan"
      Top             =   4590
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1376
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Cetak_7A3.frx":A118
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdrtf 
      Height          =   780
      Left            =   17730
      TabIndex        =   2
      ToolTipText     =   "Simpan"
      Top             =   2970
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1376
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Cetak_7A3.frx":D2FF
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdxls 
      Height          =   780
      Left            =   17730
      TabIndex        =   3
      ToolTipText     =   "Simpan"
      Top             =   3780
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1376
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Cetak_7A3.frx":10945
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1485
      TabIndex        =   8
      Top             =   10440
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   661
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16777215
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Cetak_7A3.frx":13E24
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   8400
      Left            =   360
      TabIndex        =   9
      Top             =   1710
      Width           =   17130
      _ExtentX        =   30215
      _ExtentY        =   14817
      SectionData     =   "Cetak_7A3.frx":1A686
   End
   Begin VB.Label lblbarang_R 
      Height          =   330
      Left            =   10440
      TabIndex        =   12
      Top             =   2925
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rekap Pinjaman dan Sewa Outlet"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   1170
      TabIndex        =   11
      Top             =   135
      Width           =   7665
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PER TANGGAL :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   405
      TabIndex        =   10
      Top             =   1260
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   10905
      Index           =   0
      Left            =   0
      Picture         =   "Cetak_7A3.frx":1A6C2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18690
   End
End
Attribute VB_Name = "Cetak_7A3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rs As ADODB.Recordset
Dim sqlT, sql1 As String
Dim sqlA As String
Dim color As Long, flag As Byte
Dim kategori As String





Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub



Private Sub Form_Activate()
    On Error GoTo err
    color = vbBlue
    flag = flag Or LWA_COLORKEY
    SetTransparan1 Me.hWnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub




Private Sub total()

sql1 = "select kdcustomer,kdbarang,sum(pjm) as pjm,sum(swa) as swa from (" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,sum(b.unit) as pjm,0 as swa from pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam where a.tglpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union ALL" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,-sum(b.unit) as pjm,0 as swa from Rpinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam where a.tglRpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union ALL" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,sum(b.unit) as swa from sewa a left join sewa_d b on a.kdsewa=b.kdsewa where a.tglsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union all" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,-sum(b.unit) as swa from Rsewa a left join Rsewa_d b on a.kdRsewa=b.kdRsewa where a.tglRsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       ") a group by kdcustomer,kdbarang"



sqlA = "select a.kdcustomer,b.nmcustomer,b.alamat,a.kdbarang,c.nmbarang,c.kd1,c.kdkategori,a.pjm,a.swa,(a.pjm + a.swa) as total from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
      "left join barang c on a.kdbarang=c.kdbarang where a.pjm+a.swa <>0 and c.kdkategori between '04' and '10'"


sqlB = "select count(kdcustomer) as jmlcust,nmcustomer,sum(case kdkategori when '04' then pjm else 0 end) as P1,sum(case kdkategori when '05' then pjm else 0 end) as P2,sum(case kdkategori when '06' then pjm else 0 end) as P3,sum(case kdkategori when '07' then pjm else 0 end) as P4,sum(case kdkategori when '08' then pjm else 0 end) as P5,sum(case kdkategori when '09' then pjm else 0 end) as P6,sum(case kdkategori when '10' then pjm else 0 end) as P7," & vbCrLf & _
      "sum(case kdkategori when '04' then swa else 0 end) as S1,sum(case kdkategori when '05' then swa else 0 end) as S2,sum(total) as total from (" & sqlA & ") a group by nmcustomer"

sqlX = "select '1' as kode,a.*,(a.p1+a.p2+a.p3+a.p4+a.p5+a.p6+p7) as P_total,(a.S1+a.S2) as S_total from (" & sqlB & ") a "


sqlT = "select kode,sum(jmlcust) as jmlcust,sum(p1) as p1,sum(p2) as p2,sum(p3) as p3,sum(p4) as p4,sum(p5) as p5,sum(p6) as p6,sum(p7) as p7,sum(p_total) as p_total,sum(S1) as S1,sum(S2) as S2,sum(S_total) as S_total, sum(total) as total from (" & sqlX & ") a group by kode"
Set rs = con.Execute(sqlT)

End Sub





 





Private Sub ARV1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub


Private Sub CHK1_Click()

Call Cetak

End Sub

Private Sub CHK1_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
    If Chk1.Value = 1 Then
    Chk1.Value = 0
    Else
    Chk1.Value = 1
    End If
    
    Call Cetak
        
ElseIf KeyAscii = 27 Then
Unload Me
End If

End Sub


Private Sub Cetak()
On Error GoTo hell


Unload AR_7A3

sql1 = "select kdcustomer,kdbarang,sum(pjm) as pjm,sum(swa) as swa from (" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,sum(b.unit) as pjm,0 as swa from pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam where a.tglpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union ALL" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,-sum(b.unit) as pjm,0 as swa from Rpinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam where a.tglRpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union ALL" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,sum(b.unit) as swa from sewa a left join sewa_d b on a.kdsewa=b.kdsewa where a.tglsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union all" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,-sum(b.unit) as swa from Rsewa a left join Rsewa_d b on a.kdRsewa=b.kdRsewa where a.tglRsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       ") a group by kdcustomer,kdbarang"



sqlA = "select a.kdcustomer,b.nmcustomer,b.alamat,a.kdbarang,c.nmbarang,c.kd1,c.kdkategori,a.pjm,a.swa,(a.pjm + a.swa) as total from (" & sql1 & ") a left join customer b on a.kdcustomer=b.kdcustomer" & vbCrLf & _
      "left join barang c on a.kdbarang=c.kdbarang where a.pjm+a.swa <>0 and c.kdkategori between '04' and '10'"


sqlB = "select 1 as jmlcust,kdcustomer,nmcustomer,sum(case kdkategori when '04' then pjm else 0 end) as P1,sum(case kdkategori when '05' then pjm else 0 end) as P2,sum(case kdkategori when '06' then pjm else 0 end) as P3,sum(case kdkategori when '07' then pjm else 0 end) as P4,sum(case kdkategori when '08' then pjm else 0 end) as P5,sum(case kdkategori when '09' then pjm else 0 end) as P6,sum(case kdkategori when '10' then pjm else 0 end) as P7," & vbCrLf & _
      "sum(case kdkategori when '04' then swa else 0 end) as S1,sum(case kdkategori when '05' then swa else 0 end) as S2,sum(total) as total from (" & sqlA & ") a group by nmcustomer,kdcustomer"
 
sqlC = "select nmcustomer,sum(jmlcust) as jmlCust,sum(p1) as p1,sum(p2) as p2,sum(p3) as p3,sum(p4) as p4,sum(p5) as p5,sum(p6) as p6,sum(p7) as p7,sum(S1) as S1,sum(S2) as S2,sum(total) as total from (" & sqlB & ") a group by nmcustomer"
 
 
sql = "select a.*,(a.p1+a.p2+a.p3+a.p4+a.p5+a.p6+p7) as P_total,(a.S1+a.S2) as S_total from (" & sqlC & ") a order by nmcustomer"

With AR_7A3.DC1
.ConnectionString = koneksi
.Source = sql
End With
'
With AR_7A3
.fldjmlcust.DataField = "jmlcust"
.fldnmcus.DataField = "nmcustomer"
.fldP1.DataField = "p1"
.fldP2.DataField = "p2"
.fldP3.DataField = "p3"
.fldP4.DataField = "p4"
.fldP5.DataField = "p5"
.fldP6.DataField = "p6"
.fldP7.DataField = "p7"
.fldS1.DataField = "S1"
.fldS2.DataField = "S2"
.fldP_total.DataField = "P_total"
.fldS_total.DataField = "S_total"
.fldtotal.DataField = "Total"


.lblcetak = Format(Date, "dd/MM/yyyy")
.lbltgl1 = txttgl1
.lbljudul = "REKAP PINJAMAN & SEWA"


Call total
If rs.RecordCount <> 0 Then
.lblP1 = Format(rs!p1, "#,###0")
.lblP2 = Format(rs!p2, "#,###0")
.lblP3 = Format(rs!p3, "#,###0")
.lblP4 = Format(rs!p4, "#,###0")
.lblP5 = Format(rs!p5, "#,###0")
.lblP6 = Format(rs!p6, "#,###0")
.lblP7 = Format(rs!p7, "#,###0")
.lblP_total = Format(rs!p_total, "#,###0")
.lblS1 = Format(rs!S1, "#,###0")
.lblS2 = Format(rs!S2, "#,###0")
.lblS_total = Format(rs!S_total, "#,###0")
.lbltotal = Format(rs!total, "#,###0")
.lbljmlCust = Format(rs!jmlcust, "#,###0")
Else
.lblP1 = 0
.lblP2 = 0
.lblP3 = 0
.lblP4 = 0
.lblP5 = 0
.lblP6 = 0
.lblP7 = 0
.lblP_total = 0
.lblS1 = 0
.lblS2 = 0
.lblS_total = 0
.lbltotal = 0
.lbljmlCust = 0
End If
'
.GroupHeader1.Visible = False
'
'
If Chk1.Value = 1 Then

.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False
.GroupHeader1.Visible = True
.GroupFooter1.Visible = True


.fldnmcus.WordWrap = False
.fldjmlcust.WordWrap = False
.fldP1.WordWrap = False
.fldP2.WordWrap = False
.fldP3.WordWrap = False
.fldP4.WordWrap = False
.fldP5.WordWrap = False
.fldP6.WordWrap = False
.fldP7.WordWrap = False
.fldS1.WordWrap = False
.fldS2.WordWrap = False
.fldP_total.WordWrap = False
.fldS_total.WordWrap = False
.fldtotal.WordWrap = False
.fldNO.WordWrap = False


End If

Set Me.ARV1.ReportSource = AR_7A3
End With


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"

End Sub



Private Sub cmdfs_Click()
AR_7A3.Zoom = 110
AR_7A3.Show vbModal
End Sub

Private Sub cmdfs_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdOK_Click()
Call Cetak
ARV1.ToolbarVisible = False
ARV1.ToolbarVisible = True
End Sub

Private Sub cmdGO_Click()

Call Cetak

End Sub

Private Sub cmdGO_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdPDF_Click()
TimerPdf.Interval = 10
End Sub

Private Sub cmdPDF_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdrtf_Click()
TimerRtf.Interval = 10
End Sub

Private Sub cmdrtf_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdOK_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub


Private Sub cmdxls_Click()
Timerxls.Interval = 10
End Sub


Private Sub cmdxls_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub



Private Sub Form_Load()
GradientForm Me, 0


txttgl1 = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub


Private Sub TimerPDF_Timer()
On Error GoTo hell
Dim pdf As New ActiveReportsPDFExport.ARExportPDF

out2 = out2 + 1

Call save_out
pdf.filename = alamat_save & "\outfile" & CStr(out2) & ".pdf"
pdf.Export ARV1.Pages

Call EX_PDF(Me)
TimerPdf.Interval = 0

Exit Sub
hell:
TimerPdf.Interval = 0
If out2 < 10 Then
cmdPDF_Click
End If

End Sub

Private Sub Timerrtf_Timer()
On Error GoTo hell
Dim rtf As New ActiveReportsRTFExport.ARExportRTF
out = out + 1

Call save_out
rtf.filename = alamat_save & "\outfile" & CStr(out) & ".rtf"
rtf.Export ARV1.Pages

Call EX_WORD(Me)
TimerRtf.Interval = 0

Exit Sub
hell:
TimerRtf.Interval = 0
If out < 10 Then
cmdrtf_Click
End If
End Sub

Private Sub Timerxls_Timer()
On Error GoTo hell
Dim xls As New ActiveReportsExcelExport.ARExportExcel



out1 = out1 + 1

Call save_out
xls.filename = alamat_save & "\outfile" & CStr(out1) & ".xls"
xls.Export ARV1.Pages

Call EX_EXEL(Me)
Timerxls.Interval = 0

Exit Sub
hell:
Timerxls.Interval = 0
If out1 < 10 Then
cmdxls_Click
End If
End Sub





Private Sub txttgl1_Change()
Call nul(txttgl1)
End Sub

Private Sub txttgl1_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttgl1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If

End Sub

Private Sub txttgl1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii <> vbKeyBack Then
    cekTBL = InStr("1234567890-/", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If

End If

End Sub

Private Sub txttgl1_LostFocus()
On Error GoTo hell

txttgl1 = FormatDateTime(txttgl1, vbGeneralDate)

Exit Sub
hell:
MsgBox "Format Tanggal tidak sesuai !", vbCritical, "Error !"
txttgl1.SetFocus
End Sub

Private Sub txttgl2_Change()
Call nul(txttgl2)
End Sub









