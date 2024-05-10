VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Cetak_7A6 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   10920
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18765
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   10920
   ScaleWidth      =   18765
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Opt2 
      BackColor       =   &H00000000&
      Caption         =   "DETAIL PEMBAHARUAN SPP"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   7110
      TabIndex        =   16
      Top             =   1305
      Width           =   2760
   End
   Begin VB.OptionButton OPT1 
      BackColor       =   &H00000000&
      Caption         =   "LIST PEMBAHARUAN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   4950
      TabIndex        =   15
      Top             =   1305
      Width           =   2085
   End
   Begin VB.TextBox txttgl2 
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
      Left            =   3240
      TabIndex        =   2
      Top             =   1305
      Width           =   1590
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
      Left            =   1215
      TabIndex        =   1
      Top             =   1305
      Width           =   1590
   End
   Begin VB.Timer TimerPdf 
      Left            =   14895
      Top             =   2295
   End
   Begin VB.Timer TimerRtf 
      Left            =   13950
      Top             =   2295
   End
   Begin VB.Timer Timerxls 
      Left            =   14400
      Top             =   2295
   End
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
      Left            =   9810
      TabIndex        =   0
      Top             =   2115
      Width           =   555
   End
   Begin Threed.SSCommand cmdfs 
      Height          =   300
      Left            =   16065
      TabIndex        =   3
      Top             =   2070
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
      Picture         =   "Cetak_7A6.frx":0000
      Caption         =   "&Full Screen"
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   8220
      Left            =   360
      TabIndex        =   4
      Top             =   1980
      Width           =   17130
      _ExtentX        =   30215
      _ExtentY        =   14499
      SectionData     =   "Cetak_7A6.frx":6862
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   405
      TabIndex        =   5
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
      Left            =   17730
      TabIndex        =   6
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
      Picture         =   "Cetak_7A6.frx":689E
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdPdf 
      Height          =   780
      Left            =   17820
      TabIndex        =   7
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
      Picture         =   "Cetak_7A6.frx":A154
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdrtf 
      Height          =   780
      Left            =   17820
      TabIndex        =   8
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
      Picture         =   "Cetak_7A6.frx":D33B
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdxls 
      Height          =   780
      Left            =   17820
      TabIndex        =   9
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
      Picture         =   "Cetak_7A6.frx":10981
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1575
      TabIndex        =   10
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
      Picture         =   "Cetak_7A6.frx":13E60
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "SD"
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
      Left            =   2835
      TabIndex        =   14
      Top             =   1350
      Width           =   375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TANGGAL :"
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
      Left            =   315
      TabIndex        =   13
      Top             =   1350
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cek Pembaharuan SPP dan Lampiran SPP"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   1260
      TabIndex        =   12
      Top             =   135
      Width           =   8700
   End
   Begin VB.Label lblbarang_R 
      Height          =   330
      Left            =   10530
      TabIndex        =   11
      Top             =   2925
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   10905
      Index           =   0
      Left            =   0
      Picture         =   "Cetak_7A6.frx":1A6C2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "Cetak_7A6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rs As ADODB.Recordset
Dim sqlT, sql1 As String
Dim sqlA As String
Dim kata As String
Dim color As Long, flag As Byte
Dim sqlA1, sqlA2, sqlA3, sqlB1, sqlB2, sqlB3, sqlC, sqlD As String

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


Private Sub Q_Dasar()
sqlA1 = "select a.kdcustomer,a.kdbarang,b.kdkategori,sum(pjm) as pjm,sum(swa) as swa from (" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,sum(b.unit) as pjm,0 as swa from pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam where a.tglpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union all" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,-sum(b.unit) as pjm,0 as swa from Rpinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam where a.tglRpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union all" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,sum(b.unit) as swa from sewa a left join sewa_d b on a.kdsewa=b.kdsewa where a.tglsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union all" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,-sum(b.unit) as swa from Rsewa a left join Rsewa_d b on a.kdRsewa=b.kdRsewa where a.tglRsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       ") a left join barang b on a.kdbarang=b.kdbarang group by a.kdcustomer,a.kdbarang,b.kdkategori"


sqlA2 = "select a.kdcustomer,sum(case when kdkategori in ('04','05','06','07') then pjm else 0 end) as Pjm_Disp," & vbCrLf & _
       "sum(case when kdkategori in ('08','09') then pjm else 0 end) as Pjm_shw,sum(case when kdkategori ='10' then pjm else 0 end) as Pjm_RG, 0 as swa_disp" & vbCrLf & _
       "from (" & sqlA1 & ") a where a.pjm <> 0 and kdkategori between '04' and '10' group by a.kdcustomer union all" & vbCrLf & _
       "select kdcustomer,0 as pjm_disp,0 as pjm_shw, 0 as pjm_rg,sum(case when kdkategori in ('04','05','06','07') then swa else 0 end) as Swa_Disp " & vbCrLf & _
       "from (" & sqlA1 & ") a where a.swa <> 0 and kdkategori between '04' and '10' group by a.kdcustomer"

sqlA3 = "select kdcustomer,sum(pjm_disp) as pjm_disp, sum(pjm_shw) as pjm_shw, sum(pjm_rg) as pjm_Rg,sum(swa_disp) as swa_disp,0 as pjm_disp1,0 as pjm_shw1,0 as pjm_Rg1,0 as swa_disp1  from (" & sqlA2 & ") x group by kdcustomer"


sqlB1 = "select a.kdcustomer,a.kdbarang,b.kdkategori,sum(pjm) as pjm,sum(swa) as swa from (" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,sum(b.unit) as pjm,0 as swa from pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam where a.tglpinjam <= '" & Format(txttgl2, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union all" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,-sum(b.unit) as pjm,0 as swa from Rpinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam where a.tglRpinjam <= '" & Format(txttgl2, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union all" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,sum(b.unit) as swa from sewa a left join sewa_d b on a.kdsewa=b.kdsewa where a.tglsewa <= '" & Format(txttgl2, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       "Union all" & vbCrLf & _
       "select a.kdcustomer,b.kdbarang,0 as pjm,-sum(b.unit) as swa from Rsewa a left join Rsewa_d b on a.kdRsewa=b.kdRsewa where a.tglRsewa <= '" & Format(txttgl2, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang" & vbCrLf & _
       ") a left join barang b on a.kdbarang=b.kdbarang group by a.kdcustomer,a.kdbarang,b.kdkategori"


sqlB2 = "select a.kdcustomer,sum(case when kdkategori in ('04','05','06','07') then pjm else 0 end) as Pjm_Disp," & vbCrLf & _
       "sum(case when kdkategori in ('08','09') then pjm else 0 end) as Pjm_shw,sum(case when kdkategori ='10' then pjm else 0 end) as Pjm_RG, 0 as swa_disp" & vbCrLf & _
       "from (" & sqlB1 & ") a where a.pjm <> 0 and kdkategori between '04' and '10' group by a.kdcustomer union all" & vbCrLf & _
       "select kdcustomer,0 as pjm_disp,0 as pjm_shw, 0 as pjm_rg,sum(case when kdkategori in ('04','05','06','07') then swa else 0 end) as Swa_Disp " & vbCrLf & _
       "from (" & sqlB1 & ") a where a.swa <> 0 and kdkategori between '04' and '10' group by a.kdcustomer"

sqlB3 = "select kdcustomer,0 as pjm_disp,0 as pjm_shw,0 as pjm_Rg,0 as swa_disp,sum(pjm_disp) as pjm_disp1, sum(pjm_shw) as pjm_shw1, sum(pjm_rg) as pjm_Rg1,sum(swa_disp) as swa_disp1 from (" & sqlB2 & ") x group by kdcustomer"

sqlC = "select kdcustomer,sum(pjm_disp) as pjm_disp, sum(pjm_shw) as pjm_shw, sum(pjm_rg) as pjm_Rg,sum(swa_disp) as swa_disp,sum(pjm_disp1) as pjm_disp1, sum(pjm_shw1) as pjm_shw1, sum(pjm_rg1) as pjm_Rg1,sum(swa_disp1) as swa_disp1 from (" & sqlA3 & " union all " & sqlB3 & ") x group by kdcustomer"



End Sub



Private Sub ARV1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub


Private Sub CHK1_Click()
If OPT1.Value = True Then
Call Cetak1
Else
Call Cetak
End If
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
'On Error GoTo hell

MousePointer = vbHourglass

Unload AR_7A6
Unload AR_7A6_A

Call Q_Dasar


sql1 = "select a.*,b.nmcustomer,b.alamat from (" & sqlC & ") a left join customer b on a.kdcustomer=b.kdcustomer where (pjm_disp<>0 or pjm_shw<>0 or pjm_rg<>0 or swa_disp<>0 or pjm_disp1<>0 or pjm_shw1<>0 or pjm_rg1<>0 or swa_disp1<>0) and (pjm_disp <> pjm_disp1 or pjm_shw<>pjm_shw1 or pjm_rg<>pjm_rg1 or swa_disp<>swa_disp1)"

sql = "select * from (" & sql1 & ") z where pjm_disp1 + pjm_shw1 + pjm_rg1 + swa_disp1 <>0 order by nmcustomer"


With AR_7A6_A.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_7A6_A
.fldkdcustomer.DataField = "kdcustomer"
.fldnmcustomer.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.fldP_disp.DataField = "Pjm_DISP"
.fldP_SHW.DataField = "Pjm_SHW"
.fldP_RG.DataField = "pjm_RG"

.fldP_disp1.DataField = "Pjm_DISP1"
.fldP_SHW1.DataField = "Pjm_SHW1"
.fldp_rg1.DataField = "pjm_RG1"

.fldS_disp.DataField = "swa_disp"
.fldS_disp1.DataField = "swa_disp1"

.lblcetak = Format(Now, "dd/MM/yyyy HH:mm")
.lbltgl1 = txttgl1
.lbltgl2 = txttgl2

.GroupHeader1.Visible = False
If Chk1.Value = 1 Then

.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False
.GroupHeader1.Visible = True
.GroupFooter1.Visible = False


.fldkdcustomer.WordWrap = False
.fldnmcustomer.WordWrap = False
.fldalamat.WordWrap = False
.fldP_disp.WordWrap = False
.fldP_SHW.WordWrap = False
.fldP_RG.WordWrap = False

.fldP_disp1.WordWrap = False
.fldP_SHW1.WordWrap = False
.fldp_rg1.WordWrap = False
.fldS_disp.WordWrap = False
.fldS_disp1.WordWrap = False


.fldno.WordWrap = False

End If
'
Set Me.ARV1.ReportSource = AR_7A6_A
End With

MousePointer = vbDefault
'
'Exit Sub
'hell:
'MsgBox err.Description, vbCritical, "Error !"
'Text1 = sql
End Sub



Private Sub Cetak1()
On Error GoTo hell


Unload AR_7A6

Call Q_Dasar


sqlD = "select a.kdcustomer,b.nmcustomer,b.alamat,'SPP' as Pembaharuan,a.pjm_disp1 + a.pjm_shw1 + a.pjm_Rg1 + a.swa_disp1 as X from (" & sqlC & ") a left join customer b on a.kdcustomer=b.kdcustomer where (pjm_disp<>0 or pjm_shw<>0 or pjm_rg<>0 or swa_disp<>0 or pjm_disp1<>0 or pjm_shw1<>0 or pjm_rg1<>0 or swa_disp1<>0) and (pjm_disp <> pjm_disp1 or pjm_shw<>pjm_shw1 or pjm_rg<>pjm_rg1 or swa_disp<>swa_disp1)"

sqlX1 = "select kdcustomer,kdbarang, pjm + swa as qty,kdcustomer + '_' + kdbarang as kode from (" & sqlA1 & ") x where pjm + swa <> 0"

sqlX2 = "select kdcustomer,kdbarang, pjm + swa as qty from (" & sqlB1 & ") x where pjm + swa <>0 and kdcustomer + '_' + kdbarang not in (select kode from (" & sqlX1 & ") a ) and kdcustomer not in (select kdcustomer from (" & sqlD & ") b) and kdkategori between '04' and '10'"

sql1 = "select a.kdcustomer,b.nmcustomer,b.alamat,'Lampiran SPP' as pembaharuan,1 as x from (" & sqlX2 & ") a left join customer b on a.kdcustomer=b.kdcustomer group by a.kdcustomer,b.nmcustomer,b.alamat union all " & sqlD & "" & vbCrLf

sql = "select kdcustomer,nmcustomer,alamat,case when x=0 then 'Retur' else pembaharuan end as pembaharuan,x from (" & sql1 & ") z order by nmcustomer"

With AR_7A6.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_7A6
.fldkdcustomer.DataField = "kdcustomer"
.fldnmcustomer.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.fldpembaharuan.DataField = "pembaharuan"


.lblcetak = Format(Now, "dd/MM/yyyy HH:mm")
.lbltgl1 = txttgl1
.lbltgl2 = txttgl2


.GroupHeader1.Visible = False

If Chk1.Value = 1 Then

.ReportFooter.Visible = False
.ReportHeader.Visible = False
.PageFooter.Visible = False
.PageHeader.Visible = False
.GroupHeader1.Visible = True
.GroupFooter1.Visible = False


.fldkdcustomer.WordWrap = False
.fldnmcustomer.WordWrap = False
.fldalamat.WordWrap = False
.fldpembaharuan.WordWrap = False
.fldno.WordWrap = False

End If
'
Set Me.ARV1.ReportSource = AR_7A6
End With
'
''

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
Text1 = sql
End Sub





Private Sub cmdBRKr_Click()
Karyawan_BR.LBLKODE = "LAD"
Karyawan_BR.Show vbModal

End Sub

Private Sub cmdBRKr_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub







Private Sub cmdfs_Click()
If OPT1.Value = True Then
AR_7A6.Show vbModal
End If
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
If OPT1.Value = True Then
Call Cetak1
Else
Call Cetak
End If

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



Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

OPT1.Value = True

txttgl1 = Date
txttgl2 = Date

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

Private Sub txttgl2_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttgl2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If

End Sub

Private Sub txttgl2_KeyPress(KeyAscii As Integer)
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

Private Sub txttgl2_LostFocus()
On Error GoTo hell

txttgl2 = FormatDateTime(txttgl2, vbGeneralDate)

Exit Sub
hell:
MsgBox "Format Tanggal tidak sesuai !", vbCritical, "Error !"
txttgl2.SetFocus
End Sub













