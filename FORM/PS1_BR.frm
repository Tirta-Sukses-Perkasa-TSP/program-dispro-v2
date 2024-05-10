VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form PS1_BR 
   BorderStyle     =   0  'None
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17070
   LinkTopic       =   "Form1"
   ScaleHeight     =   9120
   ScaleWidth      =   17070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtcari 
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
      Left            =   225
      TabIndex        =   1
      Top             =   1305
      Width           =   2490
   End
   Begin VB.Timer TimerALL 
      Left            =   6075
      Top             =   1665
   End
   Begin VB.Timer TimerG 
      Left            =   5535
      Top             =   1665
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   2
      Top             =   855
      Width           =   15855
      _Version        =   524288
      _ExtentX        =   27966
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   900
      TabIndex        =   3
      Top             =   8460
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
      Picture         =   "PS1_BR.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   6360
      Left            =   135
      TabIndex        =   0
      Top             =   1755
      Width           =   15810
      _cx             =   27887
      _cy             =   11218
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16744576
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12648447
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"PS1_BR.frx":6862
      ScrollTrack     =   -1  'True
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   0   'False
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   1
      WordWrap        =   -1  'True
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   4
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin Threed.SSCommand cmdALL 
      Height          =   870
      Left            =   16065
      TabIndex        =   8
      ToolTipText     =   "Pilih Semua"
      Top             =   1755
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
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
      Picture         =   "PS1_BR.frx":69BB
      ButtonStyle     =   4
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Cari Data :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   270
      TabIndex        =   7
      Top             =   990
      Width           =   1500
   End
   Begin VB.Label lblkdkategori 
      Caption         =   "lblkategori"
      Height          =   315
      Left            =   1575
      TabIndex        =   6
      Top             =   9135
      Width           =   1155
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   16065
      Picture         =   "PS1_BR.frx":B406
      Stretch         =   -1  'True
      Top             =   270
      Width           =   285
   End
   Begin VB.Label lbljudul 
      BackStyle       =   0  'Transparent
      Caption         =   "Pinjam Pakai"
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
      Left            =   675
      TabIndex        =   5
      Top             =   135
      Width           =   7755
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   135
      TabIndex        =   4
      Top             =   9135
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   9060
      Left            =   0
      Picture         =   "PS1_BR.frx":B7C6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16980
   End
End
Attribute VB_Name = "PS1_BR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim color As Long, flag As Byte
Dim sql1, sql2, sql As String
Dim sqlALL As String

Private Sub cmdALL_Click()
On Error GoTo hell


If LBLKODE = "RPINJAM_DTU" Then
con.Execute ("insert into RPinjam_d " & sqlALL & "  ")
SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
MsgBox "Data Berhasil di Input Semua", vbInformation, "Info !"

Rpinjam_D.TimerALL.Interval = 10
Unload Me
Unload Rpinjam_DTU
Else
con.Execute ("insert into Rsewa_d " & sqlALL & "  ")
SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
MsgBox "Data Berhasil di Input Semua", vbInformation, "Info !"

RSewa_d.TimerALL.Interval = 10
Unload Me
Unload Rsewa_DTU
End If


Exit Sub
hell:
SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
MsgBox err.Description, vbCritical, "Error !"
End Sub

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


Private Sub LG()

End Sub

Private Sub all()
On Error GoTo hell

If LBLKODE = "RPINJAM_DTU" Then

sql1 = "select a.kdpinjam,b.tglpinjam,b.kdcustomer,a.kdbarang,SUM(a.unit) as unit from pinjam_d a left join Pinjam b on a.kdpinjam =b.kdpinjam group by a.kdpinjam,b.tglpinjam,b.kdcustomer,a.kdbarang Union All" & vbCrLf & _
       "select a.kdpinjam,b.tglpinjam,b.kdcustomer,a.kdbarang,-SUM(a.unit) as unit from Rpinjam_d a left join Pinjam b on a.kdpinjam =b.kdpinjam group by a.kdpinjam,b.tglpinjam,b.kdcustomer,a.kdbarang"
 
sql2 = "select a.kdpinjam,a.tglpinjam,a.kdbarang,b.kd1,b.kdsap,b.nmbarang,b.merk,sum(a.unit) as unit,c.harga,c.rupiah,b.satuan from (" & sql1 & ") a left join barang b on a.kdbarang=b.kdbarang left join pinjam_d c on a.kdpinjam = c.kdpinjam and a.kdbarang=c.kdbarang where a.kdcustomer='" & Rpinjam_D.lblkdcustomer & "' group by a.kdpinjam,a.tglpinjam,a.kdbarang,b.kd1,b.kdsap,b.nmbarang,b.merk,b.satuan,c.harga,c.rupiah "

    If txtcari = "" Then
    sql = "select * from (" & sql2 & ") a where unit <> 0 order by tglpinjam"
    Else
    sql = "select * from (" & sql2 & ") a where unit <> 0 and (kdbarang like '%" & txtcari & "%' or kdPinjam like '%" & txtcari & "%' or kd1 like '%" & txtcari & "%' or kdsap like '%" & txtcari & "%' or merk like '%" & txtcari & "%' ) order by tglpinjam"
    End If


    If txtcari = "" Then
    sqlALL = "select kdbarang + '_' + '" & Rpinjam_D.lblKDRPinjam & "','" & Rpinjam_D.lblKDRPinjam & "',kdbarang,1,0,0,'',kdpinjam,getdate() from (" & sql2 & ") a where unit <> 0"
    Else
    sqlALL = "select kdbarang + '_' + '" & Rpinjam_D.lblKDRPinjam & "','" & Rpinjam_D.lblKDRPinjam & "',kdbarang,1,0,0,'',kdpinjam,getdate() from (" & sql2 & ") a where unit <> 0 and (kdbarang like '%" & txtcari & "%' or kdPinjam like '%" & txtcari & "%' or kd1 like '%" & txtcari & "%' or kdsap like '%" & txtcari & "%' or merk like '%" & txtcari & "%' )"
    End If

ElseIf LBLKODE = "RSEWA_DTU" Then

sql1 = "select a.kdsewa,b.tglsewa,b.kdcustomer,a.kdbarang,SUM(a.unit) as unit from sewa_d a left join sewa b on a.kdsewa =b.kdsewa group by a.kdsewa,b.tglsewa,b.kdcustomer,a.kdbarang Union All" & vbCrLf & _
       "select a.kdsewa,b.tglsewa,b.kdcustomer,a.kdbarang,-SUM(a.unit) as unit from Rsewa_d a left join sewa b on a.kdsewa =b.kdsewa group by a.kdsewa,b.tglsewa,b.kdcustomer,a.kdbarang"
 
sql2 = "select a.kdsewa,a.tglsewa,a.kdbarang,b.kd1,b.kdsap,b.nmbarang,b.merk,sum(a.unit) as unit,c.harga,c.rupiah,b.satuan from (" & sql1 & ") a left join barang b on a.kdbarang=b.kdbarang left join sewa_d c on a.kdsewa = c.kdsewa and a.kdbarang=c.kdbarang where a.kdcustomer='" & RSewa_d.lblkdcustomer & "' group by a.kdsewa,a.tglsewa,a.kdbarang,b.kd1,b.kdsap,b.nmbarang,b.merk,b.satuan,c.harga,c.rupiah "

    If txtcari = "" Then
    sql = "select * from (" & sql2 & ") a where unit <> 0 order by tglsewa"
    Else
    sql = "select * from (" & sql2 & ") a where unit <> 0 and (kdbarang like '%" & txtcari & "%' or kdsewa like '%" & txtcari & "%' or kd1 like '%" & txtcari & "%' or kdsap like '%" & txtcari & "%' or merk like '%" & txtcari & "%') order by tglsewa"
    End If

    If txtcari = "" Then
    sqlALL = "select kdbarang + '_' + '" & RSewa_d.lblKDRsewa & "','" & RSewa_d.lblKDRsewa & "',kdbarang,1,0,0,'',kdsewa,getdate() from (" & sql2 & ") a where unit <> 0"
    Else
    sqlALL = "select kdbarang + '_' + '" & RSewa_d.lblKDRsewa & "','" & RSewa_d.lblKDRsewa & "',kdbarang,1,0,0,'',kdsewa,getdate() from (" & sql2 & ") a where unit <> 0 and (kdbarang like '%" & txtcari & "%' or kdSewa like '%" & txtcari & "%' or kd1 like '%" & txtcari & "%' or kdsap like '%" & txtcari & "%' or merk like '%" & txtcari & "%' )"
    End If
    
Else

End If




Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs
Call LG

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub



Private Sub datagrid1_DblClick()
On Error GoTo hell
If LBLKODE = UCase("RPINJAM_DTU") Then
Rpinjam_DTU.lblkdbarang = rs!kdbarang
Rpinjam_DTU.lblnmbarang = rs!nmbarang
Rpinjam_DTU.txtunit = rs!unit
Rpinjam_DTU.lblmaxunit = rs!unit
Rpinjam_DTU.txtharga = rs!harga
Rpinjam_DTU.lblrupiah = rs!rupiah
Rpinjam_DTU.lblsatuan = rs!satuan
Rpinjam_DTU.lblkdPinjam = rs!kdpinjam
ElseIf LBLKODE = UCase("RSEWA_DTU") Then
Rsewa_DTU.lblkdbarang = rs!kdbarang
Rsewa_DTU.lblnmbarang = rs!nmbarang
Rsewa_DTU.txtunit = rs!unit
Rsewa_DTU.lblmaxunit = rs!unit
Rsewa_DTU.txtharga = rs!harga
Rsewa_DTU.lblrupiah = rs!rupiah
Rsewa_DTU.lblsatuan = rs!satuan
Rsewa_DTU.lblkdsewa = rs!kdsewa

End If
Unload Me

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub

Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
TimerG.Interval = 10

If KeyCode = vbKeyUp Then

    If rs.AbsolutePosition = 1 Then
    txtcari.SetFocus
    End If

ElseIf KeyCode = vbKeyEnd Then
rs.MoveLast
ElseIf KeyCode = vbKeyHome Then
rs.MoveFirst
End If
End Sub

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
TimerG.Interval = 10

On Error GoTo hell

If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
    If LBLKODE = UCase("RPINJAM_DTU") Then
    Rpinjam_DTU.lblkdbarang = rs!kdbarang
    Rpinjam_DTU.lblnmbarang = rs!nmbarang
    Rpinjam_DTU.txtunit = rs!unit
    Rpinjam_DTU.lblmaxunit = rs!unit
    Rpinjam_DTU.txtharga = rs!harga
    Rpinjam_DTU.lblrupiah = rs!rupiah
    Rpinjam_DTU.lblsatuan = rs!satuan
    Rpinjam_DTU.lblkdPinjam = rs!kdpinjam

    ElseIf LBLKODE = UCase("RSEWA_DTU") Then
    Rsewa_DTU.lblkdbarang = rs!kdbarang
    Rsewa_DTU.lblnmbarang = rs!nmbarang
    Rsewa_DTU.txtunit = rs!unit
    Rsewa_DTU.lblmaxunit = rs!unit
    Rsewa_DTU.txtharga = rs!harga
    Rsewa_DTU.lblrupiah = rs!rupiah
    Rsewa_DTU.lblsatuan = rs!satuan
    Rsewa_DTU.lblkdsewa = rs!kdsewa

    
    End If
    Unload Me

ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then

 Call all
End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"


End Sub

Private Sub Form_Load()
GradientForm Me, 0



TimerALL.Interval = 10
End Sub




Private Sub TimerAll_Timer()
On Error Resume Next
Call all

TimerALL.Interval = 0
End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
End Sub

Private Sub TXTCARI_Change()
TimerALL.Interval = 10
End Sub

Private Sub TXTCARI_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub TXTCARI_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
    If rs.RecordCount <> 0 Then
    datagrid1.SetFocus
    Call LG
'    Else
'    CMBCARI.SetFocus
    End If
End If

End Sub

Private Sub TXTCARI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If rs.RecordCount <> 0 Then
    datagrid1.SetFocus
    Call LG
'    Else
'    CMBCARI.SetFocus
    End If

ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
End If

End Sub










