VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form List_Customer_SPP 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   10200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   10200
   ScaleWidth      =   20370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timerxls 
      Left            =   19665
      Top             =   3105
   End
   Begin VB.TextBox TXTCARI 
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
      Height          =   300
      Left            =   15930
      TabIndex        =   9
      Top             =   585
      Width           =   2850
   End
   Begin VB.Timer TimerG 
      Left            =   5535
      Top             =   1665
   End
   Begin VB.Timer TimerALL 
      Left            =   6075
      Top             =   1665
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   495
      TabIndex        =   4
      Top             =   540
      Width           =   18600
      _Version        =   524288
      _ExtentX        =   32808
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   945
      TabIndex        =   5
      Top             =   9540
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
      Picture         =   "List_Customer_SPP.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSOption Opt1 
      Height          =   330
      Left            =   495
      TabIndex        =   1
      Top             =   585
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   582
      _Version        =   262144
      ForeColor       =   65280
      BackColor       =   0
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "ALL Customer"
   End
   Begin Threed.SSOption Opt2 
      Height          =   330
      Left            =   2250
      TabIndex        =   2
      Top             =   585
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   582
      _Version        =   262144
      ForeColor       =   65280
      BackColor       =   0
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Blom Ada SPP"
   End
   Begin Threed.SSOption Opt3 
      Height          =   330
      Left            =   4410
      TabIndex        =   3
      Top             =   585
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   582
      _Version        =   262144
      ForeColor       =   65280
      BackColor       =   0
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sudah Ada SPP"
   End
   Begin Threed.SSCommand cmdT 
      Height          =   870
      Index           =   0
      Left            =   19305
      TabIndex        =   11
      ToolTipText     =   "Buka Folder SPP"
      Top             =   1125
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1535
      _Version        =   262144
      CaptionStyle    =   1
      ForeColor       =   16744576
      BackColor       =   -2147483643
      PictureMaskColor=   -2147483643
      PictureFrames   =   1
      BackStyle       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "List_Customer_SPP.frx":6862
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdxls 
      Height          =   915
      Left            =   19305
      TabIndex        =   12
      ToolTipText     =   "Simpan"
      Top             =   2025
      Width           =   870
      _ExtentX        =   1535
      _ExtentY        =   1614
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
      Picture         =   "List_Customer_SPP.frx":AF2B
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   8475
      Left            =   315
      TabIndex        =   0
      Top             =   900
      Width           =   18915
      _cx             =   33364
      _cy             =   14949
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
      BackColorAlternate=   16777152
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
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"List_Customer_SPP.frx":E40A
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
      Editable        =   2
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
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   2775
      Left            =   6660
      TabIndex        =   13
      Top             =   4050
      Width           =   7140
      _ExtentX        =   12594
      _ExtentY        =   4895
      SectionData     =   "List_Customer_SPP.frx":E529
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "CARI :"
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
      Left            =   15300
      TabIndex        =   10
      Top             =   630
      Width           =   555
   End
   Begin VB.Label lblalamat_SPP 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "sssssss"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   330
      Left            =   11340
      TabIndex        =   8
      Top             =   135
      Width           =   4650
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   5670
      TabIndex        =   7
      Top             =   9450
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Monitoring SPP"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   510
      Left            =   1575
      TabIndex        =   6
      Top             =   0
      Width           =   5280
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   19215
      Picture         =   "List_Customer_SPP.frx":E565
      Stretch         =   -1  'True
      Top             =   270
      Width           =   285
   End
   Begin VB.Image Image1 
      Height          =   10185
      Left            =   0
      Picture         =   "List_Customer_SPP.frx":E925
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20355
   End
End
Attribute VB_Name = "List_Customer_SPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim color As Long, flag As Byte
Dim fso As New FileSystemObject
Dim kata As String

Private Sub Cetak()


End Sub


Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdT_Click(Index As Integer)
If rs!keterangan <> "" Then
Shell "explorer.exe " & lblalamat_SPP & rs!keterangan & "", vbMaximizedFocus
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

Private Sub DataGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
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
'On Error GoTo hell
Dim filename As String
Dim alamat_SPP As String
Dim i, j As Integer

MousePointer = vbHourglass

filename = App.Path & "\Koneksi.ini"
alamat_SPP = ReadINI("Koneksi", "alamat_SPP", filename)


lblalamat_SPP = alamat_SPP



sqlSPP2 = "select row_number() over (partition by kdcustomer order by tglspp desc) as x, kdcustomer,nospp,tglSPP from list_SPP "

sqllamp = "select max(tgllampiran) as tgllampiran,nospp from lampiran_SPP where nospp in (select nospp from (" & sqlSPP2 & ") a where x=1) group by nospp"

sqlSPP = "select a.*,b.tgllampiran from (" & sqlSPP2 & ") a left join (" & sqllamp & ") b on a.nospp=b.nospp where a.x=1"

sqlps = "select kdcustomer,sum(sisa) as sisa from rekap_pjm_sewa_SPP where kdkategori > 3 group by kdcustomer"

sql1 = "select a.kdcustomer,b.nmcustomer,b.alamat,b.keterangan,c.nospp,c.tglspp,c.tgllampiran from (" & sqlps & ") a left join customer b on a.kdcustomer=b.kdcustomer left join (" & sqlSPP & ") c on a.kdcustomer=c.kdcustomer"

If TXTCARI = "" Then
sql = "select *,'" & alamat_SPP & "' + keterangan + '\' + nospp + '.pdf' as fileSPP from (" & sql1 & ") x where " & kata & ""
Else
sql = "select *,'" & alamat_SPP & "' + keterangan + '\' + nospp + '.pdf' as fileSPP from (" & sql1 & ") x where " & kata & " and (kdcustomer like '%" & TXTCARI & "%' or nmcustomer like '%" & TXTCARI & "%' or alamat like '%" & TXTCARI & "%' or nospp like '%" & TXTCARI & "%')"
End If


'cetak di ARV

Unload AR_LIST_SPP


With AR_LIST_SPP.DC1
.ConnectionString = koneksi
.Source = sql
End With
'
With AR_LIST_SPP
.fldkdcustomer.DataField = "kdcustomer"
.fldnmcustomer.DataField = "nmcustomer"
.fldalamat.DataField = "alamat"
.fldnospp.DataField = "nospp"
.fldtglSPP.DataField = "tglspp"
.fldtgllampiran.DataField = "tgllampiran"
.fldketerangan.DataField = "keterangan"
.fldfileSPP.DataField = "fileSPP"

Set Me.ARV1.ReportSource = AR_LIST_SPP
End With

'-------------------------

Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs



If rs.RecordCount <> 0 Then
    For i = 1 To (datagrid1.Rows - 1)
    For j = 1 To (datagrid1.Cols - 1)
    
    
    If fso.FileExists(datagrid1.TextMatrix(i, 8)) = False Then
    datagrid1.Cell(flexcpForeColor, i, j) = vbRed
    End If
    
    
    datagrid1.TextMatrix(i, 0) = i
    
    Next
    Next

End If



MousePointer = vbDefault
'Exit Sub
'hell:
'SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
'MsgBox err.Description, vbCritical, "Error !!"
'Text1 = sqlps
End Sub



Private Sub DataGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
TimerG.Interval = 10

If KeyCode = vbKeyEnd Then
rs.MoveLast
ElseIf KeyCode = vbKeyHome Then
rs.MoveFirst
ElseIf KeyCode = vbKeyF5 Then
TimerALL.Interval = 10
End If
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

Opt1.Value = True

TimerALL.Interval = 10
End Sub




Private Sub OPT1_Click(Value As Integer)
kata = "kdcustomer <> '@@@'"
TimerALL.Interval = 10
End Sub

Private Sub Opt2_Click(Value As Integer)
kata = "nospp is null"
TimerALL.Interval = 10
End Sub

Private Sub Opt3_Click(Value As Integer)
kata = "nospp <> ''"
TimerALL.Interval = 10
End Sub

Private Sub TimerAll_Timer()
Call all

TimerALL.Interval = 0

End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
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
    
    TimerALL.Interval = 10
    

ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
End If

End Sub


