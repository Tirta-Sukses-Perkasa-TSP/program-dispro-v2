VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form X_Rpt_A3 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   10905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18780
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   10905
   ScaleWidth      =   18780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkPer 
      BackColor       =   &H00000000&
      Caption         =   "Tampilkan Periode Idf"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   330
      Left            =   9630
      MaskColor       =   &H00000000&
      TabIndex        =   12
      Top             =   1575
      Width           =   2400
   End
   Begin Threed.SSCommand cmdfs 
      Height          =   300
      Left            =   16200
      TabIndex        =   16
      Top             =   2835
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
      _Version        =   262144
      ForeColor       =   255
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "X_Rpt_A3.frx":0000
      Caption         =   "&Full Screen"
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin VB.CheckBox ChkGT 
      BackColor       =   &H00000000&
      Caption         =   "Jgn Tampilkan Total Per Customer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   12465
      MaskColor       =   &H00000000&
      TabIndex        =   32
      Top             =   1260
      Width           =   2850
   End
   Begin VB.ComboBox cmbASPS 
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
      Height          =   345
      Left            =   9405
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   765
      Width           =   1815
   End
   Begin VB.ComboBox cmbASPM 
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
      Height          =   345
      Left            =   6660
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   765
      Width           =   1815
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   7005
      Left            =   180
      TabIndex        =   10
      Top             =   2745
      Width           =   17400
      _ExtentX        =   30692
      _ExtentY        =   12356
      SectionData     =   "X_Rpt_A3.frx":6862
   End
   Begin VB.OptionButton OTSP3 
      BackColor       =   &H00000000&
      Caption         =   "Rekap Per Customer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   4590
      TabIndex        =   9
      Top             =   1575
      Width           =   2355
   End
   Begin VB.OptionButton OTSP2 
      BackColor       =   &H00000000&
      Caption         =   "Rincian Per Customer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   2160
      TabIndex        =   8
      Top             =   1575
      Width           =   2355
   End
   Begin VB.OptionButton OTSP1 
      BackColor       =   &H00000000&
      Caption         =   "Raw data"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   900
      TabIndex        =   7
      Top             =   1575
      Width           =   1185
   End
   Begin VB.ComboBox CMbDbase 
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
      Height          =   345
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   765
      Width           =   1500
   End
   Begin VB.Timer Timerxls 
      Left            =   14490
      Top             =   2070
   End
   Begin VB.Timer TimerRtf 
      Left            =   14040
      Top             =   2070
   End
   Begin VB.Timer TimerPdf 
      Left            =   14985
      Top             =   2070
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
      Left            =   12465
      TabIndex        =   4
      Top             =   765
      Width           =   1365
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
      Left            =   14130
      TabIndex        =   5
      Top             =   765
      Width           =   1365
   End
   Begin VB.ComboBox cmbgroup 
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
      Height          =   345
      Left            =   8055
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1575
      Width           =   1500
   End
   Begin VB.OptionButton OAIBM1 
      BackColor       =   &H00000000&
      Caption         =   "Raw data"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   900
      TabIndex        =   13
      Top             =   2295
      Width           =   1185
   End
   Begin VB.ComboBox cmBdbase1 
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
      Height          =   345
      Left            =   4275
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   765
      Width           =   1500
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   495
      TabIndex        =   17
      Top             =   630
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
      Picture         =   "X_Rpt_A3.frx":689E
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdPdf 
      Height          =   780
      Left            =   17775
      TabIndex        =   15
      ToolTipText     =   "Simpan"
      Top             =   3735
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
      Picture         =   "X_Rpt_A3.frx":A154
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdxls 
      Height          =   780
      Left            =   17775
      TabIndex        =   14
      ToolTipText     =   "Simpan"
      Top             =   2925
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
      Picture         =   "X_Rpt_A3.frx":D33B
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1575
      TabIndex        =   18
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
      Picture         =   "X_Rpt_A3.frx":1081A
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VB.Label lblTBL 
      Caption         =   "Label6"
      Height          =   285
      Left            =   13140
      TabIndex        =   31
      Top             =   315
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Report TSP"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   540
      TabIndex        =   30
      Top             =   1305
      Width           =   1545
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "ASPS :"
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
      Left            =   8775
      TabIndex        =   29
      Top             =   810
      Width           =   600
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "DBase :"
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
      Left            =   540
      TabIndex        =   28
      Top             =   810
      Width           =   735
   End
   Begin VB.Label lblbarang_R 
      Height          =   330
      Left            =   10530
      TabIndex        =   27
      Top             =   2925
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report By PIC Marketing"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   1260
      TabIndex        =   26
      Top             =   45
      Width           =   4560
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Shipt Date :"
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
      Left            =   11475
      TabIndex        =   25
      Top             =   810
      Width           =   960
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
      Left            =   13770
      TabIndex        =   24
      Top             =   810
      Width           =   420
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "ASPM :"
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
      Left            =   5985
      TabIndex        =   23
      Top             =   810
      Width           =   870
   End
   Begin VB.Label lblgroup 
      BackStyle       =   0  'Transparent
      Caption         =   "Group By :"
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
      Left            =   7110
      TabIndex        =   22
      Top             =   1620
      Width           =   960
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Report AIBM"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   540
      TabIndex        =   21
      Top             =   2025
      Width           =   1545
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "DBase :"
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
      Left            =   3600
      TabIndex        =   20
      Top             =   810
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dan"
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
      Left            =   2880
      TabIndex        =   19
      Top             =   810
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   10905
      Index           =   0
      Left            =   0
      Picture         =   "X_Rpt_A3.frx":1707C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "X_Rpt_A3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsX As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim kt_asps As String
Dim kat_group As String
Dim color As Long, flag As Byte

Private Sub cmbASPS_Click()
On Error GoTo hell
sql = "select ASPM, ASPS from " & cmBdbase1 & "..VSP_IAP where ASPS='" & cmbASPS.Text & "'"
Set rsX = con.Execute(sql)

If rsX.RecordCount <> 0 Then
cmbASPM.Text = rsX!ASPM
End If

Exit Sub
hell:
cmbASPS.ListIndex = 0
End Sub

Private Sub cmbgroup_Click()
If cmbgroup.ListIndex = 2 Then
chkPer.Value = 1
chkPer.Enabled = False
Else
chkPer.Enabled = True
End If

End Sub

Private Sub cmdGO_Click()
lblTBL = "TBL_GO"
If OTSP1.Value = True Or OAIBM1.Value = True Then
    Call Cetak_TSP1
ElseIf OTSP2.Value = True Then
    Call Cetak_TSP2
ElseIf OTSP3.Value = True Then
    Call Cetak_TSP3
End If
End Sub

Private Sub cmdGO_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdPDF_Click()
TimerPdf.Interval = 10
End Sub

Private Sub cmdxls_Click()
lblTBL = "TBL_EXCEL"

If OTSP2.Value = True Then
Call Cetak_TSP2

ElseIf OTSP3.Value = True Then
Call Cetak_TSP3

ElseIf OTSP1.Value = True Then
Call Cetak_TSP1

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

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub Form_Load()
GradientForm Me, 0

'pilih database--------------------------------------
sql = "Select * from Dbase_RPT order by urutan"
Set rs = con.Execute(sql)

rs.MoveFirst

Do While Not rs.EOF
CMbDbase.AddItem rs!nmDbase
cmBdbase1.AddItem rs!nmDbase
rs.MoveNext
Loop

CMbDbase.ListIndex = 0
cmBdbase1.ListIndex = 0
'----------------------------------------------------

'pilih aspm------------------------------------------
sql1 = "select ASPM from " & cmBdbase1 & "..VSP_IAP where ASPS<>'-' group by ASPM"
Set rs1 = con.Execute(sql1)

rs1.MoveFirst

Do While Not rs1.EOF
cmbASPM.AddItem rs1!ASPM
rs1.MoveNext
Loop

cmbASPM.ListIndex = 0
'----------------------------------------------------

'pilih asps------------------------------------------
sql2 = "select ASPS from " & cmBdbase1 & "..VSP_IAP where ASPS<>'-' group by ASPS"
Set rs2 = con.Execute(sql2)

cmbASPS.AddItem "ALL"

rs2.MoveFirst

Do While Not rs2.EOF
cmbASPS.AddItem rs2!ASPS
rs2.MoveNext
Loop

cmbASPS.ListIndex = 0
'----------------------------------------------------

txttgl1 = Date
txttgl2 = Date

cmbgroup.AddItem "BLN"
cmbgroup.AddItem "CUSTOMER"
cmbgroup.AddItem "PERIODCD"
cmbgroup.ListIndex = 0

OTSP1.Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Cetak_TSP1()
On Error Resume Next
Dim filename As String
Dim Exel_ODC As String
Dim nmview As String

sqlQ = "select * from User_m where kduser='" & UTAMA.lblkduser & "'"
Set rsQ = con.Execute(sqlQ)

filename = rsQ!alamat_save & "\Kon_rpt.ini"
Exel_ODC = ReadINI("Kon_RPT", "Exel_ODC", filename)
nmview = ReadINI("Kon_RPT", "nmview", filename)

If cmbASPS.ListIndex = 0 Then
kt_asps = "asps<> '@@@'"
Else
kt_asps = "asps= '" & cmbASPS.Text & "'"
End If

con.Execute ("drop view " & nmview & "")

If CMbDbase.Text = cmBdbase1.Text Then

    If OTSP1.Value = True Then
        sql = "create View " & nmview & " As select * from " & CMbDbase & "..V_OMSET_TSP" & " where shipdt between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' And ASPM = '" & cmbASPM.Text & "' And " & kt_asps & " "
    Else
        sql = "create View " & nmview & " As select * from " & CMbDbase & "..V_OMSET_AIBM" & " where shipdt between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' And ASPM = '" & cmbASPM.Text & "' And " & kt_asps & " "
    End If

Else

    If OTSP1.Value = True Then
        sql = "create View " & nmview & " As select * from " & CMbDbase & "..V_OMSET_TSP" & " where shipdt between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' And ASPM = '" & cmbASPM.Text & "' And " & kt_asps & " union all select * from " & cmBdbase1 & "..V_OMSET_TSP" & " where shipdt between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' And ASPM = '" & cmbASPM.Text & "' And " & kt_asps & " "
    Else
        sql = "create View " & nmview & " As select * from " & CMbDbase & "..V_OMSET_AIBM" & " where shipdt between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' And ASPM = '" & cmbASPM.Text & "' And " & kt_asps & " union all select * from " & cmBdbase1 & "..V_OMSET_AIBM" & " where shipdt between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' And ASPM = '" & cmbASPM.Text & "' And " & kt_asps & " "
    End If
    
End If


con.Execute (sql)

Shell "" & Exel_ODC & " " & rsQ!alamat_save & "\rpt.odc", vbMaximizedFocus
End Sub

Private Sub Cetak_TSP2()
On Error Resume Next
Dim filename As String
Dim Exel_ODC As String
Dim nmview As String

sqlQ = "select * from User_m where kduser='" & UTAMA.lblkduser & "'"
Set rsQ = con.Execute(sqlQ)

filename = rsQ!alamat_save & "\Kon_rpt.ini"
Exel_ODC = ReadINI("Kon_RPT", "Exel_ODC", filename)
nmview = ReadINI("Kon_RPT", "nmview", filename)

If cmbASPS.ListIndex = 0 Then
kt_asps = "asps<> '@@@'"
Else
kt_asps = "asps= '" & cmbASPS.Text & "'"
End If

con.Execute ("drop view " & nmview & "")
con.Execute ("drop view " & nmview & "R1" & "")
    
If CMbDbase.Text = cmBdbase1.Text Then
    sql = "create View " & nmview & "R1" & " as SELECT  KDCUST_IAP, PLANTCD, CABANG, SPOINTDESC, CUSTCD, CUSTNM, ADDR1, ASPM, ASPS, SHIPDT,BLN, PERIODCD, INVNUM, SUM(CASE KAT1 WHEN '120 ML' THEN QTY ELSE 0 END) " & vbCrLf & _
      "AS C120ML, SUM(CASE KAT1 WHEN '150 ML' THEN QTY ELSE 0 END) AS C150ML, SUM(CASE KAT1 WHEN '220 ML' THEN QTY ELSE 0 END) AS C220ML,SUM(CASE KAT1 WHEN 'B220 ML' THEN QTY ELSE 0 END) AS C240ML, SUM(CASE KAT1 WHEN '250 ML' THEN QTY ELSE 0 END) AS C250ML," & vbCrLf & _
      "SUM(CASE KAT1 WHEN '330 ML' THEN QTY ELSE 0 END) AS C330ML, SUM(CASE KAT1 WHEN '600 ML' THEN QTY ELSE 0 END) AS C600ML,SUM(CASE KAT1 WHEN '1500 ML' THEN QTY ELSE 0 END) AS C1500ML, SUM(CASE KAT1 WHEN '19 L' THEN QTY ELSE 0 END) AS C19L," & vbCrLf & _
      "SUM(CASE KAT WHEN 'CUP' THEN QTY ELSE 0 END) AS CUP, SUM(CASE KAT WHEN 'BTL' THEN QTY ELSE 0 END) AS BTL," & vbCrLf & _
      "SUM(CASE KAT WHEN 'GLN' THEN QTY ELSE 0 END) AS GLN From " & CMbDbase & "..V_Omset_TSP WHERE shipdt BETWEEN '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' And ASPM = '" & cmbASPM.Text & "' And " & kt_asps & " GROUP BY KDCUST_IAP, PLANTCD, CABANG, SPOINTDESC, CUSTCD, CUSTNM, ADDR1, ASPM, ASPS, SHIPDT,BLN, PERIODCD, INVNUM"
Else
    sql = "create View " & nmview & "R1" & " as SELECT  KDCUST_IAP, PLANTCD, CABANG, SPOINTDESC, CUSTCD, CUSTNM, ADDR1, ASPM, ASPS, SHIPDT,BLN, PERIODCD, INVNUM, SUM(CASE KAT1 WHEN '120 ML' THEN QTY ELSE 0 END) " & vbCrLf & _
      "AS C120ML, SUM(CASE KAT1 WHEN '150 ML' THEN QTY ELSE 0 END) AS C150ML, SUM(CASE KAT1 WHEN '220 ML' THEN QTY ELSE 0 END) AS C220ML,SUM(CASE KAT1 WHEN 'B220 ML' THEN QTY ELSE 0 END) AS C240ML, SUM(CASE KAT1 WHEN '250 ML' THEN QTY ELSE 0 END) AS C250ML," & vbCrLf & _
      "SUM(CASE KAT1 WHEN '330 ML' THEN QTY ELSE 0 END) AS C330ML, SUM(CASE KAT1 WHEN '600 ML' THEN QTY ELSE 0 END) AS C600ML,SUM(CASE KAT1 WHEN '1500 ML' THEN QTY ELSE 0 END) AS C1500ML, SUM(CASE KAT1 WHEN '19 L' THEN QTY ELSE 0 END) AS C19L," & vbCrLf & _
      "SUM(CASE KAT WHEN 'CUP' THEN QTY ELSE 0 END) AS CUP, SUM(CASE KAT WHEN 'BTL' THEN QTY ELSE 0 END) AS BTL," & vbCrLf & _
      "SUM(CASE KAT WHEN 'GLN' THEN QTY ELSE 0 END) AS GLN From " & CMbDbase & "..V_Omset_TSP WHERE shipdt BETWEEN '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' And ASPM = '" & cmbASPM.Text & "' And " & kt_asps & " GROUP BY KDCUST_IAP, PLANTCD, CABANG, SPOINTDESC, CUSTCD, CUSTNM, ADDR1, ASPM, ASPS, SHIPDT,BLN, PERIODCD, INVNUM union all" & vbCrLf & _
      "SELECT  KDCUST_IAP, PLANTCD, CABANG, SPOINTDESC, CUSTCD, CUSTNM, ADDR1, ASPM, ASPS, SHIPDT,BLN, PERIODCD, INVNUM, SUM(CASE KAT1 WHEN '120 ML' THEN QTY ELSE 0 END) " & vbCrLf & _
      "AS C120ML, SUM(CASE KAT1 WHEN '150 ML' THEN QTY ELSE 0 END) AS C150ML, SUM(CASE KAT1 WHEN '220 ML' THEN QTY ELSE 0 END) AS C220ML,SUM(CASE KAT1 WHEN 'B220 ML' THEN QTY ELSE 0 END) AS C240ML, SUM(CASE KAT1 WHEN '250 ML' THEN QTY ELSE 0 END) AS C250ML," & vbCrLf & _
      "SUM(CASE KAT1 WHEN '330 ML' THEN QTY ELSE 0 END) AS C330ML, SUM(CASE KAT1 WHEN '600 ML' THEN QTY ELSE 0 END) AS C600ML,SUM(CASE KAT1 WHEN '1500 ML' THEN QTY ELSE 0 END) AS C1500ML, SUM(CASE KAT1 WHEN '19 L' THEN QTY ELSE 0 END) AS C19L," & vbCrLf & _
      "SUM(CASE KAT WHEN 'CUP' THEN QTY ELSE 0 END) AS CUP, SUM(CASE KAT WHEN 'BTL' THEN QTY ELSE 0 END) AS BTL," & vbCrLf & _
      "SUM(CASE KAT WHEN 'GLN' THEN QTY ELSE 0 END) AS GLN From " & cmBdbase1 & "..V_Omset_TSP WHERE shipdt BETWEEN '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' And ASPM = '" & cmbASPM.Text & "' And " & kt_asps & " GROUP BY KDCUST_IAP, PLANTCD, CABANG, SPOINTDESC, CUSTCD, CUSTNM, ADDR1, ASPM, ASPS, SHIPDT,BLN, PERIODCD, INVNUM"
End If

    
con.Execute (sql)

sql1 = "select '1' as Urut,KDCUST_IAP,CABANG,SPOINTDESC,CUSTCD,CUSTNM,ADDR1,ASPM,ASPS,SHIPDT,BLN,PERIODCD,sum(C120ML) AS C120ML,SUM(C150ML) AS C150ML,SUM(C220ML) AS C220ML,SUM(C240ML) AS C240ML,SUM(C250ML) AS C250ML,SUM(C330ML) AS C330ML,SUM(C600ML) C600ML,SUM(C1500ML) AS C1500ML,SUM(C19L) AS C19L,SUM(CUP+BTL) AS SPS,SUM(GLN) AS GLN,SUM(CUP+BTL+GLN) AS TOTAL from " & nmview & "R1" & " group by KDCUST_IAP,CABANG,SPOINTDESC,CUSTCD,CUSTNM,ADDR1,SHIPDT,BLN,PERIODCD,ASPM,ASPS"

sql2 = "select '2' as Urut,KDCUST_IAP,'TOTAL' as CABANG,'' as SPOINTDESC,'' as CUSTCD,'' as CUSTNM,'' as ADDR1,'' AS ASPM,'' AS ASPS,'1900/01/01' as SHIPDT,'13' AS BLN,'13' as PERIODCD,sum(C120ML) AS C120ML,SUM(C150ML) AS C150ML,SUM(C220ML) AS C220ML,SUM(C240ML) AS C240ML,SUM(C250ML) AS C250ML,SUM(C330ML) AS C330ML,SUM(C600ML) C600ML,SUM(C1500ML) AS C1500ML,SUM(C19L) AS C19L,SUM(CUP + BTL) AS SPS,SUM(GLN) AS GLN,SUM(CUP + BTL + GLN) AS TOTAL from " & nmview & "R1" & " GROUP BY kdcust_iap"

If ChkGT.Value = 0 Then
    sqlY = "" & sql1 & " union All " & sql2 & " "
Else
    sqlY = sql1
End If

sqlX1 = "select '3' as Urut,'9999' as kdcust_iap,'GRAND TOTAL' as CABANG,'' as SPOINTDESC,'' as CUSTCD,'' as CUSTNM,'' as ADDR1,'' AS ASPM,'' AS ASPS,'1900/01/01' as SHIPDT,0 as BLN,0 as PERIODCD,sum(C120ML) AS C120ML,SUM(C150ML) AS C150ML,SUM(C220ML) AS C220ML,SUM(C240ML) AS C240ML,SUM(C250ML) AS C250ML,SUM(C330ML) AS C330ML,SUM(C600ML) C600ML,SUM(C1500ML) AS C1500ML,SUM(C19L) AS C19L,SUM(CUP+BTL) AS SPS,SUM(GLN) AS GLN,SUM(CUP+BTL+GLN) AS TOTAL from " & nmview & "R1" & " GROUP BY left(KDCUST_IAP,2)"


If lblTBL = "TBL_EXCEL" Then

    
    sql = "create View " & nmview & " As select row_number() over (partition by kdcust_iap order by urut) as x, * from (" & sqlY & " UNION ALL " & sqlX1 & ") X "
    
    con.Execute (sql)

    Shell "" & Exel_ODC & " " & rsQ!alamat_save & "\rpt.odc", vbMaximizedFocus
Else
'ke active report
    Unload X_AR_rptA1
    Unload X_AR_rptA2
    
    sql = "select row_number() over (partition by kdcust_iap order by urut) as x, * from (" & sqlY & " UNION ALL " & sqlX1 & ") X "
    
    With X_AR_rptA1.DC1
    .ConnectionString = koneksi
    .Source = sql
    End With
    
    With X_AR_rptA1
    .fldcabang.DataField = "cabang"
    .fldnmsp.DataField = "Spointdesc"
    .fldcustCD.DataField = "custcd"
    .fldcustnm.DataField = "custnm"
    .fldaddr.DataField = "addr1"
    .fldshipdt.DataField = "shipdt"
    
    .fldperCD.DataField = "periodcd"
    .fld120.DataField = "c120ml"
    .fld150.DataField = "c150ml"
    .fld220.DataField = "c220ml"
    .fld240.DataField = "c240ml"
    .fld250.DataField = "c250ml"
    .fld330.DataField = "c330ml"
    .fld600.DataField = "c600ml"
    .fld1500.DataField = "c1500ml"
    .fld19.DataField = "c19l"
    .fldsps.DataField = "sps"
    .fldtotal.DataField = "total"
    .fldgln.DataField = "gln"
    .fldx.DataField = "x"
    .fldurut.DataField = "urut"

    .lbltgl1 = txttgl1
    .lbltgl2 = txttgl2
    .lblcetak = Format(Now, "dd/MM/yyyy HH:mm")


    Set Me.ARV1.ReportSource = X_AR_rptA1
    End With
    
End If

Exit Sub
hell:
MsgBox err.Description
End Sub


Private Sub Cetak_TSP3()
On Error Resume Next
Dim filename As String
Dim Exel_ODC As String
Dim nmview As String

sqlQ = "select * from User_m where kduser='" & UTAMA.lblkduser & "'"
Set rsQ = con.Execute(sqlQ)

filename = rsQ!alamat_save & "\Kon_rpt.ini"
Exel_ODC = ReadINI("Kon_RPT", "Exel_ODC", filename)
nmview = ReadINI("Kon_RPT", "nmview", filename)

If cmbASPS.ListIndex = 0 Then
kt_asps = "asps<> '@@@'"
Else
kt_asps = "asps= '" & cmbASPS.Text & "'"
End If

con.Execute ("drop view " & nmview & "")
con.Execute ("drop view " & nmview & "R1" & "")

If CMbDbase.Text = cmBdbase1.Text Then
    sql = "create View " & nmview & "R1" & " as SELECT  KDCUST_IAP, PLANTCD, CABANG, SPOINTDESC, CUSTCD, CUSTNM, ADDR1, ASPM, ASPS, SHIPDT,BLN, PERIODCD, INVNUM, SUM(CASE KAT1 WHEN '120 ML' THEN QTY ELSE 0 END) " & vbCrLf & _
      "AS C120ML, SUM(CASE KAT1 WHEN '150 ML' THEN QTY ELSE 0 END) AS C150ML, SUM(CASE KAT1 WHEN '220 ML' THEN QTY ELSE 0 END) AS C220ML,SUM(CASE KAT1 WHEN 'B220 ML' THEN QTY ELSE 0 END) AS C240ML, SUM(CASE KAT1 WHEN '250 ML' THEN QTY ELSE 0 END) AS C250ML," & vbCrLf & _
      "SUM(CASE KAT1 WHEN '330 ML' THEN QTY ELSE 0 END) AS C330ML, SUM(CASE KAT1 WHEN '600 ML' THEN QTY ELSE 0 END) AS C600ML,SUM(CASE KAT1 WHEN '1500 ML' THEN QTY ELSE 0 END) AS C1500ML, SUM(CASE KAT1 WHEN '19 L' THEN QTY ELSE 0 END) AS C19L," & vbCrLf & _
      "SUM(CASE KAT WHEN 'CUP' THEN QTY ELSE 0 END) AS CUP, SUM(CASE KAT WHEN 'BTL' THEN QTY ELSE 0 END) AS BTL," & vbCrLf & _
      "SUM(CASE KAT WHEN 'GLN' THEN QTY ELSE 0 END) AS GLN From " & CMbDbase & "..V_Omset_TSP WHERE shipdt BETWEEN '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' And ASPM = '" & cmbASPM.Text & "' And " & kt_asps & " GROUP BY KDCUST_IAP, PLANTCD, CABANG, SPOINTDESC, CUSTCD, CUSTNM, ADDR1, ASPM, ASPS, SHIPDT,BLN, PERIODCD, INVNUM"
Else
    sql = "create View " & nmview & "R1" & " as SELECT  KDCUST_IAP, PLANTCD, CABANG, SPOINTDESC, CUSTCD, CUSTNM, ADDR1, ASPM, ASPS, SHIPDT,BLN, PERIODCD, INVNUM, SUM(CASE KAT1 WHEN '120 ML' THEN QTY ELSE 0 END) " & vbCrLf & _
      "AS C120ML, SUM(CASE KAT1 WHEN '150 ML' THEN QTY ELSE 0 END) AS C150ML, SUM(CASE KAT1 WHEN '220 ML' THEN QTY ELSE 0 END) AS C220ML,SUM(CASE KAT1 WHEN 'B220 ML' THEN QTY ELSE 0 END) AS C240ML, SUM(CASE KAT1 WHEN '250 ML' THEN QTY ELSE 0 END) AS C250ML," & vbCrLf & _
      "SUM(CASE KAT1 WHEN '330 ML' THEN QTY ELSE 0 END) AS C330ML, SUM(CASE KAT1 WHEN '600 ML' THEN QTY ELSE 0 END) AS C600ML,SUM(CASE KAT1 WHEN '1500 ML' THEN QTY ELSE 0 END) AS C1500ML, SUM(CASE KAT1 WHEN '19 L' THEN QTY ELSE 0 END) AS C19L," & vbCrLf & _
      "SUM(CASE KAT WHEN 'CUP' THEN QTY ELSE 0 END) AS CUP, SUM(CASE KAT WHEN 'BTL' THEN QTY ELSE 0 END) AS BTL," & vbCrLf & _
      "SUM(CASE KAT WHEN 'GLN' THEN QTY ELSE 0 END) AS GLN From " & CMbDbase & "..V_Omset_TSP WHERE shipdt BETWEEN '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' And ASPM = '" & cmbASPM.Text & "' And " & kt_asps & " GROUP BY KDCUST_IAP, PLANTCD, CABANG, SPOINTDESC, CUSTCD, CUSTNM, ADDR1, ASPM, ASPS, SHIPDT,BLN, PERIODCD, INVNUM union all" & vbCrLf & _
      "SELECT  KDCUST_IAP, PLANTCD, CABANG, SPOINTDESC, CUSTCD, CUSTNM, ADDR1, SHIPDT,BLN, PERIODCD, INVNUM, SUM(CASE KAT1 WHEN '120 ML' THEN QTY ELSE 0 END) " & vbCrLf & _
      "AS C120ML, SUM(CASE KAT1 WHEN '150 ML' THEN QTY ELSE 0 END) AS C150ML, SUM(CASE KAT1 WHEN '220 ML' THEN QTY ELSE 0 END) AS C220ML,SUM(CASE KAT1 WHEN 'B220 ML' THEN QTY ELSE 0 END) AS C240ML, SUM(CASE KAT1 WHEN '250 ML' THEN QTY ELSE 0 END) AS C250ML," & vbCrLf & _
      "SUM(CASE KAT1 WHEN '330 ML' THEN QTY ELSE 0 END) AS C330ML, SUM(CASE KAT1 WHEN '600 ML' THEN QTY ELSE 0 END) AS C600ML,SUM(CASE KAT1 WHEN '1500 ML' THEN QTY ELSE 0 END) AS C1500ML, SUM(CASE KAT1 WHEN '19 L' THEN QTY ELSE 0 END) AS C19L," & vbCrLf & _
      "SUM(CASE KAT WHEN 'CUP' THEN QTY ELSE 0 END) AS CUP, SUM(CASE KAT WHEN 'BTL' THEN QTY ELSE 0 END) AS BTL," & vbCrLf & _
      "SUM(CASE KAT WHEN 'GLN' THEN QTY ELSE 0 END) AS GLN From " & cmBdbase1 & "..V_Omset_TSP WHERE shipdt BETWEEN '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' And ASPM = '" & cmbASPM.Text & "' And " & kt_asps & " GROUP BY KDCUST_IAP, PLANTCD, CABANG, SPOINTDESC, CUSTCD, CUSTNM, ADDR1, ASPM, ASPS, SHIPDT,BLN, PERIODCD, INVNUM"
End If

    
    con.Execute (sql)

    If chkPer.Value = 1 Then
    sql1 = "select '1' as Urut,KDCUST_IAP,CABANG,SPOINTDESC,CUSTCD,CUSTNM,ADDR1,BLN,PERIODCD,year(shipdt) as thn,count(invnum) as EC,sum(C120ML) AS C120ML,SUM(C150ML) AS C150ML,SUM(C220ML) AS C220ML,SUM(C240ML) AS C240ML,SUM(C250ML) AS C250ML,SUM(C330ML) AS C330ML,SUM(C600ML) C600ML,SUM(C1500ML) AS C1500ML,SUM(C19L) AS C19L,SUM(CUP+BTL) AS SPS,SUM(GLN) AS GLN,SUM(CUP+BTL+GLN) AS TOTAL from " & nmview & "R1" & " GROUP BY KDCUST_IAP,CABANG,SPOINTDESC,CUSTCD,CUSTNM,ADDR1,BLN,PERIODCD,year(shipdt)"
    Else
    sql1 = "select '1' as Urut,KDCUST_IAP,CABANG,SPOINTDESC,CUSTCD,CUSTNM,ADDR1,BLN,'' as PERIODCD,year(shipdt) as thn,count(invnum) as EC,sum(C120ML) AS C120ML,SUM(C150ML) AS C150ML,SUM(C220ML) AS C220ML,SUM(C240ML) AS C240ML,SUM(C250ML) AS C250ML,SUM(C330ML) AS C330ML,SUM(C600ML) C600ML,SUM(C1500ML) AS C1500ML,SUM(C19L) AS C19L,SUM(CUP+BTL) AS SPS,SUM(GLN) AS GLN,SUM(CUP+BTL+GLN) AS TOTAL from " & nmview & "R1" & " GROUP BY KDCUST_IAP,CABANG,SPOINTDESC,CUSTCD,CUSTNM,ADDR1,BLN,year(shipdt)"
    End If
    
    
    If cmbgroup.ListIndex = 0 Then
        sql2 = "select '2' as Urut,'' as KDCUST_IAP,'TOTAL' as CABANG,'' AS SPOINTDESC,'' AS CUSTCD,'' AS CUSTNM,'' AS ADDR1,bln,'' as periodcd,year(shipdt) as thn,count(invnum) as EC,sum(C120ML) AS C120ML,SUM(C150ML) AS C150ML,SUM(C220ML) AS C220ML,SUM(C240ML) AS C240ML,SUM(C250ML) AS C250ML,SUM(C330ML) AS C330ML,SUM(C600ML) C600ML,SUM(C1500ML) AS C1500ML,SUM(C19L) AS C19L,SUM(CUP+BTL) AS SPS,SUM(GLN) AS GLN,SUM(CUP+BTL+GLN) AS TOTAL from " & nmview & "R1" & " GROUP BY bln,year(shipdt)"
    ElseIf cmbgroup.ListIndex = 1 Then
        sql2 = "select '2' as Urut,KDCUST_IAP,'TOTAL' as CABANG,'' AS SPOINTDESC,'' AS CUSTCD,'' AS CUSTNM,'' AS ADDR1,'13' as bln,'' as periodcd,'3000' as thn,count(invnum) as EC,sum(C120ML) AS C120ML,SUM(C150ML) AS C150ML,SUM(C220ML) AS C220ML,SUM(C240ML) AS C240ML,SUM(C250ML) AS C250ML,SUM(C330ML) AS C330ML,SUM(C600ML) C600ML,SUM(C1500ML) AS C1500ML,SUM(C19L) AS C19L,SUM(CUP+BTL) as SPS,SUM(GLN) AS GLN,SUM(CUP+BTL+GLN) as TOTAL from " & nmview & "R1" & " GROUP BY KDCUST_IAP"
    Else
        sql2 = "select '2' as Urut,'' as KDCUST_IAP,'TOTAL' as CABANG,'' AS SPOINTDESC,'' AS CUSTCD,'' AS CUSTNM,'' AS ADDR1,'' as bln,periodcd,year(shipdt) as thn,count(invnum) as EC,sum(C120ML) AS C120ML,SUM(C150ML) AS C150ML,SUM(C220ML) AS C220ML,SUM(C240ML) AS C240ML,SUM(C250ML) AS C250ML,SUM(C330ML) AS C330ML,SUM(C600ML) C600ML,SUM(C1500ML) AS C1500ML,SUM(C19L) AS C19L,SUM(CUP+BTL) AS SPS,SUM(GLN) AS GLN,SUM(CUP+BTL+GLN) AS TOTAL from " & nmview & "R1" & " GROUP BY periodcd,year(shipdt)"
    End If
    
    If ChkGT.Value = 0 Then
    sqlY = "" & sql1 & " union All " & sql2 & " "
    Else
    sqlY = sql1
    End If
    
    sqlX1 = "select '3' as Urut,'9999' as kdcust_iap,'GRAND TOTAL' as CABANG,'' as SPOINTDESC,'' as CUSTCD,'' as CUSTNM,'' as ADDR1,'13' as bln,'' as periodcd,'3000' as thn,count(invnum) as EC,sum(C120ML) AS C120ML,SUM(C150ML) AS C150ML,SUM(C220ML) AS C220ML,SUM(C240ML) AS C240ML,SUM(C250ML) AS C250ML,SUM(C330ML) AS C330ML,SUM(C600ML) C600ML,SUM(C1500ML) AS C1500ML,SUM(C19L) AS C19L,SUM(CUP+BTL) AS SPS,SUM(GLN) AS GLN,SUM(CUP+BTL+GLN) AS TOTAL from " & nmview & "R1" & " GROUP BY left(KDCUST_IAP,2)"

    If cmbgroup.ListIndex = 0 Then
    kat_group = "thn,bln"
    ElseIf cmbgroup.ListIndex = 1 Then
    kat_group = "kdcust_iap,thn,bln"
    Else
    kat_group = "thn,periodcd"
    End If

If lblTBL = "TBL_EXCEL" Then
     
    sql = "create View " & nmview & " As select row_number() over (partition by " & kat_group & " order by urut) as x, * from (" & sqlY & " UNION ALL " & sqlX1 & ") X "
    
    con.Execute (sql)

    Shell "" & Exel_ODC & " " & rsQ!alamat_save & "\rpt.odc", vbMaximizedFocus
Else
'ke active report
    Unload X_AR_rptA1
    Unload X_AR_rptA2
    
    MousePointer = vbHourglass
    
    sql = "select row_number() over (partition by " & kat_group & " order by urut) as x, * from (" & sqlY & " UNION ALL " & sqlX1 & ") X "
    
    With X_AR_rptA2.DC1
    .ConnectionString = koneksi
    .Source = sql
    End With
    
    With X_AR_rptA2
    If cmbgroup.ListIndex = 0 Or cmbgroup.ListIndex = 2 Then
    .lblOTSP3 = "BLN"
    Else
    .lblOTSP3 = ""
    End If
    
    .fldcabang.DataField = "cabang"
    .fldnmsp.DataField = "Spointdesc"
    .fldcustCD.DataField = "custcd"
    .fldcustnm.DataField = "custnm"
    .fldaddr.DataField = "addr1"
    .fldbln.DataField = "bln"
    .fldperiodCD.DataField = "periodCD"
    

    .fldTHN.DataField = "thn"
    .fld120.DataField = "c120ml"
    .fld150.DataField = "c150ml"
    .fld220.DataField = "c220ml"
    .fld240.DataField = "c240ml"
    .fld250.DataField = "c250ml"
    .fld330.DataField = "c330ml"
    .fld600.DataField = "c600ml"
    .fld1500.DataField = "c1500ml"
    .fld19.DataField = "c19l"
    .fldsps.DataField = "sps"
    .fldtotal.DataField = "total"
    .fldgln.DataField = "gln"
    .fldx.DataField = "x"
    .fldurut.DataField = "urut"

    .lbltgl1 = txttgl1
    .lbltgl2 = txttgl2
    .lblcetak = Format(Now, "dd/MM/yyyy HH:mm")

    Set Me.ARV1.ReportSource = X_AR_rptA2
    End With
    
    MousePointer = vbDefault
    
End If
End Sub

Private Sub OAIBM1_Click()
cmbgroup.Visible = False
lblgroup.Visible = False
chkPer.Visible = False

End Sub

Private Sub OTSP1_Click()
cmbgroup.Visible = False
lblgroup.Visible = False
chkPer.Visible = False
End Sub

Private Sub OTSP2_Click()
cmbgroup.Visible = False
lblgroup.Visible = False
chkPer.Visible = False
End Sub

Private Sub OTSP3_Click()
cmbgroup.Visible = True
lblgroup.Visible = True
chkPer.Visible = True
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