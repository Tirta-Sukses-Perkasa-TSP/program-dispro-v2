VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form PObeli_BR 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16305
   LinkTopic       =   "Form1"
   ScaleHeight     =   10035
   ScaleWidth      =   16305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerG 
      Left            =   5535
      Top             =   1665
   End
   Begin VB.Timer TimerALL 
      Left            =   6075
      Top             =   1665
   End
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
      TabIndex        =   0
      Top             =   1485
      Width           =   2490
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7215
      Left            =   225
      TabIndex        =   1
      Top             =   1935
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   12726
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      BorderStyle     =   0
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   14
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   2
      Top             =   855
      Width           =   14955
      _Version        =   524288
      _ExtentX        =   26379
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
      TabIndex        =   3
      Top             =   9405
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
      Picture         =   "PObeli_BR.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   5445
      TabIndex        =   6
      Top             =   9405
      Visible         =   0   'False
      Width           =   1155
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
      TabIndex        =   5
      Top             =   1170
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PO Pembelian Barang"
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
      Left            =   720
      TabIndex        =   4
      Top             =   180
      Width           =   7755
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   15435
      Picture         =   "PObeli_BR.frx":6862
      Stretch         =   -1  'True
      Top             =   405
      Width           =   285
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   2790
      Picture         =   "PObeli_BR.frx":6C22
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   420
   End
   Begin VB.Image Image1 
      Height          =   10005
      Left            =   0
      Picture         =   "PObeli_BR.frx":13AD2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16260
   End
End
Attribute VB_Name = "PObeli_BR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim color As Long, flag As Byte

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
    SetTransparan1 Me.hwnd, color, 0, flag

    Exit Sub
err: MsgBox err.Description & " Source : " & err.Source
End Sub


Private Sub LG()
On Error GoTo hell

With DataGrid1.Columns(0)
.Width = 120
.Caption = "KODE"
.Alignment = dbgCenter
End With

With DataGrid1.Columns(1)
.Caption = "TANGGAL"
.Width = 80
.Alignment = dbgCenter
End With

With DataGrid1.Columns(2)
.Caption = "kdgudang"
.Width = 0
.Alignment = dbgCenter
End With

With DataGrid1.Columns(3)
.Caption = "GUDANG"
.Width = 200
End With

With DataGrid1.Columns(4)
.Caption = "KETERANGAN"
.Width = 460
End With

With DataGrid1.Columns(5)
.Caption = "NO EASAP"
.Width = 100
.Alignment = dbgCenter
End With



Exit Sub
hell:

End Sub

Private Sub all()



If txtcari = "" Then
sql = "select a.kdPObeli,a.tglPObeli,a.kdgudang,b.nmgudang,a.keterangan,a.noeasap from PObeli a left join gudang b on a.kdgudang=b.kdgudang where  a.kdPObeli in (select kdPO from OTS_PO) order by a.kdPObeli"
Else
sql = "select a.kdPObeli,a.tglPObeli,a.kdgudang,b.nmgudang,a.keterangan,a.noeasap from PObeli a left join gudang b on a.kdgudang=b.kdgudang where (a.kdpobeli like '%" & txtcari & "%' or b.nmgudang like '%" & txtcari & "%' or a.keterangan like '%" & txtcari & "%' or a.noEASAP like '%" & txtcari & "%' ) and a.kdPObeli in (select kdPO from OTS_PO) order by a.kdPObeli"
End If



Set rs = con.Execute(sql)
Set DataGrid1.DataSource = rs
Call LG


End Sub



Private Sub DataGrid1_DblClick()
On Error GoTo hell
If LBLKODE = UCase("beli_D") Then
Beli_D.txtkdPO = rs!kdPObeli
Beli_D.lbltglPO = rs!tglPObeli
Beli_D.lblkdgudang = rs!kdgudang
Beli_D.lblnmgudang = rs!nmgudang
Beli_D.txtketerangan = rs!keterangan
Beli_D.lblnoEASAP = rs!noeasap
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

If KeyAscii = 13 Then
    
   If LBLKODE = UCase("beli_D") Then
    Beli_D.txtkdPO = rs!kdPObeli
    Beli_D.lbltglPO = rs!tglPObeli
    Beli_D.lblkdgudang = rs!kdgudang
    Beli_D.lblnmgudang = rs!nmgudang
    Beli_D.txtketerangan = rs!keterangan
    Beli_D.lblnoEASAP = rs!noeasap
    End If


    Unload Me

ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
txtcari = ""
 Call all
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 txtcari.SetFocus
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
On Error GoTo hell
Call all

TimerALL.Interval = 0

Exit Sub
hell:
MsgBox err.Description
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
    DataGrid1.SetFocus
    Call LG
'    Else
'    CMBCARI.SetFocus
    End If
End If

End Sub

Private Sub TXTCARI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If rs.RecordCount <> 0 Then
    DataGrid1.SetFocus
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





