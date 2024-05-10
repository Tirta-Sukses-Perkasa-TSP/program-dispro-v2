VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form X_Customer_IAP_BR 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   11115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   11115
   ScaleWidth      =   20370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtR 
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
      Left            =   17010
      TabIndex        =   17
      Text            =   "100"
      Top             =   135
      Width           =   735
   End
   Begin VB.TextBox txtcari5 
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
      Left            =   14355
      TabIndex        =   5
      Top             =   630
      Width           =   1905
   End
   Begin VB.TextBox txtcari4 
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
      Left            =   11205
      TabIndex        =   4
      Top             =   630
      Width           =   1905
   End
   Begin VB.TextBox txtcari3 
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
      Left            =   8055
      TabIndex        =   3
      Top             =   630
      Width           =   1905
   End
   Begin VB.TextBox txtcari2 
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
      Left            =   4995
      TabIndex        =   2
      Top             =   630
      Width           =   1905
   End
   Begin VB.Timer TimerG 
      Left            =   7380
      Top             =   0
   End
   Begin VB.Timer TimerALL 
      Left            =   8100
      Top             =   0
   End
   Begin VB.TextBox txtcari1 
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
      Left            =   1575
      TabIndex        =   1
      Top             =   630
      Width           =   1905
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   540
      TabIndex        =   6
      Top             =   495
      Width           =   18735
      _Version        =   524288
      _ExtentX        =   33046
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   10530
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
      Picture         =   "X_Customer_IAP_BR.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   9330
      Left            =   315
      TabIndex        =   0
      Top             =   990
      Width           =   18780
      _cx             =   33126
      _cy             =   16457
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
      HighLight       =   0
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
      FormatString    =   $"X_Customer_IAP_BR.frx":6862
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
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "TAMPILKAN :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   15885
      TabIndex        =   19
      Top             =   180
      Width           =   1185
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "RECORD"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   17775
      TabIndex        =   18
      Top             =   180
      Width           =   1185
   End
   Begin VB.Label lblkdsp 
      Caption         =   "Label7"
      Height          =   330
      Left            =   9900
      TabIndex        =   16
      Top             =   135
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Label lbldbase 
      Caption         =   "Label7"
      Height          =   285
      Left            =   8775
      TabIndex        =   15
      Top             =   135
      Visible         =   0   'False
      Width           =   780
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Alamat :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   13545
      TabIndex        =   14
      Top             =   675
      Width           =   1500
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "NM Cust :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   10305
      TabIndex        =   13
      Top             =   675
      Width           =   1500
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "SP IAP :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   7335
      TabIndex        =   12
      Top             =   675
      Width           =   1500
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cabang IAP :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   3870
      TabIndex        =   11
      Top             =   675
      Width           =   1500
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   5490
      TabIndex        =   10
      Top             =   45
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Kode Cust :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   465
      Left            =   495
      TabIndex        =   9
      Top             =   675
      Width           =   1500
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer IAP"
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
      Height          =   690
      Left            =   1215
      TabIndex        =   8
      Top             =   0
      Width           =   5280
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   19260
      Picture         =   "X_Customer_IAP_BR.frx":69A6
      Stretch         =   -1  'True
      Top             =   405
      Width           =   285
   End
   Begin VB.Image Image1 
      Height          =   11130
      Left            =   0
      Picture         =   "X_Customer_IAP_BR.frx":6D66
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20355
   End
End
Attribute VB_Name = "X_Customer_IAP_BR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim kt_cari1, kt_cari2, kt_cari3, kt_cari4, kt_cari5 As String
Dim color As Long, flag As Byte
Dim sqlL As String
Dim rsQ As ADODB.Recordset

Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub



Private Sub datagrid1_GotFocus()
datagrid1.HighLight = flexHighlightAlways
End Sub

Private Sub datagrid1_LostFocus()
datagrid1.HighLight = flexHighlightNever
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
On Error GoTo hell



Exit Sub
hell:

End Sub

Private Sub all()
On Error GoTo hell

Dim filename As String
Dim Exel_ODC As String
Dim nmview As String

If txtcari1 = "" And txtcari2 = "" And txtcari3 = "" And txtcari4 = "" And txtcari5 = "" And LBLKODE <> "X_RPT_A2" Then
SetTimer hWnd, NV_CLOSEMSGBOX, 2000, AddressOf TimerProc
MsgBox "Keyword Pencarian tidak boleh semua kosong ", vbExclamation, "Error !"
Exit Sub
End If

If txtcari1 = "" Then
kt_cari1 = "a.kdcust_iap <> '@@@@@'"
Else
kt_cari1 = "a.kdcust_iap like '%" & txtcari1 & "%'"
End If

If txtcari2 = "" Then
kt_cari2 = "a.nmplant <> '@@@@@'"
Else
kt_cari2 = "a.nmplant like '%" & txtcari2 & "%'"
End If

If txtcari3 = "" Then
kt_cari3 = "a.Spointdesc <> '@@@@@'"
Else
kt_cari3 = "a.SpointDesc like '%" & txtcari3 & "%'"
End If

If txtcari4 = "" Then
kt_cari4 = "a.custnm <> '@@@@@'"
Else
kt_cari4 = "a.custnm like '%" & txtcari4 & "%'"
End If

If txtcari5 = "" Then
kt_cari5 = "a.addr1 <> '@@@@@'"
Else
kt_cari5 = "a.addr1 like '%" & txtcari5 & "%'"
End If

If LBLKODE = "X_RPT_A2" Then

sqlQ = "select * from User_m where kduser='" & UTAMA.lblkduser & "'"
Set rsQ = con.Execute(sqlQ)

filename = rsQ!alamat_save & "\Kon_rpt.ini"
Exel_ODC = ReadINI("Kon_RPT", "Exel_ODC", filename)
nmview = ReadINI("Kon_RPT", "nmview", filename)
list_Cust = ReadINI("Kon_RPT", "LIST_CUST", filename)

sqlL = "select kdcust_iap from openrowset('microsoft.jet.OLEDB.4.0','Excel 8.0;database=" & list_Cust & X_Rpt_A2.txtnmfile & "','select * from [" & X_Rpt_A2.txtnmSheet & "$]') group by kdcust_iap"

sql = "select TOP " & CLng(txtR) & " a.* from (" & sqlL & ") b left join " & lbldbase & "..V_Mcust_iap" & " a on a.kdcust_iap=b.kdcust_iap where " & kt_cari1 & " and " & kt_cari2 & " and " & kt_cari3 & " and " & kt_cari4 & " and " & kt_cari5 & " order by a.custnm"
Else
sql = "select TOP " & CLng(txtR) & " a.* from " & lbldbase & "..V_Mcust_iap" & " a where " & kt_cari1 & " and " & kt_cari2 & " and " & kt_cari3 & " and " & kt_cari4 & " and " & kt_cari5 & " order by a.custnm"
End If




Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs
Call LG

If rs.RecordCount = 0 Then
SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
MsgBox "Data Yg dicari tidak Ada !", vbCritical, "Error !"
End If


Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub



Private Sub datagrid1_DblClick()
On Error GoTo hell
If LBLKODE = "RPT_A" Then
Rpt_A.lblsalespointcd = rs!salespointcd
Rpt_A.lblnmsp = rs!spointdesc
Rpt_A.cmbcabang.Text = rs!plantcd
Rpt_A.lblkdsp = Left(rs!kdcust_IAP, 8)

Rpt_A.lblnmCust_IAP = rs!custnm
Rpt_A.lblkdcust_IAP = rs!kdcust_IAP
Rpt_A.lblalamat_IAP = rs!addr1

ElseIf LBLKODE = "CUSTOMER_TU" Then
Customer_TU.lblkdsp = Left(rs!kdcust_IAP, 8)
Customer_TU.txtkdcustomer_IAP = rs!custcd

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
    TXTCARI.SetFocus
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
    
    If LBLKODE = "RPT_A" Then
    Rpt_A.lblsalespointcd = rs!salespointcd
    Rpt_A.lblnmsp = rs!spointdesc
    Rpt_A.cmbcabang.Text = rs!plantcd
    Rpt_A.lblkdsp = Left(rs!kdcust_IAP, 8)
    
    Rpt_A.lblnmCust_IAP = rs!custnm
    Rpt_A.lblkdcust_IAP = rs!kdcust_IAP
    Rpt_A.lblalamat_IAP = rs!addr1
    
    ElseIf LBLKODE = "CUSTOMER_TU" Then
    Customer_TU.lblkdsp = Left(rs!kdcust_IAP, 8)
    Customer_TU.txtkdcustomer_IAP = rs!custcd


    End If


    Unload Me

ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = Asc("r") Or KeyAscii = Asc("R") Then
TXTCARI = ""
 TimerALL.Interval = 10
ElseIf KeyAscii = Asc("c") Or KeyAscii = Asc("C") Then
 TXTCARI.SetFocus
End If

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"


End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then

    If LBLKODE = "X_RPT_A4" Then
      
        X_Rpt_A4.txtcari1 = txtcari1
        X_Rpt_A4.txtcari2 = txtcari2
        X_Rpt_A4.txtcari3 = txtcari3
        X_Rpt_A4.txtcari4 = txtcari4
        X_Rpt_A4.txtcari5 = txtcari5
     
    End If


Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub Form_Load()
GradientForm Me, 0



TimerALL.Interval = 10
End Sub






Private Sub TimerAll_Timer()
On Error Resume Next

MousePointer = vbHourglass

Call all

TimerALL.Interval = 0

MousePointer = vbDefault
End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
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


Private Sub txtcari1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TimerALL.Interval = 10
End If
End Sub

Private Sub txtcari2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TimerALL.Interval = 10
End If

End Sub


Private Sub txtcari3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TimerALL.Interval = 10
End If
End Sub

Private Sub txtcari4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TimerALL.Interval = 10
End If

End Sub

Private Sub txtcari5_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TimerALL.Interval = 10
End If

End Sub

Private Sub txtR_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
TimerALL.Interval = 10
End If
End Sub
