VERSION 5.00
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{0AFE7BE0-11B7-4A3E-978D-D4501E9A57FE}#1.0#0"; "c1sizer.ocx"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Copy_SPP 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   10230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   20370
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   10230
   ScaleWidth      =   20370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk1 
      BackColor       =   &H00000000&
      Caption         =   "LANGSUNG"
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
      Height          =   285
      Left            =   17145
      TabIndex        =   15
      Top             =   585
      Width           =   1860
   End
   Begin VB.Timer TimerALL 
      Left            =   6030
      Top             =   1665
   End
   Begin VB.Timer TimerG 
      Left            =   5535
      Top             =   1665
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   495
      TabIndex        =   0
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
      TabIndex        =   1
      Top             =   9540
      Width           =   3120
      _ExtentX        =   5503
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
      Picture         =   "Copy_SPP.frx":0000
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdT 
      Height          =   870
      Index           =   0
      Left            =   19305
      TabIndex        =   2
      ToolTipText     =   "Copy File SPP"
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
      Picture         =   "Copy_SPP.frx":6862
      Alignment       =   1
      ButtonStyle     =   4
   End
   Begin VSFlex8UCtl.VSFlexGrid datagrid1 
      Height          =   8430
      Left            =   270
      TabIndex        =   3
      Top             =   945
      Width           =   18870
      _cx             =   33285
      _cy             =   14870
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
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Copy_SPP.frx":B275
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
      Begin VB.Timer TimerCopy 
         Left            =   8640
         Top             =   1080
      End
   End
   Begin Threed.SSCommand cmdBF 
      Height          =   375
      Left            =   16380
      TabIndex        =   8
      ToolTipText     =   "Menuju Ke Folder SPP"
      Top             =   585
      Width           =   375
      _ExtentX        =   661
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
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Copy_SPP.frx":B339
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdLC 
      Height          =   375
      Left            =   9090
      TabIndex        =   10
      ToolTipText     =   "Menuju Ke Folder SPP"
      Top             =   585
      Width           =   375
      _ExtentX        =   661
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
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Copy_SPP.frx":11B9B
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   8550
      Top             =   45
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Threed.SSCommand cmdHps 
      Height          =   375
      Left            =   9495
      TabIndex        =   13
      ToolTipText     =   "Hapus SPP"
      Top             =   585
      Width           =   375
      _ExtentX        =   661
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
         Size            =   5.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "Copy_SPP.frx":183FD
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin C1SizerLibCtl.C1Elastic flood 
      Height          =   420
      Left            =   4140
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9540
      Visible         =   0   'False
      Width           =   14970
      _cx             =   26405
      _cy             =   741
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   0
      MousePointer    =   0
      Version         =   801
      BackColor       =   16777215
      ForeColor       =   255
      FloodColor      =   16711680
      ForeColorDisabled=   -2147483631
      Caption         =   ""
      Align           =   0
      AutoSizeChildren=   0
      BorderWidth     =   0
      ChildSpacing    =   4
      Splitter        =   0   'False
      FloodDirection  =   1
      FloodPercent    =   0
      CaptionPos      =   4
      WordWrap        =   -1  'True
      MaxChildSize    =   0
      MinChildSize    =   0
      TagWidth        =   0
      TagPosition     =   0
      Style           =   0
      TagSplit        =   2
      PicturePos      =   4
      CaptionStyle    =   0
      ResizeFonts     =   0   'False
      GridRows        =   0
      GridCols        =   0
      Frame           =   3
      FrameStyle      =   2
      FrameWidth      =   1
      FrameColor      =   -2147483628
      FrameShadow     =   -2147483632
      FloodStyle      =   1
      _GridInfo       =   ""
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   9
   End
   Begin VB.Label lbljml 
      Caption         =   "Label2"
      Height          =   330
      Left            =   5220
      TabIndex        =   16
      Top             =   180
      Width           =   645
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LIST CUSTOMER SPP YG AKAN DI COPY :"
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
      Left            =   585
      TabIndex        =   12
      Top             =   630
      Width           =   3165
   End
   Begin VB.Label lblLC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   3780
      TabIndex        =   11
      Top             =   585
      Width           =   5280
   End
   Begin VB.Label lblFC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   11070
      TabIndex        =   9
      Top             =   585
      Width           =   5280
   End
   Begin VB.Image Image4 
      Height          =   345
      Left            =   19215
      Picture         =   "Copy_SPP.frx":1EC5F
      Stretch         =   -1  'True
      Top             =   270
      Width           =   285
   End
   Begin VB.Label label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copy SPP"
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
      TabIndex        =   7
      Top             =   0
      Width           =   5280
   End
   Begin VB.Label LBLKODE 
      Caption         =   "lblkode"
      Height          =   315
      Left            =   5670
      TabIndex        =   6
      Top             =   9450
      Visible         =   0   'False
      Width           =   1155
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
      Left            =   9720
      TabIndex        =   5
      Top             =   90
      Width           =   4650
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "COPY KE :"
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
      Left            =   10080
      TabIndex        =   4
      Top             =   630
      Width           =   960
   End
   Begin VB.Image Image1 
      Height          =   10185
      Left            =   0
      Picture         =   "Copy_SPP.frx":1F01F
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20355
   End
End
Attribute VB_Name = "Copy_SPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim kategori As String
Dim color As Long, flag As Byte
Dim fso As New FileSystemObject
Dim kata As String
Dim sqlL As String
Dim rsL As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim filename As String
Dim alamat_SPP As String
Dim i As Integer



Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub cmdBF_Click()
Dim StrFolderBrowse
StrFolderBrowse = fBrowseForFolder(hWnd, "Pilih Folder")
If StrFolderBrowse <> vbNullString Then
   lblfC = StrFolderBrowse
End If

End Sub

Private Sub cmdHps_Click()
lblLC = ""
End Sub

Private Sub cmdLC_Click()
filename = App.Path & "\Koneksi.ini"
alamat_SPP = ReadINI("Koneksi", "alamat_SPP", filename)

CD1.Filter = "(*.xls;*.xlsx)|*.xls;*.xlsx"
CD1.ShowOpen
    
lblLC = CD1.filename

If Right(lblLC, 4) <> ".xls" Then
    If fso.FileExists(alamat_SPP & "/List_spp.xlsx") = True Then
    Call fso.DeleteFile(alamat_SPP & "/List_spp.xlsx", True)
    End If
    
    Call fso.CopyFile(lblLC, alamat_SPP & "/List_spp.xlsx", True)
    
Else

    If fso.FileExists(alamat_SPP & "List_spp.xls") = True Then
    Call fso.DeleteFile(alamat_SPP & "List_spp.xls", True)
    End If
    
    Call fso.CopyFile(lblLC, alamat_SPP & "/List_spp.xls", True)
End If

 

End Sub

Private Sub cmdT_Click(Index As Integer)
If lblfC <> "" Then
flood.Visible = True
lbljml = 0
TimerCopy.Interval = 10
Else
SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
MsgBox "Mohon Untuk Folder tujuan Copy diisi dulu !!", vbInformation, "Info !"
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
On Error GoTo hell

Dim i, j As Integer

MousePointer = vbHourglass

filename = App.Path & "\Koneksi.ini"
alamat_SPP = ReadINI("Koneksi", "alamat_SPP", filename)


lblalamat_SPP = alamat_SPP

If lblLC <> "" Then
    
    If Right(lblLC, 4) <> ".xls" Then
    sqlL = "select kdcustomer from openrowset('Microsoft.ACE.OLEDB.12.0','Excel 12.0 Xml;HDR=YES;database=" & alamat_SPP & "List_spp.xlsx" & " ','select * from [sheet1$]') group by kdcustomer"
    Else
    sqlL = "select kdcustomer from openrowset('Microsoft.ACE.OLEDB.12.0','Excel 12.0 Xml;HDR=YES;database=" & alamat_SPP & "List_spp.xls" & " ','select * from [sheet1$]') group by kdcustomer"
    End If
    
    Set rsL = con.Execute(sqlL)
        
    sql = "SELECT a.kdcustomer,a.nmcustomer,a.alamat,a.keterangan,b.kdmars FROM customer a left join kddispromars b on a.kdcustomer=b.kdcustomer where a.kdcustomer in (" & sqlL & ")"
    
    Set rs = con.Execute(sql)
    Set datagrid1.DataSource = rs
    
    If rs.RecordCount <> 0 Then
        For i = 1 To (datagrid1.Rows - 1)
    
        datagrid1.TextMatrix(i, 0) = i
        
        Next
    End If
      
    sql1 = "select count(kdcustomer) as jmlcust from (" & sql & ") x"
    Set rs1 = con.Execute(sql1)
Else
datagrid1.Clear
    
End If


MousePointer = vbDefault
Exit Sub
hell:
SetTimer hWnd, NV_CLOSEMSGBOX, 1000, AddressOf TimerProc
MsgBox err.Description, vbCritical, "Error !!"
MousePointer = vbDefault
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
Call nul(lblfC)

TimerALL.Interval = 10
End Sub






Private Sub lblFC_Change()
Call nul(lblfC)
End Sub

Private Sub lblLC_Change()
TimerALL.Interval = 10
End Sub

Private Sub TimerAll_Timer()
Call all

TimerALL.Interval = 0

End Sub

Private Sub TimerCopy_Timer()
Static Z As Integer

Z = Z + 1


If Z > CLng(rs1!jmlcust) Then
    Z = 0
    
    TimerCopy.Interval = 0
    
    SetTimer hWnd, NV_CLOSEMSGBOX, 3000, AddressOf TimerProc
    MsgBox "Tercopy " & lbljml & "/" & rs1!jmlcust & " Folder", vbInformation, "Info !"
    
    
    cmdT(0).Enabled = True
    flood.Visible = False
    
    TimerALL.Interval = 10
    
    

Else
     flood.FloodPercent = (Z / CLng(rs1!jmlcust)) * 100
     flood.Caption = "Proses Copy : " & FormatNumber((Z / CLng(rs1!jmlcust)) * 100, 1) & "%"
     rs.AbsolutePosition = Z
     
     If fso.FolderExists(lblalamat_SPP & rs!keterangan) = True Then
        
        If Chk1.Value = 1 Then
        Call fso.CopyFolder(lblalamat_SPP & rs!keterangan, lblfC & "/" & rs!keterangan, True)
        Else
             
            If fso.FolderExists(lblfC & "/" & rs!kdmars) = False Then
            Call fso.CreateFolder(lblfC & "/" & rs!kdmars)
            End If
            
            Call fso.CopyFolder(lblalamat_SPP & rs!keterangan, lblfC & "/" & rs!kdmars & "/" & rs!keterangan, True)
            
        End If
       lbljml = lbljml + 1
     End If
    
    
End If


End Sub

Private Sub TimerG_Timer()
Call LG
TimerG.Interval = 0
End Sub







