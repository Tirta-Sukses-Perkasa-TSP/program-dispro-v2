VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Cetak_7A5 
   BorderStyle     =   0  'None
   ClientHeight    =   3075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16710
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   16710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox CMBBLN 
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
      Left            =   1935
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   690
   End
   Begin VB.ComboBox cmbdbase 
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
      Left            =   3510
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1440
      Width           =   1365
   End
   Begin VB.ComboBox cmbDbase2 
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
      Left            =   8865
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1440
      Width           =   1365
   End
   Begin VB.ComboBox cmbbln2 
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
      Left            =   7290
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1440
      Width           =   690
   End
   Begin VB.ComboBox cmbDbase3 
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
      Left            =   14220
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1440
      Width           =   1365
   End
   Begin VB.ComboBox cmbbln3 
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
      Left            =   12645
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1440
      Width           =   690
   End
   Begin VB.ComboBox CMB1 
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
      Left            =   4185
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   990
      Width           =   1500
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
      Top             =   990
      Width           =   1590
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   315
      TabIndex        =   9
      Top             =   810
      Width           =   15405
      _Version        =   524288
      _ExtentX        =   27173
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdGO 
      Height          =   780
      Left            =   15750
      TabIndex        =   8
      ToolTipText     =   "Simpan"
      Top             =   1035
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
      Picture         =   "Cetak_7A5.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1035
      TabIndex        =   10
      Top             =   2520
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
      Picture         =   "Cetak_7A5.frx":38B6
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "OMSET PERIODE :"
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
      TabIndex        =   19
      Top             =   1530
      Width           =   1545
   End
   Begin VB.Label cmdDbase 
      BackStyle       =   0  'Transparent
      Caption         =   "DBASE :"
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
      Left            =   2790
      TabIndex        =   18
      Top             =   1530
      Width           =   1545
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "DBASE :"
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
      Left            =   8145
      TabIndex        =   17
      Top             =   1530
      Width           =   1545
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "OMSET PERIODE :"
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
      Left            =   5760
      TabIndex        =   16
      Top             =   1530
      Width           =   1545
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "DBASE :"
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
      Left            =   13500
      TabIndex        =   15
      Top             =   1530
      Width           =   1545
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "OMSET PERIODE :"
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
      Left            =   11115
      TabIndex        =   14
      Top             =   1530
      Width           =   1545
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "ANALISA :"
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
      Left            =   3375
      TabIndex        =   13
      Top             =   1035
      Width           =   870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Analisa Dispencer dan Showcase"
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
      TabIndex        =   12
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
      TabIndex        =   11
      Top             =   1035
      Width           =   1320
   End
   Begin VB.Image Image1 
      Height          =   3030
      Index           =   0
      Left            =   0
      Picture         =   "Cetak_7A5.frx":A118
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16665
   End
End
Attribute VB_Name = "Cetak_7A5"
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
Dim rsQ As ADODB.Recordset

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

Private Sub ARV1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub



Private Sub Cetak_X1()
On Error Resume Next
Dim filename As String
Dim Exel_ODC As String
Dim nmview As String

sqlQ = "select * from User_m where kduser='" & UTAMA.lblkduser & "'"
Set rsQ = con.Execute(sqlQ)

filename = rsQ!alamat_save & "\Kon_rpt.ini"
Exel_ODC = ReadINI("Kon_RPT", "Exel_ODC", filename)
nmview = ReadINI("Kon_RPT", "nmview", filename)

con.Execute ("drop view " & nmview & "")


sqlA = "select KDCUST_IAP , sum(" & "BlN_" & cmbbln.Text & ") as " & "BlN_" & cmbbln.Text & ",sum(" & "BlN_" & cmbbln2.Text & ") as " & "BlN_" & cmbbln2.Text & ",sum(" & "BlN_" & cmbbln3.Text & ") as " & "BlN_" & cmbbln3.Text & " from (" & vbCrLf & _
        "select KDCUST_IAP,sum(qty) as " & "BlN_" & cmbbln.Text & ",0 AS " & "BlN_" & cmbbln2.Text & ",0 AS " & "BlN_" & cmbbln3.Text & "  from " & CMbDbase & "..v_omset_TSP where kat='GLN' and BLN=" & cmbbln.Text & " group by KDCUST_IAP UNION ALL" & vbCrLf & _
        "select KDCUST_IAP,0 as " & "BlN_" & cmbbln.Text & ",sum(qty) AS " & "BlN_" & cmbbln2.Text & ",0 AS " & "BlN_" & cmbbln3.Text & "  from " & cmbDbase2 & "..v_omset_TSP where kat='GLN' and BLN=" & cmbbln2.Text & " group by KDCUST_IAP UNION ALL" & vbCrLf & _
        "select KDCUST_IAP,0 as " & "BlN_" & cmbbln.Text & ",0 AS " & "BlN_" & cmbbln2.Text & ",sum(qty) AS " & "BlN_" & cmbbln3.Text & "  from " & cmbDbase3 & "..v_omset_TSP where kat='GLN' and BLN=" & cmbbln3.Text & " group by KDCUST_IAP ) X group by kdcust_iap"
        
sql1 = "select a.kdcustomer,sum(case when b.kdkategori in ('04','05') then pjm+swa else 0 end) as DISP_STD,sum(case when b.kdkategori in ('06','07') then pjm+swa else 0 end) as DISP_PORT,sum(case when b.kdkategori in ('10') then pjm+swa else 0 end) as RAK_GLN,sum(pjm) as pjm,sum(swa) as swa,sum(pjm + Swa) as TOTAL_UNIT from (" & vbCrLf & _
           "select a.kdcustomer,b.kdbarang,sum(b.unit) as pjm,0 as swa from pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam where a.tglpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang Union all" & vbCrLf & _
           "select a.kdcustomer,b.kdbarang,-sum(b.unit) as pjm,0 as swa from Rpinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam where a.tglRpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang Union all" & vbCrLf & _
           "select a.kdcustomer,b.kdbarang,0 as pjm,sum(b.unit) as swa from sewa a left join sewa_d b on a.kdsewa=b.kdsewa where a.tglsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang Union all" & vbCrLf & _
           "select a.kdcustomer,b.kdbarang,0 as pjm,-sum(b.unit) as swa from Rsewa a left join Rsewa_d b on a.kdRsewa=b.kdRsewa where a.tglRsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
      ") a left join barang b on a.kdbarang=b.kdbarang where b.kdkategori IN ('04','05','06','07','10') group by a.kdcustomer"
      
sql2 = "select F.SPOINTDESC,A.KDCUSTOMER,C.NMCUSTOMER,C.ALAMAT,C.KDSP + '/' + C.KDCUSTOMER_IAP AS KDCUSTOMER_IAP,G.CUSTNM AS CUSTOMER_IAP,G.ADDR1 as ALAMAT_IAP,D.NMAREAC,E.NMTEKNISI AS CHEKER,F.ASPS,F.ASPM,A.DISP_STD,A.DISP_PORT,A.RAK_GLN,A.PJM,A.SWA,A.TOTAL_UNIT,ISNULL(" & "B.BlN_" & cmbbln.Text & ",0) AS " & "BLN_" & cmbbln.Text & ",ISNULL(" & "B.BLN_" & cmbbln2.Text & ",0) AS " & "BLN_" & cmbbln2.Text & ",ISNULL(" & "B.BlN_" & cmbbln3.Text & ",0) AS " & "BLN_" & cmbbln3.Text & " FROM (" & sql1 & ") A LEFT JOIN CUSTOMER C ON A.KDCUSTOMER=C.KDCUSTOMER LEFT JOIN (" & sqlA & ") B ON C.KDSP + '/' + C.KDCUSTOMER_IAP=B.KDCUST_IAP " & vbCrLf & _
       "LEFT JOIN  AREA_CHEKER D ON C.KDAREAC=D.KDAREAC LEFT JOIN TEKNISI E ON C.KDTEKNISI=E.KDTEKNISI LEFT JOIN " & CMbDbase & "..VSP_IAP F ON C.KDSP=F.KDSP LEFT JOIN " & CMbDbase & "..VCUSTOMER_IAP G ON C.KDSP + '/' + C.KDCUSTOMER_IAP = G.KDCUST_IAP where A.TOTAL_UNIT<>0"
      
sql3 = "select x.*,RATA2 = case when " & "BLN_" & cmbbln.Text & " = 0 and " & "BLN_" & cmbbln2.Text & " <> 0 then  round((" & "BLN_" & cmbbln2.Text & "  + " & "BLN_" & cmbbln3.Text & ") / 2,0) when " & "BLN_" & cmbbln.Text & " = 0 and " & "BLN_" & cmbbln2.Text & " = 0 and " & "BLN_" & cmbbln3.Text & " <> 0 then " & "BLN_" & cmbbln3.Text & " when " & "BLN_" & cmbbln.Text & " + " & "BLN_" & cmbbln2.Text & " + " & "BLN_" & cmbbln3.Text & " = 0 then 0 else round((" & "BLN_" & cmbbln.Text & " + " & "BLN_" & cmbbln2.Text & " + " & "BLN_" & cmbbln3.Text & ") / 3,0) end from (" & sql2 & ") x "
      
sql4 = "select y.*,RASIO=case when PJM<>0 then round(RATA2 / PJM,0) else 0 end from (" & sql3 & ") y "
      
sql = "create View " & nmview & " As " & sql4 & ""


Text1 = sql3

con.Execute (sql)

Shell "" & Exel_ODC & " " & rsQ!alamat_save & "\rpt.odc", vbMaximizedFocus
      
End Sub

Private Sub Cetak_Y1()
On Error Resume Next
Dim filename As String
Dim Exel_ODC As String
Dim nmview As String

sqlQ = "select * from User_m where kduser='" & UTAMA.lblkduser & "'"
Set rsQ = con.Execute(sqlQ)

filename = rsQ!alamat_save & "\Kon_rpt.ini"
Exel_ODC = ReadINI("Kon_RPT", "Exel_ODC", filename)
nmview = ReadINI("Kon_RPT", "nmview", filename)

con.Execute ("drop view " & nmview & "")


sqlA = "select KDCUST_IAP , sum(" & "BlN_" & cmbbln.Text & ") as " & "BlN_" & cmbbln.Text & ",sum(" & "BlN_" & cmbbln2.Text & ") as " & "BlN_" & cmbbln2.Text & ",sum(" & "BlN_" & cmbbln3.Text & ") as " & "BlN_" & cmbbln3.Text & " from (" & vbCrLf & _
        "select KDCUST_IAP,sum(qty) as " & "BlN_" & cmbbln.Text & ",0 AS " & "BlN_" & cmbbln2.Text & ",0 AS " & "BlN_" & cmbbln3.Text & "  from " & CMbDbase & "..v_omset_TSP where kat in ('CUP','BTL') and BLN=" & cmbbln.Text & " group by KDCUST_IAP UNION ALL" & vbCrLf & _
        "select KDCUST_IAP,0 as " & "BlN_" & cmbbln.Text & ",sum(qty) AS " & "BlN_" & cmbbln2.Text & ",0 AS " & "BlN_" & cmbbln3.Text & "  from " & cmbDbase2 & "..v_omset_TSP where kat in ('CUP','BTL') and BLN=" & cmbbln2.Text & " group by KDCUST_IAP UNION ALL" & vbCrLf & _
        "select KDCUST_IAP,0 as " & "BlN_" & cmbbln.Text & ",0 AS " & "BlN_" & cmbbln2.Text & ",sum(qty) AS " & "BlN_" & cmbbln3.Text & "  from " & cmbDbase3 & "..v_omset_TSP where kat in ('CUP','BTL') and BLN=" & cmbbln3.Text & " group by KDCUST_IAP ) X group by kdcust_iap"
        
sql1 = "select a.kdcustomer,sum(case when b.kdkategori in ('08') then pjm else 0 end) as SHW_KECIL,sum(case when b.kdkategori in ('09') then pjm else 0 end) as SHW_BESAR,sum(pjm + Swa) as TOTAL_UNIT from (" & vbCrLf & _
           "select a.kdcustomer,b.kdbarang,sum(b.unit) as pjm,0 as swa from pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam where a.tglpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang Union all" & vbCrLf & _
           "select a.kdcustomer,b.kdbarang,-sum(b.unit) as pjm,0 as swa from Rpinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam where a.tglRpinjam <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1 group by a.kdcustomer,b.kdbarang Union all" & vbCrLf & _
           "select a.kdcustomer,b.kdbarang,0 as pjm,sum(b.unit) as swa from sewa a left join sewa_d b on a.kdsewa=b.kdsewa where a.tglsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "'  group by a.kdcustomer,b.kdbarang Union all" & vbCrLf & _
           "select a.kdcustomer,b.kdbarang,0 as pjm,-sum(b.unit) as swa from Rsewa a left join Rsewa_d b on a.kdRsewa=b.kdRsewa where a.tglRsewa <= '" & Format(txttgl1, "yyyy/MM/dd") & "' and a.rtr=1  group by a.kdcustomer,b.kdbarang" & vbCrLf & _
      ") a left join barang b on a.kdbarang=b.kdbarang where b.kdkategori IN ('08','09') group by a.kdcustomer"
      
sql2 = "select F.SPOINTDESC,A.KDCUSTOMER,C.NMCUSTOMER,C.ALAMAT,C.KDSP + '/' + C.KDCUSTOMER_IAP AS KDCUSTOMER_IAP,G.CUSTNM AS CUSTOMER_IAP,G.ADDR1 as ALAMAT_IAP,D.NMAREAC,E.NMTEKNISI AS CHEKER,F.ASPS,F.ASPM,A.SHW_KECIL,A.SHW_BESAR,A.TOTAL_UNIT,ISNULL(" & "B.BlN_" & cmbbln.Text & ",0) AS " & "BLN_" & cmbbln.Text & ",ISNULL(" & "B.BLN_" & cmbbln2.Text & ",0) AS " & "BLN_" & cmbbln2.Text & ",ISNULL(" & "B.BlN_" & cmbbln3.Text & ",0) AS " & "BLN_" & cmbbln3.Text & " FROM (" & sql1 & ") A LEFT JOIN CUSTOMER C ON A.KDCUSTOMER=C.KDCUSTOMER LEFT JOIN (" & sqlA & ") B ON C.KDSP + '/' + C.KDCUSTOMER_IAP=B.KDCUST_IAP " & vbCrLf & _
       "LEFT JOIN  AREA_CHEKER D ON C.KDAREAC=D.KDAREAC LEFT JOIN TEKNISI E ON C.KDTEKNISI=E.KDTEKNISI LEFT JOIN " & CMbDbase & "..VSP_IAP F ON C.KDSP=F.KDSP LEFT JOIN " & CMbDbase & "..VCUSTOMER_IAP G ON C.KDSP + '/' + C.KDCUSTOMER_IAP = G.KDCUST_IAP where A.TOTAL_UNIT<>0"
      
sql3 = "select x.*,RATA2 = case when " & "BLN_" & cmbbln.Text & " = 0 and " & "BLN_" & cmbbln2.Text & " <> 0 then  round((" & "BLN_" & cmbbln2.Text & "  + " & "BLN_" & cmbbln3.Text & ") / 2,0) when " & "BLN_" & cmbbln.Text & " = 0 and " & "BLN_" & cmbbln2.Text & " = 0 and " & "BLN_" & cmbbln3.Text & " <> 0 then " & "BLN_" & cmbbln3.Text & " when " & "BLN_" & cmbbln.Text & " + " & "BLN_" & cmbbln2.Text & " + " & "BLN_" & cmbbln3.Text & " = 0 then 0 else round((" & "BLN_" & cmbbln.Text & " + " & "BLN_" & cmbbln2.Text & " + " & "BLN_" & cmbbln3.Text & ") / 3,0) end from (" & sql2 & ") x "
      
sql4 = "select y.*,RASIO=case when TOTAL_UNIT<>0 then round(RATA2 / total_unit,0) else 0 end from (" & sql3 & ") y "
      
sql = "create View " & nmview & " As " & sql4 & ""

con.Execute (sql)

Shell "" & Exel_ODC & " " & rsQ!alamat_save & "\rpt.odc", vbMaximizedFocus
      
End Sub






Private Sub cmdGO_Click()
If CMB1.ListIndex = 0 Then
Call Cetak_X1
Else
Call Cetak_Y1
End If


End Sub

Private Sub cmdGO_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0



CMB1.AddItem "DISPENCER"
CMB1.AddItem "SHOWCASE"
CMB1.ListIndex = 0

sql = "Select * from Dbase_RPT order by urutan"
Set rs = con.Execute(sql)

rs.MoveFirst

Do While Not rs.EOF
CMbDbase.AddItem rs!nmDbase
cmbDbase2.AddItem rs!nmDbase
cmbDbase3.AddItem rs!nmDbase
rs.MoveNext
Loop

CMbDbase.ListIndex = 0
cmbDbase2.ListIndex = 0
cmbDbase3.ListIndex = 0

cmbbln.AddItem "1"
cmbbln.AddItem "2"
cmbbln.AddItem "3"
cmbbln.AddItem "4"
cmbbln.AddItem "5"
cmbbln.AddItem "6"
cmbbln.AddItem "7"
cmbbln.AddItem "8"
cmbbln.AddItem "9"
cmbbln.AddItem "10"
cmbbln.AddItem "11"
cmbbln.AddItem "12"


cmbbln2.AddItem "1"
cmbbln2.AddItem "2"
cmbbln2.AddItem "3"
cmbbln2.AddItem "4"
cmbbln2.AddItem "5"
cmbbln2.AddItem "6"
cmbbln2.AddItem "7"
cmbbln2.AddItem "8"
cmbbln2.AddItem "9"
cmbbln2.AddItem "10"
cmbbln2.AddItem "11"
cmbbln2.AddItem "12"


cmbbln3.AddItem "1"
cmbbln3.AddItem "2"
cmbbln3.AddItem "3"
cmbbln3.AddItem "4"
cmbbln3.AddItem "5"
cmbbln3.AddItem "6"
cmbbln3.AddItem "7"
cmbbln3.AddItem "8"
cmbbln3.AddItem "9"
cmbbln3.AddItem "10"
cmbbln3.AddItem "11"
cmbbln3.AddItem "12"

If Month(Date) > 1 Then
cmbbln3.ListIndex = CLng(Month(Date)) - 1
ElseIf Month(Date) = 1 Then
cmbbln3.ListIndex = 11
End If



If Month(Date) > 2 Then
cmbbln2.ListIndex = CLng(Month(Date)) - 2
ElseIf Month(Date) = 2 Then
cmbbln2.ListIndex = 11
ElseIf Month(Date) = 1 Then
cmbbln2.ListIndex = 10
End If


If Month(Date) > 3 Then
cmbbln.ListIndex = CLng(Month(Date)) - 3
ElseIf Month(Date) = 3 Then
cmbbln.ListIndex = 11
ElseIf Month(Date) = 2 Then
cmbbln.ListIndex = 10
ElseIf Month(Date) = 1 Then
cmbbln.ListIndex = 9
End If



txttgl1 = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
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

