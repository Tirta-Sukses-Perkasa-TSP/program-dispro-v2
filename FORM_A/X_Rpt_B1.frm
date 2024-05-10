VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form X_Rpt_B1 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11400
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   6180
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChKDetail 
      BackColor       =   &H00000000&
      Caption         =   "DETAIL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   8685
      MaskColor       =   &H00000000&
      TabIndex        =   12
      Top             =   2610
      Width           =   1365
   End
   Begin VB.OptionButton OAIBM1 
      BackColor       =   &H00000000&
      Caption         =   "Customer Putus (Not Buy)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   315
      TabIndex        =   16
      Top             =   4500
      Width           =   2625
   End
   Begin VB.OptionButton OAIBM2 
      BackColor       =   &H00000000&
      Caption         =   "Analisa Omset"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   3060
      TabIndex        =   17
      Top             =   4500
      Width           =   1635
   End
   Begin VB.OptionButton OAIBM3 
      BackColor       =   &H00000000&
      Caption         =   "Customer Baru (Buy)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   4770
      TabIndex        =   18
      Top             =   4500
      Width           =   2355
   End
   Begin VB.OptionButton OTSP3 
      BackColor       =   &H00000000&
      Caption         =   "Customer Baru (Buy)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   4770
      TabIndex        =   15
      Top             =   3600
      Width           =   2355
   End
   Begin VB.ComboBox cmbkat1 
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
      Left            =   2925
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1755
      Width           =   1545
   End
   Begin VB.ComboBox cmbkat 
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
      TabIndex        =   5
      Top             =   1755
      Width           =   1545
   End
   Begin VB.OptionButton OTSP2 
      BackColor       =   &H00000000&
      Caption         =   "Analisa Omset"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   3060
      TabIndex        =   14
      Top             =   3600
      Width           =   1635
   End
   Begin VB.OptionButton OTSP1 
      BackColor       =   &H00000000&
      Caption         =   "Customer Putus (Not Buy)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   315
      TabIndex        =   13
      Top             =   3600
      Width           =   2625
   End
   Begin VB.TextBox txttglA1 
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
      TabIndex        =   8
      Top             =   2610
      Width           =   1365
   End
   Begin VB.TextBox txttglA2 
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
      Left            =   2880
      TabIndex        =   9
      Top             =   2610
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
      Left            =   7110
      TabIndex        =   11
      Top             =   2610
      Width           =   1365
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
      Left            =   5445
      TabIndex        =   10
      Top             =   2610
      Width           =   1365
   End
   Begin VB.Timer TimerPdf 
      Left            =   8775
      Top             =   1575
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
   Begin VB.ComboBox cmbcabang 
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
      Left            =   6255
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   765
      Width           =   1050
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
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   765
      Width           =   1500
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   495
      TabIndex        =   20
      Top             =   630
      Width           =   9015
      _Version        =   524288
      _ExtentX        =   15901
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdGO 
      Height          =   780
      Left            =   10260
      TabIndex        =   21
      ToolTipText     =   "Simpan"
      Top             =   2250
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
      Picture         =   "X_Rpt_B1.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   540
      TabIndex        =   19
      Top             =   5490
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
      Picture         =   "X_Rpt_B1.frx":38B6
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdBR1 
      Height          =   420
      Left            =   5265
      TabIndex        =   3
      Top             =   1215
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   741
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
      Picture         =   "X_Rpt_B1.frx":A118
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdC1 
      Height          =   420
      Left            =   5760
      TabIndex        =   4
      Top             =   1215
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   741
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
      Picture         =   "X_Rpt_B1.frx":C94A
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdbr2 
      Height          =   420
      Left            =   9360
      TabIndex        =   7
      Top             =   1710
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   741
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
      Picture         =   "X_Rpt_B1.frx":EF94
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "AIBM"
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
      Left            =   270
      TabIndex        =   42
      Top             =   4185
      Width           =   960
   End
   Begin VB.Label lblitemnm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6435
      TabIndex        =   41
      Top             =   1755
      Width           =   2940
   End
   Begin VB.Label lblitemcd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   5310
      TabIndex        =   40
      Top             =   1755
      Width           =   1095
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM :"
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
      Left            =   4770
      TabIndex        =   39
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "KAT BRG :"
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
      Left            =   360
      TabIndex        =   38
      Top             =   1800
      Width           =   870
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "TSP"
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
      Left            =   270
      TabIndex        =   37
      Top             =   3285
      Width           =   960
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Pembanding :"
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
      Left            =   135
      TabIndex        =   36
      Top             =   2655
      Width           =   1185
   End
   Begin VB.Label Label6 
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   270
      TabIndex        =   35
      Top             =   2295
      Width           =   960
   End
   Begin VB.Label Label3 
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
      Left            =   2520
      TabIndex        =   34
      Top             =   2655
      Width           =   420
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
      Left            =   6750
      TabIndex        =   33
      Top             =   2655
      Width           =   420
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaksi:"
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
      Left            =   4455
      TabIndex        =   32
      Top             =   2655
      Width           =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Report Per Customer"
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
      Left            =   855
      TabIndex        =   31
      Top             =   45
      Width           =   4560
   End
   Begin VB.Label lblnmcabang 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   7335
      TabIndex        =   30
      Top             =   765
      Width           =   2040
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
      TabIndex        =   29
      Top             =   810
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cabang :"
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
      Left            =   5490
      TabIndex        =   28
      Top             =   810
      Width           =   735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "SP IAP :"
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
      TabIndex        =   27
      Top             =   1305
      Width           =   735
   End
   Begin VB.Label lblsalespointcd 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1215
      TabIndex        =   26
      Top             =   1260
      Width           =   735
   End
   Begin VB.Label lblnmsp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1980
      TabIndex        =   25
      Top             =   1260
      Width           =   3300
   End
   Begin VB.Label lblkdsp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   8730
      TabIndex        =   24
      Top             =   270
      Visible         =   0   'False
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
      Left            =   2700
      TabIndex        =   23
      Top             =   810
      Width           =   600
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
      Left            =   3285
      TabIndex        =   22
      Top             =   810
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   6090
      Index           =   0
      Left            =   0
      Picture         =   "X_Rpt_B1.frx":117C6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11310
   End
End
Attribute VB_Name = "X_Rpt_B1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rs As ADODB.Recordset
Dim color As Long, flag As Byte
Dim rsQ As ADODB.Recordset
Dim sqlQ As String
Dim rsK As ADODB.Recordset
Dim ket_view As String

Private Sub Cetak_TSP1()
On Error Resume Next
Dim filename As String
Dim Exel_ODC As String
Dim nmview As String
Dim list_Cust As String


If OTSP1.Value = True Then
ket_view = "V_OMSET_TSP"
ElseIf OAIBM1.Value = True Then
ket_view = "V_OMSET_AIBM"
End If

sqlQ = "select * from User_m where kduser='" & UTAMA.lblkduser & "'"
Set rsQ = con.Execute(sqlQ)

filename = rsQ!alamat_save & "\Kon_rpt.ini"
Exel_ODC = ReadINI("Kon_RPT", "Exel_ODC", filename)
nmview = ReadINI("Kon_RPT", "nmview", filename)

con.Execute ("drop view " & nmview & "")

If cmbkat.ListIndex = 0 Then
kata = "itemcd<>'@@@'"
ElseIf cmbkat.ListIndex > 1 And cmbkat1.ListIndex = 0 Then
kata = "kat='" & cmbkat.Text & "'"
ElseIf cmbkat.ListIndex > 1 And cmbkat1.ListIndex <> 0 Then
kata = "kat1='" & cmbkat1.Text & "'"
ElseIf cmbkat.ListIndex = 1 Then
kata = "itemcd='" & lblitemcd & "'"
End If

If cmbcabang.ListIndex = 0 Then
    If CMbDbase.Text = cmBdbase1.Text Then
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where " & kata & ""
    Else
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where " & kata & " union all select * from " & cmBdbase1 & ".." & ket_view & " where " & kata & ""
    End If
ElseIf cmbcabang.ListIndex <> 0 And lblkdsp = "" Then
    If CMbDbase.Text = cmBdbase1.Text Then
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where left(kdcust_iap,4)='" & cmbcabang.Text & "' and " & kata & ""
    Else
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where left(kdcust_iap,4)='" & cmbcabang.Text & "' and " & kata & " union all select * from " & cmBdbase1 & ".." & ket_view & " where left(kdcust_iap,4)='" & cmbcabang.Text & "' and  " & kata & ""
    End If
ElseIf cmbcabang.ListIndex <> 0 And lblkdsp <> "" Then
    If CMbDbase.Text = cmBdbase1.Text Then
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where left(kdcust_iap,8)='" & lblkdsp & "' and " & kata & ""
    Else
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where left(kdcust_iap,8)='" & lblkdsp & "' and " & kata & " union all select * from " & cmBdbase1 & ".." & ket_view & " where left(kdcust_iap,8)='" & lblkdsp & "' and " & kata & ""
    End If
End If


If ChKDetail.Value = 0 Then
sqlA1 = "select KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD ,ASPS,ASPM,COUNT(invnum) as EC,SUM(qty) as QTY from (" & sql1 & ") x where shipdt between '" & Format(txttglA1, "yyyy/MM/dd") & "' and '" & Format(txttglA2, "yyyy/MM/dd") & "' group by KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD ,CUSTNM ,addr1,SLSCD,ASPS,ASPM"
sqlA2 = "select KDCUST_IAP,SUM(qty) as QTY from (" & sql1 & ") x where shipdt between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' group by kdcust_iap"

sqlB = "select a.* ,isnull(b.qty,0) as QTY1 from (" & sqlA1 & ") a left join (" & sqlA2 & ") b on a.kdcust_iap=b.kdcust_Iap"

Else
sqlA1 = "select KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD ,ASPS,ASPM,KAT1,COUNT(invnum) as EC,SUM(qty) as QTY from (" & sql1 & ") x where shipdt between '" & Format(txttglA1, "yyyy/MM/dd") & "' and '" & Format(txttglA2, "yyyy/MM/dd") & "' group by KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD ,CUSTNM ,addr1,SLSCD,ASPS,ASPM,KAT1"
sqlA2 = "select KDCUST_IAP,KAT1,SUM(qty) as QTY from (" & sql1 & ") x where shipdt between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' group by kdcust_iap,KAT1"

sqlB = "select a.* ,isnull(b.qty,0) as QTY1 from (" & sqlA1 & ") a left join (" & sqlA2 & ") b on a.kdcust_iap=b.kdcust_Iap AND A.KAT1=B.KAT1"
End If

sql = "create View " & nmview & " As select * from (" & sqlB & ") x where qty>0 and qty1<=0"

con.Execute (sql)

Shell "" & Exel_ODC & " " & rsQ!alamat_save & "\rpt.odc", vbMaximizedFocus


End Sub

Private Sub Cetak_TSP2()
On Error Resume Next
Dim filename As String
Dim Exel_ODC As String
Dim nmview As String
Dim list_Cust As String

If OTSP2.Value = True Then
ket_view = "V_OMSET_TSP"
ElseIf OAIBM2.Value = True Then
ket_view = "V_OMSET_AIBM"
End If

sqlQ = "select * from User_m where kduser='" & UTAMA.lblkduser & "'"
Set rsQ = con.Execute(sqlQ)

filename = rsQ!alamat_save & "\Kon_rpt.ini"
Exel_ODC = ReadINI("Kon_RPT", "Exel_ODC", filename)
nmview = ReadINI("Kon_RPT", "nmview", filename)

con.Execute ("drop view " & nmview & "")

If cmbkat.ListIndex = 0 Then
kata = "itemcd<>'@@@'"
ElseIf cmbkat.ListIndex > 1 And cmbkat1.ListIndex = 0 Then
kata = "kat='" & cmbkat.Text & "'"
ElseIf cmbkat.ListIndex > 1 And cmbkat1.ListIndex <> 0 Then
kata = "kat1='" & cmbkat1.Text & "'"
ElseIf cmbkat.ListIndex = 1 Then
kata = "itemcd='" & lblitemcd & "'"
End If

If cmbcabang.ListIndex = 0 Then
    If CMbDbase.Text = cmBdbase1.Text Then
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where " & kata & ""
    Else
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where " & kata & " union all select * from " & cmBdbase1 & ".." & ket_view & " where " & kata & ""
    End If
ElseIf cmbcabang.ListIndex <> 0 And lblkdsp = "" Then
    If CMbDbase.Text = cmBdbase1.Text Then
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where left(kdcust_iap,4)='" & cmbcabang.Text & "' and " & kata & ""
    Else
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where left(kdcust_iap,4)='" & cmbcabang.Text & "' and " & kata & " union all select * from " & cmBdbase1 & ".." & ket_view & " where left(kdcust_iap,4)='" & cmbcabang.Text & "' and  " & kata & ""
    End If
ElseIf cmbcabang.ListIndex <> 0 And lblkdsp <> "" Then
    If CMbDbase.Text = cmBdbase1.Text Then
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where left(kdcust_iap,8)='" & lblkdsp & "' and " & kata & ""
    Else
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where left(kdcust_iap,8)='" & lblkdsp & "' and " & kata & " union all select * from " & cmBdbase1 & ".." & ket_view & " where left(kdcust_iap,8)='" & lblkdsp & "' and " & kata & ""
    End If
End If

If ChKDetail.Value = 0 Then
sqlA1 = "select KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD,ASPS,ASPM,COUNT(invnum) as EC,SUM(qty) as QTY,0 as EC1,0 as QTY1 from (" & sql1 & ") x where shipdt between '" & Format(txttglA1, "yyyy/MM/dd") & "' and '" & Format(txttglA2, "yyyy/MM/dd") & "' group by KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD ,CUSTNM ,addr1,SLSCD,ASPS,ASPM"
sqlA2 = "select KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD,ASPS,ASPM,0 AS EC,0 as QTY,COUNT(invnum) as EC1,SUM(qty) as QTY1 from (" & sql1 & ") x where shipdt between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' group by KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD ,CUSTNM ,addr1,SLSCD,ASPS,ASPM"


sqlB1 = "select KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD,ASPS,ASPM,isnull(EC,0) as EC,isnull(qty,0) as QTY,EC1,QTY1 from (" & sqlA1 & ") x "
sqlB2 = "select KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD,ASPS,ASPM,EC,QTY,isnull(EC1,0) as EC1,isnull(qty1,0) as QTY1 from (" & sqlA2 & ") x "

sqlB = "select KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD,ASPS,ASPM,SUM(EC) as EC,SUM(qty) as QTY ,SUM(EC1) as EC1,SUM(qty1) as QTY1 from (" & sqlB1 & " UNION ALL " & sqlB2 & ") X GROUP BY KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD,ASPS,ASPM"

Else

sqlA1 = "select KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD,ASPS,ASPM,KAT1,COUNT(invnum) as EC,SUM(qty) as QTY,0 as EC1,0 as QTY1 from (" & sql1 & ") x where shipdt between '" & Format(txttglA1, "yyyy/MM/dd") & "' and '" & Format(txttglA2, "yyyy/MM/dd") & "' group by KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD ,CUSTNM ,addr1,SLSCD,ASPS,ASPM,KAT1"
sqlA2 = "select KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD,ASPS,ASPM,KAT1,0 AS EC,0 as QTY,COUNT(invnum) as EC1,SUM(qty) as QTY1 from (" & sql1 & ") x where shipdt between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' group by KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD ,CUSTNM ,addr1,SLSCD,ASPS,ASPM,KAT1"


sqlB1 = "select KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD,ASPS,ASPM,KAT1,isnull(EC,0) as EC,isnull(qty,0) as QTY,EC1,QTY1 from (" & sqlA1 & ") x "
sqlB2 = "select KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD,ASPS,ASPM,KAT1,EC,QTY,isnull(EC1,0) as EC1,isnull(qty1,0) as QTY1 from (" & sqlA2 & ") x "

sqlB = "select KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD,ASPS,ASPM,KAT1,SUM(EC) as EC,SUM(qty) as QTY ,SUM(EC1) as EC1,SUM(qty1) as QTY1 from (" & sqlB1 & " UNION ALL " & sqlB2 & ") X GROUP BY KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD,ASPS,ASPM,KAT1"

End If

sql = "create View " & nmview & " As select *, [GROW %] = case when qty<>0 then ((QTY1 - QTY) / QTY) * 100 else 0 end  from (" & sqlB & ") x "

con.Execute (sql)

Shell "" & Exel_ODC & " " & rsQ!alamat_save & "\rpt.odc", vbMaximizedFocus


End Sub


Private Sub Cetak_TSP3()
On Error Resume Next
Dim filename As String
Dim Exel_ODC As String
Dim nmview As String
Dim list_Cust As String

If OTSP3.Value = True Then
ket_view = "V_OMSET_TSP"
ElseIf OAIBM3.Value = True Then
ket_view = "V_OMSET_AIBM"
End If


sqlQ = "select * from User_m where kduser='" & UTAMA.lblkduser & "'"
Set rsQ = con.Execute(sqlQ)

filename = rsQ!alamat_save & "\Kon_rpt.ini"
Exel_ODC = ReadINI("Kon_RPT", "Exel_ODC", filename)
nmview = ReadINI("Kon_RPT", "nmview", filename)

con.Execute ("drop view " & nmview & "")

If cmbkat.ListIndex = 0 Then
kata = "itemcd<>'@@@'"
ElseIf cmbkat.ListIndex > 1 And cmbkat1.ListIndex = 0 Then
kata = "kat='" & cmbkat.Text & "'"
ElseIf cmbkat.ListIndex > 1 And cmbkat1.ListIndex <> 0 Then
kata = "kat1='" & cmbkat1.Text & "'"
ElseIf cmbkat.ListIndex = 1 Then
kata = "itemcd='" & lblitemcd & "'"
End If

If cmbcabang.ListIndex = 0 Then
    If CMbDbase.Text = cmBdbase1.Text Then
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where " & kata & ""
    Else
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where " & kata & " union all select * from " & cmBdbase1 & ".." & ket_view & " where " & kata & ""
    End If
ElseIf cmbcabang.ListIndex <> 0 And lblkdsp = "" Then
    If CMbDbase.Text = cmBdbase1.Text Then
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where left(kdcust_iap,4)='" & cmbcabang.Text & "' and " & kata & ""
    Else
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where left(kdcust_iap,4)='" & cmbcabang.Text & "' and " & kata & " union all select * from " & cmBdbase1 & ".." & ket_view & " where left(kdcust_iap,4)='" & cmbcabang.Text & "' and  " & kata & ""
    End If
ElseIf cmbcabang.ListIndex <> 0 And lblkdsp <> "" Then
    If CMbDbase.Text = cmBdbase1.Text Then
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where left(kdcust_iap,8)='" & lblkdsp & "' and " & kata & ""
    Else
    sql1 = "select * from " & CMbDbase & ".." & ket_view & " where left(kdcust_iap,8)='" & lblkdsp & "' and " & kata & " union all select * from " & cmBdbase1 & ".." & ket_view & " where left(kdcust_iap,8)='" & lblkdsp & "' and " & kata & ""
    End If
End If

If ChKDetail.Value = 0 Then
sqlA1 = "select KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD ,ASPS,ASPM,COUNT(invnum) as EC1,SUM(qty) as QTY1 from (" & sql1 & ") x where shipdt between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' group by KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD ,CUSTNM ,addr1,SLSCD,ASPS,ASPM"
sqlA2 = "select KDCUST_IAP,SUM(qty) as QTY from (" & sql1 & ") x where shipdt between '" & Format(txttglA1, "yyyy/MM/dd") & "' and '" & Format(txttglA2, "yyyy/MM/dd") & "' group by kdcust_iap"

sqlB = "select a.KDCUST_IAP,a.PLANTCD,CABANG ,a.SALESPOINTCD ,a.SPointDesc ,a.CUSTCD,a.CUSTNM ,a.addr1,a.SLSCD ,a.ASPS,a.ASPM,isnull(b.qty,0) as QTY,EC1,QTY1 from (" & sqlA1 & ") a left join (" & sqlA2 & ") b on a.kdcust_iap=b.kdcust_Iap"
Else
sqlA1 = "select KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD,CUSTNM ,addr1,SLSCD ,ASPS,ASPM,KAT1,COUNT(invnum) as EC1,SUM(qty) as QTY1 from (" & sql1 & ") x where shipdt between '" & Format(txttgl1, "yyyy/MM/dd") & "' and '" & Format(txttgl2, "yyyy/MM/dd") & "' group by KDCUST_IAP,PLANTCD,CABANG ,SALESPOINTCD ,SPointDesc ,CUSTCD ,CUSTNM ,addr1,SLSCD,ASPS,ASPM,KAT1"
sqlA2 = "select KDCUST_IAP,KAT1,SUM(qty) as QTY from (" & sql1 & ") x where shipdt between '" & Format(txttglA1, "yyyy/MM/dd") & "' and '" & Format(txttglA2, "yyyy/MM/dd") & "' group by kdcust_iap,KAT1"

sqlB = "select a.KDCUST_IAP,a.PLANTCD,CABANG ,a.SALESPOINTCD ,a.SPointDesc ,a.CUSTCD,a.CUSTNM ,a.addr1,a.SLSCD ,a.ASPS,a.ASPM,a.KAT1,isnull(b.qty,0) as QTY,EC1,QTY1 from (" & sqlA1 & ") a left join (" & sqlA2 & ") b on a.kdcust_iap=b.kdcust_Iap AND a.kat1=b.kat1"

End If


sql = "create View " & nmview & " As select * from (" & sqlB & ") x where qty<=0 and qty1>0"

con.Execute (sql)

Shell "" & Exel_ODC & " " & rsQ!alamat_save & "\rpt.odc", vbMaximizedFocus


End Sub




Private Sub Cmbcabang_Click()
On Error GoTo hell

sql = "Select * from " & CMbDbase & "..VM_plant" & "  where kdplant=" & cmbcabang.Text & " "
Set rs = con.Execute(sql)

If rs.RecordCount <> 0 Then
lblnmcabang = rs!nmplant
Else
lblnmcabang = ""
End If

Exit Sub
hell:
lblnmcabang = ""
lblsalespointcd = ""
lblkdcust_IAP = ""
End Sub

Private Sub CMbDbase_Click()
On Error GoTo hell


cmbcabang.Clear

sql = "Select * from " & CMbDbase & "..VM_plant" & "  order by nmplant"
Set rs = con.Execute(sql)

cmbcabang.AddItem "ALL"

rs.MoveFirst

Do While Not rs.EOF
cmbcabang.AddItem rs!kdplant
rs.MoveNext
Loop

cmbcabang.ListIndex = 0


'item
cmbkat.Clear

sql = "Select kat from " & CMbDbase & "..Vtblitem" & "  group by kat"
Set rs = con.Execute(sql)

cmbkat.AddItem "ALL"
cmbkat.AddItem "PILIH ITEM"

rs.MoveFirst

Do While Not rs.EOF
cmbkat.AddItem rs!kat
rs.MoveNext
Loop

cmbkat.ListIndex = 0



Exit Sub
hell:
cmbcabang.Clear
cmbkat.Clear
lblnmcabang = ""

End Sub


Private Sub cmbkat_Click()
On Error GoTo hell

If cmbkat.ListIndex = 1 Then
cmdbr2.Enabled = True
Else
cmdbr2.Enabled = False
lblitemcd = ""
lblitemnm = ""
End If



cmbkat1.Clear

cmbkat1.AddItem "ALL"

sqlK = "Select kat1 from " & CMbDbase & "..Vtblitem" & " where kat='" & cmbkat.Text & "' group by kat1"
Set rsK = con.Execute(sqlK)

rsK.MoveFirst

Do While Not rsK.EOF
cmbkat1.AddItem rsK!kat1
rsK.MoveNext
Loop

cmbkat1.ListIndex = 0



Exit Sub
hell:
cmbkat1.Clear
End Sub

Private Sub cmdBR1_Click()
X_SPIAP_BR.LBLKODE = "X_RPT_B1"
X_SPIAP_BR.lblkdcabang = cmbcabang.Text
X_SPIAP_BR.lbldbase = CMbDbase.Text
X_SPIAP_BR.Show vbModal
End Sub



Private Sub cmdBR2_Click()
X_ITEM_BR.LBLKODE = "X_RPT_B1"
X_ITEM_BR.lbldbase = CMbDbase.Text
X_ITEM_BR.Show vbModal
End Sub

Private Sub cmdC1_Click()
lblsalespointcd = ""
lblnmsp = ""
lblkdsp = ""
End Sub

Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub



Private Sub cmdGO_Click()
If OTSP1.Value = True Or OAIBM1.Value = True Then
Call Cetak_TSP1
ElseIf OTSP2.Value = True Or OAIBM2.Value = True Then
Call Cetak_TSP2
ElseIf OTSP3.Value = True Or OAIBM3.Value = True Then
Call Cetak_TSP3

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








Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0

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



txttgl1 = Date
txttgl2 = Date
txttglA1 = Date
txttglA2 = Date


OTSP1.Value = True

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


Private Sub txttglA1_Change()
Call nul(txttglA1)
End Sub

Private Sub txttglA1_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglA1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If

End Sub

Private Sub txttglA1_KeyPress(KeyAscii As Integer)
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

Private Sub txttglA1_LostFocus()
On Error GoTo hell

txttglA1 = FormatDateTime(txttglA1, vbGeneralDate)

Exit Sub
hell:
MsgBox "Format Tanggal tidak sesuai !", vbCritical, "Error !"
txttglA1.SetFocus
End Sub

Private Sub txttglA2_Change()
Call nul(txttglA2)
End Sub

Private Sub txttglA2_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglA2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If

End Sub

Private Sub txttglA2_KeyPress(KeyAscii As Integer)
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

Private Sub txttglA2_LostFocus()
On Error GoTo hell

txttglA2 = FormatDateTime(txttglA2, vbGeneralDate)

Exit Sub
hell:
MsgBox "Format Tanggal tidak sesuai !", vbCritical, "Error !"
txttglA2.SetFocus
End Sub


