VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Customer_SPP 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   10950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   18735
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   18735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TimerPdf1 
      Left            =   14040
      Top             =   5985
   End
   Begin VB.Timer TimerRtf1 
      Left            =   13095
      Top             =   5985
   End
   Begin VB.OptionButton OSPP3 
      BackColor       =   &H00000000&
      Caption         =   "Print Out Ulang SPP dan Lampiran"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   510
      Left            =   13050
      TabIndex        =   2
      Top             =   225
      Width           =   2175
   End
   Begin VB.OptionButton OSPP2 
      BackColor       =   &H00000000&
      Caption         =   "Pembaharuan Lampiran "
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   510
      Left            =   11160
      TabIndex        =   1
      Top             =   225
      Width           =   1590
   End
   Begin VB.OptionButton OSPP1 
      BackColor       =   &H00000000&
      Caption         =   "Form SPP dan Lampiran Baru"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   510
      Left            =   9180
      TabIndex        =   0
      Top             =   225
      Width           =   1590
   End
   Begin VB.TextBox txttgllampiran 
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
      Left            =   1755
      TabIndex        =   34
      Text            =   "01/01/1900"
      Top             =   5850
      Width           =   1410
   End
   Begin VB.TextBox txtomset_Krt 
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
      Left            =   13815
      TabIndex        =   29
      Text            =   "30"
      Top             =   1350
      Width           =   735
   End
   Begin VB.TextBox txtomset_GLN 
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
      Left            =   13815
      TabIndex        =   26
      Text            =   "30"
      Top             =   945
      Width           =   735
   End
   Begin VB.TextBox txttglspp 
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
      Left            =   4500
      TabIndex        =   15
      Text            =   "01/01/1900"
      Top             =   945
      Width           =   1410
   End
   Begin VB.Timer TimerRtf 
      Left            =   13950
      Top             =   2295
   End
   Begin VB.Timer TimerPdf 
      Left            =   14895
      Top             =   2295
   End
   Begin VB.Timer TimerQty 
      Left            =   8370
      Top             =   315
   End
   Begin Threed.SSCommand cmdfs 
      Height          =   300
      Left            =   16065
      TabIndex        =   3
      Top             =   2205
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
      Picture         =   "Customer_SPP.frx":0000
      Caption         =   "&Full Screen"
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
      Height          =   3585
      Left            =   315
      TabIndex        =   4
      Top             =   2115
      Width           =   17130
      _ExtentX        =   30215
      _ExtentY        =   6324
      SectionData     =   "Customer_SPP.frx":6862
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
      ToolTipText     =   "Cetak"
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
      Picture         =   "Customer_SPP.frx":689E
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdPdf 
      Height          =   780
      Left            =   17820
      TabIndex        =   7
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
      Picture         =   "Customer_SPP.frx":A154
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdrtf 
      Height          =   780
      Left            =   17820
      TabIndex        =   8
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
      Picture         =   "Customer_SPP.frx":D33B
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   1575
      TabIndex        =   9
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
      Picture         =   "Customer_SPP.frx":10981
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin Threed.SSCommand cmdFs1 
      Height          =   300
      Left            =   16065
      TabIndex        =   33
      Top             =   6300
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
      Picture         =   "Customer_SPP.frx":171E3
      Caption         =   "&Full Screen"
      Alignment       =   5
      ButtonStyle     =   3
      PictureAlignment=   1
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV2 
      Height          =   3810
      Left            =   360
      TabIndex        =   32
      Top             =   6210
      Width           =   17130
      _ExtentX        =   30215
      _ExtentY        =   6720
      SectionData     =   "Customer_SPP.frx":1DA45
   End
   Begin Threed.SSCommand cmdpdf1 
      Height          =   780
      Left            =   17820
      TabIndex        =   43
      Top             =   7065
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
      Picture         =   "Customer_SPP.frx":1DA81
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin Threed.SSCommand cmdRtf1 
      Height          =   780
      Left            =   17820
      TabIndex        =   44
      Top             =   6255
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
      Picture         =   "Customer_SPP.frx":20C68
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin VB.Label lblP_DISP_Lama 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   11025
      TabIndex        =   42
      Top             =   1755
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblP_RG_lama 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   11025
      TabIndex        =   41
      Top             =   2115
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblP_SHW_lama 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   11025
      TabIndex        =   40
      Top             =   2475
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblS_DISP_lama 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   12015
      TabIndex        =   39
      Top             =   1755
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lbltgllampiran_lama 
      Height          =   375
      Left            =   3330
      TabIndex        =   38
      Top             =   5805
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label lbltglspp_lama 
      Height          =   375
      Left            =   4500
      TabIndex        =   37
      Top             =   1395
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label lblnospp_lama 
      Height          =   375
      Left            =   2205
      TabIndex        =   36
      Top             =   1395
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL LAMPIRAN :"
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
      TabIndex        =   35
      Top             =   5895
      Width           =   1410
   End
   Begin VB.Label lblkode 
      Caption         =   "Label6"
      Height          =   375
      Left            =   16020
      TabIndex        =   31
      Top             =   1350
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "KARTON/BLN"
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
      Left            =   14625
      TabIndex        =   30
      Top             =   1395
      Width           =   1185
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "GLN/BLN"
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
      Left            =   14625
      TabIndex        =   28
      Top             =   990
      Width           =   1185
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "TARGET OMSET :"
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
      Left            =   12420
      TabIndex        =   27
      Top             =   990
      Width           =   1455
   End
   Begin VB.Label lblS_DISP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   11205
      TabIndex        =   25
      Top             =   945
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "DISPENCER :"
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
      Left            =   10170
      TabIndex        =   24
      Top             =   990
      Width           =   1185
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "SEWA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   9495
      TabIndex        =   23
      Top             =   990
      Width           =   600
   End
   Begin VB.Label lblP_SHW 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   8460
      TabIndex        =   22
      Top             =   1665
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "SHOWCASE :"
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
      Left            =   7380
      TabIndex        =   21
      Top             =   1755
      Width           =   1095
   End
   Begin VB.Label lblP_RG 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   8460
      TabIndex        =   20
      Top             =   1305
      Width           =   735
   End
   Begin VB.Label lblP_DISP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "100"
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
      Left            =   8460
      TabIndex        =   19
      Top             =   945
      Width           =   735
   End
   Begin VB.Label lblnoSPP 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "J302/0822/C00123-01"
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
      Left            =   1170
      TabIndex        =   18
      Top             =   945
      Width           =   2265
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "RAK GLN :"
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
      Left            =   7650
      TabIndex        =   17
      Top             =   1395
      Width           =   960
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "TGL SPP :"
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
      Left            =   3690
      TabIndex        =   16
      Top             =   990
      Width           =   870
   End
   Begin VB.Label lblbarang_R 
      Height          =   330
      Left            =   10530
      TabIndex        =   14
      Top             =   2925
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Surat Pernyataan Pelanggan"
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
      Left            =   1260
      TabIndex        =   13
      Top             =   90
      Width           =   7665
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "No. SPP :"
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
      TabIndex        =   12
      Top             =   990
      Width           =   870
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PINJAMAN "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   6300
      TabIndex        =   11
      Top             =   990
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "DISPENCER :"
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
      Left            =   7425
      TabIndex        =   10
      Top             =   990
      Width           =   1185
   End
   Begin VB.Image Image1 
      Height          =   10905
      Index           =   0
      Left            =   0
      Picture         =   "Customer_SPP.frx":242AE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   18735
   End
End
Attribute VB_Name = "Customer_SPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim rs As ADODB.Recordset
Dim rsQP As ADODB.Recordset
Dim rsQS As ADODB.Recordset
Dim rsL As ADODB.Recordset
Dim sqlL, sql1 As String
Dim sqlA As String
Dim color As Long, flag As Byte

Private Sub nomer()
On Error GoTo hell

sql = "Select isnull(max(right(noSPP,2)),0) as xx from LIST_SPP where month(tglspp)=" & Format(txttglspp, "MM") & " and year(tglspp)=" & Format(txttglspp, "yyyy") & " and kdcustomer='" & Customer_TU.lblkdcustomer & "'"
Set rs = con.Execute(sql)


        a = CInt(rs!xx) + 1
                
        Select Case Len(CStr(a))
        Case 1
            lblnoSPP = "J302-" & Format(txttglspp, "MMyy") & "-" & Customer_TU.lblkdcustomer & "-0" & (a)
        Case 2
            lblnoSPP = "J302-" & Format(txttglspp, "MMyy") & "-" & Customer_TU.lblkdcustomer & "-" & (a)
        
        
        End Select

Exit Sub
hell:
lblnoSPP = "J302-" & Format(txttglspp, "MMyy") & "-" & Customer_TU.lblkdcustomer & "-01"

End Sub


Private Sub qty_Unit()
sqlQP1 = "select a.kdcustomer,sum(case when kdkategori in ('04','05','06','07') then unit else 0 end ) as P_DSP," & vbCrLf & _
         "sum(case when kdkategori in ('08','09') then unit else 0 end ) as P_SHW,sum(case when kdkategori in ('10') then unit else 0 end ) as P_RG" & vbCrLf & _
         "from pinjam a left join pinjam_d b on a.kdpinjam=b.kdpinjam left join barang c on b.kdbarang=c.kdbarang where a.kdcustomer='" & Customer_TU.lblkdcustomer & "' and a.tglpinjam <= '" & Format(txttglspp, "yyyy/MM/dd") & "' group by a.kdcustomer" & vbCrLf & _
         "Union ALL" & vbCrLf & _
         "select a.kdcustomer,-sum(case when kdkategori in ('04','05','06','07') then unit else 0 end ) as P_DSP," & vbCrLf & _
         "-sum(case when kdkategori in ('08','09') then unit else 0 end ) as P_SHW,-sum(case when kdkategori in ('10') then unit else 0 end ) as P_RG" & vbCrLf & _
         "from Rpinjam a left join Rpinjam_d b on a.kdRpinjam=b.kdRpinjam left join barang c on b.kdbarang=c.kdbarang where a.kdcustomer='" & Customer_TU.lblkdcustomer & "' and a.tglRpinjam <= '" & Format(txttglspp, "yyyy/MM/dd") & "' group by a.kdcustomer"

sqlQP = "select kdcustomer,sum(P_DSP) as P_DSP,sum(P_SHW) as P_SHW,sum(P_RG) as P_RG from (" & sqlQP1 & ") x group by kdcustomer"
        

        
Set rsQP = con.Execute(sqlQP)



If rsQP.RecordCount <> 0 Then
lblP_DISP = rsQP!P_DSP
lblP_SHW = rsQP!P_SHW
lblP_RG = rsQP!P_RG
Else
lblP_DISP = 0
lblP_SHW = 0
lblP_RG = 0
End If

sqlQS1 = "select a.kdcustomer,sum(case when kdkategori in ('04','05') then unit else 0 end ) as S_DSP" & vbCrLf & _
         "from Sewa a left join Sewa_d b on a.kdSewa=b.kdSewa left join barang c on b.kdbarang=c.kdbarang where a.kdcustomer='" & Customer_TU.lblkdcustomer & "' and a.tglSewa <= '" & Format(txttglspp, "yyyy/MM/dd") & "' group by a.kdcustomer" & vbCrLf & _
         "Union ALL" & vbCrLf & _
         "select a.kdcustomer,-sum(case when kdkategori in ('04','05') then unit else 0 end ) as S_DSP" & vbCrLf & _
         "from RSewa a left join RSewa_d b on a.kdRSewa=b.kdRSewa left join barang c on b.kdbarang=c.kdbarang where a.kdcustomer='" & Customer_TU.lblkdcustomer & "' and a.tglRSewa <= '" & Format(txttglspp, "yyyy/MM/dd") & "' group by a.kdcustomer"

sqlQS = "select kdcustomer,sum(S_DSP) as S_DSP from (" & sqlQS1 & ") x group by kdcustomer"



Set rsQS = con.Execute(sqlQS)



If rsQS.RecordCount <> 0 Then
lblS_DISP = rsQS!S_DSP
Else
lblS_DISP = 0
End If



End Sub



Private Sub cmdCANCEL_Click()
Unload Me
End Sub

Private Sub cmdCANCEL_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub


Private Sub cmdFs1_Click()
AR_SPP_LAMPIRAN.Show vbModal
End Sub

Private Sub cmdPDF1_Click()
TimerPdf1.Interval = 10
End Sub

Private Sub cmdrtf1_Click()
TimerRtf1.Interval = 10
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




Private Sub Cetak()
On Error GoTo hell

'FORM SPP------------------------------------------
If OSPP1.Value = True Then
con.Execute (" update list_Spp set berlaku=0 where kdcustomer='" & Customer_TU.lblkdcustomer & "' ")
con.Execute ("insert into list_spp values ('" & lblnoSPP & "','" & Format(txttglspp, "yyyy/MM/dd") & "','" & Customer_TU.lblkdcustomer & "'," & lblP_DISP & "," & lblP_RG & "," & lblP_SHW & "," & lblS_DISP & "," & txtomset_GLN & "," & txtomset_Krt & ",1,'" & Format(txttgllampiran, "yyyy/MM/dd") & "',getdate(),'" & UTAMA.lblkduser & "')")
con.Execute ("insert into lampiran_spp values ('" & Format(txttgllampiran, "d_M_yyyy/") & lblnoSPP & "','" & lblnoSPP & "','" & Format(txttgllampiran, "yyyy/MM/dd") & "',getdate(),'" & UTAMA.lblkduser & "')")

Customer_TU.TimerSPP.Interval = 10

ElseIf OSPP2.Value = True Then
con.Execute ("insert into lampiran_spp values ('" & Format(txttgllampiran, "d_M_yyyy/") & lblnoSPP & "','" & lblnoSPP & "','" & Format(txttgllampiran, "yyyy/MM/dd") & "',getdate(),'" & UTAMA.lblkduser & "')")

Customer_TU.TimerSPP.Interval = 10

End If



Unload AR_SPP


With AR_SPP
.lblnmcustomer = Customer_TU.TXTnmcustomer
.lblalamat = Customer_TU.txtalamat
.lblnoSPP = lblnoSPP
.lbltglSPP = txttglspp
.lblP_DSP = lblP_DISP
.lblP_RG = lblP_RG
.lblP_SHW = lblP_SHW
.lblS_DSP = lblS_DISP
.lbltarget_Gln = txtomset_GLN
.lbltarget_Krt = txtomset_Krt



Set Me.ARV1.ReportSource = AR_SPP
End With


'LAMPIRAN SPP -------------------------------------------
sqlL = "exec sp_spp_lampiran @tgl='" & Format(txttgllampiran, "yyyy/MM/dd") & "',@kdcustomer='" & Customer_TU.lblkdcustomer & "'"
Set rsL = con.Execute(sqlL)



Unload AR_SPP_LAMPIRAN

With AR_SPP_LAMPIRAN.DC1
.ConnectionString = koneksi
.Source = sqlL
End With


With AR_SPP_LAMPIRAN
.fldkdbarang.DataField = "kdbarang"
.fldkd1.DataField = "kd1"
.fldkdsap.DataField = "kdsap"
.fldstatus.DataField = "jns"
.fldjns.DataField = "jnsbrg"

.lblnoSPP = "( " & lblnoSPP & " )"
.lblnmcustomer = Customer_TU.TXTnmcustomer
.lblalamat = Customer_TU.txtalamat
.lbltglSPP = txttgllampiran


Set Me.ARV2.ReportSource = AR_SPP_LAMPIRAN
End With



cmdGO.Enabled = False



Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
End Sub

Private Sub cmdfs_Click()
AR_SPP.Show vbModal
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


Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0
txttglspp = Date
txttgllampiran = Date


OSPP3.Value = True

'TimerQty.Interval = 100
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub OSPP1_Click()
txttglspp.Enabled = True
txtomset_GLN.Enabled = True
txtomset_Krt.Enabled = True
txttgllampiran.Enabled = True
txttglspp = Date
txttgllampiran = Date


TimerQty.Interval = 100
End Sub

Private Sub OSPP2_Click()
lblnoSPP = lblnospp_lama
txttglspp = lbltglspp_lama
txttgllampiran = lbltgllampiran_lama
txttglspp.Enabled = False
txtomset_GLN.Enabled = False
txtomset_Krt.Enabled = False
txttgllampiran.Enabled = True

lblP_DISP = lblP_DISP_Lama
lblP_RG = lblP_RG_lama
lblP_SHW = lblP_SHW_lama
lblS_DISP = lblS_DISP_lama
TimerQty.Interval = 100
End Sub

Private Sub OSPP3_Click()
lblnoSPP = lblnospp_lama
txttglspp = lbltglspp_lama
txttgllampiran = lbltgllampiran_lama

txttglspp.Enabled = False
txtomset_GLN.Enabled = False
txtomset_Krt.Enabled = False
txttgllampiran.Enabled = False

lblP_DISP = lblP_DISP_Lama
lblP_RG = lblP_RG_lama
lblP_SHW = lblP_SHW_lama
lblS_DISP = lblS_DISP_lama

TimerQty.Interval = 100
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

Private Sub TimerPDF1_Timer()
On Error GoTo hell
Dim pdf As New ActiveReportsPDFExport.ARExportPDF

out2 = out2 + 1

Call save_out
pdf.filename = alamat_save & "\outfile" & CStr(out2) & ".pdf"
pdf.Export ARV2.Pages

Call EX_PDF(Me)
TimerPdf1.Interval = 0

Exit Sub
hell:
TimerPdf1.Interval = 0
If out2 < 10 Then
cmdPDF1_Click
End If

End Sub

Private Sub TimerQty_Timer()
On Error GoTo hell

If OSPP1.Value = True Then
Call nomer
End If

Call qty_Unit
TimerQty.Interval = 0

Exit Sub
hell:
MsgBox err.Description
TimerQty.Interval = 0
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

Private Sub Timerrtf1_Timer()
On Error GoTo hell
Dim rtf As New ActiveReportsRTFExport.ARExportRTF
out = out + 1

Call save_out
rtf.filename = alamat_save & "\outfile" & CStr(out) & ".rtf"
rtf.Export ARV2.Pages

Call EX_WORD(Me)
TimerRtf1.Interval = 0

Exit Sub
hell:
TimerRtf1.Interval = 0
If out < 10 Then
cmdrtf1_Click
End If
End Sub



Private Sub txttgllampiran_Change()
Call nul(txttgllampiran)
End Sub

Private Sub txttgllampiran_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttgllampiran_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttgllampiran_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii <> vbKeyBack Then
    cekTBL = InStr("1234567890/-", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub txttgllampiran_LostFocus()
On Error GoTo hell

txttgllampiran = FormatDateTime(txttgllampiran, vbGeneralDate)

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttgllampiran.SetFocus

End Sub

Private Sub txttglSPP_Change()
Call nul(txttglspp)
End Sub

Private Sub txttglSPP_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txttglSPP_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
ElseIf KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub txttglSPP_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii <> vbKeyBack Then
    cekTBL = InStr("1234567890/-", Chr(KeyAscii))
    If cekTBL = 0 Then
        KeyAscii = 0
        Beep
    End If
End If
End Sub

Private Sub txttglSPP_LostFocus()
On Error GoTo hell

txttglspp = FormatDateTime(txttglspp, vbGeneralDate)

Call nomer

TimerQty.Interval = 10

Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !"
txttglspp.SetFocus

End Sub

