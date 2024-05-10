VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{C115893A-A3BF-43AF-B28D-69DB846077F3}#1.0#0"; "vsflex8u.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Print_Form_MARS 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   10575
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   19020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10575
   ScaleWidth      =   19020
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   10455
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   18780
      _ExtentX        =   33126
      _ExtentY        =   18441
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "SPP MARS"
      TabPicture(0)   =   "Print_Form_MARS.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DataGrid2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdT"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdGO_SPP"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "ARV_SPP"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txttglSPP"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "LAMPIRAN SPP"
      TabPicture(1)   =   "Print_Form_MARS.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Image3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "DataGrid3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdgo"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "ARV_lamp_SPP"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "txttglSPP1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Print SJ dan SP MARS"
      TabPicture(2)   =   "Print_Form_MARS.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Image1"
      Tab(2).Control(1)=   "Label1"
      Tab(2).Control(2)=   "Label2"
      Tab(2).Control(3)=   "cmdsimpan"
      Tab(2).Control(4)=   "datagrid1"
      Tab(2).Control(5)=   "ARV1"
      Tab(2).Control(6)=   "txtnodoc"
      Tab(2).Control(7)=   "OPT1"
      Tab(2).Control(8)=   "OPT2"
      Tab(2).Control(9)=   "txttgl1"
      Tab(2).ControlCount=   10
      Begin VB.TextBox txttglSPP1 
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
         Left            =   1530
         TabIndex        =   15
         Top             =   585
         Width           =   1590
      End
      Begin VB.TextBox txttglSPP 
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
         Left            =   -73560
         TabIndex        =   9
         Top             =   495
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
         Left            =   -73515
         TabIndex        =   4
         Top             =   630
         Width           =   1590
      End
      Begin VB.OptionButton OPT2 
         BackColor       =   &H00000000&
         Caption         =   "Surat Jalan"
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
         Left            =   -72930
         TabIndex        =   3
         Top             =   1170
         Width           =   1365
      End
      Begin VB.OptionButton OPT1 
         BackColor       =   &H00000000&
         Caption         =   "Surat Penarikan"
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
         Left            =   -74730
         TabIndex        =   2
         Top             =   1170
         Width           =   1815
      End
      Begin VB.TextBox txtnodoc 
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
         Left            =   -69870
         TabIndex        =   1
         Top             =   630
         Width           =   1590
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV1 
         Height          =   7410
         Left            =   -74820
         TabIndex        =   5
         Top             =   2700
         Width           =   18345
         _ExtentX        =   32359
         _ExtentY        =   13070
         SectionData     =   "Print_Form_MARS.frx":0054
      End
      Begin VSFlex8UCtl.VSFlexGrid datagrid1 
         Height          =   1995
         Left            =   -67800
         TabIndex        =   6
         Top             =   585
         Width           =   10185
         _cx             =   17965
         _cy             =   3519
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
         AllowUserResizing=   4
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"Print_Form_MARS.frx":0090
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
      Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV_SPP 
         Height          =   7545
         Left            =   -74865
         TabIndex        =   10
         Top             =   2520
         Width           =   18345
         _ExtentX        =   32359
         _ExtentY        =   13309
         SectionData     =   "Print_Form_MARS.frx":01A9
      End
      Begin Threed.SSCommand cmdGO_SPP 
         Height          =   780
         Left            =   -68970
         TabIndex        =   13
         ToolTipText     =   "Simpan"
         Top             =   495
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
         Picture         =   "Print_Form_MARS.frx":01E5
         Caption         =   "&s"
         ButtonStyle     =   4
      End
      Begin Threed.SSCommand cmdT 
         Height          =   780
         Left            =   -68970
         TabIndex        =   14
         ToolTipText     =   "Tambah"
         Top             =   1260
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1376
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
         Picture         =   "Print_Form_MARS.frx":3A9B
         Alignment       =   1
         ButtonStyle     =   4
      End
      Begin DDActiveReportsViewer2Ctl.ARViewer2 ARV_lamp_SPP 
         Height          =   7545
         Left            =   45
         TabIndex        =   16
         Top             =   2610
         Width           =   18345
         _ExtentX        =   32359
         _ExtentY        =   13309
         SectionData     =   "Print_Form_MARS.frx":670F
      End
      Begin Threed.SSCommand cmdgo 
         Height          =   780
         Left            =   6255
         TabIndex        =   18
         ToolTipText     =   "Simpan"
         Top             =   540
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
         Picture         =   "Print_Form_MARS.frx":674B
         Caption         =   "&s"
         ButtonStyle     =   4
      End
      Begin Threed.SSCommand cmdsimpan 
         Height          =   780
         Left            =   -69105
         TabIndex        =   19
         ToolTipText     =   "Simpan"
         Top             =   1080
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
         Picture         =   "Print_Form_MARS.frx":A001
         Caption         =   "&s"
         ButtonStyle     =   4
      End
      Begin VSFlex8UCtl.VSFlexGrid DataGrid3 
         Height          =   1995
         Left            =   7155
         TabIndex        =   20
         Top             =   540
         Width           =   11040
         _cx             =   19473
         _cy             =   3519
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
         AllowUserResizing=   4
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"Print_Form_MARS.frx":D8B7
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
      Begin VSFlex8UCtl.VSFlexGrid DataGrid2 
         Height          =   1995
         Left            =   -67845
         TabIndex        =   11
         Top             =   450
         Width           =   11040
         _cx             =   19473
         _cy             =   3519
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
         AllowUserResizing=   4
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"Print_Form_MARS.frx":DA0E
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
         Begin VB.TextBox DGUrut_SPP 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   285
            Left            =   9315
            TabIndex        =   21
            Text            =   "dgtglplan"
            Top             =   225
            Visible         =   0   'False
            Width           =   1230
         End
      End
      Begin VB.Label Label4 
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
         Left            =   315
         TabIndex        =   17
         Top             =   630
         Width           =   1320
      End
      Begin VB.Label Label3 
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
         Left            =   -74775
         TabIndex        =   12
         Top             =   540
         Width           =   1320
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
         Left            =   -74730
         TabIndex        =   8
         Top             =   675
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NO DOC TERAKHIR :"
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
         Left            =   -71490
         TabIndex        =   7
         Top             =   675
         Width           =   1725
      End
      Begin VB.Image Image1 
         Height          =   10005
         Left            =   -74955
         Picture         =   "Print_Form_MARS.frx":DB65
         Stretch         =   -1  'True
         Top             =   405
         Width           =   18690
      End
      Begin VB.Image Image2 
         Height          =   10005
         Left            =   -74910
         Picture         =   "Print_Form_MARS.frx":2AD2C
         Stretch         =   -1  'True
         Top             =   360
         Width           =   18690
      End
      Begin VB.Image Image3 
         Height          =   10005
         Left            =   45
         Picture         =   "Print_Form_MARS.frx":47EF3
         Stretch         =   -1  'True
         Top             =   360
         Width           =   18690
      End
   End
End
Attribute VB_Name = "Print_Form_MARS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset

Private Sub Cetak_SJ()
Unload AR_SP_Mars
Unload AR_SJ_Mars

sql1 = "select row_number() OVER (partition BY a.tgl ORDER BY a.kdx) AS urut,a.tgl,a.kdx,a.kdmars + ' - ' + a.nmmars as cust1,a.alamatmars + ',,' as alamatMars,a.ket_PS,a.kdmars + '-' + convert(varchar,a.urut_SPP ) as noSPP,a.urut_spp,b.jml  from V_SJ_MARS_H a left join " & vbCrLf & _
      "(select kdx,sum(unit) as jml from V_SJ_MARS_D group by kdx) b on a.kdx=b.kdx where a.tgl='" & Format(txttgl1, "yyyy/MM/dd") & "'"

sql = "select *, " & txtnodoc & " + urut as nodoc from (" & sql1 & ") x order by urut "

Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs

With AR_SJ_Mars.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_SJ_Mars
.Fldcust_mars.DataField = "cust1"
.fldalamat.DataField = "alamatMars"
.fldtgl.DataField = "tgl"
.fldno_SPP.DataField = "nospp"
.fldkdx.DataField = "kdx"
.fldnoSJ.DataField = "nodoc"

Set Me.ARV1.ReportSource = AR_SJ_Mars

'
'.Zoom = 140
'
'
'AR_SJ_MARS.Show vbModal
'
'
 End With

End Sub


Private Sub Cetak_SP()

Unload AR_SP_Mars
Unload AR_SJ_Mars

sql1 = "select row_number() OVER (partition BY a.tgl ORDER BY a.kdx) AS urut,a.tgl,a.kdx,a.kdmars + ' - ' + a.nmmars as cust1,a.alamatmars + ',,' as alamatMars,a.ket_PS,a.kdmars + '-' + convert(varchar,a.urut_SPP) as noSPP,a.urut_spp,b.jml  from V_SP_MARS_H a left join " & vbCrLf & _
      "(select kdx,sum(unit) as jml from V_SP_MARS_D group by kdx) b on a.kdx=b.kdx where a.tgl='" & Format(txttgl1, "yyyy/MM/dd") & "'"

sql = "select *, " & txtnodoc & " + urut as nodoc from (" & sql1 & ") x order by urut "

Set rs = con.Execute(sql)
Set datagrid1.DataSource = rs

With AR_SP_Mars.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_SP_Mars
.Fldcust_mars.DataField = "cust1"
.fldalamat.DataField = "alamatMars"
.fldtgl.DataField = "tgl"
.fldno_SPP.DataField = "noSPP"
.fldkdx.DataField = "kdx"
.fldnoSJ.DataField = "nodoc"
Set Me.ARV1.ReportSource = AR_SP_Mars


 End With

End Sub


Private Sub Cetak_SPP()
Unload AR_SPP_Mars

sql = "select * from V_list_cust_SPP where tgl='" & Format(txttglSPP, "yyyy/MM/dd") & "' and kdcustomer in (select kdcustomer from rekap_pjm_sewa)"

Set rs = con.Execute(sql)
Set DataGrid2.DataSource = rs

With AR_SPP_Mars.DC1
.ConnectionString = koneksi
.Source = sql
End With

With AR_SPP_Mars

.fldcustomerMARS.DataField = "customerMARS"
.fldalamatMARS.DataField = "alamatMars"
.fldtgl.DataField = "tgl"
.fldnoSPP_L.DataField = "nospp_L"
.fldnoSPP_N.DataField = "nospp_N"
.fldalamat_KTP.DataField = "Alamatmars"
.fldnotlp.DataField = "phone"
.fldkdpos.DataField = "Post"
.fldotl.DataField = "otl_type_nm"

Set Me.ARV_SPP.ReportSource = AR_SPP_Mars

End With

End Sub


Private Sub Cetak_SPP_Lampiran()
Unload AR_SPP_Lampiran_Mars


sqlA1 = "select a.*,b.jml,isnull(c.tdk_cetak,0) as tdk_cetak from V_list_cust_SPP a left join (select kdcustomer , sum(sisa) as jml from rekap_pjm_sewa group by kdcustomer) b on a.kdcustomer=b.kdcustomer left join (select * from tdk_cetak_mars where jns_T='LSPP') C on a.kdcustomer=c.kdcustomer where a.tgl='" & Format(txttglSPP1, "yyyy/MM/dd") & "' and b.jml > 0"

sql1 = "select * from (" & sqlA1 & ") a where tdk_cetak=0"

Set rs = con.Execute(sqlA1)
Set DataGrid3.DataSource = rs

With AR_SPP_Lampiran_Mars.DC1
.ConnectionString = koneksi
.Source = sql1
End With

With AR_SPP_Lampiran_Mars
.fldkdcustomer.DataField = "kdcustomer"
.fldnoSPP_Lama.DataField = "nospp_L"
.fldtgl.DataField = "tgl"

Set Me.ARV_lamp_SPP.ReportSource = AR_SPP_Lampiran_Mars

End With

End Sub




Private Sub cmdGO_Click()
Call Cetak_SPP_Lampiran
End Sub

Private Sub cmdGO_SPP_Click()
Call Cetak_SPP
End Sub

Private Sub cmdsimpan_Click()
If OPT1.Value = True Then
Call Cetak_SP
ElseIf OPT2.Value = True Then
Call Cetak_SJ
End If
End Sub


Private Sub cmdT_Click()
On Error GoTo hell
    ms = MsgBox("Apakah anda ingin MengUpdate No SPP ?", vbYesNo + vbQuestion, "Info")
    If ms = vbYes Then
        sql1 = "select kdmars  from V_list_cust_SPP where tgl='" & Format(txttglSPP, "yyyy/MM/dd") & "'"
        con.Execute ("update kddispromars set urut_SPP= urut_SPP + 1 where kdMars in (" & sql1 & ")")
         
                
        sql = "select * from V_list_cust_SPP where tgl='" & Format(txttglSPP, "yyyy/MM/dd") & "' and kdcustomer in (select kdcustomer from rekap_pjm_sewa)"
        Set rs = con.Execute(sql)
        Set DataGrid2.DataSource = rs

        
    Else
        Exit Sub
    End If


Exit Sub
hell:
MsgBox err.Description

End Sub

Private Sub DataGrid2_DblClick()
If DataGrid2.Col = 9 Then

    '    kode = 2
    '    lblpos = rs.AbsolutePosition
        DGUrut_SPP.Top = DataGrid2.CellTop
        DGUrut_SPP.Left = DataGrid2.CellLeft
        DGUrut_SPP = rs!urut_SPP_N
        
        DGUrut_SPP.Visible = True
        DGUrut_SPP.Height = DataGrid2.CellHeight
        DGUrut_SPP.Width = DataGrid2.CellWidth
        DGUrut_SPP.BackColor = vbYellow
        DGUrut_SPP.SetFocus
        SendKeys "{Home}+{End}"
End If
End Sub

Private Sub DataGrid3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF2 Then
    If rs!Tdk_cetak = 0 Then
    con.Execute ("insert into tdk_cetak_mars values('" & rs!kdcustomer & "_LSPP" & "','" & rs!kdcustomer & "','LSPP',1)")
    Else
    con.Execute ("delete from tdk_cetak_mars where kdcustomer='" & rs!kdcustomer & "' and jns_T='LSPP'")
    End If
    
    Call Cetak_SPP_Lampiran
End If

End Sub

Private Sub DGUrut_SPP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
con.Execute ("update kddisproMars set urut_spp=" & CLng(DGUrut_SPP) & " where kdcustomer='" & rs!kdcustomer & "' ")
Call Cetak_SPP
DGUrut_SPP.Visible = False
End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
End If
End Sub

Private Sub Form_Load()
txttgl1 = Date
txttglSPP = Date
txttglSPP1 = Date
OPT1.Value = True

End Sub

Private Sub OPT1_Click()
lbljudul = "SURAT PENARIKAN MARS"
End Sub

Private Sub Opt2_Click()
lbljudul = "SURAT JALAN MARS"
End Sub

Private Sub SSCommand1_Click()

End Sub

Private Sub txttgl1_Change()
Call nul(txttgl1)
End Sub

Private Sub txttgl1_GotFocus()
txttgl1.SelStart = 0
txttgl1.SelLength = Len(txttgl1)
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

Private Sub txttglspp_Change()
Call nul(txttglSPP)
End Sub

Private Sub txttglspp_GotFocus()
txttglSPP.SelStart = 0
txttglSPP.SelLength = Len(txttglSPP)
End Sub

Private Sub txttglspp_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
ElseIf KeyCode = vbKeyUp Then
SendKeys "{Home}+{Tab}"
End If

End Sub

Private Sub txttglspp_KeyPress(KeyAscii As Integer)
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

Private Sub txttglspp_LostFocus()
On Error GoTo hell

txttglSPP = FormatDateTime(txttglSPP, vbGeneralDate)

Exit Sub
hell:
MsgBox "Format Tanggal tidak sesuai !", vbCritical, "Error !"
txttglSPP.SetFocus
End Sub

