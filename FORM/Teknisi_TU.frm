VERSION 5.00
Object = "{1F81B5E0-26A8-11D0-BDCB-0020A90B183A}#8.0#0"; "PVLine.ocx"
Object = "{F20E41DE-526A-423A-B746-D860D06076B4}#4.0#0"; "IGThreed40.ocx"
Begin VB.Form Teknisi_TU 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7260
   LinkTopic       =   "Form1"
   ScaleHeight     =   3180
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkNA 
      BackColor       =   &H00000000&
      Caption         =   "NON AKTIF"
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
      Height          =   330
      Left            =   4770
      TabIndex        =   2
      Top             =   1980
      Width           =   1230
   End
   Begin VB.ComboBox CMbSTATUS 
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
      Left            =   1035
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1935
      Width           =   1500
   End
   Begin VB.TextBox txtnmteknisi 
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
      Left            =   1035
      TabIndex        =   0
      Top             =   1575
      Width           =   5010
   End
   Begin VB.TextBox lblkdteknisi 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Enabled         =   0   'False
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
      Left            =   1035
      TabIndex        =   5
      Top             =   1215
      Width           =   1005
   End
   Begin VB.Timer TimerNO 
      Left            =   2745
      Top             =   765
   End
   Begin Threed.SSCommand cmdsimpan 
      Height          =   915
      Left            =   6210
      TabIndex        =   3
      ToolTipText     =   "Simpan"
      Top             =   1260
      Width           =   960
      _ExtentX        =   1693
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
      Picture         =   "Teknisi_TU.frx":0000
      Caption         =   "&s"
      ButtonStyle     =   4
   End
   Begin PVLINE3DLib.PVLine3D PVLine3D1 
      Height          =   30
      Left            =   135
      TabIndex        =   6
      Top             =   720
      Width           =   5955
      _Version        =   524288
      _ExtentX        =   10504
      _ExtentY        =   53
      _StockProps     =   8
      ShadowHorizontal=   3
      ShadowVertical  =   3
      ShadowColor     =   8421504
      Transparent     =   -1  'True
   End
   Begin Threed.SSCommand cmdCANCEL 
      Height          =   375
      Left            =   495
      TabIndex        =   4
      Top             =   2610
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
      Picture         =   "Teknisi_TU.frx":2A6D
      Caption         =   "     &Click di sini jika ingin keluar"
      Alignment       =   1
      ButtonStyle     =   4
      PictureAlignment=   1
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "STATUS :"
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
      Left            =   225
      TabIndex        =   11
      Top             =   2025
      Width           =   780
   End
   Begin VB.Image Image4 
      Height          =   630
      Left            =   6255
      Picture         =   "Teknisi_TU.frx":92CF
      Stretch         =   -1  'True
      Top             =   135
      Width           =   735
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "TEKNISI :"
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
      Left            =   225
      TabIndex        =   10
      Top             =   1620
      Width           =   1320
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "KODE :"
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
      Left            =   225
      TabIndex        =   9
      Top             =   1260
      Width           =   645
   End
   Begin VB.Label lblkode 
      Caption         =   "lblkode"
      Height          =   285
      Left            =   945
      TabIndex        =   8
      Top             =   4275
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Detail Teknisi / Cheker"
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
      Left            =   495
      TabIndex        =   7
      Top             =   0
      Width           =   5370
   End
   Begin VB.Image Image2 
      Height          =   420
      Left            =   2205
      Picture         =   "Teknisi_TU.frx":968F
      Stretch         =   -1  'True
      Top             =   1170
      Width           =   465
   End
   Begin VB.Image Image1 
      Height          =   3165
      Left            =   0
      Picture         =   "Teknisi_TU.frx":A94C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7260
   End
End
Attribute VB_Name = "Teknisi_TU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim sql As String
Dim sql1 As String
Dim a As Integer
Dim color As Long, flag As Byte


Private Sub nomer()
On Error GoTo hell

sql = "Select isnull(max(right(kdteknisi,2)),0) as xx from teknisi"
Set rs = con.Execute(sql)


        a = CInt(rs!xx) + 1
                
        Select Case Len(CStr(a))
        Case 1
           lblkdteknisi = "T0" & (a)
        Case 2
           lblkdteknisi = "T" & (a)
        
        
        End Select

Exit Sub
hell:
lblkdteknisi = "T01"

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




Private Sub cmdsimpan_Click()
On Error GoTo hell

    If txtnmteknisi = "" Or lblkdteknisi = "" Then
    MsgBox "inputan belum lengkap !!", vbInformation, "Info !!"
    Exit Sub
    Else

         If LBLKODE = 1 Then
             sql = "insert into teknisi  values ('" & UCase(lblkdteknisi) & "','" & UCase(txtnmteknisi) & "','" & CMBStatus.Text & "'," & ChkNA.Value & ")"
             con.Execute (sql)
             MsgBox "Data Telah Tersimpan", vbInformation, "Informasi !"
     
             Teknisi.TimerAll.Interval = 10
         Else
             sql = "update teknisi set nmteknisi='" & UCase(txtnmteknisi) & "',status='" & CMBStatus & "',non_aktif=" & ChkNA.Value & " where kdteknisi='" & lblkdteknisi & "'"
             con.Execute (sql)
             MsgBox "Data Telah di Ubah", vbInformation, "Informasi !"

             Teknisi.TimerAll.Interval = 10

         End If
         
         Unload Me
    End If
Exit Sub
hell:
MsgBox err.Description, vbCritical, "Error !!"

End Sub

Private Sub cmdsimpan_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
GradientForm Me, 0
TimerNO.Interval = 10

CMBStatus.AddItem "CHEKER"
CMBStatus.AddItem "TEKNISI"
CMBStatus.AddItem "SOPIR"
CMBStatus.AddItem "HELPER"
CMBStatus.ListIndex = 0

Call nul(txtnmteknisi)
End Sub


Private Sub lblkdsupplier_Change()
Call nul(lblkdsupplier)
End Sub

Private Sub lblkdsupplier_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub lblkdsupplier_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
End If
End Sub

Private Sub lblkdsupplier_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If
End Sub

Private Sub lblkdsupplier_LostFocus()
lblkdsupplier = UCase(lblkdsupplier)
End Sub

Private Sub lblkdteknisi_Change()
Call nul(lblkdteknisi)
End Sub

Private Sub TimerNO_Timer()
If LBLKODE = 1 Then
Call nomer
End If

TimerNO.Interval = 0
End Sub

Private Sub txtnmteknisi_Change()
Call nul(txtnmteknisi)
End Sub

Private Sub txtnmteknisi_GotFocus()
SendKeys "{Home}+{End}"
End Sub

Private Sub txtnmteknisi_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDown Then
SendKeys vbTab
End If

End Sub

Private Sub txtnmteknisi_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
Unload Me
ElseIf KeyAscii = 13 Then
SendKeys vbTab
ElseIf KeyAscii = 39 Then
KeyAscii = 0
Beep
End If

End Sub

Private Sub txtnmteknisi_LostFocus()
txtnmteknisi = UCase(txtnmteknisi)
End Sub


