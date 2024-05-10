VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   600
      Left            =   4590
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   630
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   1350
      TabIndex        =   0
      Top             =   585
      Width           =   870
   End
   Begin VB.Timer Timer1 
      Left            =   1170
      Top             =   2115
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As New FileSystemObject

Private Sub Command1_Click()

Call fso.MoveFile("D:\a\cetak_qr.xls", "D:\a\dok dispro\abcde.xls")


'MsgBox fso.GetParentFolderName("D:\a\cetak_qr.xls")
End Sub
