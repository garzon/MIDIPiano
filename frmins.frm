VERSION 5.00
Begin VB.Form frmins 
   Caption         =   "使用说明"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   8370
   LinkTopic       =   "Form2"
   ScaleHeight     =   5145
   ScaleWidth      =   8370
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   4560
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   2610
      Left            =   120
      Picture         =   "frmins.frx":0000
      ScaleHeight     =   2550
      ScaleWidth      =   8175
      TabIndex        =   0
      Top             =   120
      Width           =   8235
   End
   Begin VB.Label Label1 
      Caption         =   $"frmins.frx":5511
      Height          =   1695
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   8175
   End
End
Attribute VB_Name = "frmins"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload frmins
End Sub

Private Sub Form_Load()
Me.Icon = Form1.Icon
End Sub
