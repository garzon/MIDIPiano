VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于我的应用程序"
   ClientHeight    =   3090
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2132.773
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   615
      Left            =   240
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   5
      Top             =   120
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   40
      ImageHeight     =   40
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAbout.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   345
      Left            =   4125
      TabIndex        =   0
      Top             =   2625
      Width           =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   1
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "应用程序标题"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "版本"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "警告: 版权归作者所有"
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   2
      Top             =   2625
      Width           =   3630
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "关于MIDI电子琴"
    lblVersion.Caption = "版本 V2.0"
    lblTitle.Caption = "MIDI电子琴"
    lblDescription.Caption = "作者邮箱：610204679@qq.com，发现bug可以发E-mail联系。"
    Picture1.Picture = ImageList1.ListImages(1).Picture
End Sub
