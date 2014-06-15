VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "MIDI"
   ClientHeight    =   1695
   ClientLeft      =   10110
   ClientTop       =   7005
   ClientWidth     =   5880
   ForeColor       =   &H8000000B&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   5880
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   1200
      Top             =   720
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1680
      Width           =   180
   End
   Begin VB.CheckBox Check1 
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form1.frx":08CA
      Left            =   2520
      List            =   "Form1.frx":0A4E
      Style           =   2  'Dropdown List
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComctlLib.Slider vol 
      Height          =   255
      Left            =   3000
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   1
      Max             =   127
      SelStart        =   100
      Value           =   100
   End
   Begin MSComctlLib.StatusBar Text3 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   2
      Top             =   1380
      Width           =   5880
      _ExtentX        =   10372
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   4200
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   360
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   230
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   480
      Shape           =   3  'Circle
      Top             =   1200
      Width           =   135
   End
   Begin VB.Menu mnuset 
      Caption         =   "设置"
      Begin VB.Menu mnusetup 
         Caption         =   "键盘设置"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "关于..."
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmd_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
  keypress Index, -1, Combo1.ListIndex
End Sub

Private Sub cmd_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Check1.Value = 1 Then keystop Index Else keyrelease Index
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If ifdown(KeyCode) = False Then
  keypress keyset(KeyCode) + Text1, -1, Combo1.ListIndex
  ifdown(KeyCode) = True
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If ifdown(KeyCode) Then
  ifdown(KeyCode) = False
  If Check1.Value = 1 Then keystop keyset(KeyCode) + Text1 Else keyrelease keyset(KeyCode) + Text1
End If
End Sub

Private Sub Form_Load()
loading = True
Dim rc, t, a
rc = midiOutOpen(hmidi, -1, 0, 0, 0)
If rc <> 0 Then
  MsgBox "无法打开设备，错误代号：" & rc
  End
End If
Combo1.ListIndex = 0
ins = -1
min = 36
max = 95
On Error Resume Next
frmcont.Show
For a = 1 To 255
  keyset(a) = str2id(frmcont.ctxt(a).Text)
Next
Unload frmcont
Form1.Show
loading = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
  midiOutClose hmidi
  End
End Sub

Private Sub mnuabout_Click()
  frmAbout.Show
End Sub

Sub makekeyboard()
Dim t, bb
bb = 0
For t = min To max Step 1
 Load cmd(t)
 cmd(t).Top = cmd(min).Top
 cmd(t).Visible = True
 cmd(t).TabStop = False
 If ifwhite(t) Then
  cmd(t).Left = cmd(bb).Left + cmd(0).Width
  cmd(t).BackColor = RGB(255, 255, 255)
  cmd(t).Height = cmd(0).Height
  cmd(t).Width = cmd(0).Width
  cmd(t).ZOrder 1
  bb = t
 Else
  cmd(t).BackColor = RGB(0, 0, 0)
  cmd(t).Height = 700
  cmd(t).Width = 200
  cmd(t).Left = cmd(bb).Left + cmd(bb).Width - cmd(t).Width \ 2
  cmd(t).ZOrder 0
 End If
Next
If (min <= 60) And (max >= 60) Then Shape1.Visible = True: Shape1.Left = cmd(60).Left + 50: Shape1.Top = cmd(0).Height + cmd(0).Top Else Shape1.Visible = False
Form1.Width = cmd(max).Left + cmd(max).Width + 300
If Form1.Width < 6000 Then Form1.Width = 6000
End Sub

Private Sub mnusetup_Click()
  frmcont.Show
  Form1.Hide
End Sub

Private Sub Timer1_Timer()
makekeyboard
Timer1.Enabled = False
End Sub
