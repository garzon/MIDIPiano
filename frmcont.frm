VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmcont 
   Caption         =   "º¸≈Ãøÿ÷∆…Ë÷√"
   ClientHeight    =   4590
   ClientLeft      =   7995
   ClientTop       =   5445
   ClientWidth     =   9510
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   9510
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   7800
      TabIndex        =   92
      Text            =   "0"
      Top             =   4080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frmcont.frx":0000
      Left            =   5040
      List            =   "frmcont.frx":0184
      Style           =   2  'Dropdown List
      TabIndex        =   88
      TabStop         =   0   'False
      Top             =   3000
      Width           =   3495
   End
   Begin VB.CheckBox Check1 
      Caption         =   "œ˚≥˝”‡“Ù"
      Height          =   255
      Left            =   7320
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   3360
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   7800
      TabIndex        =   86
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   2640
      Width           =   615
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      ItemData        =   "frmcont.frx":07B3
      Left            =   5760
      List            =   "frmcont.frx":07E7
      Style           =   2  'Dropdown List
      TabIndex        =   84
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "»°œ˚"
      Height          =   615
      Left            =   5280
      TabIndex        =   83
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "»∑∂®"
      Height          =   615
      Left            =   2280
      TabIndex        =   82
      Top             =   3840
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   2400
      TabIndex        =   81
      Text            =   "+++1"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1320
      TabIndex        =   79
      Text            =   "---1"
      Top             =   3000
      Width           =   735
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   107
      Left            =   9000
      TabIndex        =   78
      Text            =   "7"
      Top             =   1200
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   111
      Left            =   8160
      TabIndex        =   77
      Text            =   "0"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   106
      Left            =   8520
      TabIndex        =   76
      Text            =   "0"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   109
      Left            =   9000
      TabIndex        =   75
      Text            =   "+7"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   38
      Left            =   6840
      TabIndex        =   74
      Text            =   "0"
      Top             =   1720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   110
      Left            =   8520
      TabIndex        =   73
      Text            =   "-7"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   96
      Left            =   7920
      TabIndex        =   72
      Text            =   "0"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   39
      Left            =   7200
      TabIndex        =   71
      Text            =   "-3"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   40
      Left            =   6840
      TabIndex        =   70
      Text            =   "-2"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   37
      Left            =   6480
      TabIndex        =   69
      Text            =   "-1"
      Top             =   2040
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   34
      Left            =   7200
      TabIndex        =   68
      Text            =   "3"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   35
      Left            =   6840
      TabIndex        =   67
      Text            =   "2"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   46
      Left            =   6480
      TabIndex        =   66
      Text            =   "1"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   33
      Left            =   7200
      TabIndex        =   65
      Text            =   "+3"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   36
      Left            =   6840
      TabIndex        =   64
      Text            =   "+2"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   45
      Left            =   6480
      TabIndex        =   63
      Text            =   "+1"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   105
      Left            =   8520
      TabIndex        =   62
      Text            =   "+6"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   104
      Left            =   8160
      TabIndex        =   61
      Text            =   "+5"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   103
      Left            =   7800
      TabIndex        =   60
      Text            =   "+4"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   102
      Left            =   8520
      TabIndex        =   59
      Text            =   "6"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   101
      Left            =   8160
      TabIndex        =   58
      Text            =   "5"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   100
      Left            =   7800
      TabIndex        =   57
      Text            =   "4"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   99
      Left            =   8520
      TabIndex        =   56
      Text            =   "-6"
      Top             =   1720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   98
      Left            =   8160
      TabIndex        =   55
      Text            =   "-5"
      Top             =   1720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   97
      Left            =   7800
      TabIndex        =   54
      Text            =   "-4"
      Top             =   1720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   32
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   53
      Text            =   "0"
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   222
      Left            =   5160
      TabIndex        =   52
      Text            =   "4"
      Top             =   1365
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   221
      Left            =   5400
      TabIndex        =   51
      Text            =   "+5"
      Top             =   1035
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   219
      Left            =   5040
      TabIndex        =   50
      Text            =   "+4"
      Top             =   1035
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   191
      Left            =   4920
      TabIndex        =   49
      Text            =   "-3"
      Top             =   1720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   190
      Left            =   4560
      TabIndex        =   48
      Text            =   "-2"
      Top             =   1720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   188
      Left            =   4080
      TabIndex        =   47
      Text            =   "-1"
      Top             =   1720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   77
      Left            =   3720
      TabIndex        =   46
      Text            =   "--7"
      Top             =   1720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   78
      Left            =   3360
      TabIndex        =   45
      Text            =   "--6"
      Top             =   1720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   66
      Left            =   2880
      TabIndex        =   44
      Text            =   "--5"
      Top             =   1720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   86
      Left            =   2520
      TabIndex        =   43
      Text            =   "--4"
      Top             =   1720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   67
      Left            =   2160
      TabIndex        =   42
      Text            =   "--3"
      Top             =   1720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   88
      Left            =   1680
      TabIndex        =   41
      Text            =   "--2"
      Top             =   1720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   90
      Left            =   1320
      TabIndex        =   40
      Text            =   "--1"
      Top             =   1720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   186
      Left            =   4680
      TabIndex        =   39
      Text            =   "3"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   76
      Left            =   4320
      TabIndex        =   38
      Text            =   "2"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   75
      Left            =   3960
      TabIndex        =   37
      Text            =   "1"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   74
      Left            =   3600
      TabIndex        =   36
      Text            =   "-7"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   72
      Left            =   3120
      TabIndex        =   35
      Text            =   "-6"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   71
      Left            =   2760
      TabIndex        =   34
      Text            =   "-5"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   70
      Left            =   2400
      TabIndex        =   33
      Text            =   "-4"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   68
      Left            =   1920
      TabIndex        =   32
      Text            =   "-3"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   83
      Left            =   1560
      TabIndex        =   31
      Text            =   "-2"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   65
      Left            =   1080
      TabIndex        =   30
      Text            =   "-1"
      Top             =   1400
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   80
      Left            =   4560
      TabIndex        =   29
      Text            =   "+3"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   79
      Left            =   4200
      TabIndex        =   28
      Text            =   "+2"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   73
      Left            =   3720
      TabIndex        =   27
      Text            =   "+1"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   85
      Left            =   3360
      TabIndex        =   26
      Text            =   "7"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   89
      Left            =   3000
      TabIndex        =   25
      Text            =   "6"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   84
      Left            =   2520
      TabIndex        =   24
      Text            =   "5"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   82
      Left            =   2160
      TabIndex        =   23
      Text            =   "4"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   69
      Left            =   1800
      TabIndex        =   22
      Text            =   "3"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   87
      Left            =   1320
      TabIndex        =   21
      Text            =   "2"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   81
      Left            =   960
      TabIndex        =   20
      Text            =   "1"
      Top             =   1080
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   119
      Left            =   4200
      TabIndex        =   18
      Text            =   "0"
      Top             =   300
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   118
      Left            =   3720
      TabIndex        =   17
      Text            =   "0"
      Top             =   300
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   117
      Left            =   3360
      TabIndex        =   16
      Text            =   "0"
      Top             =   300
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   116
      Left            =   3000
      TabIndex        =   15
      Text            =   "0"
      Top             =   300
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   115
      Left            =   2280
      TabIndex        =   14
      Text            =   "0"
      Top             =   300
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   114
      Left            =   1920
      TabIndex        =   13
      Text            =   "0"
      Top             =   300
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   113
      Left            =   1560
      TabIndex        =   12
      Text            =   "0"
      Top             =   300
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   112
      Left            =   1080
      TabIndex        =   11
      Text            =   "0"
      Top             =   300
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   48
      Left            =   4320
      TabIndex        =   10
      Text            =   "++3"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   57
      Left            =   3960
      TabIndex        =   9
      Text            =   "++2"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   56
      Left            =   3480
      TabIndex        =   8
      Text            =   "++1"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   55
      Left            =   3120
      TabIndex        =   7
      Text            =   "+7"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   54
      Left            =   2760
      TabIndex        =   6
      Text            =   "+6"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   53
      Left            =   2280
      TabIndex        =   5
      Text            =   "+5"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   52
      Left            =   1920
      TabIndex        =   4
      Text            =   "+4"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   51
      Left            =   1560
      TabIndex        =   3
      Text            =   "+3"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   50
      Left            =   1080
      TabIndex        =   2
      Text            =   "+2"
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox ctxt 
      Height          =   270
      Index           =   49
      Left            =   720
      TabIndex        =   1
      Text            =   "+1"
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   2415
      Left            =   120
      Picture         =   "frmcont.frx":0850
      ScaleHeight     =   2355
      ScaleWidth      =   9195
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
   Begin MSComctlLib.Slider vol 
      Height          =   255
      Left            =   840
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   3360
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   393216
      Max             =   127
      SelStart        =   100
      Value           =   100
   End
   Begin VB.Label Label4 
      Caption         =   "“Ùµ˜:"
      Height          =   255
      Left            =   5280
      TabIndex        =   91
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label5 
      Caption         =   "“Ù¡ø£∫"
      Height          =   375
      Left            =   240
      TabIndex        =   90
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "¿÷∆˜:"
      Height          =   255
      Left            =   4560
      TabIndex        =   89
      Top             =   3000
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "»Ìº¸≈Ã∑∂Œß£∫         µΩ "
      Height          =   255
      Left            =   240
      TabIndex        =   80
      Top             =   3000
      Width           =   2895
   End
   Begin VB.Label Label1 
      Caption         =   "ø…”√∑˚∫≈£∫ + - 0 1 2 3 4 5 6 7 # b£¨«“±ÿ–Î“‘0~7Ω· ¯"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2640
      Width           =   9255
   End
End
Attribute VB_Name = "frmcont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Combo2_Click()
If Combo2.ListIndex <> -1 Then Text3 = Combo2.ItemData(Combo2.ListIndex)
End Sub

Private Sub Command1_Click()
  Dim tmin, tmax, a
  tmin = min
  tmax = max
  min = str2id(Text2)
  max = str2id(Text4)
  If (min = 0) Or (max = 0) Or (min > 255) Or (max > 255) Or (min > max) Then MsgBox "º¸≈Ã∑∂Œß”–ŒÛ£¨«Î÷ÿ–¬…Ë÷√°£", , "¥ÌŒÛ": min = tmin: max = tmax: Exit Sub
  For a = tmin To tmax
    Unload Form1.cmd(a)
  Next
  Form1.makekeyboard
  On Error Resume Next
  For a = 1 To 255
    keyset(a) = str2id(ctxt(a))
  Next
  clean
  Form1.Text1 = Text1
  Form1.vol.Value = vol.Value
  Form1.Combo1.ListIndex = Combo1.ListIndex
  Form1.Check1.Value = Check1.Value
  Unload frmcont
End Sub

Private Sub Command2_Click()
Unload frmcont
End Sub

Private Sub Form_Load()
If loading = False Then
  On Error Resume Next
  Dim a
  For a = 1 To 255
    ctxt(a) = id2str(keyset(a), True)
  Next
End If
  Text2 = id2str(min, True)
  Text4 = id2str(max, True)
  Text1 = Form1.Text1
  vol.Value = Form1.vol.Value
  Combo1.ListIndex = Form1.Combo1.ListIndex
  Check1.Value = Form1.Check1.Value
  Text3 = Text1
  If Text3 = 0 Then Combo2.ListIndex = 0 Else Combo2.ListIndex = -1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Form1.Show
End Sub

Private Sub Text1_Change()
If Combo2.ListIndex <> -1 Then If Text1 <> Combo2.ItemData(Combo2.ListIndex) Then Combo2.ListIndex = -1
End Sub

Private Sub Text3_Change()
On Error GoTo here
Text1 = Val(Text3)
Exit Sub
here: Text1 = 0
Combo2.ListIndex = 0
End Sub
