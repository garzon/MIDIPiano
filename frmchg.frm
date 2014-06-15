VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmchg 
   Caption         =   "转换mrd文件"
   ClientHeight    =   1245
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "frmchg.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   1245
   ScaleWidth      =   4680
   StartUpPosition =   2  '屏幕中心
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frmchg.frx":08CA
      Left            =   720
      List            =   "frmchg.frx":08FE
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   840
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog cd3 
      Left            =   3120
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.txt|*.txt"
   End
   Begin VB.CommandButton Command4 
      Caption         =   "浏览..."
      Height          =   255
      Left            =   3600
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   720
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "转换"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "浏览..."
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog cd2 
      Left            =   3960
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.mrd|*.mrd"
   End
   Begin VB.Label Label2 
      Caption         =   "转换至                   大调简谱"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   885
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "mrd文件"
      Height          =   375
      Left            =   50
      TabIndex        =   6
      Top             =   150
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "保存至"
      Height          =   375
      Left            =   150
      TabIndex        =   5
      Top             =   525
      Width           =   735
   End
End
Attribute VB_Name = "frmchg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
cd2.FileName = ""
cd2.ShowOpen
Text1 = cd2.FileName
End Sub

Private Sub Command3_Click()
If Text1 = "" Then MsgBox "文件路径无效": Exit Sub
If Text3 = "" Then MsgBox "文件路径无效": Exit Sub
Dim a() As Byte, f() As Long, b, r As String, tmp, p, s, ifplus
If Combo1.ListIndex <= 7 Then ifplus = False Else ifplus = True
Open Text1 For Binary As #1
ReDim a(0 To LOF(1) - 1)
Get #1, , a
Close #1
ReDim f(1 To (UBound(a) + 1) \ 4)
For b = 1 To UBound(f)
  f(b) = a(b * 4 - 4) * &H1000000 + a(b * 4 - 3) * &H10000 + a(b * 4 - 2) * &H100 + a(b * 4 - 1)
Next
Erase a
b = 2
p = 0
s = Combo1.List(Combo1.ListIndex)
r = Left$(s, InStr(s, "(") - 1) & "大调简谱：" & vbCrLf
While b <= UBound(f)
  If f(b) \ &H1000000 = 0 Then
    tmp = f(b) Mod &H100
    tmp = tmp - Combo1.ItemData(Combo1.ListIndex)
    r = r & id2str(tmp, ifplus) & " "
    p = p + 1
    If p = 8 Then
      p = 0
      r = r & vbCrLf
    End If
  End If
  b = b + 2
Wend
Erase f
Open Text3 For Output As #1
Print #1, r
Close #1
MsgBox "转换成功"
End Sub

Private Sub Command4_Click()
cd3.FileName = ""
cd3.ShowSave
Text3 = cd3.FileName
End Sub

Private Sub Form_Load()
  Combo1.ListIndex = 0
End Sub
