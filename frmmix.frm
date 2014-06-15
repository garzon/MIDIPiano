VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmmix 
   Caption         =   "文件混音"
   ClientHeight    =   1650
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   4515
   Icon            =   "frmmix.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4515
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command4 
      Caption         =   "浏览..."
      Height          =   255
      Left            =   3480
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   720
      TabIndex        =   5
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "混音"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "浏览..."
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "浏览..."
      Height          =   255
      Left            =   3480
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog cd2 
      Left            =   3480
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "*.mrd|*.mrd"
   End
   Begin VB.Label Label2 
      Caption         =   "文件2"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "文件1"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "保存至"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   885
      Width           =   735
   End
End
Attribute VB_Name = "frmmix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
cd2.FileName = ""
cd2.ShowOpen
Text1 = cd2.FileName
End Sub

Private Sub Command2_Click()
cd2.FileName = ""
cd2.ShowOpen
Text2 = cd2.FileName
End Sub


Private Sub Command3_Click()
Dim a1() As Byte, a2() As Byte, p1, p2, a, f1() As Long, f2() As Long, r() As Long, wr() As Byte
Erase a1
Erase a2
If Text1 = "" Then MsgBox "文件路径无效": Exit Sub
If Text2 = "" Then MsgBox "文件路径无效": Exit Sub
Open Text1 For Binary As #1
ReDim a1(0 To LOF(1) - 1)
Get #1, , a1
Close #1
Open Text2 For Binary As #1
ReDim a2(0 To LOF(1) - 1)
Get #1, , a2
Close #1

ReDim f1(1 To (UBound(a1) + 1) \ 4)
ReDim f2(1 To (UBound(a2) + 1) \ 4)
For a = 1 To UBound(f1)
  f1(a) = a1(a * 4 - 4) * &H1000000 + a1(a * 4 - 3) * &H10000 + a1(a * 4 - 2) * &H100 + a1(a * 4 - 1)
Next
For a = 1 To UBound(f2)
  f2(a) = a2(a * 4 - 4) * &H1000000 + a2(a * 4 - 3) * &H10000 + a2(a * 4 - 2) * &H100 + a2(a * 4 - 1)
Next
Erase a1
Erase a2
ReDim r(1 To UBound(f1) + UBound(f2))
p1 = 1
p2 = 1
a = 1
While a <= UBound(r)
  If p1 > UBound(f1) Then
    r(a) = f2(p2)
    r(a + 1) = f2(p2 + 1)
    p2 = p2 + 2
  Else
    If p2 > UBound(f2) Then
      r(a) = f1(p1)
      r(a + 1) = f1(p1 + 1)
      p1 = p1 + 2
    Else
      If f1(p1) <= f2(p2) Then
        r(a) = f1(p1)
        r(a + 1) = f1(p1 + 1)
        p1 = p1 + 2
      Else
        r(a) = f2(p2)
        r(a + 1) = f2(p2 + 1)
        p2 = p2 + 2
      End If
    End If
  End If
  a = a + 2
Wend
ReDim wr(0 To UBound(r) * 4 - 1)
For a = 1 To UBound(r)
  wr(a * 4 - 4) = r(a) \ &H1000000
  wr(a * 4 - 3) = (r(a) \ &H10000) Mod &H100
  wr(a * 4 - 2) = (r(a) \ &H100) Mod &H100
  wr(a * 4 - 1) = r(a) Mod &H100
Next
Open Text3 For Output As #1
Close #1
Open Text3 For Binary As #1
Put #1, , wr
Close #1
MsgBox "混音成功！"
End Sub

Private Sub Command4_Click()
cd2.FileName = ""
cd2.ShowSave
Text3 = cd2.FileName
End Sub
