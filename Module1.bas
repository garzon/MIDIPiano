Attribute VB_Name = "Module1"
Option Explicit

Public Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Public Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long

Public keyset(1 To 255) As Long, ifsound(0 To 255) As Boolean, ifdown(1 To 255) As Boolean, loading As Boolean, arr_record() As Byte, play() As Long, aplay() As Byte
Public min, max, hmidi, ins, recordp, playerp, recordtime As Long, playtime As Long

Public Function color(t)
  If ifwhite(t) Then color = RGB(255, 255, 255) Else color = RGB(0, 0, 0)
End Function

Public Function ifctxtex(ByVal p) As Boolean
On Error GoTo lbl1
If frmcont.ctxt(p) <> "" Then
End If
ifctxtex = True
Exit Function
lbl1:
ifctxtex = False
End Function

Public Function ifwhite(ByVal t) As Boolean
t = t + 1
While t > 12
  t = t - 12
Wend
While t < 1
  t = t + 12
Wend
Select Case t
  Case 1, 3, 5, 6, 8, 10, 12
    ifwhite = True
  Case Else
    ifwhite = False
End Select
End Function

Public Function str2id(ByVal s As String) As Long
Dim st(1 To 7)
Dim a, tmp
  st(1) = 0
  st(2) = 2
  st(3) = 4
  st(4) = 5
  st(5) = 7
  st(6) = 9
  st(7) = 11
  tmp = 0
  If s = "0" Then
    str2id = 0
    Exit Function
  End If
  For a = 1 To Len(s) - 1
    Select Case Mid(s, a, 1)
      Case "+"
        tmp = tmp + 12
      Case "-"
        tmp = tmp - 12
      Case "#"
        tmp = tmp + 1
      Case "b"
        tmp = tmp - 1
      Case Else
        tmp = 0
        GoTo a1
    End Select
  Next
a1:
  str2id = Val(Mid(s, Len(s), 1))
  If (str2id <= 0) Or (str2id > 7) Then
    str2id = 0
    Exit Function
  End If
  str2id = st(str2id) + tmp + 60
  If str2id < 0 Then str2id = 0
End Function

Public Function id2str(ByVal tmp, ByVal ifplus As Boolean) As String
    Dim p, s, a
    p = 0
    s = 0
    If tmp = 0 Then
      id2str = "0"
      Exit Function
    End If
    While tmp < 60
      tmp = tmp + 12
      s = s + 1
    Wend
    While tmp > 71
      tmp = tmp - 12
      p = p + 1
    Wend
    Select Case tmp
      Case 60
        id2str = "1"
      Case 61
        If ifplus Then id2str = "#1" Else id2str = "b2"
      Case 62
        id2str = "2"
      Case 63
        If ifplus Then id2str = "#2" Else id2str = "b3"
      Case 64
        id2str = "3"
      Case 65
        id2str = "4"
      Case 66
        If ifplus Then id2str = "#4" Else id2str = "b5"
      Case 67
        id2str = "5"
      Case 68
        If ifplus Then id2str = "#5" Else id2str = "b6"
      Case 69
        id2str = "6"
      Case 70
        If ifplus Then id2str = "#6" Else id2str = "b7"
      Case 71
        id2str = "7"
    End Select
    For a = 1 To p
      id2str = "+" & id2str
    Next
    For a = 1 To s
      id2str = "-" & id2str
    Next
End Function

Public Sub rec(ByVal time As Long, ByVal action As Byte, ByVal inst As Byte, ByVal vol As Byte, ByVal tone As Byte)
Dim tmp(1 To 8) As Byte, b
tmp(5) = action
tmp(6) = inst
tmp(7) = vol
tmp(8) = tone
tmp(1) = time \ &H1000000
tmp(2) = (time Mod &H1000000) \ &H10000
tmp(3) = (time Mod &H10000) \ &H100
tmp(4) = time Mod &H100
ReDim Preserve arr_record(0 To recordp + 8) As Byte
For b = recordp + 1 To recordp + 8
  arr_record(b) = tmp(b - recordp)
Next
recordp = recordp + 8
End Sub

Public Sub fplay(ByVal action As Byte, ByVal inst As Byte, ByVal vol As Byte, ByVal tone As Byte)
  Select Case action
    Case 0:
      keypress tone, vol, inst
    Case 1:
      keyrelease tone
    Case 2
      keystop tone
  End Select
End Sub

Public Sub keypress(ByVal s As Long, ByVal volu As Long, ByVal inst As Long)
  If s <= 0 Then clean: Exit Sub
  If (volu = -1) And (s = Form1.Text1) Then clean: Exit Sub
  If inst <> ins Then
    ins = inst
    midiOutShortMsg hmidi, &HC0 + inst * &H100
  End If
  If (s < 0) Or (s > 255) Then Exit Sub
  If volu = -1 Then volu = Form1.vol.Value
  midiOutShortMsg hmidi, &H90 + volu * &H10000 + &H100 * s
  ifsound(s) = True
  If (min <= s) And (max >= s) Then Form1.cmd(s).BackColor = RGB(0, 255, 0)
End Sub

Public Sub keyrelease(ByVal s As Long)
  If (s <= 0) Or (s > 255) Then Exit Sub
  If (min <= s) And (max >= s) Then Form1.cmd(s).BackColor = color(s)
End Sub

Public Sub keystop(ByVal s As Long)
  If (s < 0) Or (s > 255) Then Exit Sub
  ifsound(s) = False
  midiOutShortMsg hmidi, &H80 + s * &H100
  If (min <= s) And (max >= s) Then Form1.cmd(s).BackColor = color(s)
End Sub

Public Sub clean()
Dim a
For a = 1 To 255
  If ifsound(a) Then keystop a
Next
For a = min To max
  Form1.cmd(a).BackColor = color(a)
Next
End Sub
