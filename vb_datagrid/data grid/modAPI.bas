Attribute VB_Name = "modAPI"
Option Explicit

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" _
        (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, _
        lpRect As RECT, ByVal wFormat As Long) As Long

'Public Declare Function DrawFocusRect Lib "user32" _
'        (ByVal hdc As Long, lpRect As RECT) As Long
'
'Public Declare Function FillRect Lib "user32" _
'        (ByVal hdc As Long, lpRect As RECT, hBrush As Long) As Long


Global Const DT_TOP = &H0
Global Const DT_LEFT = &H0
Global Const DT_CENTER = &H1
Global Const DT_RIGHT = &H2
Global Const DT_VCENTER = &H4
Global Const DT_BOTTOM = &H8
Global Const DT_WORDBREAK = &H10
Global Const DT_SINGLELINE = &H20
Global Const DT_EXPANDTABS = &H40
Global Const DT_TABSTOP = &H80
Global Const DT_NOCLIP = &H100
Global Const DT_EXTERNALLEADING = &H200
Global Const DT_CALCRECT = &H400
Global Const DT_NOPREFIX = &H800
Global Const DT_INTERNAL = &H1000
Global Const DT_EDITCONTROL = &H2000
Global Const DT_PATH_ELLIPSIS = &H4000
Global Const DT_END_ELLIPSIS = &H8000
Global Const DT_MODIFYSTRING = &H10000
Global Const DT_RTLREADING = &H20000
Global Const DT_WORD_ELLIPSIS = &H40000

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'Private Sub Command1_Click()
'    Dim lSuccess As Long
'    Dim sPrintText As String
'    Dim MyRect As RECT
'    Form1.Font.Size = 12
'    Form1.ScaleMode = vbPixels
'    MyRect.Left = 0
'    MyRect.Right = Form1.ScaleWidth
'    MyRect.Top = 20
'    MyRect.Bottom = 60
'    sPrintText = "Print this text"
'    lSuccess = DrawText(Form1.hdc, sPrintText, Len(sPrintText), _
'    MyRect, DT_CENTER Or DT_WORDBREAK)
'End Sub

