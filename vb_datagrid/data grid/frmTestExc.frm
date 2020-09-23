VERSION 5.00
Begin VB.Form frmTestExc 
   Caption         =   "Runtime editing"
   ClientHeight    =   5175
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8520
   LinkTopic       =   "Form2"
   ScaleHeight     =   5175
   ScaleWidth      =   8520
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbAlign 
      Height          =   315
      Left            =   5880
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton buttUnderline 
      Caption         =   "U"
      Height          =   375
      Left            =   4320
      TabIndex        =   5
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton buttItalic 
      Caption         =   "I"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.CommandButton buttBold 
      Caption         =   "B"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
   Begin VB.ComboBox cmbFont 
      Height          =   315
      Left            =   720
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   120
      Width           =   2295
   End
   Begin Project1.ucGrid ucGrid1 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7858
      BackColorSelected=   12648447
      RowFixedData    =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Align:"
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Font:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu mnuInsert 
      Caption         =   "Insert"
      Begin VB.Menu buttInsRow 
         Caption         =   "Row"
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu buttInsCol 
         Caption         =   "Column"
         Shortcut        =   ^{F2}
      End
   End
End
Attribute VB_Name = "frmTestExc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub buttBold_Click()

    'we dont need to change fixed fileds(row and col name fields)
    If Me.ucGrid1.SelectedCol > 0 And Me.ucGrid1.SelectedRow > 0 Then
    
        Me.ucGrid1.FldFontBold(Me.ucGrid1.SelectedRow, Me.ucGrid1.SelectedCol) = Not Me.ucGrid1.FldFontBold(Me.ucGrid1.SelectedRow, Me.ucGrid1.SelectedCol)
        Me.buttBold.FontBold = Me.ucGrid1.FldFontBold(Me.ucGrid1.SelectedRow, Me.ucGrid1.SelectedCol)
        
    End If
    
End Sub

Private Sub buttInsCol_Click()
    
    Me.ucGrid1.AddCol
    Me.ucGrid1.Data(0, Me.ucGrid1.Cols) = Chr(Asc("A") - 1 + Me.ucGrid1.Cols)
    Me.ucGrid1.FldAlign(0, Me.ucGrid1.Cols) = Center
    
End Sub

Private Sub buttInsRow_Click()
    
    Me.ucGrid1.AddRow
    
End Sub

Private Sub buttItalic_Click()

    'we dont need to change fixed fileds(row and col name fields)
    If Me.ucGrid1.SelectedCol > 0 And Me.ucGrid1.SelectedRow > 0 Then
    
        Me.ucGrid1.FldFontItalic(Me.ucGrid1.SelectedRow, Me.ucGrid1.SelectedCol) = Not Me.ucGrid1.FldFontItalic(Me.ucGrid1.SelectedRow, Me.ucGrid1.SelectedCol)
        Me.buttItalic.FontItalic = Me.ucGrid1.FldFontItalic(Me.ucGrid1.SelectedRow, Me.ucGrid1.SelectedCol)
        
    End If

End Sub

Private Sub buttUnderline_Click()
    
    'we dont need to change fixed fileds(row and col name fields)
    If Me.ucGrid1.SelectedCol > 0 And Me.ucGrid1.SelectedRow > 0 Then
    
        Me.ucGrid1.FldFontUnderline(Me.ucGrid1.SelectedRow, Me.ucGrid1.SelectedCol) = Not Me.ucGrid1.FldFontUnderline(Me.ucGrid1.SelectedRow, Me.ucGrid1.SelectedCol)
        Me.buttUnderline.FontUnderline = Me.ucGrid1.FldFontUnderline(Me.ucGrid1.SelectedRow, Me.ucGrid1.SelectedCol)
        
    End If
    
End Sub

Private Sub cmbAlign_Click()
    
    'we dont need to change fixed fileds(row and col name fields)
    If Me.ucGrid1.SelectedCol > 0 And Me.ucGrid1.SelectedRow > 0 Then
    
        Me.ucGrid1.FldAlign(Me.ucGrid1.SelectedRow, Me.ucGrid1.SelectedCol) = Me.cmbAlign.ListIndex
        
    End If
    
End Sub

Private Sub cmbFont_Click()
    
    'we dont need to change fixed fileds(row and col name fields)
    If Me.ucGrid1.SelectedCol > 0 And Me.ucGrid1.SelectedRow > 0 Then
    
        Me.ucGrid1.FldFontName(Me.ucGrid1.SelectedRow, Me.ucGrid1.SelectedCol) = Me.cmbFont.List(Me.cmbFont.ListIndex)
        
    End If
    
End Sub


Private Sub Form_Load()
    
    Dim i As Integer
    
    For i = 0 To Screen.FontCount - 1
        Me.cmbFont.AddItem Screen.Fonts(i)
    Next i
    Me.cmbFont.ListIndex = 0
    
    Me.cmbAlign.AddItem "Left"
    Me.cmbAlign.AddItem "Center"
    Me.cmbAlign.AddItem "Righth"
    Me.cmbAlign.ListIndex = 0
    
    
    Me.ucGrid1.Data(0, 1) = "A"
    
    'Me.ucGrid1.FldAlign(1, 0) = Right
    Me.ucGrid1.FldAlign(0, 1) = Center
End Sub

Private Sub ucGrid1_Click(LastRow As Long, LastColumn As Long, NewRow As Long, NewColumn As Long)
    
    If NewRow < 1 Or NewColumn < 1 Then Exit Sub
    
    Me.cmbFont.Text = ucGrid1.FldFontName(NewRow, NewColumn)
    Me.buttBold.FontBold = ucGrid1.FldFontBold(NewRow, NewColumn)
    Me.buttItalic.FontItalic = ucGrid1.FldFontItalic(NewRow, NewColumn)
    Me.buttUnderline.FontUnderline = ucGrid1.FldFontUnderline(NewRow, NewColumn)
    
    Me.cmbAlign.ListIndex = ucGrid1.FldAlign(NewRow, NewColumn)
    
End Sub
