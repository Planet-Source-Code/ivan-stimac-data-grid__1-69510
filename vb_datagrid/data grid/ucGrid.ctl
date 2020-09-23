VERSION 5.00
Begin VB.UserControl ucGrid 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000C&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4095
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   ScaleHeight     =   158
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   273
   Begin VB.HScrollBar scrHoriz 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   2295
   End
   Begin VB.VScrollBar scrVert 
      Height          =   975
      Left            =   2280
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox lblCorner 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   2280
      ScaleHeight     =   735
      ScaleWidth      =   855
      TabIndex        =   1
      Top             =   960
      Width           =   855
   End
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Line mLine1 
      BorderStyle     =   2  'Dash
      BorderWidth     =   2
      DrawMode        =   6  'Mask Pen Not
      Visible         =   0   'False
      X1              =   88
      X2              =   88
      Y1              =   8
      Y2              =   112
   End
End
Attribute VB_Name = "ucGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'   # project name  :   data grid
'
'   # project type  :   ActiveX
'
'   # author        :   ivan stimac
'
'   # contact       :   ivan.stimac@po.htnet.hr
'
'   # web           :   http://mysource.50webs.com/
'                           please visit my web page for more of my VB projects
'
'   # price $       :   this project is free for any kind of use
'
'   # other info    :   please vote (if you downloaded this from PSC) if
'                           you like this, and please visit my web page
'                           if you wont see more of my projects


Private Type tp_filed

    fld_data        As String
    fld_isData      As Boolean      '0 - row/col name, 1 - data field
'    fld_isSizable   As Boolean
    fldSz_width     As Long
    fldSz_height    As Long
    fldFnt_bold     As Boolean
    fldFnt_italic   As Boolean
    fldFnt_underln  As Boolean
    fldFnt_name     As String
    fldFnt_align    As Byte         '0 - left, 1 - center, 2 - right
    
End Type

Private selectedRect As RECT, sel_row As Long, sel_col As Long
Private drawedCols As Integer, drawedRows As Integer
Private res_col As Long, res_row As Long, isMDown As Boolean, res_x As Long, res_y As Long
Private txtFocused As Boolean


Private isEditable As Boolean, isSizable As Boolean, autoAddRow As Boolean
Private mRowData As eRowData
Private mFnt As New StdFont
Private gridCol As OLE_COLOR, gridColFix As OLE_COLOR
Private FC As OLE_COLOR, FCSel As OLE_COLOR, fcFixed As OLE_COLOR
Private BC As OLE_COLOR, BCContainer As OLE_COLOR, BCFixed As OLE_COLOR, BCSel As OLE_COLOR

Private lng_cols As Long, lng_rows As Long, first_col As Long, first_row As Long, lng_rows1 As Long

Private arr_grid() As tp_filed

'.........................................................
'....... events ..........................................

Public Event Click(ByRef LastRow As Long, ByRef LastColumn As Long, ByRef NewRow As Long, ByRef NewColumn As Long)
Public Event DblClick()
Public Event ChangeData(ByRef ChangedRow As Long, ByRef ChangedColumn As Long)
Public Event MouseDown(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
Public Event MouseUp(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)
Public Event MouseMove(ByRef Button As Integer, ByRef Shift As Integer, ByRef X As Single, ByRef Y As Single)

Public Event ScrollColumnsToEnd()
Public Event ScrollRowsToEnd()


'----------------------------------------------------
'--- props --- #################################### -
'----------------------------------------------------

Public Property Get RowFixedData() As eRowData
    
    RowFixedData = mRowData
    
End Property

Public Property Let RowFixedData(ByVal nV As eRowData)
    
    mRowData = nV
    Call prv_drawTable
    
    PropertyChanged "RowFixedData"
End Property


Public Property Get Cols() As Long
    
    Cols = lng_cols
    
End Property

Public Property Let Cols(ByVal nV As Long)
    
    If nV < 1 Then
    
        Err.Raise 380
        
    End If
    
    lng_cols = nV
    Call prv_setArr
    Call prv_drawTable
    
    PropertyChanged "Cols"
End Property

'*********
Public Property Get Rows() As Long
    
    Rows = lng_rows
    
End Property

Public Property Let Rows(ByVal nV As Long)
    
    If nV < 1 Then
    
        Err.Raise 380
        
    End If
    
    lng_rows = nV
    Call prv_setArr
    Call prv_drawTable
    
    PropertyChanged "Rows"
    
End Property

'*********
Public Property Get Data(ByVal s_row As Long, ByVal s_col As Long) As String
    
    Data = arr_grid(s_row, s_col).fld_data
    
End Property

Public Property Let Data(ByVal s_row As Long, ByVal s_col As Long, ByVal new_value As String)
    
    If s_row < 0 Or s_row > lng_rows Then
    
        Err.Raise 380
    ElseIf s_col < 0 Or s_col > lng_cols Then
    
        Err.Raise 380
        
    End If
    
    arr_grid(s_row, s_col).fld_data = new_value
    
    If s_row = lng_rows Then
        
        If autoAddRow = True Then
        
            Call AddRow
        
        End If
        
    End If
    
    Call prv_drawTable
    PropertyChanged "Data"
    
End Property

'*********
Public Property Get BackColor() As OLE_COLOR
    
    BackColor = BC
    
End Property

Public Property Let BackColor(ByVal nV As OLE_COLOR)
    
    BC = nV
    Call prv_drawTable
    PropertyChanged "BackColor"
    
End Property

'*********
Public Property Get BackColorFixed() As OLE_COLOR
    
    BackColorFixed = BCFixed
    
End Property

Public Property Let BackColorFixed(ByVal nV As OLE_COLOR)
    
    BCFixed = nV
    Call prv_drawTable
    PropertyChanged "BackColorFixed"
    
End Property

'*********
Public Property Get BackColorSelected() As OLE_COLOR
    
    BackColorSelected = BCSel
    
End Property

Public Property Let BackColorSelected(ByVal nV As OLE_COLOR)
    
    BCSel = nV
    Call prv_drawTable
    PropertyChanged "BackColorSelected"
    
End Property

'*********
Public Property Get BackColorContainer() As OLE_COLOR
    
    BackColorContainer = BCContainer
    
End Property

Public Property Let BackColorContainer(ByVal nV As OLE_COLOR)
    
    BCContainer = nV
    Call prv_drawTable
    PropertyChanged "BackColorContainer"
    
End Property

'*********
Public Property Get ForeColor() As OLE_COLOR
    
    ForeColor = FC
    
End Property

Public Property Let ForeColor(ByVal nV As OLE_COLOR)
    
    FC = nV
    Call prv_drawTable
    PropertyChanged "ForeColor"
    
End Property

'*********
Public Property Get ForeColorFixed() As OLE_COLOR
    
    ForeColorFixed = fcFixed
    
End Property

Public Property Let ForeColorFixed(ByVal nV As OLE_COLOR)
    
    fcFixed = nV
    Call prv_drawTable
    PropertyChanged "ForeColorFixed"
    
End Property

'*********
Public Property Get ForeColorSelected() As OLE_COLOR
    
    ForeColorSelected = FCSel
    
End Property

Public Property Let ForeColorSelected(ByVal nV As OLE_COLOR)
    
    FCSel = nV
    Call prv_drawTable
    PropertyChanged "ForeColorSelected"
    
End Property

'*********
Public Property Get GridColor() As OLE_COLOR
    
    GridColor = gridCol
    
End Property

Public Property Let GridColor(ByVal nV As OLE_COLOR)
    
    gridCol = nV
    Call prv_drawTable
    PropertyChanged "GridColor"
    
End Property

'*********
Public Property Get GridColorFixed() As OLE_COLOR
    
    GridColorFixed = gridColFix
    
End Property

Public Property Let GridColorFixed(ByVal nV As OLE_COLOR)
    
    gridColFix = nV
    Call prv_drawTable
    PropertyChanged "GridColorFixed"
    
End Property

'*********
Public Property Get Editable() As Boolean
    
    Editable = isEditable
    
End Property

Public Property Let Editable(ByVal nV As Boolean)
    
    isEditable = nV
    'Call prv_drawTable
    UserControl.txtEdit.Visible = False
    PropertyChanged "Editable"
    
End Property

'*********
Public Property Get AutoAddNextRow() As Boolean
    
    AutoAddNextRow = autoAddRow
    
End Property

Public Property Let AutoAddNextRow(ByVal nV As Boolean)
    
    autoAddRow = nV
    PropertyChanged "AutoAddNextRow"
    
End Property
'*********
Public Property Get Sizable() As Boolean
    
    Sizable = isSizable
    
End Property

Public Property Let Sizable(ByVal nV As Boolean)
    
    isSizable = nV
    'Call prv_drawTable
    'UserControl.txtEdit.Visible = False
    PropertyChanged "Sizable"
    
End Property


'*********
Public Property Get Font() As StdFont
    
    Set Font = mFnt 'UserControl.Font
    
End Property

Public Property Set Font(ByVal nV As StdFont)
    
    Set mFnt = nV
    Set UserControl.Font = nV
    Call prv_drawTable
    PropertyChanged "Font"
    
End Property

'----------------------------------------------------
'--- data props ---                                 -
'----------------------------------------------------

'*********
Public Property Get SelectedRow() As Long
    
    SelectedRow = sel_row
    
End Property

Public Property Let SelectedRow(ByVal nV As Long)
    
    sel_row = nV
    Call prv_drawTable
    PropertyChanged "SelectedRow"
    
End Property

'*********
Public Property Get SelectedCol() As Long
    
    SelectedCol = sel_col
    
End Property

Public Property Let SelectedCol(ByVal nV As Long)
    
    sel_col = nV
    Call prv_drawTable
    PropertyChanged "SelectedCol"
    
End Property

'*********
'Public Property Get FldIsSizable(ByVal mRow As Long, ByVal mCol As Long) As Boolean
'
'    If mRow < 0 Or mCol < 0 Then
'
'        Err.Raise 380
'
'    ElseIf mRow > lng_rows Or mCol > lng_cols Then
'
'        Err.Raise 380
'
'    End If
'    FldIsSizable = arr_grid(mRow, mCol).fld_isSizable
'
'End Property
'
'Public Property Let FldIsSizable(ByVal mRow As Long, ByVal mCol As Long, ByVal nV As Boolean)
'
'    If mRow < 0 Or mCol < 0 Then
'
'        Err.Raise 380
'
'    ElseIf mRow > lng_rows Or mCol > lng_cols Then
'
'        Err.Raise 380
'
'    End If
'    arr_grid(mRow, mCol).fld_isSizable = nV
'    PropertyChanged "FldIsSizable"
'
'End Property

'*********
Public Property Get FldAlign(ByVal mRow As Long, ByVal mCol As Long) As eAlign
    
    If mRow < 0 Or mCol < 0 Then
    
        Err.Raise 380
        
    ElseIf mRow > lng_rows Or mCol > lng_cols Then
    
        Err.Raise 380
        
    End If
    
    FldAlign = arr_grid(mRow, mCol).fldFnt_align
    
End Property

Public Property Let FldAlign(ByVal mRow As Long, ByVal mCol As Long, ByVal nV As eAlign)
    
    Dim val1 As Byte
    If mRow < 0 Or mCol < 0 Then
    
        Err.Raise 380
        
    ElseIf mRow > lng_rows Or mCol > lng_cols Then
    
        Err.Raise 380
        
    End If
    
    If nV = Left Then
        val1 = 0
    ElseIf nV = Center Then
        val1 = 2
    ElseIf nV = Right Then
        val1 = 1
    End If
    
    arr_grid(mRow, mCol).fldFnt_align = nV
    
    If mRow = sel_row And mCol = sel_col Then
        UserControl.txtEdit.Alignment = val1
    End If
    
    Call prv_drawTable
    PropertyChanged "FldAlign"
    
End Property

'*********
Public Property Get FldFontBold(ByVal mRow As Long, ByVal mCol As Long) As Boolean
    
    If mRow < 0 Or mCol < 0 Then
    
        Err.Raise 380
        
    ElseIf mRow > lng_rows Or mCol > lng_cols Then
    
        Err.Raise 380
        
    End If
    FldFontBold = arr_grid(mRow, mCol).fldFnt_bold
    
End Property

Public Property Let FldFontBold(ByVal mRow As Long, ByVal mCol As Long, ByVal nV As Boolean)
    
    If mRow < 0 Or mCol < 0 Then
    
        Err.Raise 380
        
    ElseIf mRow > lng_rows Or mCol > lng_cols Then
    
        Err.Raise 380
        
    End If
    arr_grid(mRow, mCol).fldFnt_bold = nV
    
    If mRow = sel_row And mCol = sel_col Then
        UserControl.txtEdit.FontBold = nV
    End If
    
    Call prv_drawTable
    PropertyChanged "FldFontBold"
    
End Property

'*********
Public Property Get FldFontItalic(ByVal mRow As Long, ByVal mCol As Long) As Boolean
    
    If mRow < 0 Or mCol < 0 Then
    
        Err.Raise 380
        
    ElseIf mRow > lng_rows Or mCol > lng_cols Then
    
        Err.Raise 380
        
    End If
    FldFontItalic = arr_grid(mRow, mCol).fldFnt_italic
    
End Property

Public Property Let FldFontItalic(ByVal mRow As Long, ByVal mCol As Long, ByVal nV As Boolean)
    
    If mRow < 0 Or mCol < 0 Then
    
        Err.Raise 380
        
    ElseIf mRow > lng_rows Or mCol > lng_cols Then
    
        Err.Raise 380
        
    End If
    arr_grid(mRow, mCol).fldFnt_italic = nV
    
    If mRow = sel_row And mCol = sel_col Then
        UserControl.txtEdit.FontItalic = nV
    End If
    
    Call prv_drawTable
    PropertyChanged "FldFontItalic"
    
End Property

'*********
Public Property Get FldFontUnderline(ByVal mRow As Long, ByVal mCol As Long) As Boolean
    
    If mRow < 0 Or mCol < 0 Then
    
        Err.Raise 380
        
    ElseIf mRow > lng_rows Or mCol > lng_cols Then
    
        Err.Raise 380
        
    End If
    FldFontUnderline = arr_grid(mRow, mCol).fldFnt_underln
    
End Property

Public Property Let FldFontUnderline(ByVal mRow As Long, ByVal mCol As Long, ByVal nV As Boolean)
    
    If mRow < 0 Or mCol < 0 Then
    
        Err.Raise 380
        
    ElseIf mRow > lng_rows Or mCol > lng_cols Then
    
        Err.Raise 380
        
    End If
    arr_grid(mRow, mCol).fldFnt_underln = nV
    
    If mRow = sel_row And mCol = sel_col Then
        UserControl.txtEdit.FontUnderline = nV
    End If
    
    Call prv_drawTable
    PropertyChanged "FldFontUnderline"
    
End Property

'*********
Public Property Get FldFontName(ByVal mRow As Long, ByVal mCol As Long) As String
    
    If mRow < 0 Or mCol < 0 Then
    
        Err.Raise 380
        
    ElseIf mRow > lng_rows Or mCol > lng_cols Then
    
        Err.Raise 380
        
    End If
    FldFontName = arr_grid(mRow, mCol).fldFnt_name
    
End Property

Public Property Let FldFontName(ByVal mRow As Long, ByVal mCol As Long, ByVal nV As String)
    
    If mRow < 0 Or mCol < 0 Then
    
        Err.Raise 380
        
    ElseIf mRow > lng_rows Or mCol > lng_cols Then
    
        Err.Raise 380
        
    End If
    arr_grid(mRow, mCol).fldFnt_name = nV
    
    If mRow = sel_row And mCol = sel_col Then
        UserControl.txtEdit.FontName = nV
    End If
    
    Call prv_drawTable
    PropertyChanged "FldFontName"
    
End Property

'*********
Public Property Get FldWidth(ByVal mRow As Long, ByVal mCol As Long) As Long
    
    If mRow < 0 Or mCol < 0 Then
    
        Err.Raise 380
        
    ElseIf mRow > lng_rows Or mCol > lng_cols Then
    
        Err.Raise 380
        
        
    End If
    FldWidth = arr_grid(mRow, mCol).fldSz_width
    
End Property

Public Property Let FldWidth(ByVal mRow As Long, ByVal mCol As Long, ByVal nV As Long)
    
    If mRow < 0 Or mCol < 0 Then
    
        Err.Raise 380
        
    ElseIf mRow > lng_rows Or mCol > lng_cols Then
    
        Err.Raise 380
        
    ElseIf nV < 5 Then
    
        Err.Raise 380
        
    End If
    arr_grid(mRow, mCol).fldSz_width = nV
    Call prv_drawTable
    PropertyChanged "FldWidth"
    
End Property

'*********
Public Property Get FldHeight(ByVal mRow As Long, ByVal mCol As Long) As Long
    
    If mRow < 0 Or mCol < 0 Then
    
        Err.Raise 380
        
    ElseIf mRow > lng_rows Or mCol > lng_cols Then
    
        Err.Raise 380
        
        
    End If
    FldHeight = arr_grid(mRow, mCol).fldSz_height
    
End Property

Public Property Let FldHeight(ByVal mRow As Long, ByVal mCol As Long, ByVal nV As Long)
    
    If mRow < 0 Or mCol < 0 Then
    
        Err.Raise 380
        
    ElseIf mRow > lng_rows Or mCol > lng_cols Then
    
        Err.Raise 380
        
    ElseIf nV < 5 Then
    
        Err.Raise 380
        
    End If
    arr_grid(mRow, mCol).fldSz_height = nV
    Call prv_drawTable
    PropertyChanged "FldHeight"
    
End Property

'    arr_grid(i, j).fldFnt_bold
'    arr_grid(i, j).fldFnt_italic
'    arr_grid(i, j).fldFnt_name
'    arr_grid(i, j).fldFnt_underln
'    arr_grid(i, j).fldSz_height
'    arr_grid(i, j).fldSz_width

'----------------------------------------------------
'--- subs  --- #################################### -
'----------------------------------------------------

Private Sub prv_setArr()
    
    first_col = 1
    first_row = 1
    lng_rows1 = lng_rows + 50
    ReDim arr_grid(lng_rows1, lng_cols)
    Call prv_setDefValues(0)
    
End Sub

Private Sub prv_setDefValues(Optional first_row As Long = 0)

    Dim i As Long, j As Long
    
    For i = first_row To lng_rows
        
        For j = 0 To lng_cols
            
            With arr_grid(i, j)
            
                .fld_data = "" 'i & ":" & j
                If j > 0 Then
                    .fld_isData = True
                Else
                    .fld_isData = False
                End If
                
'                .fld_isSizable = True
                If j = 0 Then
                    .fldFnt_align = 2
                Else
                    .fldFnt_align = 0
                End If
                .fldFnt_bold = mFnt.Bold
                .fldFnt_italic = mFnt.Italic
                .fldFnt_name = mFnt.Name
                .fldFnt_underln = mFnt.Underline
                .fldSz_height = 20
                .fldSz_width = 70
                
            End With
            
        
        Next j
        
    Next i
    
    
End Sub

Private Sub prv_drawTable()
        
    Dim rowDataPrint As String
    Dim mX As Long, mY  As Long
    Dim i As Long, j As Long
    Dim mRect As RECT
    
    drawedCols = 0
    drawedRows = 0
    
    With UserControl
    
        .Cls
        .BackColor = BCContainer
        .txtEdit.BackColor = BCSel
        .txtEdit.ForeColor = FCSel
        .FillColor = BC
        .ForeColor = FC
    
    End With
    'draw data
    mY = arr_grid(0, 0).fldSz_height
    
    
    For i = first_row To lng_rows
        
        mX = arr_grid(0, 0).fldSz_width
        For j = first_col To lng_cols
            
            With UserControl
                
                .Font.Bold = arr_grid(i, j).fldFnt_bold
                .Font.Italic = arr_grid(i, j).fldFnt_italic
                .Font.Name = arr_grid(i, j).fldFnt_name
                .Font.Underline = arr_grid(i, j).fldFnt_underln
                
            End With
            
            If i <> sel_row Or j <> sel_col Then
                
                UserControl.FillColor = BC
                UserControl.ForeColor = FC
                
            Else
            
                UserControl.FillColor = BCSel
                UserControl.ForeColor = FCSel
                selectedRect.Left = mX
                selectedRect.Right = mX + arr_grid(0, j).fldSz_width
                selectedRect.Top = mY
                selectedRect.Bottom = mY + arr_grid(i, 0).fldSz_height
            
            End If
            
            With mRect
                .Left = mX + 2
                .Top = mY + 2
                .Right = mX + arr_grid(0, j).fldSz_width - 2
                .Bottom = mY + arr_grid(0, 0).fldSz_height - 2
            End With
            
            UserControl.Line (mX, mY)-(mX + arr_grid(0, j).fldSz_width, mY + arr_grid(i, 0).fldSz_height), gridCol, B
            mX = mX + arr_grid(0, j).fldSz_width
            
            If mX > UserControl.ScaleWidth Then
                Exit For
            End If
            
            Select Case arr_grid(i, j).fldFnt_align
            
                Case 0
                    Call DrawText(UserControl.hdc, arr_grid(i, j).fld_data, Len(arr_grid(i, j).fld_data), mRect, DT_LEFT Or DT_VCENTER)
                    
                Case 1
                    Call DrawText(UserControl.hdc, arr_grid(i, j).fld_data, Len(arr_grid(i, j).fld_data), mRect, DT_CENTER Or DT_VCENTER)
                    
                Case 2
                    Call DrawText(UserControl.hdc, arr_grid(i, j).fld_data, Len(arr_grid(i, j).fld_data), mRect, DT_RIGHT Or DT_VCENTER)
            
            End Select
            
            'MsgBox ""
        
        Next j
        
        mY = mY + arr_grid(i, 0).fldSz_height
        If mY > UserControl.ScaleHeight Then
            Exit For
        End If
        
    Next i
    
    mX = 0
    mY = 0
    
    'draw col names
    UserControl.FillColor = BCFixed
    UserControl.ForeColor = fcFixed
    For i = 0 To lng_cols
        
        drawedCols = drawedCols + 1
        If i = 1 Then
            i = first_col
        End If
        
        With UserControl
                
            .Font.Bold = arr_grid(0, i).fldFnt_bold
            .Font.Italic = arr_grid(0, i).fldFnt_italic
            .Font.Name = arr_grid(0, i).fldFnt_name
            .Font.Underline = arr_grid(0, i).fldFnt_underln
            
        End With
            
            
        With mRect
            .Left = mX + 2
            .Top = mY + 2
            .Right = mX + arr_grid(0, i).fldSz_width - 2
            .Bottom = mY + arr_grid(0, 0).fldSz_height - 2
        End With
        
        UserControl.Line (mX, mY)-(mX + arr_grid(0, i).fldSz_width, mY + arr_grid(0, 0).fldSz_height), gridColFix, B
        mX = mX + arr_grid(0, i).fldSz_width

            
        Select Case arr_grid(0, i).fldFnt_align
            
                Case 0
                    Call DrawText(UserControl.hdc, arr_grid(0, i).fld_data, Len(arr_grid(0, i).fld_data), mRect, DT_LEFT Or DT_VCENTER)
                    
                Case 1
                    Call DrawText(UserControl.hdc, arr_grid(0, i).fld_data, Len(arr_grid(0, i).fld_data), mRect, DT_CENTER Or DT_VCENTER)
                    
                Case 2
                    Call DrawText(UserControl.hdc, arr_grid(0, i).fld_data, Len(arr_grid(0, i).fld_data), mRect, DT_RIGHT Or DT_VCENTER)
            
        End Select
            
        If mX > UserControl.ScaleWidth - UserControl.scrVert.Width Then
            'MsgBox "IDE"
            drawedCols = drawedCols - 1
            Exit For
        End If
            
    Next i
    
    'draw row names
    mX = 0
    mY = arr_grid(0, 0).fldSz_height
    For i = first_row To lng_rows
        
        drawedRows = drawedRows + 1
        With UserControl
                
            .Font.Bold = arr_grid(i, 0).fldFnt_bold
            .Font.Italic = arr_grid(i, 0).fldFnt_italic
            .Font.Name = arr_grid(i, 0).fldFnt_name
            .Font.Underline = arr_grid(i, 0).fldFnt_underln
            
        End With
        
        With mRect
            .Left = mX + 2
            .Top = mY + 2
            .Right = mX + arr_grid(0, 0).fldSz_width - 2
            .Bottom = mY + arr_grid(i, 0).fldSz_height - 2
        End With
        
        UserControl.Line (mX, mY)-(mX + arr_grid(0, 0).fldSz_width, mY + arr_grid(i, 0).fldSz_height), gridColFix, B
        mY = mY + arr_grid(i, 0).fldSz_height
        
        If mRowData = UserDefined Then
            
            rowDataPrint = arr_grid(i, 0).fld_data
            
        ElseIf mRowData = SelectedPointer Then
            
            If i = sel_row And i > 0 Then
            
                rowDataPrint = "*"
                
            Else
            
                rowDataPrint = vbNullString
            
            End If
            
        ElseIf mRowData = Number Then
        
            If i > 0 Then
                rowDataPrint = Str$(i)
            Else
                rowDataPrint = vbNullString
            End If
            
        End If
        
        Select Case arr_grid(i, 0).fldFnt_align
            
                Case 0
                    Call DrawText(UserControl.hdc, rowDataPrint, Len(rowDataPrint), mRect, DT_LEFT Or DT_VCENTER)
                    
                Case 1
                    Call DrawText(UserControl.hdc, rowDataPrint, Len(rowDataPrint), mRect, DT_CENTER Or DT_VCENTER)
                    
                Case 2
                    Call DrawText(UserControl.hdc, rowDataPrint, Len(rowDataPrint), mRect, DT_RIGHT Or DT_VCENTER)
            
        End Select
        
        If mY > UserControl.ScaleHeight - UserControl.scrHoriz.Height Then
            drawedRows = drawedRows - 1
            Exit For
        End If
            
    Next i
    
    'set scrollbars
    
    'MsgBox drawedCols
    If drawedCols - 1 < lng_cols Then
        
        UserControl.scrHoriz.Enabled = True
        UserControl.scrHoriz.Min = 1 '- first_col
        UserControl.scrHoriz.Max = lng_cols '- drawedCols + 1 + (1 - first_col)
        UserControl.scrHoriz.Value = first_col
        
    Else
    
        UserControl.scrHoriz.Enabled = False
        
    End If
    
'    MsgBox drawedRows
    If drawedRows < lng_rows Then
        'MsgBox drawedRows
        UserControl.scrVert.Enabled = True
        UserControl.scrVert.Min = 1 '- first_row
        UserControl.scrVert.Max = lng_rows '- drawedRows + (1 - first_row)
        UserControl.scrVert.Value = first_row
        
    Else
    
        UserControl.scrVert.Enabled = False
        
    End If
    
    
End Sub

Private Sub prv_setEditingBox()

    If isEditable = True Then
        If arr_grid(sel_row, sel_col).fld_isData = True Then
            
            With UserControl.txtEdit
                
                .Left = selectedRect.Left + 2
                .Top = selectedRect.Top + 2
                .Height = selectedRect.Bottom - selectedRect.Top - 2
                .Width = selectedRect.Right - selectedRect.Left - 2
                .Text = arr_grid(sel_row, sel_col).fld_data
                .SelStart = 0
                .SelLength = Len(.Text)
                .Visible = True
                .SetFocus
                .Font.Bold = arr_grid(sel_row, sel_col).fldFnt_bold
                .Font.Name = arr_grid(sel_row, sel_col).fldFnt_name
                .Font.Italic = arr_grid(sel_row, sel_col).fldFnt_italic
                .Font.Underline = arr_grid(sel_row, sel_col).fldFnt_underln
                
            End With
            
        End If
    End If

End Sub


Public Sub AddRow()

    Dim arr_tmp() As tp_filed
    Dim i As Long, j As Long

    lng_rows = lng_rows + 1
    
    If lng_rows1 < lng_rows Then
        ReDim arr_tmp(lng_rows, lng_cols)
    
        For i = 0 To lng_rows - 1
        
            For j = 0 To lng_cols
            
                arr_tmp(i, j) = arr_grid(i, j)
                
            Next j
        
        Next i
        
        lng_rows1 = lng_rows + 50
        ReDim arr_grid(lng_rows1, lng_cols)
        
        For i = 0 To lng_rows - 1
        
            For j = 0 To lng_cols
            
                arr_grid(i, j) = arr_tmp(i, j)
                
            Next j
        
        Next i
        
        
        
    End If
'    MsgBox lng_rows & vbCrLf & UBound(arr_grid, 1)
    Call prv_setDefValues(lng_rows)
    Call prv_drawTable
    Erase arr_tmp
    DoEvents
End Sub

Public Sub AddCol()
    
    Dim i As Long
    lng_cols = lng_cols + 1
    'MsgBox UBound(arr_grid, 2)
    ReDim Preserve arr_grid(lng_rows1, lng_cols)
    
    For i = 0 To lng_rows
    
        arr_grid(i, lng_cols).fldFnt_name = mFnt.Name
        
        If i > 0 Then
            arr_grid(i, lng_cols).fld_isData = True
        Else
            arr_grid(i, lng_cols).fld_isData = False
        End If
        
        arr_grid(i, lng_cols).fldSz_height = 20
        arr_grid(i, lng_cols).fldSz_width = 70
        
    Next i
    Call prv_drawTable
    DoEvents
    
End Sub

'Public Sub FillWithFields()
'
'    Do While scrHoriz.Enabled = False
'
'        Call AddCol
'
'    Loop
'
'
'    Do While scrVert.Enabled = False
'
'        Call AddRow
'
'    Loop
'
'End Sub


Private Sub scrHoriz_Change()
    On Error Resume Next
    'MsgBox scrHoriz.Value
    first_col = scrHoriz.Value  '+ 1
    Call prv_drawTable
    UserControl.txtEdit.Visible = False
    UserControl.lblCorner.SetFocus
    
    If scrHoriz.Value = scrHoriz.Max Then
    
        RaiseEvent ScrollColumnsToEnd
        
    End If
    
End Sub


Private Sub scrHoriz_Scroll()
    
    scrHoriz_Change
    
End Sub

Private Sub scrVert_Change()
    
    On Error Resume Next
    first_row = scrVert.Value '+ 1
    Call prv_drawTable
    UserControl.txtEdit.Visible = False
    UserControl.lblCorner.SetFocus
    
    If scrVert.Value = scrVert.Max Then
    
        RaiseEvent ScrollRowsToEnd
        
    End If
    
    
End Sub

Private Sub scrVert_Scroll()
    
    scrVert_Change
    
End Sub

Private Sub txtEdit_GotFocus()

    txtFocused = True
    
End Sub

Private Sub txtEdit_LostFocus()

    txtFocused = False
    
End Sub

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode < 37 Or KeyCode > 40 Then
        arr_grid(sel_row, sel_col).fld_data = txtEdit.Text
        
        If sel_row = lng_rows Then
            
            If autoAddRow = True Then
            
                Call AddRow
            
            End If
            
        End If
        
        RaiseEvent ChangeData(sel_row, sel_col)
    End If
    
End Sub





Private Sub UserControl_DblClick()
    
    RaiseEvent DblClick
    
End Sub


Private Sub UserControl_Initialize()
    
    sel_row = 0
    sel_row = 0
    
    first_col = 1
    first_row = 1
    lng_cols = 1
    lng_rows = 1
    UserControl_InitProperties
    
    Call prv_setArr
    Call prv_drawTable
    
End Sub

'----------------------------------------------------
'--- func  --- #################################### -
'----------------------------------------------------

Private Function prv_getColFromPos(ByVal x_pos As Long) As Long
    
    Dim i As Long, mX As Long
    mX = 0
    If x_pos >= mX Then
        
        If x_pos <= mX + arr_grid(0, i).fldSz_width Then
            
            selectedRect.Left = mX
            selectedRect.Right = mX + arr_grid(0, i).fldSz_width
            prv_getColFromPos = i
            Exit Function
            
        End If
        mX = mX + arr_grid(0, i).fldSz_width
        
    End If
        
    mX = arr_grid(0, 0).fldSz_width
    
    For i = first_col To lng_cols
    
        If x_pos >= mX Then
        
            If x_pos <= mX + arr_grid(0, i).fldSz_width Then
                
                selectedRect.Left = mX
                selectedRect.Right = mX + arr_grid(0, i).fldSz_width
                prv_getColFromPos = i
                Exit Function
                
            End If
            mX = mX + arr_grid(0, i).fldSz_width
            
        End If
        
    Next i
    prv_getColFromPos = -1
    
End Function

Private Function prv_getColForResize(ByVal x_pos As Long) As Long
    
    Dim i As Long, mX As Long
    mX = 0
    
    i = 0
    If x_pos >= mX Then
        
        If x_pos <= mX + arr_grid(0, i).fldSz_width + 2 Then
            
            If x_pos >= mX + arr_grid(0, i).fldSz_width - 2 Then
                prv_getColForResize = i
                res_x = mX
                'MsgBox mX
            
            Else
                prv_getColForResize = -1
                
            End If
            Exit Function
            
        End If
        'mX = mX + arr_grid(0, i).fldSz_width
        
    End If
        
    mX = arr_grid(0, 0).fldSz_width
    For i = first_col To lng_cols
    
        If x_pos >= mX Then
        
            If x_pos <= mX + arr_grid(0, i).fldSz_width + 2 Then
                
                If x_pos >= mX + arr_grid(0, i).fldSz_width - 2 Then
                    prv_getColForResize = i
                    res_x = mX
                
                Else
                    prv_getColForResize = -1
                    
                End If
                Exit Function
                
            End If
            mX = mX + arr_grid(0, i).fldSz_width
            
        End If
        
    Next i
    prv_getColForResize = -1
    
End Function

Private Function prv_getRowFromPos(ByVal y_pos As Long) As Long
    
    Dim i As Long, mY As Long
    mY = 0
    If y_pos >= mY Then
        
        If y_pos <= mY + arr_grid(i, 0).fldSz_height Then
            
            selectedRect.Top = mY
            selectedRect.Bottom = mY + arr_grid(i, 0).fldSz_height
            prv_getRowFromPos = i
            Exit Function
            
        End If
        mY = mY + arr_grid(i, 0).fldSz_height
        
    End If
        
    mY = arr_grid(0, 0).fldSz_height
    
    For i = first_row To lng_rows
    
        If y_pos >= mY Then
        
            If y_pos <= mY + arr_grid(i, 0).fldSz_height Then
                
                selectedRect.Top = mY
                selectedRect.Bottom = mY + arr_grid(i, 0).fldSz_height
                prv_getRowFromPos = i
                Exit Function
                
            End If
            mY = mY + arr_grid(i, 0).fldSz_height
            
        End If
        
    Next i
    prv_getRowFromPos = -1
    
End Function

Private Function prv_getRowForResize(ByVal y_pos As Long) As Long
    
    Dim i As Long, mY As Long
    mY = 0
    
    i = 0
    If y_pos >= mY Then
        
        If y_pos <= mY + arr_grid(i, 0).fldSz_height + 2 Then
            
            If y_pos >= mY + arr_grid(i, 0).fldSz_height - 2 Then
                prv_getRowForResize = i
                res_y = mY
                
            Else
                prv_getRowForResize = -1
                
            End If
            Exit Function
            
        End If
        'mY = mY + arr_grid(i, 0).fldSz_height
        
    End If
        
    mY = arr_grid(0, 0).fldSz_height
    For i = first_row To lng_rows
    
        If y_pos >= mY Then
        
            If y_pos <= mY + arr_grid(i, 0).fldSz_height + 2 Then
                
                If y_pos >= mY + arr_grid(i, 0).fldSz_height - 2 Then
                    prv_getRowForResize = i
                    res_y = mY
                    
                Else
                    prv_getRowForResize = -1
                    
                End If
                Exit Function
                
            End If
            mY = mY + arr_grid(i, 0).fldSz_height
            
        End If
        
    Next i
    prv_getRowForResize = -1
    
End Function


Private Sub UserControl_InitProperties()
    
    gridCol = &HC0C0C0
    gridColFix = vbBlack
    FC = vbBlack
    FCSel = vbBlack
    fcFixed = vbBlack
    BC = vbWhite
    BCSel = vbWhite
    BCFixed = vbButtonFace
    BCContainer = vbApplicationWorkspace
    isEditable = True
    isSizable = True
    autoAddRow = True
    mRowData = SelectedPointer
'    Set UserControl.Font = Ambient.Font
    'Set mFnt = Ambient.Font
    
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyLeft Then
        
        If sel_col > 1 Then
            
           ' MsgBox UserControl.txtEdit.SelLength
            If UserControl.txtEdit.Visible = True Then
                
                If UserControl.txtEdit.SelStart <= 1 Or UserControl.txtEdit.SelLength > 0 Then
                    UserControl.txtEdit.Visible = False
                    sel_col = sel_col - 1
                    UserControl.txtEdit.Text = arr_grid(sel_row, sel_col).fld_data
                    Call prv_drawTable
                    Call prv_setEditingBox
                End If
                
            Else
            
                sel_col = sel_col - 1
                UserControl.txtEdit.Text = arr_grid(sel_row, sel_col).fld_data
                Call prv_drawTable
                Call prv_setEditingBox
                
            End If
            
        End If
        
    ElseIf KeyCode = vbKeyRight Then
    
        If sel_col < lng_cols Then
            If UserControl.txtEdit.Visible = True Then
                    
                If UserControl.txtEdit.SelStart = Len(UserControl.txtEdit.Text) Or UserControl.txtEdit.SelLength > 0 Then
                    UserControl.txtEdit.Visible = False
                    sel_col = sel_col + 1
                    UserControl.txtEdit.Text = arr_grid(sel_row, sel_col).fld_data
                    Call prv_drawTable
                    Call prv_setEditingBox
                End If
                    
            Else
            
                sel_col = sel_col + 1
                UserControl.txtEdit.Text = arr_grid(sel_row, sel_col).fld_data
                Call prv_drawTable
                Call prv_setEditingBox
                
            End If
        End If
    
    ElseIf KeyCode = vbKeyUp Then
    
        If sel_row > 1 Then
            
            UserControl.txtEdit.Visible = False
            sel_row = sel_row - 1
            UserControl.txtEdit.Text = arr_grid(sel_row, sel_col).fld_data
            Call prv_drawTable
            Call prv_setEditingBox
            
        End If
        
    ElseIf KeyCode = vbKeyDown Then
    
        If sel_row < lng_rows Then
            
            UserControl.txtEdit.Visible = False
            sel_row = sel_row + 1
            UserControl.txtEdit.Text = arr_grid(sel_row, sel_col).fld_data
            Call prv_drawTable
            Call prv_setEditingBox
            
        End If
        
'    Else
'
'        If isEditable = True Then
'            If arr_grid(sel_row, sel_col).fld_isData = True Then
'
'                With UserControl.txtEdit
'
'                    .Left = selectedRect.Left + 2
'                    .Top = selectedRect.Top + 2
'                    .Height = selectedRect.Bottom - selectedRect.Top - 2
'                    .Width = selectedRect.Right - selectedRect.Left - 2
'                    .Text = arr_grid(sel_row, sel_col).fld_data
'                    .SelStart = 0
'                    .SelLength = Len(.Text)
'                    .Visible = True
'                    .SetFocus
'                    .Font.Bold = arr_grid(sel_row, sel_col).fldFnt_bold
'                    .Font.Name = arr_grid(sel_row, sel_col).fldFnt_name
'                    .Font.Italic = arr_grid(sel_row, sel_col).fldFnt_italic
'                    .Font.Underline = arr_grid(sel_row, sel_col).fldFnt_underln
'
'                End With
'
'            End If
'        End If
        
    End If
    
    
   ' If txtFocused = False Then
            
    
            
        'End If
    
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If res_col > -1 Or res_row > -1 Then
    
        If isSizable = True Then
            isMDown = True
        End If
        
    End If
    RaiseEvent MouseDown(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If isMDown = False Then
    
        res_col = prv_getColForResize(X)
        res_row = prv_getRowForResize(Y)
        
        If res_col > -1 Then
        
            UserControl.MousePointer = 9
            res_row = -1
            
        ElseIf res_row > -1 Then
            
            UserControl.MousePointer = 7
            
        Else
            
            UserControl.MousePointer = 0
            
        End If
        
    Else
    
        If res_col > -1 Then
            
            UserControl.mLine1.X1 = X '- res_x
            UserControl.mLine1.X2 = X '- res_x
            UserControl.mLine1.Y1 = 0
            UserControl.mLine1.Y2 = arr_grid(0, 0).fldSz_height
            UserControl.mLine1.Visible = True
            
            'MsgBox res_x
            If X - res_x > 10 Then
                arr_grid(0, res_col).fldSz_width = X - res_x '* 2
            End If
            
        ElseIf res_row > -1 Then
        
            UserControl.mLine1.X1 = 0 '- res_x
            UserControl.mLine1.X2 = arr_grid(0, 0).fldSz_width '- res_x
            UserControl.mLine1.Y1 = Y
            UserControl.mLine1.Y2 = Y
            UserControl.mLine1.Visible = True
            
            'MsgBox res_x
            If Y - res_y > 10 Then
                arr_grid(res_row, 0).fldSz_height = Y - res_y '* 2
            End If
            
        End If
        
        
    End If
        
        
    
    RaiseEvent MouseMove(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim old_row As Long, old_col As Long
    'resize
    If isMDown = True Then
    
        isMDown = False
        UserControl.mLine1.Visible = False
        
        UserControl.txtEdit.Visible = False
        
    Else
        
        old_row = sel_row
        old_col = sel_col
        sel_row = prv_getRowFromPos(Y)
        sel_col = prv_getColFromPos(X)
        
        'MsgBox sel_row & vbCrLf & sel_col
        RaiseEvent Click(old_row, old_col, sel_row, sel_col)
        
        UserControl.txtEdit.Visible = False
        If isEditable = True Then
            
            If sel_row > 0 And sel_col > 0 Then
                
                If arr_grid(sel_row, sel_col).fld_isData = True Then
                    
                    With UserControl.txtEdit
                        
                        .Left = selectedRect.Left + 2
                        .Top = selectedRect.Top + 2
                        .Height = selectedRect.Bottom - selectedRect.Top - 2
                        .Width = selectedRect.Right - selectedRect.Left - 2
                        .Text = arr_grid(sel_row, sel_col).fld_data
                        .SelStart = 0
                        .SelLength = Len(.Text)
                        .Visible = True
                        .SetFocus
                        .Font.Bold = arr_grid(sel_row, sel_col).fldFnt_bold
                        .Font.Name = arr_grid(sel_row, sel_col).fldFnt_name
                        .Font.Italic = arr_grid(sel_row, sel_col).fldFnt_italic
                        .Font.Underline = arr_grid(sel_row, sel_col).fldFnt_underln
                        
                    End With
                    
                End If
                
            End If
            
        End If
        
    End If
    
    
    
    
    'MsgBox sel_row & vbCrLf & sel_col
    
    
    
    Call prv_drawTable
    RaiseEvent MouseUp(Button, Shift, X, Y)
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    lng_rows = PropBag.ReadProperty("Rows", 1)
    lng_cols = PropBag.ReadProperty("Cols", 1)
    
    BC = PropBag.ReadProperty("BackColor", vbWhite)
    BCFixed = PropBag.ReadProperty("BackColorFixed", vbButtonFace)
    BCSel = PropBag.ReadProperty("BackColorSelected", vbWhite)
    BCContainer = PropBag.ReadProperty("BackColorContainer", vbApplicationWorkspace)
    
    FC = PropBag.ReadProperty("ForeColor", vbBlack)
    fcFixed = PropBag.ReadProperty("ForeColorFixed", vbBlack)
    FCSel = PropBag.ReadProperty("ForeColorSelected", vbBlack)
    
    gridCol = PropBag.ReadProperty("GridColor", &HC0C0C0)
    gridColFix = PropBag.ReadProperty("GridColorFixed", vbBlack)
    
    isEditable = PropBag.ReadProperty("Editable", True)
    isSizable = PropBag.ReadProperty("Sizable", True)
    autoAddRow = PropBag.ReadProperty("AutoAddNextRow", True)
    
    mRowData = PropBag.ReadProperty("RowFixedData", 2)
    
    
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set mFnt = PropBag.ReadProperty("Font", Ambient.Font)
    
    Call prv_setArr
    Call prv_drawTable
    
End Sub

Private Sub UserControl_Resize()

    With UserControl
        
        .scrHoriz.Top = .ScaleHeight - .scrHoriz.Height
        .scrHoriz.Width = .ScaleWidth - .scrVert.Width
        
        .scrVert.Left = .ScaleWidth - .scrVert.Width
        .scrVert.Height = .ScaleHeight - .scrHoriz.Height
        
        .lblCorner.Left = .scrHoriz.Width
        .lblCorner.Top = .scrVert.Height
        
    End With
    
End Sub

Private Sub UserControl_Terminate()
    
    Set mFnt = Nothing
    Erase arr_grid
    
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    
    PropBag.WriteProperty "Rows", lng_rows, 1
    PropBag.WriteProperty "Cols", lng_cols, 1
    
    PropBag.WriteProperty "BackColor", BC, vbWhite
    PropBag.WriteProperty "BackColorFixed", BCFixed, vbButtonFace
    PropBag.WriteProperty "BackColorSelected", BCSel, vbWhite
    PropBag.WriteProperty "BackColorContainer", BCContainer, vbApplicationWorkspace
    
    PropBag.WriteProperty "ForeColor", FC, vbBlack
    PropBag.WriteProperty "ForeColorFixed", fcFixed, vbBlack
    PropBag.WriteProperty "ForeColorSelected", FCSel, vbBlack
    
    PropBag.WriteProperty "GridColor", gridCol, &HC0C0C0
    PropBag.WriteProperty "Editable", isEditable, True
    PropBag.WriteProperty "Sizable", isSizable, True
    PropBag.WriteProperty "AutoAddNextRow", autoAddRow, True
    
    PropBag.WriteProperty "RowFixedData", mRowData, 2
    
    PropBag.WriteProperty "ForeColorSelected", FCSel, vbBlack
    
    PropBag.WriteProperty "Font", mFnt, Ambient.Font
    
End Sub
