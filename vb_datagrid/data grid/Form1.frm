VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton buttExcel 
      Caption         =   "Runtime editing"
      Height          =   495
      Left            =   7440
      TabIndex        =   1
      Top             =   240
      Width           =   2055
   End
   Begin Project1.ucGrid ucGrid1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _extentx        =   12726
      _extenty        =   7435
      cols            =   5
      backcolorselected=   12648447
      font            =   "Form1.frx":0000
      font            =   "Form1.frx":0028
   End
   Begin VB.Label Label1 
      Caption         =   "Please visit http://mysource.50webs.com/ for more of my VB projects"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   120
      MouseIcon       =   "Form1.frx":0050
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   4440
      Width           =   7095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_NORMAL = 1


Private Sub buttExcel_Click()
    frmTestExc.Show
End Sub

Private Sub Form_Load()

    Dim i As Long
    Me.ucGrid1.Cols = 5
    Me.ucGrid1.Rows = 1
    
    Me.ucGrid1.Data(0, 1) = "ID"
    Me.ucGrid1.Data(0, 2) = "Name"
    Me.ucGrid1.Data(0, 3) = "Phone"

    Me.ucGrid1.Data(1, 1) = "1"
    Me.ucGrid1.Data(1, 2) = "Ivan Stimac"
    Me.ucGrid1.Data(1, 3) = "111-222-333"
    
    
   ' Me.ucGrid1.AddRow
    
    
'    For i = 1 To Me.ucGrid1.Rows
'        Me.ucGrid1.Data(i, 0) = i
'    Next i

    
End Sub

Private Sub Label1_Click()
        ShellExecute hwnd, "open", "http://mysource.50webs.com", vbNullString, vbNullString, SW_NORMAL
End Sub
