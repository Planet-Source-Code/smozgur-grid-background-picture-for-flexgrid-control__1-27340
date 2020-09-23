VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm 
   Caption         =   "Flex Grid BackGround Picture"
   ClientHeight    =   5490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7380
   Icon            =   "frm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4200
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Change Picture"
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   4920
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "GridLines"
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   4560
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   8
      Text            =   "1"
      Top             =   4800
      Width           =   400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   5
      Text            =   "0"
      Top             =   4560
      Width           =   400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   4
      Text            =   "7"
      Top             =   4800
      Width           =   400
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Text            =   "7"
      Top             =   4560
      Width           =   400
   End
   Begin MSFlexGridLib.MSFlexGrid MyFlex 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   15
      Cols            =   7
      FixedRows       =   0
      FixedCols       =   0
      TextStyleFixed  =   4
      FormatString    =   ""
   End
   Begin PicClip.PictureClip PictureClip1 
      Left            =   1440
      Top             =   480
      _ExtentX        =   8467
      _ExtentY        =   6350
      _Version        =   393216
      Rows            =   4
      Cols            =   4
      Picture         =   "frm.frx":000C
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.bmp"
      DialogTitle     =   "Load Picture"
   End
   Begin VB.Label Label1 
      Caption         =   "Font Color"
      Height          =   255
      Index           =   4
      Left            =   3120
      TabIndex        =   12
      Top             =   4920
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "FixedRows"
      Height          =   255
      Index           =   3
      Left            =   1320
      TabIndex        =   7
      Top             =   4800
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "FixedCols"
      Height          =   255
      Index           =   2
      Left            =   1320
      TabIndex        =   6
      Top             =   4560
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Rows"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   4800
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Cols"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   600
   End
End
Attribute VB_Name = "frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub BackGround(Optional PicFileName As String)
Dim TotColWidth As Integer
Dim TotRowHeight As Integer
Dim i As Integer
Dim j As Integer
    If Not PicFileName = "" Then
        Set PictureClip1.Picture = LoadPicture(PicFileName)
    End If
    MyFlex.Visible = False
    For i = MyFlex.FixedCols To MyFlex.Cols - 1
        TotColWidth = TotColWidth + MyFlex.ColWidth(i) / Screen.TwipsPerPixelX
    Next
    For i = MyFlex.FixedRows To MyFlex.Rows - 1
        TotRowHeight = TotRowHeight + MyFlex.RowHeight(i) / Screen.TwipsPerPixelY
    Next
   
    PictureClip1.Cols = MyFlex.Cols - MyFlex.FixedCols
    PictureClip1.Rows = MyFlex.Rows - MyFlex.FixedRows
   
    PictureClip1.StretchX = TotColWidth / PictureClip1.Cols
    PictureClip1.StretchY = TotRowHeight / PictureClip1.Rows
    MyFlex.Clear
    
    For i = MyFlex.FixedRows To MyFlex.Rows - 1
        MyFlex.Row = i
        For j = MyFlex.FixedCols To MyFlex.Cols - 1
            MyFlex.Col = j
            Set MyFlex.CellPicture = Me.PictureClip1.GraphicCell(j + PictureClip1.Cols * (i - MyFlex.FixedRows) - MyFlex.FixedCols)
        Next j
    Next i
  
    MyFlex.TextMatrix(0, 0) = "Column " & 1
    For j = MyFlex.FixedCols To MyFlex.Cols - 1
        For i = MyFlex.FixedRows To MyFlex.Rows - 1
            MyFlex.TextMatrix(i, j) = 1200 + i * j
        Next i
        MyFlex.TextMatrix(0, j) = "Column " & j + 1
    Next j
    MyFlex.Col = MyFlex.FixedCols
    MyFlex.Row = MyFlex.FixedRows
    MyFlex.Visible = True
End Sub




Private Sub Check1_Click()
    If Check1.Value = Checked Then
        MyFlex.GridLines = flexGridFlat
    Else
        MyFlex.GridLines = flexGridNone
    End If
End Sub

Private Sub Combo1_Click()
    With MyFlex
        Select Case Combo1.ListIndex
        Case 0
            MyFlex.ForeColor = vbBlack
        Case 1
            MyFlex.ForeColor = vbRed
        Case 2
            MyFlex.ForeColor = vbBlue
        Case 3
            MyFlex.ForeColor = vbWhite
        End Select
     End With
End Sub


Private Sub Command1_Click()
    On Error GoTo errHandler
    cmdlg.InitDir = cmdlg.FileName
    cmdlg.FileName = ""
    cmdlg.Action = 1
    Call BackGround(cmdlg.FileName)
errHandler:
End Sub

Private Sub Form_Load()
    Text1(0).Text = MyFlex.Cols
    Text1(1).Text = MyFlex.Rows
    Text1(2).Text = MyFlex.FixedCols
    Text1(3).Text = MyFlex.FixedRows
    Call BackGround
    
    With Combo1
        .AddItem "Black"
        .AddItem "Red"
        .AddItem "Blue"
        .AddItem "White"
        .ListIndex = 0
    End With

    cmdlg.InitDir = App.Path
    cmdlg.Filter = "Pictures (*.bmp;*.jpg;*.gif;*.ico)|*.bmp;*.jpg;*.gif;*.ico"

End Sub




Private Sub Text1_Change(Index As Integer)
    On Error GoTo x:
    If Not IsNumeric(Text1(Index).Text) Or (Index < 2 And Text1(Index).Text <= 0) Or Text1(Index).Text < 0 Then
        Select Case Index
        Case 0
            Text1(Index).Text = MyFlex.Cols
        Case 1
            Text1(Index).Text = MyFlex.Rows
        Case 2
            Text1(Index).Text = MyFlex.FixedCols
        Case 3
            Text1(Index).Text = MyFlex.FixedRows
        End Select
        Exit Sub
    End If
    If Index < 2 Then
        If Val(Text1(Index).Text) < 1 Then Exit Sub
        If Val(Text1(Index).Text) <= Val(Text1(Index + 2)) Then
            Text1(Index).Text = Text1(Index + 2).Text + 1
        End If
    Else
        If Val(Text1(Index).Text) >= Val(Text1(Index - 2)) Then
            Text1(Index).Text = Text1(Index - 2) - 1
        End If
    End If
    MyFlex.FixedCols = Text1(2).Text
    MyFlex.FixedRows = Text1(3).Text
    MyFlex.Cols = Text1(0).Text
    MyFlex.Rows = Text1(1).Text
    Call BackGround
x:
End Sub

