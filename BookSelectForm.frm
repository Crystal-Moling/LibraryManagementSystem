VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form BookSelectForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   3075
      TabIndex        =   15
      Top             =   5760
      Width           =   3135
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "借阅"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   3075
      TabIndex        =   13
      Top             =   6360
      Width           =   3135
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "返回"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   120
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "图书搜索"
      Height          =   2535
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   2895
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   12
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "搜   索"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   8
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "图书分类"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "作者"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "书名"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "图书编号"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5175
      Left            =   3120
      TabIndex        =   2
      Top             =   2040
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   9128
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Height          =   5175
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
   End
End
Attribute VB_Name = "BookSelectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MSHFlexGrid1.Clear
    MSHFlexGrid1.Rows = 1
    Dim getBookListSQL As String
    getBookListSQL = "SELECT * FROM 图书表 "
    Dim Cond As String
    For i = 0 To 3
        If Not Text1(i) = "" Then
            If Cond = "" Then
                Cond = "WHERE 是否借出 = 0 " & Label3(i).Caption & " LIKE ""%" & Text1(i) & "%"""
            Else
                Cond = Cond & " AND " & Label3(i).Caption & " LIKE ""%" & Text1(i) & "%"""
            End If
        End If
    Next i
    If Cond = "" Then
        getBookListSQL = getBookListSQL & "WHERE 是否借出 = 0"
    Else
        getBookListSQL = getBookListSQL & Cond
    End If
    Set rec = SQL.Execute(getBookListSQL)
    MSHFlexGrid1.Cols = 11
    For i = 1 To rec.RecordCount
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
        For j = 0 To 10
            MSHFlexGrid1.TextMatrix(i, j) = Trim(rec.Fields(j))
        Next j
        rec.MoveNext
    Next i
End Sub

Private Sub Form_Load()
 Move 0, 0
End Sub

Private Sub Label8_Click()
    Picture2_Click
End Sub

Private Sub Picture1_Click()
    Dim SelectedBookLine As Integer: SelectedBookLine = MSHFlexGrid1.RowSel
    Dim SelectedBookId As String: SelectedBookId = MSHFlexGrid1.TextMatrix(SelectedBookLine, 0)
    Dim SelectedBookName As String: SelectedBookName = MSHFlexGrid1.TextMatrix(SelectedBookLine, 1)
    Dim SelectedBookWriter As String: SelectedBookWriter = MSHFlexGrid1.TextMatrix(SelectedBookLine, 2)
    Dim SelectedBookISBN As String: SelectedBookISBN = MSHFlexGrid1.TextMatrix(SelectedBookLine, 3)
    Dim MsgBoxPrompt As String: MsgBoxPrompt = "是否要借阅《" + SelectedBookName + "》作者：" + SelectedBookWriter
    If MsgBox(MsgBoxPrompt, vbYesNo, "借阅确认") = vbYes Then
        Dim lendBookSQL As String: lendBookSQL = "INSERT INTO 借还书表 (借书ID, 学生证号, 图书编号, 借出日期, 应还日期) "
        Dim lendBookId As Integer: lendBookId = Val(SQL.Execute("SELECT Count(*) FROM 借还书表").Fields(0)) + 1
        Dim studentId As String: studentId = Variables.GetLoginUserID
        Dim lendTime As String: lendTime = Date
        Dim backTime As String: backTime = DateAdd("m", 1, lendTime)
        lendBookSQL = lendBookSQL & "VALUES (" & lendBookId & ", """ & studentId & """, """ & SelectedBookId & """, " & lendTime & ", " & backTime & ")"
        Dim changeListSQL As String: changeListSQL = "UPDATE 图书表 SET 是否借出 = -1, 借出次数 = 借出次数 + 1 WHERE 图书编号 = """ & SelectedBookId & """"
        Dim changeStroageSQL As String: changeStroageSQL = "UPDATE 库存信息 SET 库存量 = 库存量 - 1, 已借出数量 = 已借出数量 + 1 WHERE ISBN号 = """ & SelectedBookISBN & """"
        SQL.Execute (lendBookSQL)
        SQL.Execute (changeListSQL)
        SQL.Execute (changeStroageSQL)
        MsgBox "借出完成", vbOKOnly, "借出完成"
        Command1_Click
    End If
End Sub

Private Sub Label4_Click()
    Picture1_Click
End Sub

Private Sub Picture2_Click()
    BookSelectForm.Hide
    MenuForm.Show
    Unload Me
End Sub
