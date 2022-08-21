VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form BorrowedBooksForm 
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
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF8080&
      Caption         =   "显示已归还"
      Height          =   255
      Left            =   6840
      TabIndex        =   5
      Top             =   2880
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5175
      Left            =   0
      TabIndex        =   4
      Top             =   2040
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   9128
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6600
      ScaleHeight     =   555
      ScaleWidth      =   3075
      TabIndex        =   2
      Top             =   2040
      Width           =   3135
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   375
      End
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
         TabIndex        =   3
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF8080&
      Height          =   5175
      Left            =   6600
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
Attribute VB_Name = "BorrowedBooksForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    MSHFlexGrid1.Clear
    MSHFlexGrid1.Rows = 1
    Dim getBorrowedListSQL As String
    If Check1.Value Then
        getBorrowedListSQL = "SELECT 图书表.书名, 借还书表.借出日期, 借还书表.应还日期, 借还书表.实际还书日期, 借还书表.还书是否完好 FROM 图书表 INNER JOIN 借还书表 ON 图书表.图书编号 = 借还书表.图书编号 WHERE 借还书表.学生证号 = '" & Variables.GetLoginUserID & "'"
    Else
        getBorrowedListSQL = "SELECT 图书表.书名, 借还书表.借出日期, 借还书表.应还日期, 借还书表.实际还书日期, 借还书表.还书是否完好 FROM 图书表 INNER JOIN 借还书表 ON 图书表.图书编号 = 借还书表.图书编号 WHERE 借还书表.学生证号 = '" & Variables.GetLoginUserID & "' AND 实际还书日 <> ''"
    End If
    Set rec = SQL.Execute(getBorrowedListSQL)
    MSHFlexGrid1.Cols = 5
    For i = 1 To rec.RecordCount
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
        For j = 0 To 4
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

Private Sub Picture2_Click()
    BorrowedBooksForm.Hide
    MenuForm.Show
    Unload Me
End Sub
