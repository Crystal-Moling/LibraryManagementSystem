VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form BookListForm 
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "操作"
      Height          =   4695
      Left            =   6840
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   2715
         TabIndex        =   9
         Top             =   1440
         Width           =   2775
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "图书归还"
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
            TabIndex        =   10
            Top             =   120
            Width           =   1335
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   375
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   2715
         TabIndex        =   7
         Top             =   2040
         Width           =   2775
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
            TabIndex        =   8
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   2715
         TabIndex        =   5
         Top             =   240
         Width           =   2775
         Begin VB.Shape Shape3 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "保存信息"
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
            TabIndex        =   6
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   2715
         TabIndex        =   3
         Top             =   840
         Width           =   2775
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "新建信息"
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
            TabIndex        =   4
            Top             =   120
            Width           =   1335
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   375
         End
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4575
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8070
      _Version        =   393216
      BackColor       =   16777215
      Rows            =   1
      FixedRows       =   0
      BackColorBkg    =   16777215
      ScrollTrack     =   -1  'True
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
Attribute VB_Name = "BookListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Move 0, 0
    MSHFlexGrid1.Cols = 11
    getBookListSQL = "SELECT * FROM 图书表"
    Set rec = SQL.Execute(getBookListSQL)
    For i = 1 To rec.RecordCount
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
        For j = 0 To 10
            MSHFlexGrid1.TextMatrix(i, j) = Trim(rec.Fields(j))
        Next j
        rec.MoveNext
    Next i
End Sub

Private Sub Label2_Click()
    Picture4_Click
End Sub

Private Sub Label8_Click()
    Picture2_Click
End Sub

Private Sub Picture2_Click()
    BookListForm.Hide
    MenuForm.Show
    Unload Me
End Sub

Private Sub Picture4_Click()
    BookListForm.Hide
    ReturnBookList.Show
    Unload Me
End Sub
