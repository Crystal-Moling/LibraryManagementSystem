VERSION 5.00
Begin VB.Form StudentInfoForm 
   BorderStyle     =   0  'None
   Caption         =   "图书借阅管理系统-学生信息"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   6840
      ScaleHeight     =   555
      ScaleWidth      =   2715
      TabIndex        =   18
      Top             =   2400
      Width           =   2775
      Begin VB.Label Label8 
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
         TabIndex        =   19
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
   Begin VB.Frame Frame2 
      Caption         =   "操作"
      Height          =   5655
      Left            =   6840
      TabIndex        =   2
      Top             =   1440
      Width           =   2775
      Begin VB.PictureBox Picture1 
         Height          =   615
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   2715
         TabIndex        =   16
         Top             =   360
         Width           =   2775
         Begin VB.Label Label7 
            Caption         =   "修改信息"
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
            TabIndex        =   17
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
   Begin VB.Frame Frame1 
      Caption         =   "个人信息"
      Enabled         =   0   'False
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   6615
      Begin VB.TextBox CallText 
         Height          =   270
         Left            =   1320
         TabIndex        =   15
         Top             =   2400
         Width           =   1935
      End
      Begin VB.ComboBox SignDay 
         Height          =   300
         Left            =   3120
         TabIndex        =   13
         Top             =   1920
         Width           =   615
      End
      Begin VB.ComboBox SignMonth 
         Height          =   300
         ItemData        =   "StudentInfoForm.frx":0000
         Left            =   2400
         List            =   "StudentInfoForm.frx":0028
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   1920
         Width           =   615
      End
      Begin VB.ComboBox SignYear 
         Height          =   300
         Left            =   1320
         TabIndex        =   11
         Top             =   1920
         Width           =   975
      End
      Begin VB.OptionButton SexFOption 
         Caption         =   "女"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton SexMOption 
         Caption         =   "男"
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox ClassText 
         Height          =   270
         Left            =   840
         TabIndex        =   6
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox StudentNameText 
         Height          =   270
         Left            =   840
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "联系电话："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "入学时间："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "性别："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "班级："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "姓名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
   End
End
Attribute VB_Name = "StudentInfoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LoginUserID As String
Dim IsInfoChanged As Boolean
Dim db As ADODB.Connection
Public Sub SetLoginUserID(LUID As String)
    LoginUserID = LUID
End Sub

Private Sub Form_Load()
    Dim SignDate As String
    Move 0, 0
    Set db = New ADODB.Connection
    db.Open ("provider=microsoft.jet.oledb.4.0;data source=.\data.mdb ")
    getUserSQL = "SELECT * FROM 借阅者表 WHERE 学生编号 = '" & LoginUserID & "'"
    Set rec = New ADODB.Recordset
    rec.Open Trim(getUserSQL), db
    Set ExecuteSQL = rec
    StudentNameText.Text = Trim(rec.Fields(3))
    If Trim(rec.Fields(4)) = "男" Then
        SexMOption.Value = True
    ElseIf Trim(rec.Fields(4)) = "女" Then
        SexFOption.Value = True
    End If
    ClassText.Text = Trim(rec.Fields(6))
    SignDate = Trim(rec.Fields(5))
    If Mid(SignDate, 7, 1) = "/" Then
        SignMonth.Text = Mid(SignDate, 6, 1)
        If Mid(SignDate, 9, 1) = "" Then
            SignDay.Text = Right(SignDate, 1)
        Else
            SignDay.Text = Right(SignDate, 2)
        End If
    Else
        SignMonth.Text = Mid(SignDate, 6, 2)
        If Mid(SignDate, 10, 1) = "" Then
            SignDay.Text = Right(SignDate, 1)
        Else
            SignDay.Text = Right(SignDate, 2)
        End If
    End If
    SignYear.Text = Left(SignDate, 4)
    CallText.Text = Trim(rec.Fields(7))
    db.Close
End Sub

Private Sub Label7_Click()
    Picture1_Click
End Sub

Private Sub Label8_Click()
    Picture2_Click
End Sub

Private Sub Picture1_Click()
    Frame1.Enabled = True
End Sub

Private Sub Picture2_Click()
    If IsInfoChanged Then
        If MsgBox("有未保存的更改，是否退出", vbOKCancel + vbExclamation, "提示") = vbOK Then
            StudentInfoForm.Hide
            MenuForm.SetLoginUserID LoginUserID
            MenuForm.Show
            Unload Me
        End If
    Else
        StudentInfoForm.Hide
        MenuForm.SetLoginUserID LoginUserID
        MenuForm.Show
        Unload Me
    End If
End Sub
