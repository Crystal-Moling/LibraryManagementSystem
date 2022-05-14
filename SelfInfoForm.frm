VERSION 5.00
Begin VB.Form SelfInfoForm 
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
      Height          =   5655
      Left            =   6840
      TabIndex        =   17
      Top             =   1440
      Width           =   2775
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   2715
         TabIndex        =   20
         Top             =   960
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
            TabIndex        =   21
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   2715
         TabIndex        =   18
         Top             =   360
         Width           =   2775
         Begin VB.Shape Shape4 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   375
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "修改密码"
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
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "个人信息"
      Enabled         =   0   'False
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   6615
      Begin VB.TextBox StudentNameText 
         Height          =   270
         Left            =   840
         TabIndex        =   10
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox ClassText 
         Height          =   270
         Left            =   840
         TabIndex        =   9
         Top             =   1920
         Width           =   1095
      End
      Begin VB.OptionButton SexMOption 
         BackColor       =   &H00FFFFFF&
         Caption         =   "男"
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   1440
         Width           =   495
      End
      Begin VB.OptionButton SexFOption 
         BackColor       =   &H00FFFFFF&
         Caption         =   "女"
         Height          =   255
         Left            =   1440
         TabIndex        =   7
         Top             =   1440
         Width           =   495
      End
      Begin VB.ComboBox SignYear 
         Height          =   300
         Left            =   1320
         TabIndex        =   6
         Top             =   2400
         Width           =   975
      End
      Begin VB.ComboBox SignMonth 
         Height          =   300
         ItemData        =   "SelfInfoForm.frx":0000
         Left            =   2400
         List            =   "SelfInfoForm.frx":0028
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2400
         Width           =   615
      End
      Begin VB.ComboBox SignDay 
         Height          =   300
         Left            =   3120
         TabIndex        =   4
         Top             =   2400
         Width           =   615
      End
      Begin VB.TextBox CallText 
         Height          =   270
         Left            =   1320
         TabIndex        =   3
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox StudentNumberText 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   16
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   15
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   14
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   13
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
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
         TabIndex        =   12
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "学生编号："
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
         TabIndex        =   11
         Top             =   480
         Width           =   975
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
Attribute VB_Name = "SelfInfoForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LoginUserID As String
Private Sub Form_Load()

    LoginUserID = Variables.GetLoginUserID
    
    Move 0, 0
    Set db = New ADODB.Connection
    db.Open ("provider=microsoft.jet.oledb.4.0;data source=.\data.mdb ")
    getUserSQL = "SELECT * FROM 借阅者表 WHERE 学生编号 = '" & LoginUserID & "'"
    Set rec = New ADODB.Recordset
    rec.Open Trim(getUserSQL), db
    Set ExecuteSQL = rec
    StudentNumberText.Text = Trim(rec.Fields(0))
    StudentNameText.Text = Trim(rec.Fields(3))
    If Trim(rec.Fields(4)) = "男" Then
        SexMOption.Value = True
    ElseIf Trim(rec.Fields(4)) = "女" Then
        SexFOption.Value = True
    End If
    ClassText.Text = Trim(rec.Fields(6))
    signDate = Trim(rec.Fields(5))
    SignYear.Text = Left(signDate, 4)
    If Mid(signDate, 7, 1) = "/" Then
        SignMonth.Text = Mid(signDate, 6, 1)
        If Mid(signDate, 9, 1) = "" Then
            SignDay.Text = Right(signDate, 1)
        Else
            SignDay.Text = Right(signDate, 2)
        End If
    Else
        SignMonth.Text = Mid(signDate, 6, 2)
        If Mid(signDate, 10, 1) = "" Then
            SignDay.Text = Right(signDate, 1)
        Else
            SignDay.Text = Right(signDate, 2)
        End If
    End If
    CallText.Text = Trim(rec.Fields(7))
    db.Close
End Sub

Private Sub Label14_Click()
    Picture4_Click
End Sub

Private Sub Label8_Click()
    Picture2_Click
End Sub

Private Sub Picture2_Click()
    StudentInfoForm.Hide
    MenuForm.Show
    Unload Me
End Sub

Private Sub Picture4_Click()
    PasswordChangeForm.Show
    StudentInfoForm.Hide
    Unload Me
End Sub
