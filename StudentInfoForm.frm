VERSION 5.00
Begin VB.Form StudentInfoForm 
   BackColor       =   &H00FFFFFF&
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
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6840
      ScaleHeight     =   555
      ScaleWidth      =   2715
      TabIndex        =   18
      Top             =   3600
      Width           =   2775
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "操作"
      Height          =   5655
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
         TabIndex        =   27
         Top             =   960
         Width           =   2775
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
            TabIndex        =   28
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
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   2715
         TabIndex        =   20
         Top             =   1560
         Width           =   2775
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
            TabIndex        =   21
            Top             =   120
            Width           =   1335
         End
         Begin VB.Shape Shape3 
            BackColor       =   &H00000000&
            BackStyle       =   1  'Opaque
            Height          =   375
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   120
            Width           =   375
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   2715
         TabIndex        =   16
         Top             =   360
         Width           =   2775
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "个人信息"
      Enabled         =   0   'False
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   6615
      Begin VB.ComboBox StudentNumberCombo 
         Height          =   300
         Left            =   2880
         TabIndex        =   29
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox StudentNumberText 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1200
         TabIndex        =   26
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox CallText 
         Height          =   270
         Left            =   1320
         TabIndex        =   15
         Top             =   2880
         Width           =   1935
      End
      Begin VB.ComboBox SignDay 
         Height          =   300
         Left            =   3120
         TabIndex        =   13
         Top             =   2400
         Width           =   615
      End
      Begin VB.ComboBox SignMonth 
         Height          =   300
         ItemData        =   "StudentInfoForm.frx":0000
         Left            =   2400
         List            =   "StudentInfoForm.frx":0028
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2400
         Width           =   615
      End
      Begin VB.ComboBox SignYear 
         Height          =   300
         Left            =   1320
         TabIndex        =   11
         Top             =   2400
         Width           =   975
      End
      Begin VB.OptionButton SexFOption 
         BackColor       =   &H00FFFFFF&
         Caption         =   "女"
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         Top             =   1440
         Width           =   495
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
      Begin VB.TextBox ClassText 
         Height          =   270
         Left            =   840
         TabIndex        =   6
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox StudentNameText 
         Height          =   270
         Left            =   840
         TabIndex        =   4
         Top             =   960
         Width           =   1815
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
         TabIndex        =   25
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "不得多于或少于11个字符"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   3360
         TabIndex        =   24
         Top             =   2880
         Width           =   1980
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   2040
         TabIndex        =   23
         Top             =   1440
         Width           =   90
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* 不得多于14个字符"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   2760
         TabIndex        =   22
         Top             =   960
         Width           =   1620
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
         TabIndex        =   14
         Top             =   2880
         Width           =   975
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
         TabIndex        =   10
         Top             =   2400
         Width           =   975
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
         TabIndex        =   7
         Top             =   1440
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
         TabIndex        =   5
         Top             =   1920
         Width           =   615
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
         TabIndex        =   3
         Top             =   960
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
Dim LoginUserPermission As Boolean
Dim IsInfoChanged As Boolean
Dim db As ADODB.Connection
Dim rec
Public Sub SetLoginUserID(LUID As String)
    LoginUserID = LUID
End Sub

Private Sub CallText_Change()
    IsInfoChanged = True
End Sub

Private Sub ClassText_Change()
    IsInfoChanged = True
End Sub

Private Sub Form_Load()
    Dim signDate As String
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
    IsInfoChanged = False
    Next i
End Sub

Private Sub Label14_Click()
    Picture4_Click
End Sub

Private Sub Label7_Click()
    Picture1_Click
End Sub

Private Sub Label8_Click()
    Picture2_Click
End Sub

Private Sub Label9_Click()
    Picture3_Click
End Sub

Private Sub Picture1_Click()
    Frame1.Enabled = True
End Sub

Private Sub Picture2_Click()
    If IsInfoChanged Then
        If MsgBox("有未保存的更改，是否退出", vbOKCancel + vbExclamation, "提示") = vbOK Then
            StudentInfoForm.Hide
            MenuForm.SetLoginUserInfo LoginUserID, LoginUserPermission
            MenuForm.Show
            db.Close
            Unload Me
        End If
    Else
        StudentInfoForm.Hide
        MenuForm.SetLoginUserInfo LoginUserID, LoginUserPermission
        MenuForm.Show
        db.Close
        Unload Me
    End If
End Sub

Private Sub Picture3_Click()
    Dim gender As String
    Dim signDate As String
    If StudentNameText.Text = "" Then
        MsgBox "姓名不能为空", vbOKOnly + vbExclamation, "提示"
    ElseIf Len(StudentNameText.Text) > 14 Then
        MsgBox "姓名不合规", vbOKOnly + vbExclamation, "提示"
    End If
    If SexMOption.Value Then
        gender = "男"
    ElseIf SexFOption.Value Then
        gender = "女"
    Else
        MsgBox "性别不能为空", vbOKOnly + vbExclamation, "提示"
    End If
    If Len(CallText.Text) <> 11 Then
        MsgBox "电话格式错误", vbOKOnly + vbExclamation, "提示"
    End If
    signDate = Trim(SignYear.Text) & "/" & Trim(SignMonth.Text) & "/" & Trim(SignDay.Text)
    saveChangeSQL = "UPDATE 借阅者表 SET 姓名 = '" & Trim(StudentNameText.Text) & "', 性别 = '" & gender & "', 入学时间 = #" & signDate & "#, 班级 = '" & Trim(ClassText.Text) & "', 联系电话 = '" & Trim(CallText.Text) & "' WHERE 学生编号 = '" & LoginUserID & "'"
    db.Execute (saveChangeSQL)
    IsInfoChanged = False
End Sub

Private Sub Picture4_Click()
    PasswordChangeForm.SetLoginUserID LoginUserID
    PasswordChangeForm.Show
    StudentInfoForm.Hide
    Unload Me
End Sub

Private Sub SexFOption_Click()
    IsInfoChanged = True
End Sub

Private Sub SexMOption_Click()
    IsInfoChanged = True
End Sub

Private Sub SignDay_Change()
    IsInfoChanged = True
End Sub

Private Sub SignDay_GetFocus()
    SignDay.Clear
    M0 = SignMonth.Text
    If M0 = "1" Or M0 = "3" Or M0 = "5" Or M0 = "7" Or M0 = "8" Or M0 = "10" Or M0 = "12" Then
        For m = 1 To 31
            SignDay.AddItem CStr(m)
        Next m
    ElseIf M0 = "4" Or M0 = "6" Or M0 = "9" Or M0 = "11" Then
        For m = 1 To 30
            SignDay.AddItem CStr(m)
        Next m
    ElseIf M0 = "2" Then
        Dim year As Integer
        year = Val(SignYear.Text)
        If year Mod 4 = 0 And year Mod 100 <> 0 Or year Mod 400 = 0 Then
            For m = 1 To 29
                SignDay.AddItem CStr(m)
            Next m
        Else
            For m = 1 To 28
                SignDay.AddItem CStr(m)
            Next m
        End If
    End If
End Sub

Private Sub SignMonth_Change()
    IsInfoChanged = True
End Sub

Private Sub SignMonth_LostFocus()
    SignDay_GetFocus
End Sub

Private Sub SignYear_Change()
    SignDay_GetFocus
End Sub

Private Sub StudentNameText_Change()
    IsInfoChanged = True
End Sub

