VERSION 5.00
Begin VB.Form LoginForm 
   BorderStyle     =   0  'None
   Caption         =   "登陆"
   ClientHeight    =   2025
   ClientLeft      =   7140
   ClientTop       =   5790
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2025
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ExitButton 
      Caption         =   "退出"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton LoginButton 
      Caption         =   "登陆"
      Default         =   -1  'True
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox PasswordTextbox 
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   1080
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox UsernameTextbox 
      Height          =   270
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "密码："
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "用户名："
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection

Private Sub ExitButton_Click()
    If MsgBox("确定要退出吗", vbOKCancel + vbQuestion, "注意") = vbOK Then
        End
    End If
End Sub

Private Sub Form_Load()
    Set db = New ADODB.Connection
    db.Open ("provider=microsoft.jet.oledb.4.0;data source=.\data.mdb ")
End Sub

Private Sub LoginButton_Click()
    If UsernameTextbox.Text = "" Then
        MsgBox "请输入用户名", vbExclamation + vbOKOnly, "警告"
    Else
        If PasswordTextbox.Text = "" Then
            MsgBox "请输入密码", vbExclamation + vbOKOnly, "警告"
        Else
            getUserSQL = "SELECT * FROM 借阅者表 WHERE 用户名 = '" & UsernameTextbox.Text & "'"
            Set rec = New ADODB.Recordset
            rec.Open Trim(getUserSQL), db
            Set ExecuteSQL = rec
            If rec.EOF Then
                MsgBox "用户名或密码错误", vbOKCancel + vbExclamation, "警告"
                UsernameTextbox.SetFocus
            Else
                If Trim(rec.Fields(2)) = Trim(PasswordTextbox.Text) Then
                    StudentInfoForm.SetLoginUserID Trim(rec.Fields(0))
                    'MainForm.Show
                    StudentInfoForm.Show
                    LoginForm.Hide
                    db.Close
                Else
                    MsgBox "用户名或密码错误", vbOKCancel + vbExclamation, "警告"
                End If
            End If
        End If
    End If
End Sub
