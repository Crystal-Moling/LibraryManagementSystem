VERSION 5.00
Begin VB.Form LoginForm 
   BorderStyle     =   0  'None
   Caption         =   "登陆"
   ClientHeight    =   7215
   ClientLeft      =   7140
   ClientTop       =   5790
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton ExitButton 
      Caption         =   "退出"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   4440
      Width           =   735
   End
   Begin VB.CommandButton LoginButton 
      Caption         =   "登陆"
      Default         =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   4440
      Width           =   735
   End
   Begin VB.TextBox PasswordTextbox 
      Appearance      =   0  'Flat
      Height          =   270
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   3960
      Width           =   1575
   End
   Begin VB.TextBox UsernameTextbox 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   3960
      TabIndex        =   4
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "密码："
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   3
      Top             =   3720
      Width           =   615
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "用户名："
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   2
      Top             =   3120
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF8080&
      Height          =   5175
      Left            =   0
      TabIndex        =   1
      Top             =   2040
      Width           =   9735
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF8080&
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9855
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
    Move 0, 0
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
                    MenuForm.SetLoginUserID Trim(rec.Fields(0))
                    UsernameTextbox.Text = ""
                    PasswordTextbox.Text = ""
                    MenuForm.Show
                    'StudentInfoForm.Show
                    LoginForm.Hide
                    Unload Me
                    db.Close
                Else
                    MsgBox "用户名或密码错误", vbOKCancel + vbExclamation, "警告"
                End If
            End If
        End If
    End If
End Sub
