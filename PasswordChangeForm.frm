VERSION 5.00
Begin VB.Form PasswordChangeForm 
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�޸�����"
      Height          =   3735
      Left            =   2040
      TabIndex        =   1
      Top             =   2040
      Width           =   5655
      Begin VB.CommandButton Command2 
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   9
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ȷ��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1080
         TabIndex        =   8
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox VerifyPasswordText 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Top             =   1680
         Width           =   2775
      End
      Begin VB.TextBox NewPasswordText 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox OldPasswordText 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "ȷ�����룺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "�����룺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ԭ���룺"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
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
Attribute VB_Name = "PasswordChangeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LoginUserID As String
Dim IsInfoChanged As Boolean
Dim db As ADODB.Connection
Dim rec
Public Sub SetLoginUserID(LUID As String)
    LoginUserID = LUID
End Sub

Private Sub Command1_Click()
    If OldPasswordText.Text = "" Then
        MsgBox "������ԭ����", vbOKOnly + vbExclamation, "��ʾ"
    ElseIf OldPasswordText.Text = rec.Fields(2) Then
        If NewPasswordText.Text = "" Then
            MsgBox "������������", vbOKOnly + vbExclamation, "��ʾ"
        ElseIf NewPasswordText.Text = rec.Fields(2) Then
            MsgBox "�����벻����ԭ����һ��", vbOKOnly + vbExclamation, "��ʾ"
        ElseIf VerifyPasswordText.Text = NewPasswordText.Text Then
            setPasswordSQL = "UPDATE �����߱� SET ���� = '" & NewPasswordText.Text & "' WHERE ѧ����� = '" & LoginUserID & "'"
            db.Execute (setPasswordSQL)
            MsgBox "�޸����", vbOKOnly + vbInformation, "��ʾ"
        Else
            MsgBox "�������벻һ��", vbOKOnly + vbExclamation, "��ʾ"
        End If
    Else
        MsgBox "ԭ�������", vbOKOnly + vbExclamation, "��ʾ"
    End If
End Sub

Private Sub Command2_Click()
    StudentInfoForm.SetLoginUserID LoginUserID
    StudentInfoForm.Show
    PasswordChangeForm.Hide
    Unload Me
End Sub

Private Sub Form_Load()
    Move 0, 0
    Set db = New ADODB.Connection
    db.Open ("provider=microsoft.jet.oledb.4.0;data source=.\data.mdb ")
    getUserSQL = "SELECT * FROM �����߱� WHERE ѧ����� = '" & LoginUserID & "'"
    Set rec = New ADODB.Recordset
    rec.Open Trim(getUserSQL), db
    Set ExecuteSQL = rec
End Sub
