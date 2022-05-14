VERSION 5.00
Begin VB.Form PublisherForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox PublisherContactText 
      Height          =   270
      Left            =   1080
      TabIndex        =   17
      Top             =   3360
      Width           =   1095
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6840
      ScaleHeight     =   555
      ScaleWidth      =   2715
      TabIndex        =   15
      Top             =   2880
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
         TabIndex        =   16
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "操作"
      Height          =   5655
      Left            =   6840
      TabIndex        =   10
      Top             =   1440
      Width           =   2775
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   0
         ScaleHeight     =   555
         ScaleWidth      =   2715
         TabIndex        =   13
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
            TabIndex        =   14
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
         TabIndex        =   11
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
            TabIndex        =   12
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
      Caption         =   "出版社信息"
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   6615
      Begin VB.TextBox PublisherNameText 
         Height          =   270
         Left            =   1320
         TabIndex        =   19
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox PublisherNumberText 
         Height          =   270
         Left            =   1320
         TabIndex        =   5
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox PublishAddressText 
         Height          =   270
         Left            =   840
         TabIndex        =   4
         Top             =   2400
         Width           =   2655
      End
      Begin VB.TextBox PublisherCallText 
         Height          =   270
         Left            =   1320
         TabIndex        =   3
         Top             =   1440
         Width           =   1935
      End
      Begin VB.ComboBox PublisherNameCombo 
         Height          =   300
         Left            =   1320
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "联系人："
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
         TabIndex        =   18
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "出版社编号："
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
         TabIndex        =   9
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "地址："
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
         TabIndex        =   8
         Top             =   2400
         Width           =   615
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
         TabIndex        =   7
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "出版社名称："
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
         TabIndex        =   6
         Top             =   480
         Width           =   1215
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
Attribute VB_Name = "PublisherForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Dim NewInfo As Boolean
Dim IsInfoChanged As Boolean
Private Sub Form_Load()
    Move 0, 0
    NewInfo = False
    IsInfoChanged = False
    Set db = New ADODB.Connection
    db.Open ("provider=microsoft.jet.oledb.4.0;data source=.\data.mdb ")
    getPublisherNameSQL = "SELECT 出版社名称 FROM 出版社表"
    Set rec = New ADODB.Recordset
    rec.Open Trim(getPublisherNameSQL), db
    Set ExecuteSQL = rec
    While Not rec.EOF
        PublisherNameCombo.AddItem (Trim(rec.Fields(0)))
        rec.MoveNext
    Wend
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
    NewInfo = True
    PublisherNameText.Visible = True
    PublisherNameCombo.Visible = False
End Sub

Private Sub Picture2_Click()
    If IsInfoChanged Then
        If MsgBox("有未保存的更改，是否退出", vbOKCancel + vbExclamation, "提示") = vbOK Then
            PublisherForm.Hide
            db.Close
            MenuForm.Show
            Unload Me
        End If
    Else
        PublisherForm.Hide
        db.Close
        MenuForm.Show
        Unload Me
    End If
End Sub

Private Sub Picture3_Click()
    If Trim(PublisherNumberText.Text) = "" Then
        MsgBox "出版社编号不能为空", vbOKOnly + vbExclamation, "提示"
    Else
        If Trim(PublisherNameCombo.Text) = "" Then
            MsgBox "出版社名称不能为空", vbOKOnly + vbExclamation, "提示"
        Else
            Dim saveChangeSQL As String
            If NewInfo Then
                saveChangeSQL = "INSERT INTO 出版社表 (出版社编号, 出版社名称, 联系电话, 联系人姓名, 地址) VALUES (" & Trim(PublisherNumberText.Text) & ", " & Trim(PublisherNameCombo.Text) & ", " & Trim(PublisherCallText.Text) & ", " & Trim(PublisherContactText.Text) & ", " & Trim(PublishAddressText.Text) & ")"
                NewInfo = True
                PublisherNameText.Visible = False
                PublisherNameCombo.Visible = True
            Else
                saveChangeSQL = "UPDATE 出版社表 SET 出版社编号 = '" & Trim(PublisherNumberText.Text) & "', 出版社名称 = '" & Trim(PublisherNameCombo.Text) & "', 联系电话 = '" & Trim(PublisherCallText.Text) & "', 联系人姓名 = '" & Trim(PublisherContactText.Text) & "', 地址 = '" & Trim(PublishAddressText.Text) & "'"
            End If
            db.Execute (saveChangeSQL)
            IsInfoChanged = False
            Form_Load
        End If
    End If
End Sub

Private Sub PublishAddressText_Change()
    IsInfoChanged = True
End Sub

Private Sub PublisherCallText_Change()
    IsInfoChanged = True
End Sub

Private Sub PublisherNameCombo_Change()
    IsInfoChanged = True
End Sub

Private Sub PublisherNameCombo_LostFocus()
    getPublisherSQL = "SELECT * FROM 出版社表 WHERE 出版社名称 = '" & PublisherNameCombo.Text & "'"
    Set rec = New ADODB.Recordset
    rec.Open Trim(getPublisherSQL), db
    Set ExecuteSQL = rec
    PublisherNumberText.Text = Trim(rec.Fields(0))
    If Trim(rec.Fields(2)) <> "" Then PublisherCallText.Text = Trim(rec.Fields(2))
    If Trim(rec.Fields(3)) <> "" Then PublisherContactText.Text = Trim(rec.Fields(3))
    If Trim(rec.Fields(4)) <> "" Then PublishAddressText.Text = Trim(rec.Fields(4))
    IsInfoChanged = False
End Sub

Private Sub PublisherNumberText_Change()
    IsInfoChanged = True
End Sub
