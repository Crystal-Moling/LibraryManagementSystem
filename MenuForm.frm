VERSION 5.00
Begin VB.Form MenuForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "图书借阅管理系统-主菜单"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9765
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   9765
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox BorrowedBook 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3120
      ScaleHeight     =   675
      ScaleWidth      =   6555
      TabIndex        =   14
      Top             =   3480
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Shape Shape7 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "已借阅图书"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   840
         TabIndex        =   15
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.PictureBox BooksList 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3120
      ScaleHeight     =   675
      ScaleWidth      =   6555
      TabIndex        =   12
      Top             =   2760
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "图书列表"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   840
         TabIndex        =   13
         Top             =   120
         Width           =   1815
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox BooksInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3120
      ScaleHeight     =   675
      ScaleWidth      =   6555
      TabIndex        =   10
      Top             =   3480
      Width           =   6615
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "图书管理"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   840
         TabIndex        =   11
         Top             =   120
         Width           =   1815
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox PublisherInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3120
      ScaleHeight     =   675
      ScaleWidth      =   6555
      TabIndex        =   8
      Top             =   2760
      Width           =   6615
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "出版社管理"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   840
         TabIndex        =   9
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.PictureBox PersonalInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3120
      ScaleHeight     =   675
      ScaleWidth      =   6555
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Shape Shape3 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   495
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "个人信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.PictureBox LogOut 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3120
      ScaleHeight     =   675
      ScaleWidth      =   6555
      TabIndex        =   3
      Top             =   4200
      Width           =   6615
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "退出登陆"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   840
         TabIndex        =   5
         Top             =   120
         Width           =   1815
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.PictureBox ReaderInfo 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3120
      ScaleHeight     =   675
      ScaleWidth      =   6555
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   6615
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "借阅者信息"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   840
         TabIndex        =   4
         Top             =   120
         Width           =   2175
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   495
      End
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
Attribute VB_Name = "MenuForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LoginUserPermission As Boolean

Private Sub Form_Load()

    LoginUserPermission = Variables.GetLoginUserPermission

    Move 0, 0
    If LoginUserPermission Then
        PersonalInfo.Visible = False
        BooksList.Visible = False
        BorrowedBook.Visible = False
        
        ReaderInfo.Visible = True
        BooksInfo.Visible = True
        PublisherInfo.Visible = True
    Else
        PersonalInfo.Visible = True
        BooksList.Visible = True
        BorrowedBook.Visible = True
        
        ReaderInfo.Visible = False
        BooksInfo.Visible = False
        PublisherInfo.Visible = False
    End If
End Sub

' Admins menu start

    '' Reader Info
    
        Private Sub ReaderInfo_Click()
            StudentInfoForm.Show
            MenuForm.Hide
        End Sub
        Private Sub Label3_Click()
            ReaderInfo_Click
        End Sub
    
    '' Publisher Info
    
        Private Sub PublisherInfo_Click()
            PublisherForm.Show
            MenuForm.Hide
        End Sub
        Private Sub Label6_Click()
            PublisherInfo_Click
        End Sub
    
    '' Books Info
    
        Private Sub BooksInfo_Click()
            BookListForm.Show
            MenuForm.Hide
        End Sub
        Private Sub Label7_Click()
            BooksInfo_Click
        End Sub
' Admins menu end

' Users menu start

    '' Personal Info
    
        Private Sub PersonalInfo_Click()
            SelfInfoForm.Show
            MenuForm.Hide
        End Sub
        Private Sub Label5_Click()
            PersonalInfo_Click
        End Sub
    
    '' Borrowed book
    
        Private Sub BorrowedBook_Click()
            BorrowedBooksForm.Show
            MenuForm.Hide
        End Sub
        Private Sub Label9_Click()
            BorrowedBook_Click
        End Sub
    
    '' Books list
    
        Private Sub BooksList_Click()
            BookSelectForm.Show
            MenuForm.Hide
        End Sub
        Private Sub Label8_Click()
            BooksList_Click
        End Sub
' Users menu end

' Logout
    Private Sub LogOut_Click()
        If MsgBox("确定要退出登录吗", vbOKCancel + vbQuestion, "提示") = vbOK Then
            Variables.SetLoginUserID ""
            Variables.SetLoginUserPermission False
            MenuForm.Hide
            LoginForm.Show
            Unload Me
        End If
    End Sub
    Private Sub Label4_Click()
        LogOut_Click
    End Sub
