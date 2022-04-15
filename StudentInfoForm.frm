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
Public Sub SetLoginUserID(LUID As String)
    LoginUserID = LUID
End Sub

