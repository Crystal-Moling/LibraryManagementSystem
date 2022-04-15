VERSION 5.00
Begin VB.Form StudentInfoForm 
   Caption         =   "图书借阅管理系统-学生信息"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5805
   ScaleWidth      =   6345
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
