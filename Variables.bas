Attribute VB_Name = "Variables"
Dim LoginUserID As String
Dim LoginUserPermission As Boolean

Public Function SetLoginUserID(LUID As String)
    LoginUserID = LUID
End Function

Public Function SetLoginUserPermission(Permission As Boolean)
    LoginUserPermission = Permission
End Function

Public Function GetLoginUserID() As String
    GetLoginUserID = LoginUserID
End Function

Public Function GetLoginUserPermission() As Boolean
    GetLoginUserPermission = LoginUserPermission
End Function

