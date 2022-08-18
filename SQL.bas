Attribute VB_Name = "SQL"
Dim db As ADODB.Connection

Public Function Execute(ByVal SQLSentence As String) As ADODB.Recordset
    Set db = New ADODB.Connection
    'db.Open ("provider=microsoft.jet.oledb.4.0;data source=.\data.mdb ")
    db.Open ("provider=microsoft.jet.oledb.4.0;data source=F:\LibraryManagementSystem-\data.mdb ")
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient
    rec.Open Trim(SQLSentence), db
    Set Execute = rec
End Function
