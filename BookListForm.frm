VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form BookListForm 
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   8281
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
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
Attribute VB_Name = "BookListForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As ADODB.Connection
Private Sub Form_Load()
    Move 0, 0
    MSHFlexGrid1.Cols = 11
    Set db = New ADODB.Connection
    db.Open ("provider=microsoft.jet.oledb.4.0;data source=.\data.mdb ")
    getBookListSQL = "SELECT * FROM Õº È±Ì"
    Set rec = New ADODB.Recordset
    rec.CursorLocation = adUseClient
    rec.Open Trim(getBookListSQL), db
    Set ExecuteSQL = rec
    For i = 1 To rec.RecordCount
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
        For j = 0 To 10
            MSHFlexGrid1.TextMatrix(i, j) = Trim(rec.Fields(j))
        Next j
        rec.MoveNext
    Next i
End Sub
