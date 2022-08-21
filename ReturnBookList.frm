VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form ReturnBookList 
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   3975
      Left            =   0
      TabIndex        =   4
      Top             =   3240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7011
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "�黹ͼ��"
      Height          =   1455
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   6975
      Begin VB.CommandButton Command2 
         Caption         =   "�黹"
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   840
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "��ѯ"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "ѧ�����"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6960
      ScaleHeight     =   555
      ScaleWidth      =   2715
      TabIndex        =   1
      Top             =   1800
      Width           =   2775
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   120
         Width           =   1335
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   375
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
Attribute VB_Name = "ReturnBookList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    MSHFlexGrid1.Clear
    MSHFlexGrid1.Rows = 1
    Dim getBorrowedListSQL As String
    getBorrowedListSQL = "SELECT �軹���.����ID, ͼ���.����, �軹���.ͼ����, �軹���.�������, �軹���.Ӧ������, �軹���.ʵ�ʻ�������, �軹���.�����Ƿ���� FROM ͼ��� INNER JOIN �軹��� ON ͼ���.ͼ���� = �軹���.ͼ���� WHERE �軹���.ѧ��֤�� = """ & Combo1.Text & """"
    Set rec = SQL.Execute(getBorrowedListSQL)
    MSHFlexGrid1.Cols = 7
    For i = 1 To rec.RecordCount
        MSHFlexGrid1.Rows = MSHFlexGrid1.Rows + 1
        For j = 0 To 6
            MSHFlexGrid1.TextMatrix(i, j) = IIf(IsNull(Trim(rec.Fields(j))), "-", Trim(rec.Fields(j)))
        Next j
        rec.MoveNext
    Next i
End Sub

Private Sub Command2_Click()
    Dim SelectedBookLine As Integer: SelectedBookLine = MSHFlexGrid1.RowSel
    Dim SelectedBorrowId As String: SelectedBorrowId = MSHFlexGrid1.TextMatrix(SelectedBookLine, 0)
    Dim SelectedBookId As String: SelectedBookId = MSHFlexGrid1.TextMatrix(SelectedBookLine, 2)
    Dim markAsReturned As String
    markAsReturned = "UPDATE �軹��� SET ʵ�ʻ������� = """ & Date & """ WHERE ����ID = """ & SelectedBorrowId & """"
    Dim updateListInfo As String
    updateListInfo = "UPDATE ͼ��� SET �Ƿ��� = -1 WHERE ͼ���� = """ & SelectedBookId & """"
    Dim getBookISBN As String
    getBookISBN = "SELECT ISBN FROM ͼ��� WHERE ͼ���� = """ & SelectedBookId & """"
    Dim selectedBookISBN As String
    selectedBookISBN = SQL.Execute(getBookISBN).Fields(0)
    Dim updateStroageInfo As String
    updateStroageInfo = "Update �����Ϣ SET ����� = ����� + 1, �ѽ������ = �ѽ������ - 1 WHERE ISBN�� = """ & selectedBookISBN & """"
    MsgBox updateStroageInfo
    SQL.Execute (markAsReturned)
    SQL.Execute (updateListInfo)
    SQL.Execute (updateStroageInfo)
    MsgBox "�黹���", vbOKOnly, "�黹���"
End Sub

Private Sub Form_Load()
    Move 0, 0
    getNumberSQL = "SELECT ѧ����� FROM �����߱� WHERE ����Ա <> True"
    Set rec = SQL.Execute(getNumberSQL)
    Set ExecuteSQL = rec
    While Not rec.EOF
        Combo1.AddItem (Trim(rec.Fields(0)))
        rec.MoveNext
    Wend
End Sub

Private Sub Label8_Click()
    Picture2_Click
End Sub

Private Sub Picture2_Click()
    ReturnBookList.Hide
    MenuForm.Show
    Unload Me
End Sub
