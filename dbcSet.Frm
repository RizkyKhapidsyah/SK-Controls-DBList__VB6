VERSION 5.00
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3645
   ClientLeft      =   1695
   ClientTop       =   2025
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3645
   ScaleWidth      =   5640
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Width           =   3195
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "dbcSet.frx":0000
      Height          =   1935
      Left            =   2280
      OleObjectBlob   =   "dbcSet.frx":0014
      TabIndex        =   3
      Top             =   660
      Width           =   3195
   End
   Begin MSDBCtls.DBCombo DBCombo1 
      Bindings        =   "dbcSet.frx":09E7
      Height          =   315
      Left            =   2280
      TabIndex        =   2
      Top             =   180
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   16777215
      ForeColor       =   0
      Text            =   "DBCombo1"
   End
   Begin MSDBCtls.DBList DBList1 
      Bindings        =   "dbcSet.frx":09FB
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   4577
      _Version        =   393216
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    '
    ' search for Au_ID value
    '
    If FindRec() = False Then
        MsgBox "Unable to Find Requested record!", vbInformation
        Exit Sub
    End If
    '
    DBListUpdate
    DBComboUpdate
    '
End Sub

Private Sub Form_Load()
    '
    ' set up form and controls
    '
    Me.Caption = "Data-Bound List Examples"
    Command1.Caption = "&Search"
    '
    DataControlSetUp
    DBGrid1.Refresh
    DBListSetUp "Author", "Au_ID"
    DBComboSetUp "Author", "Au_ID"
    '
End Sub

Public Function FindRec()
    '
    ' attempt to find a record
    '
    Dim bolReturn As Boolean
    Dim strSearch As String
    '
    ' get input and exit if empty
    strSearch = InputBox("Enter Search String:")
    If Trim(strSearch) = "" Then
        bolReturn = False
        GoTo FindRecExit
    End If
    '
    ' attempt search
    Data1.Recordset.FindFirst strSearch
    If Data1.Recordset.NoMatch = True Then
        bolReturn = False
    Else
        bolReturn = True
    End If
    '
FindRecExit:
    FindRec = bolReturn ' set return value
    '
End Function

Public Sub DBComboUpdate()
    '
    ' update DBCombo control with results
    '
    Dim strDBCUpdate As String
    '
    strDBCUpdate = DBCombo1.ListField
    DBCombo1.Text = Data1.Recordset.Fields(strDBCUpdate)
    '
End Sub

Public Sub DBListUpdate()
    '
    ' update dblist control
    '
    Dim strDBLUpdate As String
    '
    strDBLUpdate = DBList1.BoundColumn
    DBList1.Text = Data1.Recordset.Fields(strDBLUpdate)
    '
End Sub

Public Sub DBListSetUp(strListFld As String, strBndCol As String)
    '
    ' set up dblist control
    '
    DBList1.BoundColumn = strBndCol
    DBList1.ListField = strListFld
    DBList1.Refresh
    '
End Sub

Public Sub DBComboSetUp(strListFld As String, strBndCol As String)
    '
    ' setup dbcombo control
    '
    DBCombo1.BoundColumn = strBndCol
    DBCombo1.ListField = strListFld
    DBCombo1.Refresh
    '
End Sub

Public Sub DataControlSetUp()
    '
    ' set up data control
    '
    Dim strDBName As String
    Dim strRSName As String
    '
    ' point to location of biblio.mdb
    strDBName = App.Path + "\biblio.mdb"
    strRSName = "SELECT * FROM Authors"
    '
    ' setup data control
    Data1.DatabaseName = strDBName
    Data1.RecordSource = strRSName
    Data1.Refresh
    '
End Sub
