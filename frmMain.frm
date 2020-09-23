VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Update your tables from the old Database to the New One!"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9930
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4005
   ScaleWidth      =   9930
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   3540
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2220
      Width           =   3915
   End
   Begin VB.TextBox Text3 
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   3960
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   660
      Width           =   3915
   End
   Begin VB.TextBox Text2 
      Height          =   390
      Left            =   1980
      TabIndex        =   2
      Text            =   "db2.mdb"
      Top             =   1740
      Width           =   7875
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   2400
      TabIndex        =   1
      Text            =   "db1.mdb"
      Top             =   180
      Width           =   7455
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   3420
      Width           =   3075
   End
   Begin VB.Label Label4 
      Caption         =   "Password DB2:"
      Height          =   315
      Left            =   1980
      TabIndex        =   8
      Top             =   2220
      Width           =   1515
   End
   Begin VB.Label Label3 
      Caption         =   "Password DB1:"
      Height          =   315
      Left            =   2400
      TabIndex        =   7
      Top             =   660
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Update From OLD:"
      Height          =   315
      Left            =   60
      TabIndex        =   4
      Top             =   180
      Width           =   1875
   End
   Begin VB.Label Label2 
      Caption         =   "To NEW:"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1740
      Width           =   1755
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Code developed by Roberto Dalmonte (robdal@tiscalinet.it) and Alessio Deiana (alessio.deiana@tiscalinet.it)

Dim mstrConnectionStringDB1 As String ' Old Database Connection String
Dim mstrConnectionStringDB2 As String ' New Database Connection String

Dim mstrJetPasswordDB1 As String
Dim mstrJetPasswordDB2 As String

Dim mastrTables() As String 'this array contains the tables that are present in both db1 and db2.
Dim mastrTablesDB1() As String 'this array contains all the tables in DB1
Dim mastrTablesDB2() As String 'this array contains all the tables in DB2

Dim mstrTableToBeUpdated As String

Dim mstrOldTable As String
Dim mstrErrorMsg As String

Private Sub cmdUpdate_Click()
   Dim intLoop As Integer
   
   'initialize connection strings
   mstrJetPasswordDB1 = JetPassword & Text3.Text & ";"
   mstrJetPasswordDB2 = JetPassword & Text4.Text & ";"
   
   'do you have a password? Change the connection string accordingly.
   If Text3.Text = vbNullString Then
      mstrConnectionStringDB1 = Provider & "Data Source = " & App.Path & "\" & Text1.Text
   Else
      mstrConnectionStringDB1 = Provider & "Data Source = " & App.Path & "\" & "NoPWD" & Text1.Text & ";" & mstrJetPasswordDB1
      'Yes there is a password. Then remove it.
      Call RemovePassword(App.Path & "\" & Text1.Text, Text3.Text & ";")
   End If
   If Text4.Text = vbNullString Then
      mstrConnectionStringDB2 = Provider & "Data Source = " & App.Path & "\" & Text2.Text
   Else
      mstrConnectionStringDB2 = Provider & "Data Source = " & App.Path & "\" & "NoPWD" & Text2.Text & ";" & mstrJetPasswordDB2
      'Yes there is a password. Then remove it.
      Call RemovePassword(App.Path & "\" & Text2.Text, Text4.Text & ";")
   End If
   
   Call LoadTablesInArrays
   
   For intLoop = 0 To UBound(mastrTables())
      mstrTableToBeUpdated = mastrTables(intLoop)
      Call ImportTable 'this will import the tables from the copies of the databases without password
      Call CreateQuery
   Next intLoop
      
   Call DeleteOriginalDatabases
   
   If Text3.Text = vbNullString Then
   Else
      Call PutThePasswordBack(App.Path & "\" & "NoPWD" & Text1.Text, Text3.Text)
   End If
   
   If Text4.Text = vbNullString Then
   Else
      Call PutThePasswordBack(App.Path & "\" & "NoPWD" & Text2.Text, Text4.Text)
   End If
   
   Call DeleteNoPWDDatabases
   
   If mstrErrorMsg = vbNullString Then
      MsgBox "Update Succesfully Terminated"
   Else
      MsgBox mstrErrorMsg
   End If
End Sub

Private Sub DeleteOriginalDatabases()
   If FileExists(App.Path & "\" & Text1.Text) And FileExists(App.Path & "\" & "NoPWD" & Text1.Text) Then
      Kill App.Path & "\" & Text1.Text
   End If
   If FileExists(App.Path & "\" & Text2.Text) And FileExists(App.Path & "\" & "NoPWD" & Text2.Text) Then
      Kill App.Path & "\" & Text2.Text
   End If
End Sub

Private Sub DeleteNoPWDDatabases()
   If FileExists(App.Path & "\" & Text1.Text) And FileExists(App.Path & "\" & "NoPWD" & Text1.Text) Then
      Kill App.Path & "\" & "NoPWD" & Text1.Text
   End If
   If FileExists(App.Path & "\" & Text2.Text) And FileExists(App.Path & "\" & "NoPWD" & Text2.Text) Then
      Kill App.Path & "\" & "NoPWD" & Text2.Text
   End If
End Sub

Private Sub LoadTablesInArrays()
   Call TablesDB1
   Call TablesDB2
   Call Merge2StringArraysInOne
End Sub

Private Sub Merge2StringArraysInOne()
   'This routine merge the 2 arrays containing the table names and merge it in the 3rd array (mastrTables)
   Dim intLoop1 As Integer
   Dim intLoop2 As Integer
   Dim intFound As Integer
   Dim strTableNameToBeMatched As String
   
   For intLoop1 = LBound(mastrTablesDB1) To UBound(mastrTablesDB1)
         strTableNameToBeMatched = mastrTablesDB1(intLoop1)
         For intLoop2 = LBound(mastrTablesDB2) To UBound(mastrTablesDB2)
            If strTableNameToBeMatched = mastrTablesDB2(intLoop2) Then
               ReDim Preserve mastrTables(intFound)
               mastrTables(intFound) = strTableNameToBeMatched
               intFound = intFound + 1
               Exit For
            End If
         Next intLoop2
   Next intLoop1
End Sub

Private Sub TablesDB1()

   Dim oConn As ADODB.Connection
   Dim oCat As ADOX.Catalog
   Dim oTable As ADOX.Table
   Dim intNumberOfTables As Integer
   
   Set oConn = New Connection
   oConn.ConnectionString = mstrConnectionStringDB1
   oConn.Open
   Set oCat = New ADOX.Catalog
   oCat.ActiveConnection = oConn
   Set oTable = New Table
   
   For Each oTable In oCat.Tables
      If oTable.Type <> "SYSTEM TABLE" And oTable.Type <> "SYSTEM VIEW" And Left$(oTable.Name, 3) = "tbl" Then
         ReDim Preserve mastrTablesDB1(intNumberOfTables)
         mastrTablesDB1(intNumberOfTables) = oTable.Name
         intNumberOfTables = intNumberOfTables + 1
      End If
   Next
   Set oCat = Nothing
   Set oTable = Nothing
   oConn.Close
   Set oConn = Nothing

End Sub

Private Sub TablesDB2()

   Dim oConn As ADODB.Connection
   Dim oCat As ADOX.Catalog
   Dim oTable As ADOX.Table
   Dim intNumberOfTables As Integer
   
   Set oConn = New Connection
   oConn.ConnectionString = mstrConnectionStringDB2
   oConn.Open
   Set oCat = New ADOX.Catalog
   oCat.ActiveConnection = oConn
   Set oTable = New Table
   
   For Each oTable In oCat.Tables
      If oTable.Type <> "SYSTEM TABLE" And oTable.Type <> "SYSTEM VIEW" And Left$(oTable.Name, 3) = "tbl" Then
         ReDim Preserve mastrTablesDB2(intNumberOfTables)
         mastrTablesDB2(intNumberOfTables) = oTable.Name
         intNumberOfTables = intNumberOfTables + 1
      End If
   Next
   
   Set oCat = Nothing
   Set oTable = Nothing
   oConn.Close
   Set oConn = Nothing
End Sub

' this routine has been taken from CODE_UPLOAD851 and modified as convenient. (Sorry, I can't remember the Author's name)
Private Sub ImportTable()
   Dim cnn1 As ADODB.Connection
   Dim cmdQuery As ADODB.Command
   Dim strCnn As String
   Dim Rs1 As ADODB.Recordset
   Dim prm As ADODB.Parameter
   Dim sTable As String
   Dim First As String
   Dim UserPath As String
   Dim NewTable As String
   Dim SelectString As String
   Dim FromString As String
   Dim FromSource As String
   Dim SQLtext As String
   Dim Msg As String
   Dim Destination As String
   'error handler
   On Error GoTo ErrorHandler
   
   sTable = mstrTableToBeUpdated  ' the table we are going to copy
   ' ADO connection string that MS uses to to setup a link between
   ' the program and a JET database
   ' the string would be different for non JET DB's such as Oracle.
   
   ' create and open a connection to the databse
   Set cnn1 = New ADODB.Connection
   cnn1.Open mstrConnectionStringDB1
   
   ' declare working variables we'll need for executing SQL statements
   Set cmdQuery = New ADODB.Command
   Set Rs1 = New ADODB.Recordset
   Rs1.ActiveConnection = mstrConnectionStringDB1
   'active connection
   Set cmdQuery.ActiveConnection = cnn1
   
   ' build the query up
   SelectString = "SELECT * INTO"
   FromString = "FROM"
   If Text3.Text = vbNullString Then
      FromSource = App.Path & "\" & Text1.Text
   Else
      FromSource = App.Path & "\" & "NoPWD" & Text1.Text
   End If
   NewTable = "tmp" & mstrTableToBeUpdated
   mstrOldTable = NewTable
   If Text4.Text = vbNullString Then
      Destination = App.Path & "\" & Text2.Text
   Else
      Destination = App.Path & "\" & "NoPWD" & Text2.Text
   End If
   'builds up query that will copy tables and data
   SQLtext = SelectString & Space(1) & "[" & Destination & "]" _
               & "." & NewTable & Space(1) & FromString & Space(1) _
               & "[" & FromSource & "]" & "." & sTable

   
   ' assigns the SQL statement to the command object
   cmdQuery.CommandText = SQLtext
   
   ' runs the SQL statement
   Set Rs1 = cmdQuery.Execute()
   
   If Err.Number = 0 Then
   End If
   
   
   'close the connection to the database
   cnn1.Close
   Exit Sub
   
   
   
ErrorHandler:      ' Error-handling routine.
      Select Case Err.Number   ' Evaluate error number.
   
          Case Else
   
         Msg = "Unexpected error #" & Str(Err.Number)
         Msg = Msg & " occurred: " & Err.Description
         ' Display message box with Stop sign icon and
         ' OK button.
         MsgBox Msg, vbCritical
   
      End Select
      Resume Next  ' Resume execution at same line
               ' that caused the error.
End Sub

Private Sub CreateQuery()
    Dim oConn As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim Rs1 As ADODB.Recordset
    Dim oCat As ADOX.Catalog
    Dim strSQL As String
    
    On Error GoTo ErrorMsg
   'copy the Table selected from db1 to db2 (tblPeople from db1 in tblPeople1 in db2)
    strSQL = "INSERT INTO " & mstrTableToBeUpdated & " SELECT " & mstrOldTable & ".* FROM " & mstrTableToBeUpdated
    Set oConn = New Connection
    Set cmd = New ADODB.Command
    Set Rs1 = New ADODB.Recordset
    Set oCat = New ADOX.Catalog
    oConn.ConnectionString = mstrConnectionStringDB2
    oConn.Open
    oCat.ActiveConnection = oConn
   
   ' Clear all the records of the table in db2)
    strSQL = "DELETE " & mstrTableToBeUpdated & ".* FROM " & mstrTableToBeUpdated
    cmd.CommandText = strSQL
    oConn.Execute strSQL
    
    'Run a command that update tblPeople from tblPeople1
    strSQL = "INSERT INTO " & mstrTableToBeUpdated & " SELECT " & mstrOldTable & ".* FROM " & mstrOldTable
    cmd.CommandText = strSQL
    oConn.Execute strSQL
    
    'delete the imported table (tblPeople1)
    strSQL = "DROP TABLE " & mstrOldTable
    cmd.CommandText = strSQL
    oConn.Execute strSQL
    Set oCat = Nothing
    Set cmd = Nothing
    Set oConn = Nothing
ErrorMsg:
   If Err.Number = -2147217833 Then
      MsgBox "Check the field size of each field contained in " & mstrTableToBeUpdated & "." & vbCrLf & _
      "You're trying to insert more data than expected"
      mstrErrorMsg = "Errors occurred in " & mstrTableToBeUpdated & ". You should check this table and repeat the update"
      Resume Next
   End If
End Sub
