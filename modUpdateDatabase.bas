Attribute VB_Name = "modUpdateDatabase"
Option Explicit

Public Const Provider As String = "Provider=Microsoft.Jet.OLEDB.4.0;"
Public Const JetPassword As String = "Jet OLEDB:Database Password="

Public Function FileExists(ByVal strFile As String) As Boolean
   FileExists = False
   On Error GoTo FileExists_EH
   
   If Trim$(strFile) <> "" Then
      ' Check for Path Only
      If Right$(Trim$(strFile), 1) <> "\" Then
         ' Now look for file
         If Dir(strFile) <> "" Then
            FileExists = True
         End If
      End If
   End If
   
   Exit Function
   
FileExists_EH:
   ' The path is invalid.
   Exit Function
End Function

   'This is to remove the password from the databases
Public Sub RemovePassword(ByVal strDb As String, ByVal PWD As String)
   Dim JRO As JRO.JetEngine
   Dim strProvider As String
   Dim strDataDestination As String
   Dim strJetPassword As String
   Dim strSource As String
   Dim strDestination As String
   Dim strToBeKilled As String
   Dim strRenamed As String
   Set JRO = New JRO.JetEngine

   strProvider = "Microsoft.Jet.OLEDB.4.0"
   strJetPassword = "Jet OLEDB:Database Password=" & PWD
   strSource = "Provider = " & strProvider & ";" & "Data Source = " & strDb & ";" & strJetPassword
   strDataDestination = App.Path & "\NoPWD" & GetFileBaseName(strDb) & ".mdb"
   strDestination = "Provider = " & strProvider & ";" & "Data Source = " & strDataDestination
   On Error GoTo ErroreCompattazione
   strToBeKilled = strDataDestination
   If FileExists(strToBeKilled) Then Kill strToBeKilled
   JRO.CompactDatabase strSource, strDestination
      If Err = 0 Then
       Else
         MsgBox "A problem has occurred while compacting. (" & Err.Description & ")", vbExclamation
       End If
       Exit Sub
ErroreCompattazione:
   If Err.Number = 3356 Then
      MsgBox "Other users might be using the database." & vbCrLf & "Couldn't open it exclusively!"
   Else
      Resume Next
   End If
End Sub

Function GetFileBaseName(FileName As String, Optional ByVal IncludePath As Boolean) As String
    Dim i As Long, startPos As Long, endPos As Long
    
    startPos = 1
    
    For i = Len(FileName) To 1 Step -1
        Select Case Mid$(FileName, i, 1)
            Case "."
                ' we've found the extension
                If IncludePath Then
                    ' if we must return the path, we've done
                    GetFileBaseName = Left$(FileName, i - 1)
                    Exit Function
                End If
                ' else, just take note of where the extension begins
                If endPos = 0 Then endPos = i - 1
            Case ":", "\"
                If Not IncludePath Then startPos = i + 1
                Exit For
        End Select
    Next
    
    If endPos = 0 Then
        ' this file has no extension
        GetFileBaseName = Mid$(FileName, startPos)
    Else
        GetFileBaseName = Mid$(FileName, startPos, endPos - startPos + 1)
    End If
End Function


'This is to put the password back into the databases
Public Sub PutThePasswordBack(ByVal strDb As String, ByVal PWD As String)
   Dim JRO As JRO.JetEngine
   Dim strProvider As String
   Dim strDataDestination As String
   Dim strSource As String
   Dim strDestination As String
   Dim strToBeKilled As String
   Dim strRenamed As String
   Set JRO = New JRO.JetEngine

   strProvider = "Microsoft.Jet.OLEDB.4.0"
   strSource = "Provider = " & strProvider & ";" & "Data Source = " & strDb
   strDataDestination = GetFileBaseName(strDb)
   strDataDestination = Mid$(strDataDestination, (6), Len(strDataDestination))
   strDataDestination = App.Path & "\" & strDataDestination & ".mdb"
   strDestination = "Provider = " & strProvider & ";" & "Data Source = " & strDataDestination & ";" & "Jet OLEDB:Database Password=" & PWD & ";"
   On Error GoTo ErroreCompattazione
   JRO.CompactDatabase strSource, strDestination
      If Err = 0 Then
       Else
         MsgBox "A problem has occurred while compacting. (" & Err.Description & ")", vbExclamation
       End If
       Exit Sub
ErroreCompattazione:
   If Err.Number = 3356 Then
      MsgBox "Other users might be using the database." & vbCrLf & "Couldn't open it exclusively!"
   Else
      Resume Next
   End If
End Sub

