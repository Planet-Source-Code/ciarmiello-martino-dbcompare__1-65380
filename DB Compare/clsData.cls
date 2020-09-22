VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Variabile o variabili locali per la memorizzazione dei valori delle proprietà.
Private iID As Variant
Private PathApplicazione As String
Private AdoCmd As New ADODB.Command
Private rs As New ADODB.Recordset
Private Const MODULE_NAME = "clsData"


Public Function NewRecordset(rsSet As Recordset, Optional CurLocation As String, Optional CurType As String, Optional LockType As String)
    On Error GoTo NewRecordset_error
    Set rs = rsSet
    With rs
        If CurLocation = "" Then
            .CursorLocation = adUseClient
        Else
            .CursorLocation = CurLocation
        End If
        
        If Trim(CurType) = "" Then
            .CursorType = adOpenDynamic
        Else
            .CursorType = CurType
        End If
        
        If LockType = "" Then
            .LockType = adLockOptimistic
        Else
            .LockType = LockType
        End If
    End With

NewRecordset_Done:
  Exit Function

NewRecordset_error:
  Call Process_Error(MODULE_NAME, "NewRecordset")
  Resume NewRecordset_Done
End Function


Public Sub CleanUp(rsSet As Recordset)
    On Error GoTo CleanUp_error
    
    rsSet.Close
    Set rsSet = Nothing
    
CleanUp_Done:
  Exit Sub

CleanUp_error:
  Call Process_Error(MODULE_NAME, "CleanUp")
  Resume CleanUp_Done
End Sub


Public Function RecordCount(rs As Recordset) As Long
    RecordCount = rs.RecordCount
End Function


Function mSQL(sSQL As String, Optional cnConnection As Integer = 1, Optional cmdType As String)
    On Error GoTo mSQL_Error
    
    If cmdType = "" Then AdoCmd.CommandType = adCmdText Else AdoCmd.CommandType = cmdType
    AdoCmd.CommandText = sSQL
    If cnConnection = 1 Then rs.Open AdoCmd.CommandText, cn
    If cnConnection = 2 Then rs.Open AdoCmd.CommandText, cn2
    
mSQL_Done:
  Exit Function

mSQL_Error:
  Call Process_Error(MODULE_NAME, "mSQL: " & sSQL)
  Resume mSQL_Done
End Function


Public Sub OpenDB(DBName As String, Optional DBUserID As String, Optional DBPwd As String, Optional cnConnection As Integer = 1)
    Dim strConnect As String
    Dim strPWD As String
    Dim strUID As String
    On Error GoTo OpenDB_Error
    
    Screen.MousePointer = vbHourglass
    If cnConnection = 1 Then If cn.State = adStateOpen Then cn.Close
    If cnConnection = 2 Then If cn2.State = adStateOpen Then cn2.Close
    
    If DBUserID <> "" Then
        strUID = "USER ID=" & DBUserID & ";"
    Else
        strUID = "USER ID=;"
    End If
    
    If DBPwd <> "" Then
        strPWD = "Password=" & DBPwd & ";"
    Else
        strPWD = "Password=;"
    End If
    
    PathApplicazione = DataBaseName
    strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;" _
             & "Data Source=" & PathApplicazione + ";" _
             & "Jet OLEDB:Database " & strPWD
    
    If cnConnection = 1 Then cn.Open strConnect
    If cnConnection = 2 Then cn2.Open strConnect
    Screen.MousePointer = vbDefault

OpenDB_Done:
  Exit Sub

OpenDB_Error:
    Select Case Err
        Case -2147467259
            Call Process_Error(MODULE_NAME, "OpenDB")
            End
        Case Else
            Call Process_Error(MODULE_NAME, "OpenDB")
            Resume OpenDB_Done
    End Select
End Sub



Public Function NewAdoRs(rsSet As ADODB.Recordset, Optional sSQL As String, Optional CurLocation As String, Optional CurType As String, Optional LockType As String, Optional ByVal cnConnection As String)
    On Error Resume Next
    Set rs = rsSet
    DoEvents
    With rs
        If CurLocation = "" Then .CursorLocation = adUseClient Else .CursorLocation = CurLocation
        If Trim(CurType) = "" Then .CursorType = adOpenDynamic Else .CursorType = CurType
        If LockType = "" Then .LockType = adLockOptimistic Else .LockType = LockType
    End With
    DoEvents
    If sSQL <> "" Then
        If cmdType = "" Then AdoCmd.CommandType = adCmdText Else AdoCmd.CommandType = cmdType
        AdoCmd.CommandText = sSQL
        If cnConnection = "" Then
            rs.Open AdoCmd.CommandText, cn
        ElseIf cnConnection = "cn2" Then
            rs.Open AdoCmd.CommandText, cn2
        End If
    End If
    DoEvents
End Function
