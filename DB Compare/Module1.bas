Attribute VB_Name = "Module1"
Public cData As New clsData
Public cn As New ADODB.Connection
Public cn2 As New ADODB.Connection
Public adoGridRS As New ADODB.Recordset

Public MyWorkspace As Workspace
Public MyDb As Database
Public DataBaseName As String
Public DatabasePwd As String

Public TableList(100) As String
Public ContTable As Integer
Public FieldList(100, 500) As String
Public FieldType(100, 500) As Double
Public FieldDim(100, 500) As Double
Public ContField(500) As Integer

Public TableList2(100) As String
Public ContTable2 As Integer
Public FieldList2(100, 500) As String
Public FieldType2(100, 500) As Double
Public FieldDim2(100, 500) As Double
Public ContField2(500) As Integer

Public DatabaseSelezionato As Integer
Public TabellaSelezionata As Integer
Public Cont As Integer
Public Cont2 As Integer
Public Directory As String
Public Posizione As Integer



Public Sub ApriConnessione(Index As Integer, Optional cnConnection As Integer = 1)
On Error Resume Next
    If Index = 1 Then
        DataBaseName = Compare.Text1(0).Text
        DatabasePwd = Compare.Text1(1).Text
    End If
    If Index = 2 Then
        DataBaseName = Compare.Text2(0).Text
        DatabasePwd = Compare.Text2(1).Text
    End If
    
    cData.OpenDB DataBaseName, , DatabasePwd, cnConnection
    DoEvents
    Set MyWorkspace = Workspaces(0)
    Set MyDb = MyWorkspace.OpenDatabase(DataBaseName, False, False, ";PWD=" & DatabasePwd)
End Sub


Public Sub ChiudiConnessione()
    On Error Resume Next
    If cn.State = adStateOpen Then cn.Close
    If cn2.State = adStateOpen Then cn2.Close
    MyWorkspace.Close
End Sub


Public Sub Tabelle(Index As Integer)
    Dim rsSchema As New ADODB.Recordset
    
    Set rsSchema = cn.OpenSchema(adSchemaTables)
    If Index = 1 Then ContTable = 0 Else ContTable2 = 0
    Do Until rsSchema.EOF
        If rsSchema("TABLE_TYPE") = "TABLE" Then
            If Index = 1 Then
                TableList(ContTable) = rsSchema("TABLE_NAME")
                ContTable = ContTable + 1
            Else
                TableList2(ContTable2) = rsSchema("TABLE_NAME")
                ContTable2 = ContTable2 + 1
            End If
        End If
        rsSchema.MoveNext
    Loop
    If Index = 1 Then ContTable = ContTable - 1 Else ContTable2 = ContTable2 - 1
    rsSchema.Close
End Sub


Public Sub GetFieldList(MyTable As Integer)
    On Error GoTo errHandler
    
    If MyTable = 1 Then
        For Cont = 0 To ContTable
            For Cont2 = 0 To MyDb.TableDefs(TableList(Cont)).Fields.Count - 1
                FieldList(Cont, Cont2) = MyDb.TableDefs(TableList(Cont)).Fields(Cont2).Name
                FieldType(Cont, Cont2) = MyDb.TableDefs(TableList(Cont)).Fields(Cont2).Type
                Select Case FieldType(Cont, Cont2)
                  Case 2  'Byte: 'Intero 0-255
                      FieldDim(Cont, Cont2) = 3
                  Case 3  'Intero: -32768 +32768
                      FieldDim(Cont, Cont2) = 5
                  Case 4  'Intero Lungo: -2147483648 +2147483647
                      FieldDim(Cont, Cont2) = 10
                  Case 7  'Precisione Doppia
                      FieldDim(Cont, Cont2) = 15
                  Case 8  'Data
                      FieldDim(Cont, Cont2) = MyDb.TableDefs(TableList(Cont)).Fields(Cont2).Size
                  Case 10 'Stringa
                      FieldDim(Cont, Cont2) = MyDb.TableDefs(TableList(Cont)).Fields(Cont2).Size
                  Case Else
                      FieldDim(Cont, Cont2) = MyDb.TableDefs(TableList(Cont)).Fields(Cont2).Size
                End Select
            Next Cont2
            ContField(Cont) = Cont2
        Next Cont
        Exit Sub
    End If
    
    If MyTable = 2 Then
        For Cont = 0 To ContTable2
            For Cont2 = 0 To MyDb.TableDefs(TableList2(Cont)).Fields.Count - 1
                FieldList2(Cont, Cont2) = MyDb.TableDefs(TableList2(Cont)).Fields(Cont2).Name
                FieldType2(Cont, Cont2) = MyDb.TableDefs(TableList2(Cont)).Fields(Cont2).Type
                Select Case FieldType2(Cont, Cont2)
                  Case 2  'Byte: 'Intero 0-255
                      FieldDim2(Cont, Cont2) = 3
                  Case 3  'Intero: -32768 +32768
                      FieldDim2(Cont, Cont2) = 5
                  Case 4  'Intero Lungo: -2147483648 +2147483647
                      FieldDim2(Cont, Cont2) = 10
                  Case 7  'Precisione Doppia
                      FieldDim2(Cont, Cont2) = 15
                  Case 8  'Data
                      FieldDim2(Cont, Cont2) = MyDb.TableDefs(TableList2(Cont)).Fields(Cont2).Size
                  Case 10 'Stringa
                      FieldDim2(Cont, Cont2) = MyDb.TableDefs(TableList2(Cont)).Fields(Cont2).Size
                  Case Else
                      FieldDim2(Cont, Cont2) = MyDb.TableDefs(TableList2(Cont)).Fields(Cont2).Size
                End Select
            Next Cont2
            ContField2(Cont) = Cont2
        Next Cont
        Exit Sub
    End If

errHandler:
    'MsgBox Err.Description
    Resume Next
End Sub


Public Sub Espandi(Index As Integer, Optional Vero As Boolean = True)
On Error Resume Next
    If Vero = True Then
        For Cont = Compare.trwControlli(Index).Nodes.Count To 1 Step -1
            Compare.trwControlli(Index).Nodes(Cont).Expanded = True
        Next Cont
    Else
        For Cont = Compare.trwControlli(Index).Nodes.Count To 1 Step -1
            Compare.trwControlli(Index).Nodes(Cont).Expanded = False
        Next Cont
    End If
End Sub


Sub AccodaEliminaCampo(tdfTemp As TableDef, strComando As String, strNome As String, Optional varTipo, Optional varDimensione, Optional Posizione As Integer)
    On Error Resume Next
    With tdfTemp
        If .Updatable = False Then
            MsgBox "TableDef non aggiornabile! " & "Impossibile completare il task."
            Exit Sub
        End If
        If strComando = "APPEND" Then
            .Fields.Append .CreateField(strNome, varTipo, varDimensione)
            If Posizione > 0 Then .Fields(strNome).OrdinalPosition = Posizione
            Call ChiudiConnessione
        End If
        If strComando = "DELETE" Then
            .Fields.Delete strNome
            Call ChiudiConnessione
        End If
    End With
End Sub


Sub AccodaEliminaTabella(strComando As String, strNome As String)
On Error Resume Next
    If strComando = "APPEND" Then
        Call ApriConnessione(2)
        MyDb.Execute "SELECT ' ' AS CampoTest INTO " & strNome
        MyDb.TableDefs.Refresh
        MyDb.TableDefs(strNome).Fields.Delete "CampoTest"
        Call ChiudiConnessione
        For Cont = 0 To ContField(TabellaSelezionata)
            Call ApriConnessione(2)
            Call AccodaEliminaCampo(MyDb.TableDefs(strNome), "APPEND", FieldList(TabellaSelezionata, Cont), Val(FieldType(TabellaSelezionata, Cont)), Val(FieldDim(TabellaSelezionata, Cont)))
        Next Cont
    End If
    If strComando = "DELETE" Then
        Call ApriConnessione(2)
        MyDb.TableDefs.Delete (strNome)
    End If
End Sub


Public Function GetAccessPassWord(DataBaseName As String) As String
On Error GoTo errHandler
    Dim ch(18) As Byte, x As Integer
    Dim sec
    If Trim(DataBaseName) = "" Then Exit Function
    sec = Array(0, 134, 251, 236, 55, 93, 68, 156, 250, 198, 94, 40, 230, 19, 182, 138, 96, 84)
    Open DataBaseName For Binary Access Read As #1 Len = 18
    Get #1, &H42, ch
    Close #1
    For x = 1 To 17
        If (ch(x) Xor sec(x)) <> 0 Then GetAccessPassWord = GetAccessPassWord & Chr(ch(x) Xor sec(x))
    Next x
    Exit Function
    
errHandler:
    MsgBox "ERROR occcured:" & vbCrLf & Err.Number & ":  " & Err.Description, vbCritical, "ERROR"
    Exit Function
End Function



Public Sub SaveXML(DestinationXML As String, Optional QueryXML As String, Optional rsSetXML As ADODB.Recordset, Optional ConnectionXML As String = "")
    On Error Resume Next
    Dim AdorsXML As New ADODB.Recordset
    If QueryXML <> "" Then
        cData.NewAdoRs AdorsXML, QueryXML, , , , ConnectionXML
    Else
        If rsSetXML.RecordCount > 0 Then Set AdorsXML = rsSetXML
    End If
    DoEvents
    AdorsXML.Save DestinationXML, adPersistXML
    DoEvents
    Set AdorsXML = Nothing
End Sub


Public Sub OpenXML(PercorsoXML As String, rsSetXML As ADODB.Recordset)
    On Error Resume Next
    Dim AdorsXML As New ADODB.Recordset
    AdorsXML.Open PercorsoXML, "Provider=mspersist"
    DoEvents
    Set rsSetXML = AdorsXML
End Sub

