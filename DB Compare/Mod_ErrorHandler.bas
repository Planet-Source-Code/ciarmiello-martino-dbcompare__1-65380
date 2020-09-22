Attribute VB_Name = "Mod_ErrHandler"
Option Explicit
Private Const MODULE_NAME = "Error"
Public Const VALIDATION = "Validation"
Dim modname As String
Dim procname As String
Dim errnumb As String
Dim errdesc As String
Dim utente As String

Public Sub Process_Error(ByVal vstrModuleName As String, _
                         ByVal vstrProcName As String)
Dim strMsg As String

    modname = vstrModuleName
    procname = vstrProcName
    errnumb = CStr(Err.Number)
    errdesc = Err.Description
    
    'Nel caso genero volontariamente un errore con error.raise setto il source a validation
    If Err.Source = VALIDATION Then
        Beep
        MsgBox Err.Description, vbExclamation
    Else
        strMsg = "Is encourred an error in the module " & vbCrLf & _
                 " '" & vstrModuleName & "' in the sub '" & vstrProcName & "'." & vbCrLf & vbCrLf & _
                 "Error Number: " & CStr(Err.Number) & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
                 "Contact marte1977@hotmail.com"
        Beep
        MsgBox strMsg, vbCritical, App.Title
        Call WriteErrLog
        Screen.MousePointer = vbDefault
    End If
End Sub

Public Sub WriteErrLog()
    Dim strRow As String
    Dim strTmp1 As String
    Dim strPrimaRiga As String
    
    On Error GoTo fileopenerror
    
    'apre il file ascii; nel caso non esista lo crea tramite l'errore gestito (fileopenerror)
    Open App.Path + "\err.log" For Input As #1
    Line Input #1, strPrimaRiga
    'se non c'e' l'intestazione nel log la crea
    If Left(strPrimaRiga, 4) <> "Data" Then
        Close #1
        Open App.Path + "\err.log" For Append As #1
        strRow = ""
        strRow = strRow & FnRpad("Data", 11, " ")
        strRow = strRow & FnRpad("Utente", 8, " ")
        strRow = strRow & FnRpad("Modulo", 35, " ")
        strRow = strRow & FnRpad("Procedura", 35, " ")
        strRow = strRow & FnRpad("Error Num.", 15, " ")
        strRow = strRow & FnRpad("Descrizione Errore", 100, " ")
        Print #1, strRow
    End If
    Close #1
        
'        Open "A:\fileb.105" For Output As #1
    Open App.Path + "\err.log" For Append As #1
    strRow = ""
    strRow = FnRpad(Format(Now, "DD/MM/YYYY"), 11, " ")
    strRow = strRow & FnRpad(utente, 8, " ")
    strRow = strRow & FnRpad(modname, 35, " ")
    strRow = strRow & FnRpad(procname, 35, " ")
    strRow = strRow & FnRpad(errnumb, 15, " ")
    strRow = strRow & FnRpad(errdesc, 100, " ")
    Print #1, strRow
    Close #1

Exit Sub

fileopenerror:
    Select Case Err
        'Crea il file err.log
        Case 53
            Open App.Path + "\err.log" For Append As #1
            Close #1
            Open App.Path + "\err.log" For Input As #1
    End Select
    Resume Next
End Sub

Function FnLpad(Stringa As Variant, Lung As Integer, Carattere As String)
    'FnLpad=formattazione della stringa con n caratteri a sinistra
    'Lung=lunghezza della stringa da formattare
    'Carattere=carattere di riepimento della stringa
    'es FnLpad "1",6,"0": stringa="000001"
    
    On Error GoTo NullStringa
    If Len(Stringa) > Lung Then
        Stringa = Left(Stringa, Lung)
    End If
    FnLpad = String(Lung - Len(Stringa), Carattere) & Stringa
NullStringa:
    If Err = 13 Then
        Stringa = Str(Stringa)
        Resume
    ElseIf Err Then
        Stringa = ""
        Resume
    End If

End Function

Function FnRpad(Stringa As Variant, Lung As Integer, Carattere As String)
    'FnLpad = formattazione della stringa con n caratteri a sinistra
    'Lung = lunghezza della stringa da formattare
    'Carattere = carattere di riepimento della stringa, può essere anche ""
    'es FnRpad "Casa",10," ": stringa="Casa      "
    
    On Error GoTo NullStringa
    If Len(Stringa) > Lung Then
        Stringa = Left(Stringa, Lung)
    End If
    If Not Carattere = "" Then
        FnRpad = Stringa + String(Lung - Len(Stringa), Carattere)
    Else
        FnRpad = Stringa
    End If
NullStringa:
    If Err = 13 Then
        Stringa = Str(Stringa)
        Resume
    ElseIf Err Then
        Stringa = ""
        Resume
    End If
    
End Function
