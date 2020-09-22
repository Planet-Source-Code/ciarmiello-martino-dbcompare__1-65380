VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Compare 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DB Compare"
   ClientHeight    =   6945
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   9375
   Icon            =   "Compare.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6945
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "&Open Table"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7140
      TabIndex        =   23
      Top             =   6540
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Analize Database &2"
      Height          =   435
      Index           =   1
      Left            =   8220
      TabIndex        =   12
      Top             =   540
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   4260
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   6600
      Width           =   435
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3720
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   6600
      Width           =   435
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   3180
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   6600
      Width           =   435
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   6600
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   6600
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "&End"
      Height          =   375
      Left            =   8280
      TabIndex        =   24
      Top             =   6540
      Width           =   1035
   End
   Begin VB.CommandButton Command4 
      Caption         =   "<-- &Remove"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6000
      TabIndex        =   22
      Top             =   6540
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Insert -->"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4860
      TabIndex        =   21
      Top             =   6540
      Width           =   1035
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Compare"
      Height          =   315
      Left            =   8220
      TabIndex        =   13
      Top             =   1020
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Analize Database &1"
      Height          =   435
      Index           =   0
      Left            =   8220
      TabIndex        =   11
      Top             =   60
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Height          =   4875
      Left            =   60
      TabIndex        =   29
      Top             =   1440
      Width           =   9255
      Begin MSComctlLib.TreeView trwControlli 
         Height          =   4695
         Index           =   1
         Left            =   4680
         TabIndex        =   15
         Top             =   120
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   8281
         _Version        =   393217
         HideSelection   =   0   'False
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.TreeView trwControlli 
         Height          =   4695
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   8281
         _Version        =   393217
         HideSelection   =   0   'False
         Style           =   7
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Database  2  -  Destination"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   60
      TabIndex        =   27
      Top             =   720
      Width           =   7935
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "Compare.frx":0442
         Left            =   120
         List            =   "Compare.frx":044C
         TabIndex        =   38
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   9
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox Text2 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   6840
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox TxtFile 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.Label Label2 
         Caption         =   "Password :"
         Height          =   195
         Index           =   2
         Left            =   6000
         TabIndex        =   28
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Database  1  -  Origin"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   615
      Left            =   60
      TabIndex        =   25
      Top             =   60
      Width           =   7935
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "Compare.frx":0464
         Left            =   120
         List            =   "Compare.frx":046E
         TabIndex        =   36
         Top             =   240
         Width           =   1515
      End
      Begin VB.TextBox Text1 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   6840
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   7
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox TxtFile 
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   37
         Top             =   240
         Visible         =   0   'False
         Width           =   6135
      End
      Begin VB.Label Label2 
         Caption         =   "Password :"
         Height          =   195
         Index           =   1
         Left            =   5940
         TabIndex        =   26
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   6855
      Left            =   60
      TabIndex        =   35
      Top             =   60
      Visible         =   0   'False
      Width           =   9255
      Begin VB.CommandButton CmdXml1 
         Caption         =   "&Open  XML"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   6360
         Width           =   1155
      End
      Begin VB.CommandButton CmdXml2 
         Caption         =   "&Save to XML"
         Height          =   375
         Left            =   1380
         TabIndex        =   4
         Top             =   6360
         Width           =   1155
      End
      Begin VB.CommandButton CmdXml3 
         Caption         =   "Puts Data XML in line to Table"
         Height          =   375
         Left            =   2640
         TabIndex        =   5
         Top             =   6360
         Visible         =   0   'False
         Width           =   4095
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Execute SQL"
         Height          =   315
         Left            =   7980
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TxtSql 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   7755
      End
      Begin VB.CommandButton Command7 
         Caption         =   "&End"
         Height          =   375
         Left            =   8100
         TabIndex        =   6
         Top             =   6360
         Width           =   1035
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5595
         Left            =   120
         TabIndex        =   2
         Top             =   660
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   9869
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1040
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            ScrollBars      =   3
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   7440
         Top             =   6300
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Selezione del file XML"
         Filter          =   "*.xml"
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Pos."
      Height          =   195
      Left            =   4260
      TabIndex        =   34
      Top             =   6420
      Width           =   435
   End
   Begin VB.Label Label7 
      Caption         =   "Dim."
      Height          =   195
      Left            =   3720
      TabIndex        =   33
      Top             =   6420
      Width           =   435
   End
   Begin VB.Label Label6 
      Caption         =   "Tipo"
      Height          =   195
      Left            =   3180
      TabIndex        =   32
      Top             =   6420
      Width           =   435
   End
   Begin VB.Label Label5 
      Caption         =   "Campo"
      Height          =   195
      Left            =   1620
      TabIndex        =   31
      Top             =   6420
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Tabella"
      Height          =   195
      Left            =   60
      TabIndex        =   30
      Top             =   6420
      Width           =   975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuEsci 
         Caption         =   "End"
      End
   End
   Begin VB.Menu mnuStrumenti 
      Caption         =   "&Struments"
      Begin VB.Menu mnuTrovaPwd 
         Caption         =   "Find Password"
      End
      Begin VB.Menu mnuCopiaDatiDB 
         Caption         =   "Copy DB1 to DB2"
      End
      Begin VB.Menu mnuEsportaStruttura 
         Caption         =   "Esport Structure"
      End
   End
   Begin VB.Menu mnuDati 
      Caption         =   "&Data"
      Enabled         =   0   'False
      Begin VB.Menu mnuSvuotaTabella 
         Caption         =   "Clean Table"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuAccodaDati 
         Caption         =   "Puts Data in line"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "?"
      Begin VB.Menu mnuTipo 
         Caption         =   "Field type"
      End
      Begin VB.Menu mnuDimensione 
         Caption         =   "Field length"
      End
   End
End
Attribute VB_Name = "Compare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mNode As Node
Dim strMsg As String
Dim intMsg As Integer
Dim TipoCampo As Integer
Dim CurrentT2 As Integer
Dim TabellaXML As String



Private Sub Combo1_Click()
    If Combo1.Text = "Database" Then
        TxtFile(0).Text = ""
        Text1(0).Text = ""
        Text1(1).Text = ""
        TxtFile(0).Visible = False
        Text1(0).Visible = True
        Text1(1).Visible = True
        mnuTrovaPwd.Enabled = True
    Else
        TxtFile(0).Text = App.Path & "\DBCompareDB1.DBC"
        Text1(0).Text = ""
        Text1(1).Text = ""
        TxtFile(0).Visible = True
        Text1(0).Visible = False
        Text1(1).Visible = False
        mnuTrovaPwd.Enabled = False
    End If
End Sub

Private Sub Combo2_Click()
    If Combo2.Text = "Database" Then
        TxtFile(1).Text = ""
        Text2(0).Text = ""
        Text2(1).Text = ""
        TxtFile(1).Visible = False
        Text2(0).Visible = True
        Text2(1).Visible = True
        mnuTrovaPwd.Enabled = True
    Else
        TxtFile(1).Text = App.Path & "\DBCompareDB2.DBC"
        Text2(0).Text = ""
        Text2(1).Text = ""
        TxtFile(1).Visible = True
        Text2(0).Visible = False
        Text2(1).Visible = False
        mnuTrovaPwd.Enabled = False
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    If TxtFile(Index).Visible = False Then Call Esegui(Index) Else Call Esegui2(Index)
End Sub


Private Sub Command2_Click()
    Call Analizza
End Sub


Private Sub Command7_Click()
    Frame4.Visible = False
    Frame4.ZOrder 1
    Set adoGridRS = Nothing
    Call HideXml
    Call ChiudiConnessione
End Sub


Public Sub Esegui(Index As Integer)
    If Index = 0 Then
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
        DoEvents
        Call Resetta(1)
        If Dir(Text1(0).Text) <> "" Then
            Call ApriConnessione(1)
            Call Tabelle(1)
            Call GetFieldList(1)
            Call InitTreeView(1)
            Call ChiudiConnessione
        End If
        DoEvents
        Screen.MousePointer = vbDefault
        Me.Enabled = True
        MsgBox "Analysis finished", vbOKOnly, "Attention"
        DoEvents
    End If
    If Index = 1 Then
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
        DoEvents
        Call Resetta(2)
        If Dir(Text2(0).Text) <> "" Then
            Call ApriConnessione(2)
            Call Tabelle(2)
            Call GetFieldList(2)
            Call InitTreeView(2)
            Call ChiudiConnessione
        End If
        Screen.MousePointer = vbDefault
        Me.Enabled = True
        MsgBox "Analysis finished", vbOKOnly, "Attention"
        DoEvents
    End If
End Sub


Public Sub Esegui2(Index As Integer)
    Me.Enabled = False
    Screen.MousePointer = vbHourglass
    DoEvents
    Call Resetta(Index + 1)
    If Dir(TxtFile(Index).Text) <> "" Then
        Call LoadDBC(Index + 1)
        Call InitTreeView(Index + 1)
    End If
    DoEvents
    Screen.MousePointer = vbDefault
    Me.Enabled = True
    MsgBox "Analysis finished", vbOKOnly, "Attention"
    DoEvents
End Sub


Public Sub Analizza()
    Me.Enabled = False
    If ContTable <> ContTable2 Then
        MsgBox "The Tables number of the two database is not correspondent:" & vbCrLf & "Execute Updates and still execute analysis !", vbOKOnly, "Attention"
        Me.Enabled = True
        Exit Sub
    End If
    For Cont = 0 To ContTable
        If TableList(Cont) <> TableList2(Cont) Then
            MsgBox "To the Table " & TableList(Cont) & " correspond another table:" & vbCrLf & "Execute Updates and still execute analysis !", vbOKOnly, "Attention"
            Me.Enabled = True
            Exit Sub
        End If
    Next Cont
    For Cont = 0 To ContTable
        For Cont2 = 0 To 200
            If FieldList(Cont, Cont2) <> FieldList2(Cont, Cont2) Then
                MsgBox "In the Table " & TableList(Cont) & ", to the field " & FieldList(Cont, Cont2) & " corriespond another field:" & vbCrLf & "Execute Updates and still execute analysis !", vbOKOnly, "Attention"
                Me.Enabled = True
                Exit Sub
            End If
            If FieldType(Cont, Cont2) <> FieldType2(Cont, Cont2) Then
                MsgBox "In the Table " & TableList(Cont) & ", the field " & FieldList(Cont, Cont2) & " have a different type :" & vbCrLf & "Execute Updates and still execute analysis !", vbOKOnly, "Attention"
                Me.Enabled = True
                Exit Sub
            End If
            If FieldDim(Cont, Cont2) <> FieldDim2(Cont, Cont2) Then
                MsgBox "In the Table " & TableList(Cont) & ", the field " & FieldList(Cont, Cont2) & " have a different length:" & vbCrLf & "Execute Updates and still execute analysis !", vbOKOnly, "Attention"
                Me.Enabled = True
                Exit Sub
            End If
        Next Cont2
    Next Cont
    Me.Enabled = True
    MsgBox "Compared was completed: the two Databases have equal structure!", vbOKOnly, "Attention"
End Sub


Public Sub Resetta(Index As Integer)
    On Error Resume Next
    If Index = 1 Then
        ContTable = 0
        For Cont = 0 To 100
            TableList(Cont) = ""
            For Cont2 = 0 To 500
                ContField(Cont2) = 0
                FieldList(Cont, Cont2) = ""
                FieldType(Cont, Cont2) = 0
                FieldDim(Cont, Cont2) = 0
            Next Cont2
        Next Cont
    End If
    
    If Index = 2 Then
        ContTable2 = 0
        For Cont = 0 To 100
            TableList2(Cont) = ""
            For Cont2 = 0 To 500
                ContField2(Cont2) = 0
                FieldList2(Cont, Cont2) = ""
                FieldType2(Cont, Cont2) = 0
                FieldDim2(Cont, Cont2) = 0
            Next Cont2
        Next Cont
    End If
End Sub


Public Sub InitTreeView(Index As Integer)
    On Error Resume Next
    
    If Index = 1 Then
        With trwControlli(0)
            .Nodes.Clear
            .Sorted = False
            .LabelEdit = False
            .LineStyle = tvwRootLines
            .Checkboxes = False
        End With

        For Cont = 0 To ContTable
            Set mNode = trwControlli(0).Nodes.Add(, , "R" & Trim(Format(Cont, 0#)))
            mNode.Text = TableList(Cont)
            mNode.Tag = TableList(Cont)
            DoEvents
            For Cont2 = 0 To ContField(Cont)
                If FieldList(Cont, Cont2) <> "" Then
                    Set mNode = trwControlli(0).Nodes.Add("R" & Trim(Format(Cont, 0#)), tvwChild, Trim(Format(Cont, 0#)) & "F" & Trim(Format(Cont2, 0#)), FieldList(Cont, Cont2))
                    mNode.Text = FieldList(Cont, Cont2)
                    DoEvents
                    Set mNode = trwControlli(0).Nodes.Add(Trim(Format(Cont, 0#)) & "F" & Trim(Format(Cont2, 0#)), tvwChild, Trim(Format(Cont, 0#)) & "T" & Trim(Format(Cont2, 0#)), FieldType(Cont, Cont2))
                    mNode.Text = FieldType(Cont, Cont2)
                    DoEvents
                    Set mNode = trwControlli(0).Nodes.Add(Trim(Format(Cont, 0#)) & "F" & Trim(Format(Cont2, 0#)), tvwChild, Trim(Format(Cont, 0#)) & "L" & Trim(Format(Cont2, 0#)), FieldDim(Cont, Cont2))
                    mNode.Text = FieldDim(Cont, Cont2)
                    DoEvents
                End If
            Next Cont2
        Next Cont
    End If
    
    If Index = 2 Then
        With trwControlli(1)
            .Nodes.Clear
            .Sorted = False
            .LabelEdit = False
            .LineStyle = tvwRootLines
            .Checkboxes = False
        End With

        For Cont = 0 To ContTable2
            Set mNode = trwControlli(1).Nodes.Add(, , "R" & Trim(Format(Cont, 0#)))
            mNode.Text = TableList2(Cont)
            mNode.Tag = TableList2(Cont)
            DoEvents
            For Cont2 = 0 To ContField(Cont)
                If FieldList2(Cont, Cont2) <> "" Then
                    Set mNode = trwControlli(1).Nodes.Add("R" & Trim(Format(Cont, 0#)), tvwChild, Trim(Format(Cont, 0#)) & "F" & Trim(Format(Cont2, 0#)), FieldList2(Cont, Cont2))
                    mNode.Text = FieldList2(Cont, Cont2)
                    DoEvents
                    Set mNode = trwControlli(1).Nodes.Add(Trim(Format(Cont, 0#)) & "F" & Trim(Format(Cont2, 0#)), tvwChild, Trim(Format(Cont, 0#)) & "T" & Trim(Format(Cont2, 0#)), FieldType2(Cont, Cont2))
                    mNode.Text = FieldType2(Cont, Cont2)
                    DoEvents
                    Set mNode = trwControlli(1).Nodes.Add(Trim(Format(Cont, 0#)) & "F" & Trim(Format(Cont2, 0#)), tvwChild, Trim(Format(Cont, 0#)) & "L" & Trim(Format(Cont2, 0#)), FieldDim2(Cont, Cont2))
                    mNode.Text = FieldDim2(Cont, Cont2)
                    DoEvents
                End If
            Next Cont2
        Next Cont
    End If
End Sub

Private Sub Command3_Click()
    If Text4.Text <> "" Then
        Call ApriConnessione(2)
        AccodaEliminaCampo MyDb.TableDefs(Text3.Text), "APPEND", Text4.Text, Val(Text5.Text), Val(Text6.Text), Val(Text7.Text)
        Call ChiudiConnessione
    End If
    If Text4.Text = "" Then
        AccodaEliminaTabella "APPEND", Text3.Text
    End If
    Call Blocca
End Sub

Private Sub Command4_Click()
    If Text4.Text <> "" Then
        Call ApriConnessione(2)
        AccodaEliminaCampo MyDb.TableDefs(Text3.Text), "DELETE", Text4.Text
        Call ChiudiConnessione
    End If
    If Text4.Text = "" Then
        AccodaEliminaTabella "DELETE", Text3.Text
    End If
    Call Blocca
End Sub


Private Sub Command5_Click()
    If TabellaSelezionata >= 0 And DatabaseSelezionato > 0 Then
        If DatabaseSelezionato = 1 And TxtFile(0).Visible = True Then Exit Sub
        If DatabaseSelezionato = 2 And TxtFile(1).Visible = True Then Exit Sub
        Call ApriConnessione(DatabaseSelezionato)
        Frame4.Visible = True
        Frame4.ZOrder 0
        Frame4.Caption = " EDITOR  Database" & Trim(DatabaseSelezionato) & "  -  Tabella =  '" & TableList(TabellaSelezionata) & "'"
        Set adoGridRS = Nothing
        cData.NewRecordset adoGridRS
        If DatabaseSelezionato = 1 Then cData.mSQL ("SELECT * FROM " & TableList(TabellaSelezionata))
        If DatabaseSelezionato = 2 Then cData.mSQL ("SELECT * FROM " & TableList2(TabellaSelezionata))
        Set DataGrid1.DataSource = adoGridRS
    End If
End Sub


Private Sub Command8_Click()
    On Error GoTo SqlError
    Dim AdoRsTest As New ADODB.Recordset
    Dim Prova As String
    cData.NewRecordset AdoRsTest
    cData.mSQL TxtSql.Text
    DoEvents
    Prova = AdoRsTest.Fields.Count
    cData.CleanUp AdoRsTest
    cData.CleanUp adoGridRS
    DoEvents
    cData.NewRecordset adoGridRS
    cData.mSQL TxtSql.Text
    DoEvents
    Set DataGrid1.DataSource = adoGridRS
    DataGrid1.Refresh
    Exit Sub
    
SqlError:
    MsgBox "Query Sql not Valid !", vbOKOnly + vbCritical, "Attention"
    Exit Sub
End Sub


Private Sub Command6_Click()
    Call ChiudiConnessione
    End
End Sub


Private Sub Form_Load()
    Posizione = 1
    Combo1.ListIndex = 0
    Combo2.ListIndex = 0
End Sub

Private Sub mnuDimensione_Click()
    MsgBox "2 = Byte: 3 " & vbCrLf & "3 = Intero: 5 " & vbCrLf & "4 = Intero Lungo: 10 " & vbCrLf & "7 = Precisione Doppia: 15 " & vbCrLf & "8 = Data: Definita dal database " & vbCrLf & "10 = Stringa: Definita dall'utente", vbInformation + vbOKOnly, "Field Dim"
End Sub
    
Private Sub mnuTipo_Click()
    MsgBox "2 = Byte: Intero 0-255 " & vbCrLf & "3 = Intero: -32768 +32768 " & vbCrLf & "4 = Intero Lungo: -2147483648 +2147483647 " & vbCrLf & "7 = Precisione Doppia " & vbCrLf & "8 = Data " & vbCrLf & "10 = Stringa", vbInformation + vbOKOnly, "Field Type"
End Sub

Private Sub mnuEsci_Click()
    Call ChiudiConnessione
    End
End Sub




Private Sub mnuTrovaPwd_Click()
    Me.Enabled = False
    frm_MAPwd.Show
End Sub

Private Sub mnuEsportaStruttura_Click()
    On Error Resume Next
    Dim Percorso As String
    If Posizione = 1 Then If ContTable <= 0 Then Exit Sub
    If Posizione = 2 Then If ContTable2 <= 0 Then Exit Sub
    Percorso = App.Path & "\DBCompareDB" & Posizione & ".DBC"
    intMsg = MsgBox("Export the structure of the database " & Posizione & " in the file " & vbCrLf & Percorso & "?", vbYesNo + vbInformation + vbDefaultButton1, "Exporting DB structure")
    If intMsg = vbYes Then
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
        If Dir(Percorso) <> "" Then Kill (Percorso)
        Open Percorso For Append As #1
        Close #1
        Open Percorso For Append As #1
        If Posizione = 1 Then
            Print #1, "ContTable"
            Close #1
            Open Percorso For Append As #1
            Print #1, ContTable
            Close #1
            Open Percorso For Append As #1
            Print #1, "TableList"
            Close #1
            For Cont = 0 To ContTable
                Open Percorso For Append As #1
                Print #1, Cont & "." & TableList(Cont)
                Close #1
            Next Cont
            Open Percorso For Append As #1
            Print #1, "ContField"
            Close #1
            For Cont = 0 To ContTable
                Open Percorso For Append As #1
                Print #1, Cont & "." & "ContField"
                Close #1
                Open Percorso For Append As #1
                Print #1, ContField(Cont)
                Close #1
            Next Cont
            Open Percorso For Append As #1
            Print #1, "FieldList"
            Close #1
            For Cont = 0 To ContTable
                For Cont2 = 0 To ContField(Cont) - 1
                    Open Percorso For Append As #1
                    Print #1, Cont & "." & Cont2 & "." & "FieldList"
                    Close #1
                    Open Percorso For Append As #1
                    Print #1, FieldList(Cont, Cont2)
                    Close #1
                Next Cont2
            Next Cont
            Open Percorso For Append As #1
            Print #1, "FieldType"
            Close #1
            For Cont = 0 To ContTable
                For Cont2 = 0 To ContField(Cont) - 1
                    Open Percorso For Append As #1
                    Print #1, Cont & "." & Cont2 & "." & "FieldType"
                    Close #1
                    Open Percorso For Append As #1
                    Print #1, FieldType(Cont, Cont2)
                    Close #1
                Next Cont2
            Next Cont
            Open Percorso For Append As #1
            Print #1, "FieldDim"
            Close #1
            For Cont = 0 To ContTable
                For Cont2 = 0 To ContField(Cont) - 1
                    Open Percorso For Append As #1
                    Print #1, Cont & "." & Cont2 & "." & "FieldDim"
                    Close #1
                    Open Percorso For Append As #1
                    Print #1, Val(FieldDim(Cont, Cont2))
                    Close #1
                Next Cont2
            Next Cont
        End If
        If Posizione = 2 Then
            Print #1, "ContTable"
            Close #1
            Open Percorso For Append As #1
            Print #1, ContTable2
            Close #1
            Open Percorso For Append As #1
            Print #1, "TableList"
            Close #1
            For Cont = 0 To ContTable2
                Open Percorso For Append As #1
                Print #1, Cont & "." & TableList2(Cont)
                Close #1
            Next Cont
            Open Percorso For Append As #1
            Print #1, "ContField"
            Close #1
            For Cont = 0 To ContTable2
                Open Percorso For Append As #1
                Print #1, Cont & "." & "ContField"
                Close #1
                Open Percorso For Append As #1
                Print #1, ContField2(Cont)
                Close #1
            Next Cont
            Open Percorso For Append As #1
            Print #1, "FieldList"
            Close #1
            For Cont = 0 To ContTable2
                For Cont2 = 0 To ContField2(Cont) - 1
                    Open Percorso For Append As #1
                    Print #1, Cont & "." & Cont2 & "." & "FieldList"
                    Close #1
                    Open Percorso For Append As #1
                    Print #1, FieldList2(Cont, Cont2)
                    Close #1
                Next Cont2
            Next Cont
            Open Percorso For Append As #1
            Print #1, "FieldType"
            Close #1
            For Cont = 0 To ContTable2
                For Cont2 = 0 To ContField2(Cont) - 1
                    Open Percorso For Append As #1
                    Print #1, Cont & "." & Cont2 & "." & "FieldType"
                    Close #1
                    Open Percorso For Append As #1
                    Print #1, FieldType2(Cont, Cont2)
                    Close #1
                Next Cont2
            Next Cont
            Open Percorso For Append As #1
            Print #1, "FieldDim"
            Close #1
            For Cont = 0 To ContTable2
                For Cont2 = 0 To ContField2(Cont) - 1
                    Open Percorso For Append As #1
                    Print #1, Cont & "." & Cont2 & "." & "FieldDim"
                    Close #1
                    Open Percorso For Append As #1
                    Print #1, Val(FieldDim2(Cont, Cont2))
                    Close #1
                Next Cont2
            Next Cont
        End If
        Me.Enabled = True
        Screen.MousePointer = vbDefault
        MsgBox "Export the structure of the database " & Posizione & " in the file " & vbCrLf & Percorso & vbCrLf & "Completed", vbOKOnly + vbInformation, "Exporting DB structure"
    End If
End Sub

Private Sub mnuAccodaDati_Click()
    On Error Resume Next
    strMsg = ""
    strMsg = strMsg + vbCr + " The operation will modify the content of the table"
    strMsg = strMsg + vbCr + " leaving intact all the records already present."
    strMsg = strMsg + vbCr
    strMsg = strMsg + vbCr + "    Continue with the operation ?"
    intMsg = MsgBox(strMsg, vbCritical + vbYesNo + vbDefaultButton2, " Putting in line of Data")
    If intMsg = vbYes Then
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
        If TabellaSelezionata >= 0 And DatabaseSelezionato = 1 Then
            Dim AdoRsDati As New ADODB.Recordset
            Dim AdoRsDati2 As New ADODB.Recordset
            Dim ContCampi As Integer
            Call ApriConnessione(1)
            Call cData.NewRecordset(AdoRsDati)
            Call cData.mSQL("SELECT * FROM " & TableList2(TabellaSelezionata))
            DoEvents
            Call ApriConnessione(2, 2)
            Call cData.NewRecordset(AdoRsDati2)
            Call cData.mSQL("SELECT * FROM " & TableList2(TabellaSelezionata), 2)
            AdoRsDati.MoveFirst
            While AdoRsDati.EOF = False
                AdoRsDati2.AddNew
                For ContCampi = 0 To AdoRsDati2.Fields.Count - 1
                    AdoRsDati2.Fields(ContCampi) = AdoRsDati.Fields(ContCampi)
                Next ContCampi
                AdoRsDati2.Update
                AdoRsDati.MoveNext
                DoEvents
            Wend
        Set AdoRsDati = Nothing
        Set AdoRsDati2 = Nothing
        Call ChiudiConnessione
        End If
        Me.Enabled = True
        Screen.MousePointer = vbDefault
        MsgBox "Operation terminated correctly!", vbOKOnly, " Putting in line of Data"
    End If
    Call Blocca
End Sub


Private Sub mnuSvuotaTabella_Click()
    On Error Resume Next
    strMsg = ""
    strMsg = strMsg + vbCr + " The operation will delete in permanent way"
    strMsg = strMsg + vbCr + " all the records present in this table ."
    strMsg = strMsg + vbCr
    strMsg = strMsg + vbCr + "    Continue with the operation ?"
    intMsg = MsgBox(strMsg, vbCritical + vbYesNo + vbDefaultButton2, " Deleting Data")
    If intMsg = vbYes Then
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
        If TabellaSelezionata >= 0 And DatabaseSelezionato = 2 Then
            Dim AdoRsDati As New ADODB.Recordset
            Call ApriConnessione(2)
            cData.NewRecordset AdoRsDati
            cData.mSQL ("SELECT * FROM " & TableList2(TabellaSelezionata))
            AdoRsDati.MoveFirst
            While AdoRsDati.EOF = False
                AdoRsDati.Delete
                AdoRsDati.Update
                AdoRsDati.MoveNext
            Wend
        Set AdoRsDati = Nothing
        Call ChiudiConnessione
        End If
        Me.Enabled = True
        Screen.MousePointer = vbDefault
        MsgBox "Operation terminated correctly!", vbOKOnly, " Deleting Data"
    End If
    Call Blocca
End Sub




Private Sub Text1_GotFocus(Index As Integer)
    Posizione = 1
    mnuTrovaPwd.Enabled = True
End Sub
Private Sub Text2_GotFocus(Index As Integer)
    Posizione = 2
    mnuTrovaPwd.Enabled = True
End Sub
Private Sub TxtFile_GotFocus(Index As Integer)
    Posizione = Index + 1
    mnuTrovaPwd.Enabled = False
End Sub

Private Sub trwControlli_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
    Call Blocca
    If Index = 0 And Mid(Node.Key, 1, 1) = "R" Then
        If trwControlli(1).Nodes.Count > 0 Then Command3.Enabled = True
        Text3.Text = Node.Text
        TabellaSelezionata = Val(Mid(Node.Key, 2, 2))
        DatabaseSelezionato = 1
        If TxtFile(Index).Visible = False Then
            Command5.Enabled = True
            mnuDati.Enabled = True
            mnuAccodaDati.Enabled = True
            mnuCopiaDatiDB.Enabled = True
            mnuTrovaPwd.Enabled = True
        Else
            mnuCopiaDatiDB.Enabled = False
            mnuTrovaPwd.Enabled = False
        End If
    End If
    If Index = 0 And Mid(Node.Key, 2, 1) = "F" Then
        If trwControlli(1).Nodes.Count > 0 Then Command3.Enabled = True
        Text3.Text = Node.Parent.Text
        Text4.Text = Node.Text
        Text5.Text = Node.Child.Text
        Text6.Text = Node.Child.Next.Text
        Text7.Text = Mid(Node.Key, 3)
        DatabaseSelezionato = 1
        TabellaSelezionata = Val(Mid(Node.Parent.Key, 2, 2))
        If TxtFile(Index).Visible = False Then
            Command5.Enabled = True
            mnuDati.Enabled = True
            mnuAccodaDati.Enabled = True
            mnuCopiaDatiDB.Enabled = True
            mnuTrovaPwd.Enabled = True
         Else
            mnuCopiaDatiDB.Enabled = False
            mnuTrovaPwd.Enabled = False
        End If
    End If
    If Index = 1 And Mid(Node.Key, 1, 1) = "R" Then
        Command4.Enabled = True
        Text3.Text = Node.Text
        TabellaSelezionata = Val(Mid(Node.Key, 2, 2))
        DatabaseSelezionato = 2
        Command5.Enabled = True
        mnuDati.Enabled = True
        mnuSvuotaTabella.Enabled = True
    End If
    If Index = 1 And Mid(Node.Key, 2, 1) = "F" Then
        Command4.Enabled = True
        Text3.Text = Node.Parent.Text
        Text4.Text = Node.Text
        Text5.Text = Node.Child.Text
        Text6.Text = Node.Child.Next.Text
        Text7.Text = Mid(Node.Key, 3)
        DatabaseSelezionato = 2
        TabellaSelezionata = Val(Mid(Node.Parent.Key, 2, 2))
        Command5.Enabled = True
        mnuDati.Enabled = True
        mnuSvuotaTabella.Enabled = True
    End If
End Sub


Public Sub Blocca()
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Command3.Enabled = False
    Command4.Enabled = False
    Command5.Enabled = False
    mnuDati.Enabled = False
    mnuAccodaDati.Enabled = False
    mnuSvuotaTabella.Enabled = False
    TabellaSelezionata = -1
    DatabaseSelezionato = 0
End Sub


Private Sub mnuCopiaDatiDB_Click()
    On Error Resume Next
    If TxtFile(0).Visible = True Or TxtFile(1).Visible = True Then Exit Sub
    If ContTable <= 0 Or ContTable2 <= 0 Then Exit Sub
    strMsg = ""
    strMsg = strMsg + vbCr + " The operation will copy all the content of the Database 1 "
    strMsg = strMsg + vbCr + " in the Database 2, demanding some minute."
    strMsg = strMsg + vbCr
    strMsg = strMsg + vbCr + "    Continue with the operation ?"
    intMsg = MsgBox(strMsg, vbExclamation + vbYesNo + vbDefaultButton2, " Copy Database ")
    If intMsg = vbYes Then
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
        Dim AdoRsDati As New ADODB.Recordset
        Dim AdoRsDati2 As New ADODB.Recordset
        Dim ContCampi As Integer
        Dim ContTabelle As Integer
        Dim ContRigaRec As Long
        Dim ContFieldTable As Integer
        For ContTabelle = 0 To ContTable
            If VerificaTabella(ContTabelle) = True Then
                Call ApriConnessione(1)
                Call cData.NewRecordset(AdoRsDati)
                Call cData.mSQL("SELECT * FROM " & TableList(ContTabelle))
                DoEvents
                Call ApriConnessione(2, 2)
                Call cData.NewRecordset(AdoRsDati2)
                Call cData.mSQL("SELECT * FROM " & TableList(ContTabelle), 2)
                If AdoRsDati2.Fields.Count > 0 Then
                    If AdoRsDati2.RecordCount <= 0 Then
                        Text3.Text = TableList(ContTabelle)
                        DoEvents
                        AdoRsDati.MoveFirst
                        While AdoRsDati.EOF = False
                            AdoRsDati2.AddNew
                            ContRigaRec = ContRigaRec + 1
                            Text4.Text = ContRigaRec
                            DoEvents
                            ContFieldTable = AdoRsDati2.Fields.Count - 1
                            If ContFieldTable < AdoRsDati.Fields.Count - 1 Then ContFieldTable = AdoRsDati.Fields.Count - 1
                            For ContCampi = 0 To ContFieldTable
                                If VerificaCampo(ContTabelle, ContCampi) = True Then
                                    If TipoCampo < 8 Then AdoRsDati2.Fields(AdoRsDati.Fields(ContCampi).Name) = Val(AdoRsDati.Fields(ContCampi))
                                    If TipoCampo = 8 Then AdoRsDati2.Fields(AdoRsDati.Fields(ContCampi).Name) = Format(AdoRsDati.Fields(ContCampi), "DD/MM/YYYY")
                                    If TipoCampo = 10 Then AdoRsDati2.Fields(AdoRsDati.Fields(ContCampi).Name) = AdoRsDati.Fields(ContCampi)
                                End If
                            Next ContCampi
                            AdoRsDati2.Update
                            AdoRsDati.MoveNext
                            DoEvents
                        Wend
                        Text3.Text = ""
                        Text4.Text = ""
                        ContRigaRec = 0
                        DoEvents
                    End If
                End If
            End If
        Next ContTabelle
        Set AdoRsDati = Nothing
        Set AdoRsDati2 = Nothing
        Call ChiudiConnessione
        MsgBox "Operation terminated correctly!", vbOKOnly, " Copy Database"
    End If
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    Call Blocca
End Sub


Public Function VerificaTabella(Numero As Integer) As Boolean
    Dim contVerif As Integer
    VerificaTabella = False
    CurrentT2 = 1000
    For contVerif = 0 To ContTable2
        If TableList2(contVerif) = TableList(Numero) Then
            VerificaTabella = True
            CurrentT2 = contVerif
            Exit Function
        End If
    Next contVerif
End Function


Public Function VerificaCampo(NumeroT As Integer, NumeroC As Integer) As Boolean
    Dim contVerif As Integer
    VerificaCampo = False
    If CurrentT2 = 1000 Then Exit Function
    For contVerif = 0 To ContField2(CurrentT2)
        If FieldList2(CurrentT2, contVerif) = FieldList(NumeroT, NumeroC) Then
            VerificaCampo = True
            TipoCampo = FieldType2(CurrentT2, contVerif)
            Exit Function
        End If
    Next contVerif
End Function


Public Function LoadDBC(Index As Integer)
    Dim Percorso As String
    Dim riga As String
    Dim CurVariabile As String
    Dim CurCont As Double
    Dim CurCont2 As Double
    Dim Operazione As String
    Percorso = TxtFile(Index - 1).Text
    If Dir(Percorso) <> "" Then
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
        Open Percorso For Input As #1
        Do While Not EOF(1)
            DoEvents
            Line Input #1, riga
            If riga = "ContTable" Or riga = "TableList" Or riga = "ContField" Or riga = "FieldList" Or riga = "FieldType" Or riga = "FieldDim" Then
                CurVariabile = riga
                GoTo Successivo
            End If
            Select Case CurVariabile
                Case "ContTable"
                    If Index = 1 Then ContTable = Val(riga)
                    If Index = 2 Then ContTable2 = Val(riga)
                Case "TableList"
                    CurCont = ExtNumero(1, riga)
                    If Index = 1 Then TableList(CurCont) = ExtStringa(riga)
                    If Index = 2 Then TableList2(CurCont) = ExtStringa(riga)
                Case "ContField"
                    If IsNumeric(riga) = True Then
                        If Index = 1 Then ContField(CurCont) = Val(riga)
                        If Index = 2 Then ContField2(CurCont) = Val(riga)
                    Else
                        CurCont = ExtNumero(1, riga)
                    End If
                Case "FieldList"
                    If InStr(1, riga, ".") > 0 Then
                        CurCont = ExtNumero(1, riga)
                        CurCont2 = ExtNumero(2, riga)
                    Else
                        If Index = 1 Then FieldList(CurCont, CurCont2) = riga
                        If Index = 2 Then FieldList2(CurCont, CurCont2) = riga
                    End If
                Case "FieldType"
                    If IsNumeric(riga) = True Then
                        If Index = 1 Then FieldType(CurCont, CurCont2) = Val(riga)
                        If Index = 2 Then FieldType2(CurCont, CurCont2) = Val(riga)
                    Else
                        CurCont = ExtNumero(1, riga)
                        CurCont2 = ExtNumero(2, riga)
                    End If
                Case "FieldDim"
                    If IsNumeric(riga) = True Then
                        If Index = 1 Then FieldDim(CurCont, CurCont2) = Val(riga)
                        If Index = 2 Then FieldDim2(CurCont, CurCont2) = Val(riga)
                    Else
                        CurCont = ExtNumero(1, riga)
                        CurCont2 = ExtNumero(2, riga)
                    End If
            End Select
Successivo:
        Loop
        Close #1
    End If
End Function


Public Function ExtNumero(Index As Integer, ExStringa As String) As Double
    Dim ContExt As Long
    Dim StartExt As Long
    For ContExt = 1 To Len(ExStringa)
        If Mid(ExStringa, ContExt, 1) = "." Then
            If Index = 1 Then
                ExtNumero = Val(Left(ExStringa, ContExt - 1))
                Exit Function
            End If
            If Index = 2 Then
                If StartExt = 0 Then
                    StartExt = ContExt + 1
                Else
                    ExtNumero = Val(Mid(ExStringa, StartExt, ContExt - StartExt))
                    Exit Function
                End If
            End If
        End If
    Next ContExt
End Function


Public Function ExtStringa(ExStringa As String) As String
    Dim ContExt As Long
    For ContExt = Len(ExStringa) To 1 Step -1
        If Mid(ExStringa, ContExt, 1) = "." Then
            ExtStringa = Right(ExStringa, Len(ExStringa) - ContExt)
            Exit Function
        End If
    Next ContExt
End Function





Private Sub CmdXml1_Click()
    Dim xmlPath As String
    xmlPath = PercorsoXML(False)
    If xmlPath = "" Then Exit Sub
    Set adoGridRS = Nothing
    Call OpenXML(xmlPath, adoGridRS)
    DoEvents
    Set DataGrid1.DataSource = Nothing
    Set DataGrid1.DataSource = adoGridRS
    DataGrid1.Refresh
    If adoGridRS.RecordCount > 0 Then
        Call VerTabellaXML(xmlPath)
        If TabellaXML <> "" Then
            CmdXml3.Caption = "Puts Data XML in line to Table " & TabellaXML
            CmdXml3.Visible = True
        End If
    End If
End Sub

Private Sub CmdXml2_Click()
    Dim xmlPath As String
    xmlPath = PercorsoXML(True)
    Call SaveXML(xmlPath, , adoGridRS)
    MsgBox "Operation terminated correctly!", vbOKOnly + vbInformation, "DbCompare"
End Sub

Private Sub CmdXml3_Click()
    On Error Resume Next
    Dim AdorsUpdateXML As New ADODB.Recordset
    Dim ContFieldXml As Integer
    Dim ContFieldAdo As Integer
    If adoGridRS.RecordCount > 0 Then
        If DatabaseSelezionato = 1 Then cData.NewAdoRs AdorsUpdateXML, "SELECT Top 1 * FROM " & TabellaXML
        If DatabaseSelezionato = 2 Then cData.NewAdoRs AdorsUpdateXML, "SELECT Top 1 * FROM " & TabellaXML, , , , "cn2"
        ContFieldXml = AdorsUpdateXML.Fields.Count
        ContFieldAdo = adoGridRS.Fields.Count
        If ContFieldXml = ContFieldAdo Then
            adoGridRS.MoveFirst
            While adoGridRS.EOF = False
                AdorsUpdateXML.AddNew
                For ContFieldXml = 0 To ContFieldAdo - 1
                    If (AdorsUpdateXML.Fields(ContFieldXml).Name = adoGridRS.Fields(ContFieldXml).Name) And (AdorsUpdateXML.Fields(ContFieldXml).Type = adoGridRS.Fields(ContFieldXml).Type) Then
                        AdorsUpdateXML.Fields(ContFieldXml) = adoGridRS.Fields(ContFieldXml)
                    End If
                Next ContFieldXml
                AdorsUpdateXML.Update
                DoEvents
                adoGridRS.MoveNext
                DoEvents
            Wend
            adoGridRS.MoveFirst
            MsgBox "Operation terminated correctly!", vbOKOnly + vbInformation, "DbCompare"
        Else
            MsgBox "Impossible to carry out the demanded operation !" & vbCrLf & "The tables have not the same structure.", vbOKOnly + vbCritical, "DbCompare"
        End If
        Set AdorsUpdateXML = Nothing
    Else
        MsgBox "Impossible to carry out the demanded operation !" & vbCrLf & "There are not data to copy.", vbOKOnly + vbCritical, "DbCompare"
    End If
End Sub

Private Function PercorsoXML(Optional SaveFileXml As Boolean = True) As String
    On Error Resume Next
    With CommonDialog1
        .FileName = ""
        .CancelError = False
        .Filter = "File XML (*.xml)|*.xml"
        If SaveFileXml = True Then .ShowSave
        If SaveFileXml = False Then .ShowOpen
        If .FileName = "" Then Exit Function
        PercorsoXML = .FileName
        If UCase(Right(PercorsoXML, 4)) <> ".XML" Then PercorsoXML = PercorsoXML & ".xml"
    End With
End Function

Private Function VerTabellaXML(Percorso As String)
    On Error Resume Next
    Dim RigaXML As String
    Dim PosStart As Integer
    Dim PosEnd As Integer
    Call HideXml
    Open Percorso For Input As #1
    Do While Not EOF(1)
        DoEvents
        Line Input #1, RigaXML
        PosStart = InStr(1, RigaXML, "rs:basetable=")
        If PosStart > 0 Then
            PosStart = PosStart + 14
            PosEnd = PosStart + 1
            While Mid(RigaXML, PosEnd, 1) <> "'"
                PosEnd = PosEnd + 1
            Wend
            TabellaXML = Mid(RigaXML, PosStart, PosEnd - PosStart)
            Close #1
            Exit Function
        End If
    Loop
    Close #1
End Function

Private Sub HideXml()
    CmdXml3.Visible = False
    CmdXml3.Caption = "Puts Data XML in line to Table"
    TabellaXML = ""
End Sub

