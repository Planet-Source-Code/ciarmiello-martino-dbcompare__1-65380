VERSION 5.00
Begin VB.Form frm_MAPwd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Find Access Password"
   ClientHeight    =   1890
   ClientLeft      =   1095
   ClientTop       =   1515
   ClientWidth     =   3885
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "&Confirm"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1620
      TabIndex        =   4
      Top             =   1440
      Width           =   1035
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Find Password"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1395
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   1440
      Width           =   1035
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   900
      Width           =   2835
   End
   Begin VB.CommandButton Command1 
      Caption         =   "?"
      Height          =   315
      Index           =   0
      Left            =   3540
      TabIndex        =   1
      Top             =   360
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Password"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Path and Name of  Database MS Access 97"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "frm_MAPwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MODULE_NAME = "frm_MAPwd"


Private Sub Command2_Click()
    Unload Me
    Compare.Enabled = True
End Sub

Private Sub Command3_Click()
    If Dir(Text1.Text) <> "" Then
        Text3.Text = GetAccessPassWord(Text1.Text)
        Command4.Enabled = True
    Else
        MsgBox "Operation cancelled : The select directory it does not exist!", vbCritical, "DBCompare"
        Command4.Enabled = False
        Exit Sub
    End If
End Sub

Private Sub Command4_Click()
    Compare.Enabled = True
    If Posizione = 1 Then
        Compare.Text1(0).Text = Text1.Text
        Compare.Text1(1).Text = Text3.Text
    ElseIf Posizione = 2 Then
        Compare.Text2(0).Text = Text1.Text
        Compare.Text2(1).Text = Text3.Text
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    If Directory <> "" Then
        If UCase(Right(Directory, 4)) = ".MDB" Then Text1.Text = Directory
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Compare.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Directory = ""
    Compare.Enabled = True
End Sub

Private Sub Command1_Click(Index As Integer)
    Me.Enabled = False
    Text3.Text = ""
    frm_File.Show
End Sub


