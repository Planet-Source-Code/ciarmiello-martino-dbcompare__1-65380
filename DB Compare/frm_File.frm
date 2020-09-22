VERSION 5.00
Begin VB.Form frm_File 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Select Database MS Access"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   Icon            =   "frm_File.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&SET"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   60
      TabIndex        =   3
      Top             =   3480
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   60
      TabIndex        =   2
      Top             =   2340
      Width           =   4815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Chiudi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   3480
      Width           =   1935
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   4815
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4815
   End
End
Attribute VB_Name = "frm_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
    frm_MAPwd.Enabled = True
End Sub

Private Sub Command2_Click()
    If File1.FileName = "" Then
        MsgBox "Attention: the selected path is incorrect."
        Exit Sub
    Else
        If Right(Dir1.Path, 1) <> "\" Then
            Directory = Dir1.Path & "\" & File1.FileName
        Else
            Directory = Dir1.Path & File1.FileName
        End If
    End If
    Unload Me
    frm_MAPwd.Enabled = True
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
    On Error GoTo errore
    Dir1.Path = Drive1.Drive
    Exit Sub
    
errore:
    Drive1.Drive = "C:"
    Resume
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Drive1.Drive = "C:\"
    Dir1.Path = "C:\"
    File1.FileName = "C:\"
    File1.Pattern = "*.MDB"
    Me.Tag = "File"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frm_MAPwd.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frm_MAPwd.Enabled = True
End Sub
