VERSION 5.00
Begin VB.Form saveA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Save Sheet"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6885
   Icon            =   "saveA.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox sName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3720
      TabIndex        =   6
      Text            =   "Sheet1"
      Top             =   3720
      Width           =   3015
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2775
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   3915
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2775
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   3540
      Left            =   3000
      Pattern         =   "*.s+"
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Filename:"
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   3720
      Width           =   1335
   End
End
Attribute VB_Name = "savea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim la
la = "\"
If File1.Path = "C:\" Then la = ""

doc.saveSheet File1.Path & la & sName.Text & ".s+"

Unload Me
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
Dir1.Path = Drive1.Drive

End Sub
