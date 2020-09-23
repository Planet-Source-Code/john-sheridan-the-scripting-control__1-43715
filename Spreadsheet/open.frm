VERSION 5.00
Begin VB.Form openA 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open Sheet"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   Icon            =   "open.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   4
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   3480
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   3150
      Left            =   3000
      Pattern         =   "*.s+"
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   3465
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   2775
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "openA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error GoTo a
Dim la
la = "\"
If Dir1.Path = "C:\" Then la = ""

doc.openSheet Dir1.Path & la & File1.fileName
Unload Me
a:
'just an escape to not unload
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

