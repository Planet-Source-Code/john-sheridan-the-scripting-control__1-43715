VERSION 5.00
Begin VB.Form addCells 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add cells"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "addCells.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Clear list"
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox ans 
      Height          =   285
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Compute"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.ListBox lst 
      Height          =   2595
      ItemData        =   "addCells.frx":030A
      Left            =   120
      List            =   "addCells.frx":030C
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label label2 
      Caption         =   "Answer:"
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Click a cell to add it to the list"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "addCells"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const HWND_TOPMOST = -&H1
Private Const HWND_NOTOPMOST = -&H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Declare Sub SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)


Dim total As Long

Function addNum(num As Long)
On Error GoTo a

total = total + num
lst.AddItem num
a:

End Function

Private Sub Command1_Click()
ans.Text = total
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
ans.Text = 0
total = 0
lst.clear
End Sub

Private Sub Form_Load()
    Call SetWindowPos(Me.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, (SWP_NOSIZE Or SWP_NOMOVE))

total = 0
ans.Text = 0
lst.clear
doc.addC = True
doc.MousePointer = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
doc.addC = False
doc.MousePointer = 0
End Sub
