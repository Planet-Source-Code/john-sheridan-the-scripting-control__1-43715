VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Using the MS Scripting control"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Code Samples"
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   4695
      Begin VB.CommandButton Command5 
         Caption         =   "Pythagorean"
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "If example"
         Height          =   255
         Left            =   2280
         TabIndex        =   14
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Loop"
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Default"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run code"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtRet 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   1695
   End
   Begin VB.TextBox txtCode 
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   1920
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "Form1.frx":0000
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox txt2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "End Function"
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Function returnStuff(x, y)"
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "Code: (x is value #1, y is #2)"
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Returns:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Value #2:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Value #1:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error GoTo youMessedUp
'error handling

Dim scriptingCTL
Set scriptingCTL = CreateObject("ScriptControl")
'create the scripting control

'the code for our little program:
Dim program As String
program = "Function returnStuff(x, y)" & _
            vbCrLf & txtCode.Text & vbCrLf & _
            "End Function"


scriptingCTL.language = "VBScript"
'be SURE to do this!
scriptingCTL.addcode program
'add our program to the Scripting Object

Dim strOutput As String
txtRet.Text = scriptingCTL.run("returnStuff", _
             txt1.Text, txt2.Text)
             
'ok this does two things:
'  1. runs our function with txt1.Text as the x argument
'     and txt2.Text as the y argument
'
'  2. it then puts the function's output
'     into a textbox


'OK! We're done. This is a very simple
'example, and the scripting object opens
'up many ways for the user to customize
'your program using VBScript.

'This has many practical uses (like in
' MS Excel, where you enter your own
' formulas for cells).

Exit Sub
youMessedUp:

If txt1.Text = "" Then GoTo esc
'if its just a simple mistake with no
'data, then skip the warning

MsgBox "Make sure your syntax is correct, and that you have the MSScripting Control!"

esc:
txtRet.Text = "Error occured."
End Sub

Private Sub Command2_Click()
txtCode = "returnStuff = Len(x) + InStr(y,""a"")"
End Sub

Private Sub Command3_Click()
MsgBox "This will find the factorial of Value #1 using a For... Next loop."

txtCode = vbCrLf & "Dim tmp" & vbCrLf & _
 "tmp = x" & vbCrLf & vbCrLf & "For i = x - 1 To 1 Step -1" & _
  vbCrLf & vbCrLf & "tmp = tmp * i" & _
  vbCrLf & vbCrLf & "Next" & vbCrLf & _
  vbCrLf & "returnStuff = tmp" & vbCrLf
End Sub

Private Sub Command4_Click()
MsgBox "Remember that the Scripting control is not case-sensitive, which is cool!"

txtCode = vbCrLf & "Dim tmp" & vbCrLf & _
          "tmp = ""my value""" & vbCrLf & _
          vbCrLf & "if tmp = ""my value"" then msgbox ""If's work!"""
          
End Sub

Private Sub Command5_Click()
MsgBox "This will find the length of the hypotenus of a right triangle, where" & vbCrLf & "value #1 and value #2 are the lengths of the legs."

txtCode = vbCrLf & "Dim tmp" & _
        vbCrLf & "tmp = (x * x) + (y * y)" & _
        vbCrLf & "'add the squares of the legs" & _
        vbCrLf & vbCrLf & _
        "returnStuff = sqr(tmp)" & vbCrLf & _
        "'Hey look! You can also use all your favorite VB functions, like Sqr(), Rnd() and trig stuff!"




End Sub
