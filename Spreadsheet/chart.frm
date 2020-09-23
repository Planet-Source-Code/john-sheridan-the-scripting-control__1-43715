VERSION 5.00
Begin VB.Form chart 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Chart"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   Icon            =   "chart.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox chrt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8490
      Left            =   0
      ScaleHeight     =   8490
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   0
      Width           =   11880
      Begin VB.Line line 
         Index           =   0
         X1              =   600
         X2              =   1320
         Y1              =   3000
         Y2              =   3000
      End
   End
End
Attribute VB_Name = "chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
