VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5475
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4680
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      Picture         =   "frmAbout.frx":0614
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "by JS Works v.1.0"
      ForeColor       =   &H0000C000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "JSWIM - JS Works Instant Messanger"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   2520
      Width           =   5535
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "JSDB - JS DataBase+"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   5535
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "JSSM - JS Site Maker+"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   2040
      Width           =   5535
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "JSI - JS Imaging+"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "JSD - JS Draw+"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   5535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "JSJS - JS JohnScript+"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "JSS - JS Sheet+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1080
      Width           =   5535
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5640
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Part of the JS+ series."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "JS Sheet+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
