VERSION 5.00
Begin VB.MDIForm Form1 
   BackColor       =   &H8000000C&
   Caption         =   "JS Sheet+"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      WindowList      =   -1  'True
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "Close"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuCellAdd 
         Caption         =   "Add up cells"
      End
      Begin VB.Menu mnuGraph 
         Caption         =   "Graph column"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fileOpen As Boolean
Public Function fixpath(lzpath As String)
    If Right$(lzpath, 1) = "\" Then fixpath = lzpath Else fixpath = lzpath & "\"
    
End Function


Function assocTypes()
Dim IconPath As String, ProgPath As String
   
    IconPath = fixpath(App.Path) & "file.ico" ' location of icon
    ProgPath = fixpath(App.Path) & App.EXEName & ".exe"   ' location of the program to load
    
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, ".s+"  ' your new file type
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, ".s+\DefaultIcon"  ' your new file types icon root
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, ".s+\shell"
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, ".s+\shell\open"
    Reg32Mod.SaveKey HKEY_CLASSES_ROOT, ".s+\shell\open\command"
    
    Reg32Mod.SaveString HKEY_CLASSES_ROOT, ".s+\DefaultIcon", "", IconPath ' your new filetype icon to use
    Reg32Mod.SaveString HKEY_CLASSES_ROOT, ".s+\shell\open\command", "", Chr(34) & ProgPath & Chr(34) & " %1"

End Function


Private Sub MDIForm_Load()
assocTypes
If Command = "" Then Exit Sub
doc.openSheet Command
End Sub





Private Sub mnuCellAdd_Click()
addCells.Show vbApplicationModal
End Sub

Private Sub mnuFileAbout_Click()
frmAbout.Show vbModal
End Sub

Private Sub mnuFileClose_Click()
If fileOpen = False Then Exit Sub
Unload Me.ActiveForm
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End Sub

Private Sub mnuFileNew_Click()
If fileOpen = True Then

Save% = MsgBox("Save your work?", vbYesNoCancel, "JS Sheet+")
Select Case Save%
  Case vbYes:
   If doc.fileName = "" Then
   savea.Show vbModal
   Else
   doc.saveSheet doc.fileName
   End If
  doc.clear
  
  Case vbNo:
  doc.clear
  
  Case Else:
  Exit Sub
End Select


Else
doc.Show
fileOpen = True
End If
End Sub

Private Sub mnuFileOpen_Click()
If fileOpen = True Then

Save% = MsgBox("Save your work?", vbYesNoCancel, "JS Sheet+")
Select Case Save%
  Case vbYes:
  
  If doc.fileName = "" Then
  savea.Show vbModal
  Else
  doc.saveSheet doc.fileName
  End If
  
  openA.Show vbModal
  
  Case vbNo:
  openA.Show vbModal
  
  Case Else:
  Exit Sub
End Select


Else
openA.Show vbModal
doc.Show
fileOpen = True
End If

End Sub

Private Sub mnuFileSave_Click()
If fileOpen = False Then Exit Sub

If doc.fileName = "" Then
savea.Show vbModal
Else
doc.saveSheet doc.fileName
End If

End Sub

Private Sub mnuFileSaveAs_Click()
If fileOpen = False Then Exit Sub

savea.Show vbModal

End Sub

Private Sub mnuGraph_Click()
MsgBox "This feature will be in the next version!", vbInformation, "JS Sheet+"
End Sub
