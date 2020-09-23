<div align="center">

## The Scripting Control


</div>

### Description

I've seen too many submissions that say "~!!~~ WOW A++ NEW PROGRAMMING LANGUAGE MUST SEE!!". Then you open it, and its not a new language, it just uses the MS Scripting Control. And they use it WRONG too. I've written a little tutorial about this powerful control, and included two very good examples in the .zip file.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2003-03-03 21:09:32
**By**             |[john sheridan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/john-sheridan.md)
**Level**          |Beginner
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[The\_Script155400332003\.zip](https://github.com/Planet-Source-Code/john-sheridan-the-scripting-control__1-43715/archive/master.zip)





### Source Code

The Scripting Control allows users to put VBScript into your programs, to let them expand on the functions of the program.<br><br>
In the zip, I included a program that teaches you how to let users use VBScript in your programs and return values based on other values. Second, I included a spreadsheet program, which 1) saves and opens its own file type 2) associates its file type so it opens when you click the file 3) lets you type in formulas into the cells like in Excel. Now for the tutorial!<br><br><br>
First, I'll teach you how to create an instance of the Scripting Control without loading it as a component.<br><br>
<font face="Courier New">
Sub scriptTut()<br>
Dim scriptCTL 'dim the control<br>
Set scriptCTL = CreateObject("ScriptControl")<br>
'create the control<br><br>
scriptCTL.language = "VBScript" 'set the language<br><br>
'now you can run commands on the<br> ScriptControl just like a normal one<br>'ill teach you how to do that now<br><br>
End Sub<br><br></font>
<hr>
There are many ways to use the ScriptControl. The simplest way is to use ExecuteStatement. But, it isn't as functional as other methods. Here's how to use it:<br><br><font face="Courier New">
ScriptCTL.executestatement "MsgBox ""Hi!"""<br><br>'you can also set the scriptcontrol's variables quickly<br>
ScriptCTL.executestatement "myVar = 6"<br><br>
</font>
<hr>
That's the simplest way. The good way is to add functions to the control and then run them. This is very good because it lets people add their own functions and then evaluate them depending on other variables. To do that, you'd do something like this:<br><br><font face="Courier New">
Dim strProgram As String<br>
strProgram = "Function popUp(str)" & vbcrlf & _<br>
 "MsgBox str" & vbCrlf & _<br>
 "End Function"<br><br>
'now add the code to the control<br>
scriptCTL.addcode strProgram<br>
'note: you can also add the values of textBoxes, etc.
<br><br></font>
Now, to run our program's "popUp" function, simply do this:<br><br><font face="Courier">
scriptCTL.run "popUp", "penguins are cool"<br>
'this runs "popUp", with the first<br> byVal as what you want to pop up!<br><br><br>
</font>
Now here's a basic summary of this method that takes two numbers you write in textboxes and multiplies them. (In the zip, the first example does a similar thing, but lets the user input their own function!):<br><br><font face="Courier">
Private Sub Command1_Click()<br>
<br>'create the scriptcontrol here i'm too lazy
<br><br>Dim strProgram As String<br>
strProgram = "Function multiply(x,y)" & vbCrlf & _<br>
 "multiply=x*y" & vbcrlf & "End Sub"
<br><br>
ScriptCTL.addcode strProgram<br>
MsgBox ScriptCTL.run("multiply", Text1.Text, Text2.Text)<br>
'multiplies value of 2 textBoxes<br><br>
End Sub<br><br></font><hr>
Another useful thing is evaluating statements. This function just returns a boolean value from the ScriptControl. Very simple. Here's an example:<br><br><font face="Courier">
ScriptCTL.executestatement "x = 1"<br>
MsgBox ScriptCTL.eval("x=1") 'returns true<br>
MsgBox ScriptCTL.eval("x-5=x*x+2") 'false<br><br><br></font><hr>
Error Messages: This can let you (the deugger), or your user know when an error has occured. This is how to do it:<br><br><font face="Courier New">
Private Sub Command1_Click() <br>
'create scriptcontrol here<br><br>
ScriptCTL.executestatement "x=3/0"<br>
'(dividing by zero is not allowed in math)<br><br>
On Error Goto errHan<br>Exit Sub<br><br>
errHan:<br>
Debug.Print ScriptControl1.Error.Number & _<br>
	":" & ScriptControl1.Error.Description & _<br>
	" in line " & ScriptControl1.Error.Line<br><br>
End Sub<br><br><br></font><hr>Well, that concludes this basic tutorial of the Scripting Control. There is A LOT more you can do (like add code from modules in your VB project). This can all be found in the help file that comes with windows (at least in XP). It can be found in Windows\System32 and is named "MSScript.hlp". I hope you enjoyed and learned from this tutorial. Please download the sample code, it is very good. Please vote and give feedback, I spent a really long time on the programs and this tutorial. :-)

