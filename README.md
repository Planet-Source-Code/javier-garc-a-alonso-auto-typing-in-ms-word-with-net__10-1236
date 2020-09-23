<div align="center">

## Auto\-Typing in MS Word with \.NET\!


</div>

### Description

This code opens MS Word, creates a new document, and starts writing a text in it letter by letter.Just create a new proyect, a normal windows aplication, add a reference to Microsoft Word Object Library, and copy/paste this code. Then set form opacity to 0%. Easy to implement and Cool effect, isn't?
 
### More Info
 
Don´t forget to add a reference to Microsoft Word 9.0 Object Library! or 10.0 if you are Office XP user!


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Javier García Alonso](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/javier-garc-a-alonso.md)
**Level**          |Beginner
**User Rating**    |4.0 (8 globes from 2 users)
**Compatibility**  |VB\.NET
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__10-1.md)
**World**          |[\.Net \(C\#, VB\.net\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/net-c-vb-net.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/javier-garc-a-alonso-auto-typing-in-ms-word-with-net__10-1236/archive/master.zip)





### Source Code

```
'Message that will appear in MS Word
 Private Shared message As String = "Hey, How are you doing?"
 'Variable to iterate
 Private Shared i As Integer
 'Timer Object
 Private Shared WithEvents temp As New Timer()
 'Word Objects: application object and document object
 Private Shared aplicationword As New Word.Application()
 Private Shared documentword As New Word.Document()
 'Boolean variable to escape of while loop
 Private Shared exiting As Boolean = True
 Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
  'Word is visible and activate
  aplicationword.Visible = True
  aplicationword.Activate()
  'One document added to Word
  documentword = aplicationword.Documents.Add
  'Added an event handler
  AddHandler temp.Tick, AddressOf TimerEventProcessor
  'Setting timer properties (delay)
  temp.Interval = 350
  'Yeah!, let's go
  temp.Start()
  i = 0
  While exiting = False
   Application.DoEvents()
  End While
 End Sub
 Private Shared Sub TimerEventProcessor(ByVal sender As Object, ByVal e As System.EventArgs)
  'If I've finished writting
  If i = message.Length Then
   temp.Stop()
   exiting = True
   Exit Sub
  End If
  'else writes a new letter
  aplicationword.Selection.TypeText(message.Chars(i))
  i += 1
 End Sub
```

