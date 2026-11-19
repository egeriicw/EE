VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form ESEditor 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "ES-Editor"
   ClientHeight    =   4950
   ClientLeft      =   1710
   ClientTop       =   2370
   ClientWidth     =   7320
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
   HelpContextID   =   310
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4950
   ScaleWidth      =   7320
   Tag             =   "ES-Editor"
   Begin RichTextLib.RichTextBox text1 
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8705
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"Eseditor.frx":0000
   End
   Begin VB.Menu FileCmd 
      Caption         =   "&File"
      HelpContextID   =   610
      Begin VB.Menu OpenCmd 
         Caption         =   "&Open"
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
      Begin VB.Menu NewCmd 
         Caption         =   "&New"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu SaveCmd 
         Caption         =   "&Save"
      End
      Begin VB.Menu SaveAsCmd 
         Caption         =   "Save &As"
      End
      Begin VB.Menu s3 
         Caption         =   "-"
      End
      Begin VB.Menu PrintAllCmd 
         Caption         =   "&Print All"
      End
      Begin VB.Menu PrintSelCmd 
         Caption         =   "P&rint Selection"
      End
      Begin VB.Menu z5 
         Caption         =   "-"
      End
      Begin VB.Menu LoadCmd 
         Caption         =   "Save and &Load"
      End
      Begin VB.Menu f1 
         Caption         =   "-"
      End
      Begin VB.Menu ExitCmd 
         Caption         =   "&Close"
      End
   End
   Begin VB.Menu EditCmd 
      Caption         =   "&Edit"
      HelpContextID   =   620
      Begin VB.Menu CutCmd 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu CopyCmd 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu PasteCmd 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu DeleteCmd 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu s4 
         Caption         =   "-"
      End
      Begin VB.Menu FindCmd 
         Caption         =   "&Find..."
      End
      Begin VB.Menu FindNextCmd 
         Caption         =   "Find Next"
         Enabled         =   0   'False
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "ESEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim txtchanged, filesaved


Private Sub CopyCmd_Click()

Clipboard.Clear
Clipboard.SetText text1.SelText

End Sub

Private Sub CutCmd_Click()

Clipboard.Clear
Clipboard.SetText text1.SelText
text1.SelText = ""

End Sub

Private Sub DeleteCmd_Click()

text1.SelText = ""

End Sub

Private Sub ExitCmd_Click()

If txtchanged = True And filesaved = False Then
  NotSave.Show 1
  If nsyes = True Then
     SaveAsCmd_Click
     If tnp$ = "" And tfn$ = "" Then Exit Sub Else filesaved = True
  ElseIf nscancel = True Then
     Exit Sub
  Else 'nsno = true
     filesaved = True
  End If
End If

fileselected = False

text1.Text = ""
Unload ESEditor

End Sub

Private Sub FindCmd_Click()
Dim txtposition

'call find dialog box
FindDB.Show 1

If findcancel = False And Len(findtxt$) > 0 Then
   'find string in text.
   txtposition = InStr(text1.Text, findtxt$)
   
   If txtposition Then  'findtxt$ is found
      'select the found text
      text1.SelStart = txtposition - 1
      text1.SelLength = Len(findtxt$)
      lastposition = txtposition
   Else 'findtxt$ is not found
      MsgBox "String not found.", , "Enertel Analysis"
   End If

   FindNextCmd.Enabled = True
End If


End Sub

Private Sub FindNextCmd_Click()
Dim txtposition

If Len(findtxt$) > 0 Then
   'find string in text.
   txtposition = InStr(lastposition + 1, text1.Text, findtxt$)
   
   If txtposition Then  'findtxt$ is found
      'select the found text
      text1.SelStart = txtposition - 1
      text1.SelLength = Len(findtxt$)
      lastposition = txtposition
   Else 'findtxt$ is not found
      MsgBox "String not found.", , "EModel"
   End If

   FindNextCmd.Enabled = True
End If

End Sub

Private Sub Form_Load()

'Move (Screen.Width - Editor.Width) / 2, (Screen.Height - Editor.Height) / 2
Move Main.Left + (Main.Width - ESEditor.Width) / 2, Main.Top + (Main.Height - ESEditor.Height) / 2

filesaved = False
txtchanged = False
newfile = True
opath$ = ""

'ofilename$ = ""
ofilename$ = opencode

'Text1.FontName = "FixedSys"
'Text1.FontBold = False
text1.Font.Name = "FixedSys"
text1.Font.Bold = False

'this bypasses the fileopen db routine in Editor.OpenCmd
If fileselected = True Then
   DoEvents
   Call ReadFile
End If

End Sub

Private Sub Form_Resize()

text1.Move 0, 0, ESEditor.ScaleWidth, ESEditor.ScaleHeight

End Sub

Private Sub Form_Unload(Cancel As Integer)
If txtchanged = True And filesaved = False Then
  NotSave.Show 1
  If nsyes = True Then
     SaveAsCmd_Click
     If tnp$ = "" And tfn$ = "" Then Exit Sub Else filesaved = True
  ElseIf nscancel = True Then
     Exit Sub
  Else 'nsno = true
     filesaved = True
  End If
End If

fileselected = False

text1.Text = ""
'Unload ESEditor

End Sub


Private Sub LoadCmd_Click()
Dim errorflag

SaveCmd_Click


End Sub

Private Sub NewCmd_Click()

ESEditor.Caption = "ES-Editor"
text1.Text = ""
newfile = True

End Sub

Private Sub OpenCmd_Click()
On Error GoTo vocce

'show fileopen db
Close
'opencode = "*.*"
FileOpen.Show 1
DoEvents
If tnp$ = "" And tfn$ = "" Then Exit Sub
opath$ = tnp$
ofilename$ = tfn$

Call ReadFile

vocce:
Close
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error in Open File Procedure"
   Resume voces
Else
   ESEditor.Caption = "ES-Editor " + tnp$ + tfn$
End If

voces:
End Sub

Private Sub PasteCmd_Click()

text1.SelText = Clipboard.GetText()

End Sub

Private Sub PrintAllCmd_Click()
Dim txt$
On Error GoTo epace

Screen.MousePointer = 11
txt$ = text1.Text
text1.Text = "Now printing..."
DoEvents
Printer.Print txt$
Printer.EndDoc
text1.Text = txt$

epace:
Screen.MousePointer = 0
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error in Printing Procedure"
   Resume epaes
End If
epaes:
End Sub

Private Sub PrintSelCmd_Click()
Dim seltxt$, alltxt$
On Error GoTo epsce

If text1.SelText <> "" Then

   Screen.MousePointer = 11
   seltxt$ = text1.SelText
   alltxt$ = text1.Text
   text1.Text = "Now printing selected text..."
   DoEvents
   Printer.Print seltxt$
   Printer.EndDoc
   text1.Text = alltxt$

End If

epsce:
Screen.MousePointer = 0
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error in Printing Procedure"
   Resume epses
End If
epses:
End Sub

Private Sub SaveAsCmd_Click()
On Error GoTo vsace

'call file save db
opencode = "*.*"
FileSave.Show 1
opencode = "*.*"
If tnp$ = "" And tfn$ = "" Then GoTo vsaes

'save text
Open tnp$ + tfn$ For Output As #1
Print #1, text1.Text
Close #1

If LCase$(Right$(tfn$, 3)) = "szb" Or LCase$(Right$(tfn$, 3)) = "mzb" Or LCase$(Right$(tfn$, 3)) = "utl" Then ESEditor.LoadCmd.Enabled = True


vsace:
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error in Save Procedure"
   Resume vsaes
Else
   'file saved properly
   ESEditor.Caption = "Editor " + tnp$ + tfn$
   filesaved = True
   newfile = False
End If

vsaes:
End Sub

Private Sub SaveCmd_Click()
On Error GoTo vsce

If newfile = True Then
   'user must designate a file name
   SaveAsCmd_Click
   Exit Sub
Else
   'save with present file name
   Open tnp$ + tfn$ For Output As #1
   Print #1, text1.Text
   Close #1
   filesaved = True
End If

If LCase$(Right$(tfn$, 3)) = "szb" Or LCase$(Right$(tfn$, 3)) = "mzb" Or LCase$(Right$(tfn$, 3)) = "utl" Then ESEditor.LoadCmd.Enabled = True

vsce:
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error in Save Procedure"
   Resume vses
End If


vses:

End Sub

Private Sub Text1_Change()

txtchanged = True

End Sub


Public Sub ReadFile()
On Error GoTo rfce

Dim txtfile$, newline$
'opens file and check length
'Open tnp$ + tfn$ For Input As #1
'If LOF(1) > 30000 Then
'   'file too big
'   Close
'   MsgBox "Sorry, this file is too large to edit.", , "Error"
'   Exit Sub
'Else
'   'reads file and places it in textbox
'   While Not EOF(1)
'      Line Input #1, newline$
'      txtfile$ = txtfile$ + newline$ + Chr(13) + Chr(10)
'   Wend
'   Text1.Text = txtfile$
'   DoEvents
'   txtchanged = False
'   newfile = False
'End If


'reads file and places it in textbox
Screen.MousePointer = 11
Open tnp$ + tfn$ For Input As #1
While Not EOF(1)
   Line Input #1, newline$
   txtfile$ = txtfile$ + newline$ + Chr(13) + Chr(10)
Wend
text1.Text = txtfile$
DoEvents
txtchanged = False
newfile = False


rfce:
Close
Screen.MousePointer = 0
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error in Open File Procedure"
   Resume rfes
Else
   ESEditor.Caption = "ES-Editor " + tnp$ + tfn$
End If

rfes:
End Sub
