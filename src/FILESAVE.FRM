VERSION 5.00
Begin VB.Form FileSave 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Save As"
   ClientHeight    =   4905
   ClientLeft      =   735
   ClientTop       =   1155
   ClientWidth     =   8430
   ControlBox      =   0   'False
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
   HelpContextID   =   330
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4905
   ScaleWidth      =   8430
   Begin VB.CommandButton DlgBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   6840
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   2790
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   4215
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   4320
      Width           =   4215
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   2760
      Left            =   4560
      TabIndex        =   2
      Top             =   1440
      Width           =   3735
   End
   Begin VB.CommandButton DlgBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   5160
      TabIndex        =   1
      Top             =   4320
      Width           =   1215
   End
   Begin VB.TextBox TextBox 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   4560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label PathSpec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "FileSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LastChange As Integer
Const tbc = 1, dbc = 2

Private Sub Dir1_Change()

PathSpec.Caption = Dir1.Path
File1.Path = Dir1.Path

End Sub

Private Sub Dir1_Click()

LastChange = dbc

End Sub

Private Sub DlgBut_Click(Index As Integer)

If Index = 0 Then 'if ok button then ...
   Select Case LastChange
      Case tbc 'last change was in the text box
      
         ColonPos% = InStr(TextBox.Text, ":")
         BSlashPos% = InStr(TextBox.Text, "\")
         StarPos% = InStr(TextBox.Text, "*")
         QueryPos% = InStr(TextBox.Text, "?")
         If ColonPos% Then
             Drive1.Drive = Mid$(TextBox.Text, ColonPos% - 1, 2)
         End If
         If BSlashPos% Then
             Test$ = TextBox.Text
             Do Until InStr(lastone% + 1, Test$, "\") = 0
               lastone% = InStr(lastone% + 1, Test$, "\")
             Loop
            
             If lastone% > 3 Then
               Path$ = Left$(TextBox.Text, lastone% - 1)
             Else
               Path$ = TextBox.Text
             End If

             Dir1.Path = Path$
         End If
        
         'assign path and file name
         A$ = TextBox.Text
         tnp$ = Dir1.Path
         If Right$(tnp$, 1) <> "\" Then tnp$ = tnp$ + "\"
         tfn$ = Right$(A$, Len(A$) - lastone%)
         FileSave.Hide
         File1.Pattern = "*.*"
         
      Case dbc
         Dir1.Path = Dir1.List(Dir1.ListIndex)
   End Select

   TextBox.SelStart = 0
   TextBox.SelLength = Len(TextBox.Text)
   If Visible Then TextBox.SetFocus
Else
   'cancel button
   tnp$ = ""
   tfn$ = ""
   FileSave.Hide
End If


End Sub

Private Sub Drive1_Change()

On Error Resume Next
Dir1.Path = Drive1.Drive

If Err Then
    MsgBox (Error$)
    Drive1.Drive = Dir1.Path
End If

End Sub

Private Sub File1_Click()

TextBox.Text = File1.filename

End Sub

Private Sub File1_DblClick()

tnp$ = Dir1.Path
If Right$(tnp$, 1) <> "\" Then tnp$ = tnp$ + "\"
tfn$ = TextBox.Text
File1.Pattern = "*.*"
FileSave.Hide

End Sub

Private Sub File1_PathChange()

Dir1.Path = File1.Path
Drive1.Drive = File1.Path

End Sub

Private Sub Form_Load()

FileSave.Move Main.Left + (Main.Width - FileSave.Width) / 2, Main.Top + (Main.Height - FileSave.Height) / 2

'this form is coded from pages 305 to 313
'of the VB programers guide and can be used in
'any application

If tnp$ <> "" Then
   If Right$(tnp$, 1) = "\" Then
      tnp$ = Left$(tnp$, (Len(tnp$) - 1))
   End If
   If tnp$ <> File1.Path Then
      File1.Path = tnp$
   End If
End If

If tfn$ <> "" Then
   For ind% = 0 To File1.ListCount
      If File1.List(ind%) = tfn$ Then
         TextBox.Text = tfn$
         Exit For
      End If
   Next ind%
End If

PathSpec.Caption = Dir1.Path
If File1.ListIndex = -1 Then
   TextBox.Text = File1.Pattern
End If
TextBox.SelStart = 0
TextBox.SelLength = Len(TextBox.Text)


End Sub

Private Sub TextBox_Change()

LastChange = tbc

End Sub

