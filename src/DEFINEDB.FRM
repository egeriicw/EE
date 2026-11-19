VERSION 5.00
Begin VB.Form DefineDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Define Variable Names"
   ClientHeight    =   2985
   ClientLeft      =   1605
   ClientTop       =   1710
   ClientWidth     =   6285
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
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2985
   ScaleWidth      =   6285
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   2760
      ScaleHeight     =   1305
      ScaleWidth      =   3345
      TabIndex        =   12
      Top             =   4200
      Width           =   3375
      Begin VB.TextBox Group2Box 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Group1Box 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Group 2 Name"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   675
         Width           =   1335
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Group 1 Name"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   375
         TabIndex        =   15
         Top             =   285
         Width           =   1335
      End
   End
   Begin VB.CommandButton UseSavedBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Use Saved Names"
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton UseDefaultBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Use Default Names"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox NDFlagBox 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1680
      TabIndex        =   9
      Text            =   "-99"
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton SaveNameBut 
      BackColor       =   &H80000005&
      Caption         =   " Save and Load Names"
      Height          =   495
      Left            =   3840
      TabIndex        =   6
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton ChangeBut 
      BackColor       =   &H80000005&
      Caption         =   "Change"
      Default         =   -1  'True
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton CancelBut 
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3840
      TabIndex        =   7
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox NewNameBox 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton OKBut 
      BackColor       =   &H80000005&
      Caption         =   "Load Names"
      Height          =   495
      Left            =   3840
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
   Begin VB.ComboBox CurNameBox 
      Appearance      =   0  'Flat
      Height          =   2520
      Left            =   120
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "No-Data Flag"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "New Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Current Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "DefineDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim udvn$()

Private Sub CancelBut_Click()
   Unload DefineDB
End Sub

Private Sub ChangeBut_Click()

If CurNameBox.ListIndex = -1 Then
   MsgBox "Remember to click on a Current Variable.", , "Error"
   GoTo cbes:
ElseIf NewNameBox.Text = "" Then
   MsgBox "Remember to enter a name in the New Name text box.", , "Error"
   GoTo cbes:
End If

vn$(CurNameBox.ListIndex + 1) = NewNameBox.Text
CurNameBox.AddItem NewNameBox.Text, CurNameBox.ListIndex
CurNameBox.RemoveItem CurNameBox.ListIndex
NewNameBox.Text = ""
cbes:
End Sub

Private Sub CurNameBox_Click()

NewNameBox.Text = ""

End Sub

Private Sub Form_Load()
Dim i As Integer
ReDim udvn$(50)

DefineDB.Move Main.Left + (Main.Width - DefineDB.Width) / 2, Main.Top + (Main.Height - DefineDB.Height) / 2

'Sets default variable names
'SysDefBut.Value = False
'ChangeBut.Enabled = False

If vn$(1) = "" Then
    For i = 1 To numfldsd - 3
        vn$(i) = "V" + Format$(i, "00")
        CurNameBox.AddItem vn$(i)
    Next i
Else
    For i = 1 To numflds - 3
        CurNameBox.AddItem vn$(i)
    Next i
End If

If g1name$ = "" Then g1name$ = "G1"
If g2name$ = "" Then g2name$ = "G2"
Group1Box.Text = g1name$
Group2Box.Text = g2name$

definees:
End Sub

Private Sub OKBut_Click()
Dim i As Integer, numv As Integer
Dim yrflag As Integer, moflag As Integer, dyflag As Integer, hrflag As Integer

'sets ndflag
ndflag = Val(NDFlagBox.Text)

'sets group names
If Group1Box.Text = "" Then g1name$ = "G1" Else g1name$ = Group1Box.Text
If Group2Box.Text = "" Then g2name$ = "G2" Else g2name$ = Group2Box.Text

Unload DefineDB

'initializes variable boxes
Do While Main.XVarBox.ListCount > 0
   Main.XVarBox.RemoveItem 0
   Main.YVarBox.RemoveItem 0
Loop

For i = 1 To numfldsd
   'fills variable boxes with new names
   Main.XVarBox.AddItem vn$(i)
   Main.YVarBox.AddItem vn$(i)

   'finds fields of time stamp
   If LCase$(vn$(i)) = "yr" Then
      yrfld = i: yrflag = 1
   ElseIf LCase$(vn$(i)) = "mo" Then
      mofld = i: moflag = 1
   ElseIf LCase$(vn$(i)) = "dy" Then
      dyfld = i: dyflag = 1
   ElseIf LCase$(vn$(i)) = "hr" Then
      hrfld = i: hrflag = 1
   End If
Next i

'defines resid variables
vn$(numfldsd + 1) = "SG_Resids"
vn$(numfldsd + 2) = g1name$ + "_Resids"
vn$(numfldsd + 3) = g2name$ + "_Resids"
For i = 1 To 3
    Main.XVarBox.AddItem vn$(numfldsd + i)
    Main.YVarBox.AddItem vn$(numfldsd + i)
Next i
For i = 1 To numrecs
   For j = numfldsd + 1 To numfldsd + 3
      d(i, j) = ndflag
   Next j
Next i

'adds any remaining variable names to the list
For i = numfldsd + 4 To numflds
   Main.XVarBox.AddItem vn$(i)
   Main.YVarBox.AddItem vn$(i)
Next i

'finds timeint
timeint = yrflag + moflag + dyflag + hrflag
If timeint > 0 Then
    Main.SelDatesCmd.Enabled = True
    Main.ViewDatesCmd.Enabled = True
End If
   
'enables hourly profiles cmd
If timeint = 4 Then
   Main.HourlyProfilesBut.Enabled = True
Else
   Main.HourlyProfilesBut.Enabled = False
End If

'enables mod and create var commands
Main.CreateVariableCmd.Enabled = True
Main.DeleteVariableCmd.Enabled = True
Main.AvgIntDataCmd.Enabled = True
Main.ModValuesCmd.Enabled = True
Main.SaveDataCmd.Enabled = True
Main.SaveNamesCmd.Enabled = True
Main.GroupByDateCmd.Enabled = True
If yrflag + moflag + dyflag = 3 Then Main.GroupByDOWCmd.Enabled = True

End Sub

Private Sub SaveNameBut_Click()
Dim i As Integer

'tnp$ = ""
'tfn$ = ""
FileSave.Show MODAL

If tfn$ = "*.*" Or tfn$ = "" Then GoTo snbes:
Open tnp$ + tfn$ For Output As #1

For i = 1 To numfldsd
    Print #1, vn$(i)
Next i
For i = numfldsd + 4 To numflds
    Print #1, vn$(i)
Next i

Close #1

OKBut_Click

snbes:
End Sub

Private Sub SysDefBut_Click()
Dim i As Integer

'initializes variable boxes
Do While CurNameBox.ListCount > 0
   CurNameBox.RemoveItem 0
Loop

For i = 1 To numfldsd
    vn$(i) = "V" + Format$(i, "00")
    CurNameBox.AddItem vn$(i)
Next i

End Sub

Private Sub UseDefaultBut_Click()
Dim i As Integer

'initializes variable boxes
Do While CurNameBox.ListCount > 0
   CurNameBox.RemoveItem 0
Loop

For i = 1 To numfldsd
    vn$(i) = "V" + Format$(i, "00")
    CurNameBox.AddItem vn$(i)
Next i

End Sub

Private Sub UsePreviousBut_Click()
Dim i As Integer

FileOpen.Show MODAL

If tnp$ + tfn$ = "" Then Exit Sub

Open tnp$ + tfn$ For Input As #1

i = 0
Do While Not EOF(1)
    i = i + 1
    Input #1, vn$(i)
    If i = numfldsd Then Exit Do
Loop

If i <> numfldsd Then
    MsgBox "Caution: " + Str$(numfldsd) + "fields in the data file and " + Str$(i) + " variable names in the variable name file."
End If

Close #1

'initializes variable boxes
Do While CurNameBox.ListCount > 0
   CurNameBox.RemoveItem 0
Loop

For j = 1 To i
    CurNameBox.AddItem vn$(j)
Next j

End Sub

Private Sub UserDefBut_Click()
Dim i As Integer

FileOpen.Show MODAL

If tnp$ + tfn$ = "" Then Exit Sub

Open tnp$ + tfn$ For Input As #1

i = 0
Do While Not EOF(1)
    i = i + 1
    Input #1, vn$(i)
    If i = numfldsd Then Exit Do
Loop

If i <> numfldsd Then
    MsgBox "Caution: " + Str$(numfldsd) + "fields in the data file and " + Str$(i) + " variable names in the variable name file."
End If

Close #1

'initializes variable boxes
Do While CurNameBox.ListCount > 0
   CurNameBox.RemoveItem 0
Loop

For j = 1 To i
    CurNameBox.AddItem vn$(j)
Next j


End Sub

Private Sub UseSavedBut_Click()
Dim i As Integer

FileOpen.Show MODAL

If tnp$ + tfn$ = "" Then Exit Sub


Open tnp$ + tfn$ For Input As #1

i = 0
Do While Not EOF(1)
    i = i + 1
    Input #1, vn$(i)
    If i = numfldsd Then Exit Do
Loop

If i <> numfldsd Then
    MsgBox "Caution: " + Str$(numfldsd) + "fields in the data file and " + Str$(i) + " variable names in the variable name file."
End If

Close #1

'initializes variable boxes
Do While CurNameBox.ListCount > 0
   CurNameBox.RemoveItem 0
Loop

For j = 1 To i
    CurNameBox.AddItem vn$(j)
Next j

End Sub

