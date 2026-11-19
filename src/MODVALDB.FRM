VERSION 5.00
Begin VB.Form ModValDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Modify Variable"
   ClientHeight    =   2520
   ClientLeft      =   600
   ClientTop       =   2490
   ClientWidth     =   7215
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
   ScaleHeight     =   2520
   ScaleWidth      =   7215
   Begin VB.TextBox CnstBox 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4095
      TabIndex        =   7
      Top             =   435
      Width           =   1215
   End
   Begin VB.CommandButton SetBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Set"
      Height          =   495
      Left            =   2265
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton CancelBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5745
      TabIndex        =   4
      Top             =   1065
      Width           =   1215
   End
   Begin VB.CommandButton OKBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   5745
      TabIndex        =   3
      Top             =   360
      Width           =   1215
   End
   Begin VB.ComboBox ModBox 
      Appearance      =   0  'Flat
      Height          =   960
      Left            =   1920
      Style           =   1  'Simple Combo
      TabIndex        =   1
      Top             =   450
      Width           =   1800
   End
   Begin VB.ComboBox VarBox 
      Appearance      =   0  'Flat
      Height          =   1935
      Left            =   360
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   450
      Width           =   1215
   End
   Begin VB.Label CnstLabel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Constant"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4305
      TabIndex        =   8
      Top             =   165
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Modify by:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2340
      TabIndex        =   6
      Top             =   135
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Variable Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   90
      Width           =   1335
   End
End
Attribute VB_Name = "ModValDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelBut_Click()

Unload ModValDB

End Sub

Private Sub Form_Load()
ReDim cnst(numflds), f$(numflds)
Dim i As Integer

ModValDB.Move Main.Left + (Main.Width - ModValDB.Width) / 2, Main.Top + (Main.Height - ModValDB.Height) / 2
CnstBox.Visible = False
CnstLabel.Visible = False

'fill varbox and modbox
For i = 1 To numflds
   VarBox.AddItem vn$(i)
Next i
ModBox.AddItem "F to C"
ModBox.AddItem "C to F"
ModBox.AddItem "( ) + Constant"
ModBox.AddItem "( ) - Constant"
ModBox.AddItem "( ) * Constant"
ModBox.AddItem "( ) / Constant"
ModBox.AddItem "( ) ^ Constant"
ModBox.AddItem "exp( )"
ModBox.AddItem "ln( )"
ModBox.AddItem "abs( )"
ModBox.AddItem "sin( )"
ModBox.AddItem "cos( )"
ModBox.AddItem "Degs to Rads"
ModBox.AddItem "Rads to Degs"
ModBox.AddItem "MBtu to kWh"
ModBox.AddItem "kWh to MBtu"
ModBox.AddItem "( )+"
'initialize constant matrix Cnst() and flag matrix F()
For i = 1 To numflds
   cnst(i) = 1
   f$(i) = ""
Next i


End Sub

Private Sub ModBox_Click()

If ModBox.ListIndex = 2 Or ModBox.ListIndex = 3 Or ModBox.ListIndex = 4 Or ModBox.ListIndex = 5 Or ModBox.ListIndex = 6 Then
    CnstBox.Visible = True
    CnstBox.SetFocus
    CnstLabel.Visible = True
End If

End Sub

Private Sub OKBut_Click()
Dim i As Integer, j As Integer, enote As Integer, ftest As Integer
On Error GoTo mvce

'tests if the set button has been clicked
ftest = False
Do While j <= numflds - 1 And ftest = False
   j = j + 1
   If f$(j) <> "" Then ftest = True
Loop

If ftest = False Then
   MsgBox "No variable or modifier has been selected.  Remember to click the 'Set' button.", , "Error"
Else
    Unload ModValDB
    Screen.MousePointer = 11
    
    For i = 1 To numrecs
       For j = 1 To numflds
          If f$(j) = "ftoc" And d(i, j) <> ndflag Then
             d(i, j) = (d(i, j) - 32) / 1.8
          ElseIf f$(j) = "ctof" And d(i, j) <> ndflag Then
             d(i, j) = d(i, j) * 1.8 + 32
          ElseIf f$(j) = "+C" And d(i, j) <> ndflag Then
             d(i, j) = d(i, j) + cnst(j)
          ElseIf f$(j) = "-C" And d(i, j) <> ndflag Then
             d(i, j) = d(i, j) - cnst(j)
          ElseIf f$(j) = "*C" And d(i, j) <> ndflag Then
             d(i, j) = d(i, j) * cnst(j)
          ElseIf f$(j) = "/C" And d(i, j) <> ndflag Then
             d(i, j) = d(i, j) / cnst(j)
          ElseIf f$(j) = "^C" And d(i, j) <> ndflag Then
             d(i, j) = d(i, j) ^ cnst(j)
          ElseIf f$(j) = "exp" And d(i, j) <> ndflag Then
             d(i, j) = Exp(d(i, j))
          ElseIf f$(j) = "ln" And d(i, j) <> ndflag Then
             If d(i, j) > 0 Then
                d(i, j) = Log(d(i, j))
             Else
                d(i, j) = ndflag
             End If
          ElseIf f$(j) = "sqr" And d(i, j) <> ndflag Then
             d(i, j) = d(i, j) * d(i, j)
          ElseIf f$(j) = "sqrroot" And d(i, j) <> ndflag Then
             If d(i, j) > 0 Then
                d(i, j) = d(i, j) ^ 0.5
             Else
                d(i, j) = ndflag
             End If
          ElseIf f$(j) = "sin" And d(i, j) <> ndflag Then
             d(i, j) = Sin(d(i, j))
          ElseIf f$(j) = "cos" And d(i, j) <> ndflag Then
             d(i, j) = Cos(d(i, j))
          ElseIf f$(j) = "abs" And d(i, j) <> ndflag Then
             d(i, j) = Abs(d(i, j))
          ElseIf f$(j) = "dtor" And d(i, j) <> ndflag Then
             d(i, j) = d(i, j) * 2 * 3.1428 / 360
          ElseIf f$(j) = "rtod" And d(i, j) <> ndflag Then
             d(i, j) = d(i, j) * 360 / (2 * 3.1428)
          ElseIf f$(j) = "mbtutokwh" And d(i, j) <> ndflag Then
             d(i, j) = d(i, j) * 10 ^ 6 / 3412
          ElseIf f$(j) = "kwhtombtu" And d(i, j) <> ndflag Then
             d(i, j) = d(i, j) * 3412 / 10 ^ 6
          ElseIf f$(j) = "()+" And d(i, j) <> ndflag Then
             If d(i, j) < 0 Then d(i, j) = 0
          End If
       Next j
    Next i
End If

mvce:
Screen.MousePointer = 0
If Err = 5 Or Err = 6 Then
   d(i, j) = ndflag
   enote = True
   Resume Next
ElseIf Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume mves
End If
mves:
If enote = True Then MsgBox "An error has been encountered when modifying a value.  The errorroneus value(s) has been replaced with the no-data flag."
End Sub

Private Sub SetBut_Click()

If VarBox.ListIndex = -1 Or ModBox.ListIndex = -1 Then
   MsgBox "Both a variable and a modifier must be selected before clicking the 'Set' button.", , "Error"
Else
    Select Case ModBox.ListIndex
        Case 0
            f$(VarBox.ListIndex + 1) = "ftoc"
        Case 1
            f$(VarBox.ListIndex + 1) = "ctof"
        Case 2
            If CnstBox.Text = "" Then
                MsgBox "Specify the value of the constant in the Constant text box.", , "Error"
                GoTo sbes:
            End If
            cnst(VarBox.ListIndex + 1) = Val(CnstBox.Text)
            f$(VarBox.ListIndex + 1) = "+C"
            CnstBox.Text = ""
            CnstBox.Visible = False
            CnstLabel.Visible = False
        Case 3
            If CnstBox.Text = "" Then
                MsgBox "Specify the value of the constant in the Constant text box.", , "Error"
                GoTo sbes:
            End If
            cnst(VarBox.ListIndex + 1) = Val(CnstBox.Text)
            f$(VarBox.ListIndex + 1) = "-C"
            CnstBox.Text = ""
            CnstBox.Visible = False
            CnstLabel.Visible = False
        Case 4
            If CnstBox.Text = "" Then
                MsgBox "Specify the value of the constant in the Constant text box.", , "Error"
                GoTo sbes:
            End If
            cnst(VarBox.ListIndex + 1) = Val(CnstBox.Text)
            f$(VarBox.ListIndex + 1) = "*C"
            CnstBox.Text = ""
            CnstBox.Visible = False
            CnstLabel.Visible = False
        Case 5
            If CnstBox.Text = "" Then
                MsgBox "Specify the value of the constant in the Constant text box.", , "Error"
                GoTo sbes:
            End If
            cnst(VarBox.ListIndex + 1) = Val(CnstBox.Text)
            If cnst(VarBox.ListIndex + 1) = 0 Then
                MsgBox "Cannot divide by 0.", , "Error"
                GoTo sbes:
            End If
            f$(VarBox.ListIndex + 1) = "/C"
            CnstBox.Text = ""
            CnstBox.Visible = False
            CnstLabel.Visible = False
        Case 6
            If CnstBox.Text = "" Then
                MsgBox "Specify the value of the constant in the Constant text box.", , "Error"
                GoTo sbes:
            End If
            cnst(VarBox.ListIndex + 1) = Val(CnstBox.Text)
            f$(VarBox.ListIndex + 1) = "^C"
            CnstBox.Text = ""
            CnstBox.Visible = False
            CnstLabel.Visible = False
        Case 7
            f$(VarBox.ListIndex + 1) = "exp"
        Case 8
            f$(VarBox.ListIndex + 1) = "ln"
        Case 9
            f$(VarBox.ListIndex + 1) = "abs"
        Case 10
            f$(VarBox.ListIndex + 1) = "sin"
        Case 11
            f$(VarBox.ListIndex + 1) = "cos"
        Case 12
            f$(VarBox.ListIndex + 1) = "dtor"
        Case 13
            f$(VarBox.ListIndex + 1) = "rtod"
        Case 14
            f$(VarBox.ListIndex + 1) = "mbtutokwh"
        Case 15
            f$(VarBox.ListIndex + 1) = "kwhtombtu"
        Case 16
            f$(VarBox.ListIndex + 1) = "()+"
        Case Else
            MsgBox "Specify a modifier before clicking on the 'Set' button.", , "Error"
    End Select
End If

ModBox.Text = ""
VarBox.Text = ""

sbes:
End Sub

