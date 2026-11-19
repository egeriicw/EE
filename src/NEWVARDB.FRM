VERSION 5.00
Begin VB.Form NewVarDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Create A New Variable"
   ClientHeight    =   2790
   ClientLeft      =   3045
   ClientTop       =   3525
   ClientWidth     =   7635
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2790
   ScaleWidth      =   7635
   Begin VB.CommandButton OKBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   360
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton CancelBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   360
      TabIndex        =   9
      Top             =   1800
      Width           =   1215
   End
   Begin VB.ComboBox OperatorBox 
      Appearance      =   0  'Flat
      Height          =   1155
      Left            =   5625
      Style           =   1  'Simple Combo
      TabIndex        =   6
      Top             =   1215
      Width           =   1215
   End
   Begin VB.ComboBox CurrentVarBox 
      Appearance      =   0  'Flat
      Height          =   1350
      Left            =   3120
      Style           =   1  'Simple Combo
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox EquationBox 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   480
      Width           =   4815
   End
   Begin VB.TextBox NewVarBox 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Operators"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5760
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Current Variables"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   960
      Width           =   1695
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Equation"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4440
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "="
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "New Variable Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "NewVarDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim numobjs As Integer, t$
Dim Obj$(20), VarNum(20) As Integer

Private Sub CancelBut_Click()

Unload NewVarDB

End Sub

Private Sub CurrentVarBox_Click()

numobjs = numobjs + 1
Obj$(numobjs) = CurrentVarBox.Text
VarNum(numobjs) = CurrentVarBox.ListIndex + 1
t$ = t$ + Obj$(numobjs) + " "
EquationBox.Text = t$

End Sub

Private Sub Form_Load()
Dim i As Integer

NewVarDB.Move Main.Left + (Main.Width - NewVarDB.Width) / 2, Main.Top + (Main.Height - NewVarDB.Height) / 2

'initialize variables
numobjs = 0
t$ = ""
For i = 1 To 20
    VarNum(i) = 0
    Obj$(i) = ""
Next i

'fill currentvarbox and operatorbox
For i = 1 To numflds
   CurrentVarBox.AddItem vn$(i)
Next i
OperatorBox.AddItem "*"
OperatorBox.AddItem "/"
OperatorBox.AddItem "+"
OperatorBox.AddItem "-"

End Sub

Private Sub OKBut_Click()
ReDim modvar(numrecs), dtemp(numrecs, numflds), vntemp$(numflds)
On Error GoTo nvce

'reads in new variable name and tests
nvn$ = NewVarBox.Text
If nvn$ = "" Then
    MsgBox "Remember to specify the new variable's name.", , "Error"
    GoTo nves:
ElseIf numobjs <> 1 And numobjs <> 3 Then
    MsgBox "The new variable can only be 1) a duplicate of a current variable or 2) created from two current variables seperated by one operator.  More complex variables can be constructed by multiple use of this and the Data, Modify Variable procedures.", , "Error"
    GoTo nves:
End If
Screen.MousePointer = 11

'calcs new variable and puts it into DTemp()
ReDim dtemp(numrecs, numflds + 1)
If numobjs = 1 Then
    'new variable is a duplicate of a current variable
    If VarNum(1) <= 0 Or VarNum(1) > numflds Then
        MsgBox "The new variable can only be 1) a duplicate of a current variable or 2) created from two current variables seperated by one operator.  More complex variables can be constructed by multiple use of this and the Data, Modify Values procedures.", , "Error"
        GoTo nves:
    End If
    For i = 1 To numrecs
        For j = 1 To numflds
           dtemp(i, j) = d(i, j)
        Next j
        
        If d(i, VarNum(1)) = ndflag Then
            'dtemp(i, numflds + 1) = ndflag
            modvar(i) = ndflag
        Else
            'dtemp(i, numflds + 1) = d(i, VarNum(1))
            modvar(i) = d(i, VarNum(1))
        End If
    Next i
Else
    'new variable is a combination of current variables
    If VarNum(1) <= 0 Or VarNum(1) > numflds Or VarNum(3) <= 0 Or VarNum(3) > numflds Then
        MsgBox "The new variable can only be 1) a duplicate of a current variable or 2) created from two current variables seperated by one operator.  More complex variables can be constructed by multiple use of this and the Data, Modify Values procedures.", , "Error"
        GoTo nves:
    ElseIf Obj$(2) <> "+" And Obj$(2) <> "-" And Obj$(2) <> "*" And Obj$(2) <> "/" Then
        MsgBox "The new variable can only be 1) a duplicate of a current variable or 2) created from two current variables seperated by one operator.  More complex variables can be constructed by multiple use of this and the Data, Modify Values procedures.", , "Error"
        GoTo nves:
    End If
    For i = 1 To numrecs
        For j = 1 To numflds
           dtemp(i, j) = d(i, j)
        Next j
        
        If d(i, VarNum(1)) = ndflag Or d(i, VarNum(3)) = ndflag Then
            'dtemp(i, numflds + 1) = ndflag
            modvar(i) = ndflag

        Else
            Select Case Obj$(2)
                Case "*"
                    'dtemp(i, numflds + 1) = d(i, VarNum(1)) * d(i, VarNum(3))
                    modvar(i) = d(i, VarNum(1)) * d(i, VarNum(3))
                Case "/"
                    'dtemp(i, numflds + 1) = d(i, VarNum(1)) / d(i, VarNum(3))
                    modvar(i) = d(i, VarNum(1)) / d(i, VarNum(3))
                Case "+"
                    'dtemp(i, numflds + 1) = d(i, VarNum(1)) + d(i, VarNum(3))
                    modvar(i) = d(i, VarNum(1)) + d(i, VarNum(3))
                Case "-"
                    'dtemp(i, numflds + 1) = d(i, VarNum(1)) - d(i, VarNum(3))
                    modvar(i) = d(i, VarNum(1)) - d(i, VarNum(3))
                Case Else
                    MsgBox "Operator not identified.", , "Error"
                    GoTo nves:
            End Select
        End If
    Next i
End If

If 1 = 2 Then 'old way which adds var after resid vars
   'fills D()
   numflds = numflds + 1
   ReDim d(numrecs, numflds)
   For i = 1 To numrecs
        For j = 1 To numflds
           d(i, j) = dtemp(i, j)
        Next j
   Next i
   
   'fills VN$()
   ReDim vntemp$(numflds)
   For i = 1 To numflds - 1
      vntemp$(i) = vn$(i)
   Next i
   vntemp$(i) = nvn$
   ReDim vn$(numflds)
   For i = 1 To numflds
      vn$(i) = vntemp$(i)
   Next i
   
   Main.YVarBox.AddItem nvn$
   Main.XVarBox.AddItem nvn$
End If

'redims and fills d() [moving 3 resid vars to last]
ReDim d(numrecs, numflds + 1)
For i = 1 To numrecs
   For j = 1 To numflds - 3
      d(i, j) = dtemp(i, j)
   Next j
   d(i, numflds - 2) = modvar(i)
   For j = numflds - 1 To numflds + 1
      d(i, j) = dtemp(i, j - 1)
   Next j
Next i

'fills vn$() [moving 3 resid vars to last]
For i = 1 To numflds
   vntemp$(i) = vn$(i)
Next i
ReDim vn$(numflds + 1)
For i = 1 To numflds - 3
   vn$(i) = vntemp$(i)
Next i
vn$(numflds - 2) = nvn$
For i = numflds - 1 To numflds + 1
   vn$(i) = vntemp$(i - 1)
Next i

'resets numflds
numflds = numflds + 1
numfldsd = numfldsd + 1

'initializes variable boxes
Do While Main.XVarBox.ListCount > 0
   Main.XVarBox.RemoveItem 0
   Main.YVarBox.RemoveItem 0
Loop

'fills variable boxes with new names
For i = 1 To numflds
   Main.XVarBox.AddItem vn$(i)
   Main.YVarBox.AddItem vn$(i)
Next i

'reinitializes timescale variables
timeinthr = False
timeintdy = False
timeintmo = False
timeintyr = False

nvce:
Unload NewVarDB
If Err = 6 Then
    dtemp(i, numflds + 1) = ndflag
    Resume Next
ElseIf Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume nves
End If
nves:
Screen.MousePointer = 0
End Sub

Private Sub OperatorBox_Click()

numobjs = numobjs + 1
Obj$(numobjs) = OperatorBox.Text
t$ = t$ + Obj$(numobjs) + " "
EquationBox.Text = t$

End Sub

