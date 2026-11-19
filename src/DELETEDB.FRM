VERSION 5.00
Begin VB.Form DeleteDB 
   Caption         =   "Delete Variables"
   ClientHeight    =   1785
   ClientLeft      =   285
   ClientTop       =   7200
   ClientWidth     =   3210
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1785
   ScaleWidth      =   3210
   Begin VB.ListBox VarBox 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      ItemData        =   "DELETEDB.frx":0000
      Left            =   240
      List            =   "DELETEDB.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton CancelBut 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton DeleteBut 
      Caption         =   "Delete"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "DeleteDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelBut_Click()
Unload DeleteDB
End Sub

Private Sub DeleteBut_Click()
Dim numdel, newfldnum, i, j, k
ReDim dvar(numflds)
On Error GoTo delbutce

Screen.MousePointer = 11

'finds which vars to delete and counts number of vars to delete
numdel = 0
For i = 0 To numfldsd - 1
   If VarBox.Selected(i) = True Then
      numdel = numdel + 1
      dvar(i + 1) = True
   Else
      dvar(i + 1) = False
   End If
Next i

If numdel > 0 Then 'fill dtemp() and vntemp$()
   ReDim dtemp(numrecs, numflds), vntemp$(numflds)
   For j = 1 To numrecs
      For k = 1 To numflds
         dtemp(j, k) = d(j, k)
      Next k
   Next j
   For k = 1 To numflds
      vntemp$(k) = vn$(k)
   Next k
Else 'show error message then exit sub
   MsgBox "Select the variables from the combo box that you want to delete.", , "Error"
   GoTo delbutes:
End If

'refill d() and vn$()
ReDim d(numrecs, numflds - numdel), vn$(numflds - numdel)
newfldnum = 0
For j = 1 To numflds
   If dvar(j) = False Then
      newfldnum = newfldnum + 1
      For i = 1 To numrecs
         d(i, newfldnum) = dtemp(i, j)
      Next i
      vn$(newfldnum) = vntemp$(j)
   End If
Next j
numflds = numflds - numdel    'note: numflds = numfldsd + 3
numfldsd = numfldsd - numdel

'calls open var to refill main.xvarbox and main.yvarbox
getvnsfrommerge = True
Call OpenVar("dummy")
getvnsfrommerge = False

delbutce:
ReDim dtemp(1, 1), vntemp$(1)
Unload DeleteDB
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume delbutes
End If
delbutes:
Screen.MousePointer = 0

End Sub

Private Sub Form_Load()
Dim i As Integer

DeleteDB.Move Main.Left + (Main.Width - DeleteDB.Width) / 2, Main.Top + (Main.Height - DeleteDB.Height) / 2

'fill currentvarbox and operatorbox
For i = 1 To numfldsd
   VarBox.AddItem vn$(i)
Next i

End Sub


