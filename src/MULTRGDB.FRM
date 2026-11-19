VERSION 5.00
Begin VB.Form MultRgDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Multiple Regression"
   ClientHeight    =   3510
   ClientLeft      =   1425
   ClientTop       =   2325
   ClientWidth     =   4680
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
   ScaleHeight     =   3510
   ScaleWidth      =   4680
   Begin VB.CommandButton SetBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Set"
      Height          =   495
      Left            =   840
      TabIndex        =   3
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton CancelBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox VarBox 
      Appearance      =   0  'Flat
      Height          =   1935
      Left            =   1440
      Style           =   1  'Simple Combo
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.ComboBox RegBox 
      Appearance      =   0  'Flat
      Height          =   1935
      Left            =   120
      Style           =   1  'Simple Combo
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton OKBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Variable Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Regressor"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "MultRgDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelBut_Click()

Unload MultRgDB
cancelmlr = True

End Sub

Private Sub Form_Load()

MultRgDB.Move Main.Left + (Main.Width - MultRgDB.Width) / 2, Main.Top + (Main.Height - MultRgDB.Height) / 2

'initializes xfld()
For i = 1 To 7
   xfld(i) = 0
Next i

'initializes boxes
Do While RegBox.ListCount > 0
   RegBox.RemoveItem 0
   VarBox.RemoveItem 0
Loop
   
'fills variable box
For i = 1 To numflds
   VarBox.AddItem vn$(i)
Next i

'fills RegBox
RegBox.AddItem "Y"
RegBox.AddItem "X1"
RegBox.AddItem "X2"
RegBox.AddItem "X3"
RegBox.AddItem "X4"
RegBox.AddItem "X5"
RegBox.AddItem "X6"
RegBox.AddItem "X7"

End Sub

Private Sub OKBut_Click()

cancelmlr = False
Unload MultRgDB

End Sub

Private Sub SetBut_Click()
Dim i As Integer

If RegBox.ListIndex = 0 Then
   yfld = VarBox.ListIndex + 1
Else
   xfld(RegBox.ListIndex) = VarBox.ListIndex + 1
End If



End Sub

