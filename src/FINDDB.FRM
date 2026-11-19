VERSION 5.00
Begin VB.Form FindDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Find"
   ClientHeight    =   1425
   ClientLeft      =   2580
   ClientTop       =   2385
   ClientWidth     =   6240
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
   ScaleHeight     =   1425
   ScaleWidth      =   6240
   Begin VB.CommandButton CancelBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   372
      Left            =   4800
      TabIndex        =   3
      Top             =   840
      Width           =   1212
   End
   Begin VB.CommandButton FindBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Find Next"
      Default         =   -1  'True
      Height          =   372
      Left            =   4800
      TabIndex        =   2
      Top             =   240
      Width           =   1212
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   372
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   3012
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Find What"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1212
   End
End
Attribute VB_Name = "FindDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelBut_Click()

findcancel = True
Unload FindDB

End Sub

Private Sub FindBut_Click()

If Len(Text1.Text) > 0 Then
   findtxt$ = Text1.Text
   Unload FindDB
End If

End Sub

Private Sub Form_Activate()

FindBut.Enabled = False
findtxt$ = ""
findcancel = False

If ESEditor.Text1.SelLength > 0 Then Text1.Text = ESEditor.Text1.SelText

End Sub

Private Sub Form_Load()

Move (Screen.Width - FindDB.Width) / 2, (Screen.Height - FindDB.Height) / 2

End Sub

Private Sub Text1_Change()

FindBut.Enabled = True

End Sub

