VERSION 5.00
Begin VB.Form NotSave 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Editor"
   ClientHeight    =   1680
   ClientLeft      =   3150
   ClientTop       =   2310
   ClientWidth     =   5040
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
   ScaleHeight     =   1680
   ScaleWidth      =   5040
   Begin VB.CommandButton CancelBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   492
      Left            =   3600
      TabIndex        =   2
      Top             =   960
      Width           =   1212
   End
   Begin VB.CommandButton NoBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "No"
      Height          =   492
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   1212
   End
   Begin VB.CommandButton YesBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Yes"
      Default         =   -1  'True
      Height          =   492
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1212
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Loaded file is not saved.  Save it now?"
      ForeColor       =   &H80000008&
      Height          =   372
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   4452
   End
End
Attribute VB_Name = "NotSave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelBut_Click()

nscancel = True
Unload NotSave

End Sub

Private Sub Form_Load()
nsyes = False
nsno = False
nscancel = False
Move (Screen.Width - NotSave.Width) / 2, (Screen.Height - NotSave.Height) / 2
End Sub

Private Sub NoBut_Click()

nsno = True
Unload NotSave

End Sub

Private Sub YesBut_Click()

nsyes = True
Unload NotSave

End Sub

