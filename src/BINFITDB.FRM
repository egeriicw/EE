VERSION 5.00
Begin VB.Form BinFitDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Bin Fit"
   ClientHeight    =   1695
   ClientLeft      =   2295
   ClientTop       =   3090
   ClientWidth     =   4800
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
   ScaleHeight     =   1695
   ScaleWidth      =   4800
   Begin VB.TextBox XMaxTB 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   972
   End
   Begin VB.TextBox XMinTB 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   972
   End
   Begin VB.CommandButton CancelBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton OKBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox NumIntsTB 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Xmin"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   7
      Top             =   360
      Width           =   1692
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Xmax"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   1572
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Num Intervals"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1572
   End
End
Attribute VB_Name = "BinFitDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelBut_Click()

binfitcancel = True

BinFitDB.Hide

End Sub

Private Sub Form_Load()
BinFitDB.Move Main.Left + (Main.Width - BinFitDB.Width) / 2, Main.Top + (Main.Height - BinFitDB.Height) / 2
End Sub

Private Sub OKBut_Click()

setminmax = False
If (XMaxTB.Text = "" And XMinTB.Text <> "") Or (XMaxTB.Text <> "" And XMinTB.Text = "") Then
   MsgBox "Values for both XMax and Xmin must be entered.", , "Error"
   GoTo sokes:
ElseIf XMaxTB.Text <> "" Then
   If Val(XMaxTB.Text) <= Val(XMinTB.Text) Then
      MsgBox "XMax must be greater than Xmin.", , "Error"
      GoTo sokes:
   End If
   xmaxset = Val(XMaxTB.Text)
   xminset = Val(XMinTB.Text)
   setminmax = True
End If

numints = Val(BinFitDB.NumIntsTB.Text)
If numints < 0 Then
   MsgBox "Enter a positive integer number of intervals.", , "Error"
   GoTo sokes:
ElseIf numints = 0 Then 'set default
   numints = 5
End If

BinFitDB.Hide

sokes:
End Sub

