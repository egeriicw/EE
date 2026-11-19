VERSION 5.00
Begin VB.Form SetScalesDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Set Scales"
   ClientHeight    =   3840
   ClientLeft      =   1125
   ClientTop       =   1125
   ClientWidth     =   5025
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
   ScaleHeight     =   3840
   ScaleWidth      =   5025
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "With Time Series-2"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   240
      TabIndex        =   17
      Top             =   3000
      Width           =   4575
      Begin VB.CheckBox SingleScalesCB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Use same y-scales for left and right axes "
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.CommandButton ClearBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Clear"
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox NumYLabelsTB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton OKBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   495
      Left            =   3600
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox NumXLabelsTB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   5
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox YmaxTB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox YminTB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox XmaxTB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox XminTB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Value"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   16
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Number Y Labels"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Number X Labels"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Y max"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Y min"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   30
      Left            =   330
      TabIndex        =   11
      Top             =   1050
      Width           =   15
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "X max"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "X min"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   735
   End
End
Attribute VB_Name = "SetScalesDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelBut_Click()
    SingleScalesCB.Value = 0
    Unload SetScalesDB
End Sub

Private Sub ClearBut_Click()

XMinTB.Text = ""
XMaxTB.Text = ""
NumXLabelsTB.Text = ""
YminTB.Text = ""
YmaxTB.Text = ""
NumYLabelsTB.Text = ""
SingleScalesCB.Value = 0

End Sub

Private Sub Form_Load()
SetScalesDB.Move Main.Left + (Main.Width - SetScalesDB.Width) / 2, Main.Top + (Main.Height - SetScalesDB.Height) / 2
End Sub

Private Sub OKBut_Click()
   
If NumXLabelsTB.Text <> "" And (XMaxTB.Text = "" Or XMinTB.Text = "") Then
   Msg$ = "If the number of X labels, or X-min, or X-max is set, then all three of these parameters must be set."
   MsgBox Msg$, , "Error"
ElseIf XMaxTB.Text <> "" And (NumXLabelsTB.Text = "" Or XMinTB.Text = "") Then
   Msg$ = "If the number of X labels, or X-min, or X-max is set, then all three of these parameters must be set."
   MsgBox Msg$, , "Error"
ElseIf XMinTB.Text <> "" And (XMaxTB.Text = "" Or NumXLabelsTB.Text = "") Then
   Msg$ = "If the number of X labels, or X-min, or X-max is set, then all three of these parameters must be set."
   MsgBox Msg$, , "Error"
ElseIf NumYLabelsTB.Text <> "" And (YmaxTB.Text = "" Or YminTB.Text = "") Then
   Msg$ = "If the number of Y labels, or Y-min, or Y-max is set, then all three of these parameters must be set."
   MsgBox Msg$, , "Error"
ElseIf YmaxTB.Text <> "" And (NumYLabelsTB.Text = "" Or YminTB.Text = "") Then
   Msg$ = "If the number of Y labels, or Y-min, or Y-max is set, then all three of these parameters must be set."
   MsgBox Msg$, , "Error"
ElseIf YminTB.Text <> "" And (YmaxTB.Text = "" Or NumYLabelsTB.Text = "") Then
   Msg$ = "If the number of Y labels, or Y-min, or Y-max is set, then all three of these parameters must be set."
   MsgBox Msg$, , "Error"
ElseIf (YminTB.Text <> "" And YmaxTB.Text <> "" And NumYLabelsTB.Text <> "") And (Val(YminTB.Text) > Val(YmaxTB.Text)) Then
   Msg$ = "Y-min must be less than Y-max."
   MsgBox Msg$, , "Error"
ElseIf (YminTB.Text <> "" And YmaxTB.Text <> "" And NumYLabelsTB.Text <> "") And (Val(NumYLabelsTB.Text) > 20) Then
   Msg$ = "Number of Y-labels must be less than 21."
   MsgBox Msg$, , "Error"
ElseIf (XMinTB.Text <> "" And XMaxTB.Text <> "" And NumXLabelsTB.Text <> "") And (Val(XMinTB.Text) > Val(XMaxTB.Text)) Then
   Msg$ = "X-min must be less than X-max."
   MsgBox Msg$, , "Error"
ElseIf (XMinTB.Text <> "" And XMaxTB.Text <> "" And NumXLabelsTB.Text <> "") And (Val(NumXLabelsTB.Text) > 20) Then
   Msg$ = "Number of X-labels must be less than 21."
   MsgBox Msg$, , "Error"
Else
   SetScalesDB.Hide
   If SingleScalesCB.Value = 0 And gtype$ = "TS" Then
      Main.TimeSeriesBut.Value = True
   ElseIf SingleScalesCB.Value = 0 And gtype$ = "XY" Then
      Main.XYBut.Value = True
   End If
End If

End Sub

