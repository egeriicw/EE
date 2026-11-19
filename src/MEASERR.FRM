VERSION 5.00
Begin VB.Form MeasErrDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Define Measurement Error"
   ClientHeight    =   1650
   ClientLeft      =   2535
   ClientTop       =   2475
   ClientWidth     =   3930
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
   ScaleHeight     =   1650
   ScaleWidth      =   3930
   Begin VB.CommandButton CancelBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2160
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton OKBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox MeasErrTB 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Text            =   "5"
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Measurement Error (±%) ="
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   2295
   End
End
Attribute VB_Name = "MeasErrDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelBut_Click()

MeasErrTB.Text = Format(PerMeasErr)

Unload MeasErrDB

End Sub

Private Sub Form_Load()
MeasErrDB.Move Main.Left + (Main.Width - MeasErrDB.Width) / 2, Main.Top + (Main.Height - MeasErrDB.Height) / 2
MeasErrTB.Text = Format(PerMeasErr)

End Sub

Private Sub OKBut_Click()

If Val(MeasErrTB.Text) >= 0 And Val(MeasErrTB.Text) <= 200 Then
   
   PerMeasErr = Val(MeasErrTB.Text)

Else
   
   MsgBox "Measurement error must be between 0 and 200%.", , "Error"

End If

Unload MeasErrDB

End Sub

