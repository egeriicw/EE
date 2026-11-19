VERSION 5.00
Begin VB.Form RotateDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Hourly Profiles Rotation Angle"
   ClientHeight    =   1560
   ClientLeft      =   2505
   ClientTop       =   2370
   ClientWidth     =   3960
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
   ScaleHeight     =   1560
   ScaleWidth      =   3960
   Begin VB.CommandButton CancelBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   492
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   1332
   End
   Begin VB.CommandButton OKBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   492
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1332
   End
   Begin VB.TextBox RotAngleTB 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   972
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Rotation Angle (0-90 deg)"
      ForeColor       =   &H80000008&
      Height          =   252
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2532
   End
End
Attribute VB_Name = "RotateDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelBut_Click()

Unload RotateDB

End Sub

Private Sub Form_Load()

RotateDB.Move Main.Left + (Main.Width - RotateDB.Width) / 2, Main.Top + (Main.Height - RotateDB.Height) / 2

RotAngleTB.Text = Format(thetad, "##0")

End Sub

Private Sub OKBut_Click()

If Val(RotAngleTB.Text) < 0 Or Val(RotAngleTB.Text) > 90 Then
   MsgBox "Rotation angle must be between 0 and 90 degrees.", , "Error"
Else
   thetad = Val(RotAngleTB.Text)
   Unload RotateDB
End If

End Sub

