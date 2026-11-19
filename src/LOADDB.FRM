VERSION 5.00
Begin VB.Form LoadDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Define Building Load"
   ClientHeight    =   4065
   ClientLeft      =   1320
   ClientTop       =   2115
   ClientWidth     =   5745
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
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4065
   ScaleWidth      =   5745
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Occupied Zone"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      TabIndex        =   11
      Top             =   2280
      Width           =   3735
      Begin VB.TextBox NumOccsTB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox TempZoneTB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   12
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Number Occupants"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Set-Point Temp (F)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CheckBox LatOnCB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Latent Load On"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   3735
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Shell"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   3735
      Begin VB.TextBox UATB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Text            =   "20000"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Load Coef. (Btu/hr F)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Floor Area"
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   3735
      Begin VB.TextBox FracIntTB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   8
         Text            =   "0.65"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox AreaTotTB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Text            =   "300000"
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fraction Interior"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Total Area (ft2)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton CancelBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4200
      TabIndex        =   1
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton OKBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "LoadDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelBut_Click()

Unload LoadDB

End Sub

Private Sub Form_Load()

'center db
FileOpen.Move Main.Left + (Main.Width - LoadDB.Width) / 2, Main.Top + (Main.Height - LoadDB.Height) / 2

'fills text boxes with current values
AreaTotTB.Text = Format(floorarea)
FracIntTB.Text = Format(fi)
UATB.Text = Format(ua)
TempZoneTB.Text = Format(tr)
NumOccsTB.Text = Format(numoccs)
If laton = True Then LatOnCB.Value = 1 Else LatOnCB.Value = 0



End Sub

Private Sub OKBut_Click()
Dim ecode

ecode = False
If Val(AreaTotTB.Text) <= 0 Then
   ecode = True
   Msg$ = "The total conditioned area of the building must be entered."
ElseIf FracIntTB.Text = "" Or Val(FracIntTB.Text) < 0 Or Val(FracIntTB.Text) > 1 Then
   ecode = True
   Msg$ = "The fraction of the total conditioned area that lies in the interior zone of the building must be entered."
ElseIf Val(UATB.Text) <= 0 Then
   ecode = True
ElseIf Val(TempZoneTB.Text) <= 0 Then
   ecode = True
   Msg$ = "A value for Zone Set-Point Temperature must be entered."
ElseIf Val(NumOccsTB.Text) < 0 Then
   ecode = True
   Msg$ = "A zero or positive value for the number of building occupants must be entered."
End If

If ecode = True Then
   MsgBox Msg$, , "Error"
Else
   floorarea = Val(AreaTotTB.Text)
   fi = Val(FracIntTB.Text)
   ua = Val(UATB.Text)
   tr = Val(TempZoneTB.Text)
   numoccs = Val(NumOccsTB.Text)
   If LatOnCB.Value = 1 Then laton = True Else laton = False
   Unload LoadDB
End If

End Sub

