VERSION 5.00
Begin VB.Form SystemDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Define Air-Side System"
   ClientHeight    =   5265
   ClientLeft      =   2025
   ClientTop       =   1680
   ClientWidth     =   6675
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
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5265
   ScaleWidth      =   6675
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Heating-Coil Set Point"
      ForeColor       =   &H80000008&
      Height          =   1452
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   3732
      Begin VB.OptionButton OpHDResetBut 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Optimized Reset"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   2892
      End
      Begin VB.OptionButton HDFixedBut 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fixed"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton HDResetBut 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Standard Outdoor Air Reset"
         ForeColor       =   &H80000008&
         Height          =   252
         Left            =   120
         TabIndex        =   11
         Top             =   375
         Width           =   3252
      End
      Begin VB.TextBox TempHCTB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   17
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label TempHCLab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   " Set-Point Temp (F)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1155
         TabIndex        =   18
         Top             =   1125
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cooling-Coil Set Point"
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   240
      TabIndex        =   6
      Top             =   4200
      Width           =   3735
      Begin VB.TextBox TempCCTB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Set-Point Temp (F)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Outside-Air Intake"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   1320
      Width           =   3735
      Begin VB.OptionButton OAFixedBut 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fixed"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton EconOnBut 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Economizer"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1815
      End
      Begin VB.TextBox FracOATB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3000
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
      Begin VB.Label FracOALab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Fraction Outside Air"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   765
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "System Type"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   6255
      Begin VB.TextBox VAVFlowTB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         TabIndex        =   20
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox CAVFlowTB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         TabIndex        =   19
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton VAVOnBut 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Variable-Air-Volume"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton CAVOnBut 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Constant-Air-Volume"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   2175
      End
      Begin VB.Label VAVFlowLab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Min Air Flow to Zones (cfm/ft2)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2370
         TabIndex        =   9
         Top             =   765
         Width           =   2895
      End
      Begin VB.Label CAVFlowLab 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Air Flow to Zones (cfm/ft2)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   21
         Top             =   405
         Width           =   2895
      End
   End
   Begin VB.CommandButton CancelBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton OKBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   1560
      Width           =   1335
   End
End
Attribute VB_Name = "SystemDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CancelBut_Click()

Unload SystemDB

End Sub

Private Sub CAVOnBut_Click()

If CAVOnBut.Value = True Then
   CAVFlowLab.Enabled = True
   CAVFlowTB.Enabled = True
   VAVFlowLab.Enabled = False
   VAVFlowTB.Enabled = False
End If
      
      
End Sub

Private Sub EconOnBut_Click()

If EconOnBut.Value = True Then
   FracOALab.Enabled = False
   FracOATB.Enabled = False
End If

End Sub

Private Sub Form_Load()

'center db
FileOpen.Move Main.Left + (Main.Width - SystemDB.Width) / 2, Main.Top + (Main.Height - SystemDB.Height) / 2

'fill textboxes
CAVFlowTB.Text = Format(cavflow)
VAVFlowTB.Text = Format(vavflow)
FracOATB.Text = Format(foafixed)
TempHCTB.Text = Format(thcfixed)
TempCCTB.Text = Format(tcc)

'enables or disables as required
If cavon = True Then
   CAVOnBut.Value = True
   CAVFlowLab.Enabled = True
   CAVFlowTB.Enabled = True
   VAVFlowLab.Enabled = False
   VAVFlowTB.Enabled = False
Else 'vav system
   VAVOnBut.Value = True
   CAVFlowLab.Enabled = False
   CAVFlowTB.Enabled = False
   VAVFlowLab.Enabled = True
   VAVFlowTB.Enabled = True
End If

If econon = True Then
   EconOnBut.Value = True
   FracOALab.Enabled = False
   FracOATB.Enabled = False
Else
   OAFixedBut.Value = True
   FracOALab.Enabled = True
   FracOATB.Enabled = True
End If

If hdreseton = True Then
   HDResetBut.Value = True
   TempHCLab.Enabled = False
   TempHCTB.Enabled = False
ElseIf ophdreseton = True Then
   OpHDResetBut.Value = True
   TempHCLab.Enabled = False
   TempHCTB.Enabled = False
Else
   HDFixedBut.Value = True
   TempHCLab.Enabled = True
   TempHCTB.Enabled = True
End If

End Sub

Private Sub HDFixedBut_Click()

If HDFixedBut.Value = True Then
   TempHCLab.Enabled = True
   TempHCTB.Enabled = True
End If

End Sub

Private Sub HDResetBut_Click()

If HDResetBut.Value = True Then
   TempHCLab.Enabled = False
   TempHCTB.Enabled = False
End If

End Sub

Private Sub OAFixedBut_Click()
   
If OAFixedBut.Value = True Then
   FracOALab.Enabled = True
   FracOATB.Enabled = True
End If

End Sub

Private Sub OKBut_Click()
Dim ecode

'checks values and fills global variables when necessary
If CAVOnBut.Value = True Then
   If CAVFlowTB.Text = "" Then
      ecode = True
      Msg$ = "A value for Air Flow To Zones must be entered."
   Else
      cavflow = Val(CAVFlowTB.Text)
      cavon = True
   End If
Else 'If VAVBut.Value = True Then
   If VAVFlowTB.Text = "" Then
      ecode = True
      Msg$ = "A value for Min Air Flow To Zones must be entered."
      GoTo emsg
   Else
      vavflow = Val(VAVFlowTB.Text)
      cavon = False
   End If
End If

If OAFixedBut.Value = True Then
   If FracOATB.Text = "" Then
      ecode = True
      Msg$ = "A value for Fraction Outside Air must be entered."
      GoTo emsg
   Else
      foafixed = Val(FracOATB.Text)
      econon = False
   End If
Else
   econon = True
End If

If HDFixedBut.Value = True Then
   If TempHCTB.Text = "" Then
      ecode = True
      Msg$ = "A value for Set-Point Temperature must be entered."
      GoTo emsg
   Else
      thcfixed = Val(TempHCTB.Text)
      hdreseton = False
      ophdreseton = False
   End If
ElseIf HDResetBut.Value = True Then
   hdreseton = True
   ophdreseton = False
Else 'OpHDResetBut.Value = true
   ophdreseton = True
   hdreseton = False
End If

If TempCCTB.Text = "" Then
   ecode = True
   Msg$ = "A value for Cooling-Coil Set-Point Temp must be entered."
   GoTo emsg
Else
   tcc = Val(TempCCTB.Text)
End If

emsg:
If ecode = True Then
   MsgBox Msg$, , "Error"
Else
   Unload SystemDB
End If

End Sub

Private Sub OpHDResetBut_Click()

If HDResetBut.Value = True Then
   TempHCLab.Enabled = False
   TempHCTB.Enabled = False
End If

End Sub

Private Sub VAVOnBut_Click()

If VAVOnBut.Value = True Then
   CAVFlowLab.Enabled = False
   CAVFlowTB.Enabled = False
   VAVFlowLab.Enabled = True
   VAVFlowTB.Enabled = True
End If

End Sub

