VERSION 5.00
Begin VB.Form TSOptsDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Time Scale Options"
   ClientHeight    =   3990
   ClientLeft      =   1380
   ClientTop       =   1995
   ClientWidth     =   6120
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
   ScaleHeight     =   3990
   ScaleWidth      =   6120
   Begin VB.OptionButton AvailOpt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "If NoData Then Calculate Avg/Tot Using Available Data"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2640
      Width           =   5655
   End
   Begin VB.OptionButton NDOpt 
      BackColor       =   &H00FFFFFF&
      Caption         =   "If NoData Then Set Avg/Tot To NoData"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2280
      Value           =   -1  'True
      Width           =   5655
   End
   Begin VB.CommandButton ToTotBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "<---------"
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton ToAvgBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "---------->"
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.ListBox AvgBox 
      Appearance      =   0  'Flat
      Height          =   1590
      Left            =   3840
      MultiSelect     =   2  'Extended
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.ListBox TotBox 
      Appearance      =   0  'Flat
      Height          =   1590
      ItemData        =   "TSOPTSDB.frx":0000
      Left            =   240
      List            =   "TSOPTSDB.frx":0002
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton CancelBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton OKBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Average"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "TSOptsDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelBut_Click()

tsoptscancel = True
TSOptsDB.Hide

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
Dim i
ReDim totvars(numflds)

TSOptsDB.Move Main.Left + (Main.Width - TSOptsDB.Width) / 2, Main.Top + (Main.Height - TSOptsDB.Height) / 2

'fills totbox with non-timestamp variables
For i = 1 To numflds
   If LCase$(vn$(i)) <> "mo" And LCase$(vn$(i)) <> "dy" And LCase$(vn$(i)) <> "yr" And LCase$(vn$(i)) <> "hr" And LCase$(vn$(i)) <> "mn" Then
      TotBox.AddItem vn$(i)
   End If
Next i

End Sub

Private Sub OKBut_Click()
Dim i, j

'initialize totvars()
'if totvars(i) = 0 then average the variable
'elseif totvars(i) = 1 then total the variable
For i = 1 To numflds
   totvars(i) = 0
Next i

'sets totvar(i) = 1 for all variables in totbox
For i = 0 To TotBox.ListCount - 1
   For j = 1 To numflds
      If TotBox.List(i) = vn$(j) Then totvars(j) = 1
   Next j
Next i

'sets global variable to decide how to process no data flags
If NDOpt.Value = True Then
   makend = True
Else
   makend = False
End If

tsoptscancel = False
'sets tsoptsshown if tsoptdb is used for timescale changes
'so that tsoptsdb will only be shown the first time
'someone changes timescales
If NDOpt.Visible = True Then tsoptsshown = True

'hides tsoptsdb
TSOptsDB.Hide
DoEvents

End Sub

Private Sub ToAvgBut_Click()
Dim i

'moves selected variable names from totbox to avgbox
For i = 0 To TotBox.ListCount - 1
   If i > TotBox.ListCount - 1 Then Exit For
   If TotBox.Selected(i) = True Then
      AvgBox.AddItem TotBox.List(i)
      TotBox.RemoveItem i
      i = i - 1
   End If
Next i

End Sub

Private Sub ToTotBut_Click()
Dim i

'moves selected variable names from avgbox to totbox
For i = 0 To AvgBox.ListCount - 1
   If i > AvgBox.ListCount - 1 Then Exit For
   If AvgBox.Selected(i) = True Then
      TotBox.AddItem AvgBox.List(i)
      AvgBox.RemoveItem i
      i = i - 1
   End If
Next i

End Sub

