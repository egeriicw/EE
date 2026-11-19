VERSION 5.00
Begin VB.Form NoDataDB 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Set No-data Flag"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3285
   LinkTopic       =   "Form1"
   ScaleHeight     =   1665
   ScaleWidth      =   3285
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CancelBut 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton OKBut 
      Caption         =   "OK"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox NDFlagBox 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Text            =   "-99"
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "No-Data Flag"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "NoDataDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelBut_Click()
NoDataDB.Hide
End Sub

Private Sub Form_Activate()

'sets textbox value
If ndflag = Null Or ndflag = Empty Then ndflag = -99
NDFlagBox.Text = Str$(ndflag)

End Sub

Private Sub Form_Load()
NoDataDB.Move Main.Left + (Main.Width - NoDataDB.Width) / 2, Main.Top + (Main.Height - NoDataDB.Height) / 2
End Sub

Private Sub OKBut_Click()

'sets ndflag
ndflag = Val(NDFlagBox.Text)

NoDataDB.Hide

End Sub


