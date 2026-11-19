VERSION 5.00
Begin VB.Form AboutDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "About Energy Explorer"
   ClientHeight    =   2835
   ClientLeft      =   1485
   ClientTop       =   1965
   ClientWidth     =   5280
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
   ScaleHeight     =   2835
   ScaleWidth      =   5280
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   2085
      Left            =   -15
      MousePointer    =   99  'Custom
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   5310
   End
   Begin VB.CommandButton OKBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "AboutDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetFreeSystemResources Lib "User" (ByVal fuSysResource As Integer) As Integer
Private Declare Function GetFreeSpace Lib "Kernel" (ByVal wFlags As Integer) As Long

Private Sub Form_Activate()
Dim txt As String
Screen.MousePointer = 1


'loads message into txt
txt = Chr(13) + Chr(10)
txt = txt + "Energy Explorer 1.0" + Chr(13) + Chr(10)
txt = txt + Chr(13) + Chr(10)
'txt = txt + "DEVELOPMENT VERSION: NOT FOR DISTRIBUTION" + Chr(13) + Chr(10)
txt = txt + "Developed by:" + Chr(13) + Chr(10)
txt = txt + Chr(13) + Chr(10)
txt = txt + "Energy Information Systems" + Chr(13) + Chr(10)
txt = txt + "208 Lonsdale Ave." + Chr(13) + Chr(10)
txt = txt + "Dayton, Ohio 45419" + Chr(13) + Chr(10)
txt = txt + "937-294-6101" + Chr(13) + Chr(10)
txt = txt + Chr(13) + Chr(10)
txt = txt + "Copyright 2000" + Chr(13) + Chr(10)
txt = txt + "All rights reserved."

'prints message
Text1.Text = txt

End Sub

Private Sub Form_Load()
AboutDB.Move Main.Left + (Main.Width - AboutDB.Width) / 2, Main.Top + (Main.Height - AboutDB.Height) / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
Screen.MousePointer = 0
End Sub

Private Sub OKBut_Click()
   Unload AboutDB
End Sub

