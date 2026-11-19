VERSION 5.00
Begin VB.Form AddGrpDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Group By Date"
   ClientHeight    =   3090
   ClientLeft      =   2145
   ClientTop       =   3195
   ClientWidth     =   6600
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
   ScaleHeight     =   3090
   ScaleWidth      =   6600
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   1425
      TabIndex        =   10
      Top             =   1695
      Width           =   3270
      Begin VB.TextBox Group2Box 
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   1410
         TabIndex        =   3
         Top             =   585
         Width           =   1440
      End
      Begin VB.TextBox Group1Box 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   1410
         TabIndex        =   2
         Top             =   225
         Width           =   1440
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Group 2 Name"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   75
         TabIndex        =   12
         Top             =   630
         Width           =   1305
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Group 1 Name"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   60
         TabIndex        =   11
         Top             =   285
         Width           =   1410
      End
   End
   Begin VB.TextBox EndDateBox 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3105
      TabIndex        =   1
      Top             =   1095
      Width           =   1560
   End
   Begin VB.CommandButton CancelBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5160
      TabIndex        =   5
      Top             =   1785
      Width           =   1215
   End
   Begin VB.CommandButton OKBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   5160
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox StartDateBox 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   1425
      TabIndex        =   0
      Top             =   1110
      Width           =   1590
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Start Date"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1815
      TabIndex        =   9
      Top             =   825
      Width           =   1005
   End
   Begin VB.Label TimeIntLbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   105
      TabIndex        =   8
      Top             =   240
      Width           =   6300
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "End Date"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3375
      TabIndex        =   7
      Top             =   825
      Width           =   915
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Group 1"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   165
      TabIndex        =   6
      Top             =   1170
      Width           =   855
   End
End
Attribute VB_Name = "AddGrpDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelBut_Click()
   agcancelflag = True
   Unload AddGrpDB
End Sub

Private Sub Form_Load()
Dim i As Integer

AddGrpDB.Move Main.Left + (Main.Width - AddGrpDB.Width) / 2, Main.Top + (Main.Height - AddGrpDB.Height) / 2

'displays time interval message
Select Case timeint
   Case 1
      TimeIntLbl.Caption = " For yearly data, use date format: '1999'"
   Case 2
      TimeIntLbl.Caption = " For monthly data, use date format: '1/1999'"
   Case 3
      TimeIntLbl.Caption = " For daily data, use date format: '1/2/1999'"
   Case 4
      TimeIntLbl.Caption = " For hourly or subhourly data, use date format: '1/2/1999/13'"
End Select

agcancelflag = False

If g1name$ = "" Then g1name$ = "G1"
If g2name$ = "" Then g2name$ = "G2"
Group1Box.Text = g1name$
Group2Box.Text = g2name$


End Sub

Private Sub OKBut_Click()
Dim txt, i, numslash
On Error GoTo okbutce

'init msg$
Msg$ = ""

'checks for proper data formats
If EndDateBox.Text = "" And StartDateBox.Text = "" Then
   Msg$ = "Must enter Start Date, End Date, or both."
   MsgBox Msg$, , "Error"
   GoTo okbutce:

ElseIf EndDateBox.Text <> "" Then
   txt = Trim(EndDateBox.Text)
   numslash = 0
   For i = Len(txt) To 1 Step -1
      If Mid$(txt, i, 1) = "/" Then numslash = numslash + 1
   Next i
   If timeint = 4 Then 'hourly data
      If numslash <> 3 Then
         Msg$ = "For hourly or sub-hourly data, please enter dates using the format: '7/1/1999/1'."
         MsgBox Msg$, , "Error"
         GoTo okbutce:
      End If
   ElseIf timeint = 3 Then 'daily data
      If numslash <> 2 Then
         Msg$ = "For daily data, please enter dates using the format: '7/1/1999'."
         MsgBox Msg$, , "Error"
         GoTo okbutce:
      End If
   ElseIf timeint = 2 Then 'monthly data
      If numslash <> 1 Then
         Msg$ = "For monthly data, please enter dates using the format: '7/1999'."
         MsgBox Msg$, , "Error"
         GoTo okbutce:
      End If
   ElseIf timeint = 1 Then 'yearly data
      If numslash <> 0 Then
         Msg$ = "For yearly data, please enter dates using the format: '1999'."
         MsgBox Msg$, , "Error"
         GoTo okbutce:
      End If
   End If

ElseIf StartDateBox.Text <> "" Then
   txt = Trim(StartDateBox.Text)
   numslash = 0
   For i = Len(txt) To 1 Step -1
      If Mid$(txt, i, 1) = "/" Then numslash = numslash + 1
   Next i
   If timeint = 4 Then 'hourly data
      If numslash <> 3 Then
         Msg$ = "For hourly or sub-hourly data, please enter dates using the format: '7/1/1999/1'."
         MsgBox Msg$, , "Error"
         GoTo okbutce:
      End If
   ElseIf timeint = 3 Then 'daily data
      If numslash <> 2 Then
         Msg$ = "For daily data, please enter dates using the format: '7/1/1999'."
         MsgBox Msg$, , "Error"
         GoTo okbutce:
      End If
   ElseIf timeint = 2 Then 'monthly data
      If numslash <> 1 Then
         Msg$ = "For monthly data, please enter dates using the format: '7/1999'."
         MsgBox Msg$, , "Error"
         GoTo okbutce:
      End If
   ElseIf timeint = 1 Then 'yearly data
      If numslash <> 0 Then
         Msg$ = "For yearly data, please enter dates using the format: '1999'."
         MsgBox Msg$, , "Error"
         GoTo okbutce:
      End If
   End If
End If

'close db if no error
If Msg$ = "" Then AddGrpDB.Hide

'error handling
okbutce:
Screen.MousePointer = 0
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume okbutes
End If
okbutes:
End Sub

