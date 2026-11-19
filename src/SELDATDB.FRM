VERSION 5.00
Begin VB.Form SelDatesDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Select Dates"
   ClientHeight    =   3855
   ClientLeft      =   990
   ClientTop       =   2310
   ClientWidth     =   5850
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
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3855
   ScaleWidth      =   5850
   Begin VB.CommandButton SetDelBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Set"
      Height          =   495
      Left            =   1920
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.TextBox EDelBox 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox SDelBox 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1080
      TabIndex        =   7
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton CancelBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cancel"
      Height          =   495
      Left            =   4440
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox EDateBox 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton OKBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OK"
      Height          =   495
      Left            =   4440
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox SDateBox 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Delete:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   735
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "To:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "From:"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1440
      TabIndex        =   9
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "End Date"
      ForeColor       =   &H80000008&
      Height          =   260
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Start Date"
      ForeColor       =   &H80000008&
      Height          =   260
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.Label TimeIntLbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   120
      TabIndex        =   12
      Top             =   165
      Width           =   5535
   End
End
Attribute VB_Name = "SelDatesDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelBut_Click()

Unload SelDatesDB

End Sub

Private Sub Form_Load()
Dim i As Integer

SelDatesDB.Move Main.Left + (Main.Width - SelDatesDB.Width) / 2, Main.Top + (Main.Height - SelDatesDB.Height) / 2

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
End Sub

Private Sub OKBut_Click()
Dim i As Integer, j As Integer, k As Integer, numdata As Integer
Dim txt, numslash
Dim sd$, ed$, sdate
Dim errflag As Integer
ReDim dtemp(numrecs, numflds)
On Error GoTo sdes

Screen.MousePointer = 11

'finds start date from sdatebox
sd$ = SDateBox.Text
If sd$ = "" Then
   startd = 0
Else
   txt = Trim(sd$)
   numslash = 0
   For i = Len(txt) To 1 Step -1
      If Mid$(txt, i, 1) = "/" Then numslash = numslash + 1
   Next i
   If timeint = 4 Then 'hourly data
      If numslash <> 3 Then
         Msg$ = "For hourly or sub-hourly data, please enter dates using the format: '7/1/1999/1'."
         MsgBox Msg$, , "Error"
         GoTo sdes:
      End If
   ElseIf timeint = 3 Then 'daily data
      If numslash <> 2 Then
         Msg$ = "For daily data, please enter dates using the format: '7/1/1999'."
         MsgBox Msg$, , "Error"
         GoTo sdes:
      End If
   ElseIf timeint = 2 Then 'monthly data
      If numslash <> 1 Then
         Msg$ = "For monthly data, please enter dates using the format: '7/1999'."
         MsgBox Msg$, , "Error"
         GoTo sdes:
      End If
   ElseIf timeint = 1 Then 'yearly data
      If numslash <> 0 Then
         Msg$ = "For yearly data, please enter dates using the format: '1999'."
         MsgBox Msg$, , "Error"
         GoTo sdes:
      End If
   End If
    startd = SerialDate(sd$, errflag)
End If

'finds end date from edatebox
ed$ = EDateBox.Text
If ed$ = "" Then
   endd = 9999999999#
Else
   txt = Trim(ed$)
   numslash = 0
   For i = Len(txt) To 1 Step -1
      If Mid$(txt, i, 1) = "/" Then numslash = numslash + 1
   Next i
   If timeint = 4 Then 'hourly data
      If numslash <> 3 Then
         Msg$ = "For hourly or sub-hourly data, please enter dates using the format: '7/1/1999/1'."
         MsgBox Msg$, , "Error"
         GoTo sdes:
      End If
   ElseIf timeint = 3 Then 'daily data
      If numslash <> 2 Then
         Msg$ = "For daily data, please enter dates using the format: '7/1/1999'."
         MsgBox Msg$, , "Error"
         GoTo sdes:
      End If
   ElseIf timeint = 2 Then 'monthly data
      If numslash <> 1 Then
         Msg$ = "For monthly data, please enter dates using the format: '7/1999'."
         MsgBox Msg$, , "Error"
         GoTo sdes:
      End If
   ElseIf timeint = 1 Then 'yearly data
      If numslash <> 0 Then
         Msg$ = "For yearly data, please enter dates using the format: '1999'."
         MsgBox Msg$, , "Error"
         GoTo sdes:
      End If
   End If
      
   endd = SerialDate(ed$, errflag)
End If

If errflag = True Then
   MsgBox "Enter dates exactly as shown", , "Error"
   GoTo sdes:
End If


For i = 1 To numrecs
   'finds the date for each record
   Call SerDate(i, sdate)
   
   'skips dates before and after start and end dates
   If sdate < startd Then GoTo NEXTREC:
   If sdate > endd Then GoTo DONE:
      
   'skips deleted dates
   For j = 1 To numbdate
      If sdate >= SDD(j) And sdate <= EDD(j) Then
          For k = 1 To numflds
             If LCase$(vn$(k)) <> "mo" And LCase$(vn$(k)) <> "dy" And LCase$(vn$(k)) <> "yr" And LCase$(vn$(k)) <> "hr" And vn$(k) <> "DOW" Then
                d(i, k) = ndflag
             End If
          Next k
       End If
   Next j
   
   'fills dtemp() with accepted data
   numdata = numdata + 1
   For k = 1 To numflds
      dtemp(numdata, k) = d(i, k)
   Next k
NEXTREC: Next i

DONE: 'Transfers dtemp() into d() and finds new numrecs
numrecs = numdata
For i = 1 To numrecs
    For j = 1 To numflds
       d(i, j) = dtemp(i, j)
    Next j
Next i

'display file name and numrecs in listbox
Main.ValueBar.Cls
Main.ValueBar.Print " " + UCase$(filename$) + "     Number Observations = " + Format$(numrecs, "####") + "     Number Columns = " + Format$(numflds, "###")

'error handling
Unload SelDatesDB
sdes:
Screen.MousePointer = 0
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume sdeses
End If
sdeses:
End Sub

Private Sub SetDelBut_Click()
Dim sdb$, edb$, errflag As Integer


'counts number of dates to delete
numbdate = numbdate + 1
 
'stores start and end dates
sdb$ = SDelBox.Text
edb$ = EDelBox.Text
'Call TxtToSer(sdb$, SDD(numbdate), errflag)
'Call TxtToSer(edb$, EDD(numbdate), errflag)
SDD(numbdate) = SerialDate(sdb$, errflag)
EDD(numbdate) = SerialDate(edb$, errflag)

If errflag = True Then
   MsgBox "Enter dates exactly as shown in Time Interval text box.", , "Error"
   GoTo sdbes:
End If

'makes end date = startdate if enddate is not entered
If SDD(numbdate) <> 0 And EDD(numbdate) = 0 Then
 EDD(numbdate) = SDD(numbdate)
End If

'clears text boxes
SDelBox.Text = ""
EDelBox.Text = ""
SDelBox.SetFocus

sdbes:
End Sub

