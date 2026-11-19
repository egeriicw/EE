VERSION 5.00
Begin VB.Form Main 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Energy Explorer"
   ClientHeight    =   6570
   ClientLeft      =   1230
   ClientTop       =   2865
   ClientWidth     =   10455
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
   Icon            =   "MAIN.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6570
   ScaleWidth      =   10455
   Begin VB.CommandButton FivePMVRBut 
      Caption         =   "M"
      Height          =   495
      Left            =   4680
      TabIndex        =   30
      Top             =   6000
      Width           =   255
   End
   Begin VB.CommandButton FourPMVRBut 
      Caption         =   "M"
      Height          =   495
      Left            =   3720
      TabIndex        =   29
      Top             =   6000
      Width           =   255
   End
   Begin VB.CommandButton ThreePMVRBut 
      Caption         =   "M"
      Height          =   255
      Index           =   1
      Left            =   2760
      TabIndex        =   28
      Top             =   6240
      Width           =   255
   End
   Begin VB.CommandButton ThreePMVRBut 
      Caption         =   "M"
      Height          =   255
      Index           =   0
      Left            =   2760
      TabIndex        =   27
      Top             =   6000
      Width           =   255
   End
   Begin VB.CommandButton FivePBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "5P"
      Height          =   495
      Left            =   3960
      TabIndex        =   26
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton HourlyProfilesBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hourly Profiles"
      Enabled         =   0   'False
      Height          =   400
      Left            =   7785
      TabIndex        =   25
      Top             =   5640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton MedianBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Median"
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   6240
      Width           =   855
   End
   Begin VB.CommandButton HistBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Histogram"
      Height          =   400
      Left            =   7800
      TabIndex        =   23
      Top             =   5160
      Width           =   2535
   End
   Begin VB.CommandButton BinFitBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Bin Fit"
      Height          =   500
      Left            =   5880
      TabIndex        =   22
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton ClearGraphBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Height          =   195
      Left            =   7440
      TabIndex        =   21
      Top             =   1320
      Width           =   255
   End
   Begin VB.CommandButton ClearValueBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Height          =   195
      Left            =   7440
      TabIndex        =   20
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton TimeSeries2But 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "2"
      Height          =   400
      Left            =   9840
      TabIndex        =   19
      Top             =   3720
      Width           =   495
   End
   Begin VB.CommandButton XYAnimateBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "XY Animate"
      Height          =   400
      Left            =   7785
      TabIndex        =   18
      Top             =   4695
      Width           =   2535
   End
   Begin VB.CommandButton ThreePBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "3P-H"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   17
      Top             =   6240
      Width           =   735
   End
   Begin VB.OptionButton G1G2But 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "WD/WE"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8280
      TabIndex        =   11
      Top             =   6120
      Width           =   255
   End
   Begin VB.OptionButton SingleGroupBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Single Group"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6720
      TabIndex        =   10
      Top             =   6120
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton ThreePBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "3P-C"
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   16
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton FourPBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "4P"
      Height          =   500
      Left            =   3000
      TabIndex        =   2
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton MLRBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "M"
      Height          =   500
      Left            =   1920
      TabIndex        =   9
      Top             =   6120
      Width           =   255
   End
   Begin VB.CommandButton TwoPBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "2P"
      Height          =   500
      Left            =   1080
      TabIndex        =   1
      Top             =   6000
      Width           =   735
   End
   Begin VB.CommandButton MeanBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Mean"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   6000
      Width           =   855
   End
   Begin VB.CommandButton XYBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "XY"
      Height          =   400
      Left            =   7800
      TabIndex        =   8
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CommandButton TimeSeriesBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "TimeSeries"
      Height          =   400
      Left            =   7800
      TabIndex        =   7
      Top             =   3720
      Width           =   2055
   End
   Begin VB.ComboBox XVarBox 
      Appearance      =   0  'Flat
      Height          =   1350
      Left            =   7800
      Style           =   1  'Simple Combo
      TabIndex        =   3
      Top             =   2175
      Width           =   2535
   End
   Begin VB.PictureBox GraphBox 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   0
      ScaleHeight     =   -125
      ScaleLeft       =   -15
      ScaleMode       =   0  'User
      ScaleTop        =   110
      ScaleWidth      =   125
      TabIndex        =   0
      Top             =   1320
      Width           =   7695
   End
   Begin VB.ComboBox YVarBox 
      Appearance      =   0  'Flat
      Height          =   1350
      Left            =   7800
      Style           =   1  'Simple Combo
      TabIndex        =   4
      Top             =   375
      Width           =   2535
   End
   Begin VB.PictureBox ValueBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   0
      ScaleHeight     =   1305
      ScaleWidth      =   7665
      TabIndex        =   12
      Top             =   0
      Width           =   7695
   End
   Begin VB.Label G2Lab 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "G2"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8880
      TabIndex        =   15
      Top             =   6150
      Width           =   375
   End
   Begin VB.Label G1Lab 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      Caption         =   "G1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8520
      TabIndex        =   14
      Top             =   6150
      Width           =   375
   End
   Begin VB.Label XVarLab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "X Variable"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7905
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label YVarLab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Y Variable"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7920
      TabIndex        =   6
      Top             =   135
      Width           =   1215
   End
   Begin VB.Menu FileM 
      Caption         =   "&File"
      Begin VB.Menu OpenDataCmd 
         Caption         =   "Open &Data from Text File..."
      End
      Begin VB.Menu OpenNamesCmd 
         Caption         =   "Open &Names from Text File..."
      End
      Begin VB.Menu zf1 
         Caption         =   "-"
      End
      Begin VB.Menu PrintCmd 
         Caption         =   "&Print"
      End
      Begin VB.Menu zf2 
         Caption         =   "-"
      End
      Begin VB.Menu SaveDataCmd 
         Caption         =   "Save Da&ta..."
      End
      Begin VB.Menu SaveNamesCmd 
         Caption         =   "Save Na&mes..."
      End
      Begin VB.Menu SaveLogCmd 
         Caption         =   "Save &Log..."
         Visible         =   0   'False
      End
      Begin VB.Menu zf3 
         Caption         =   "-"
      End
      Begin VB.Menu ExitCmd 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu EditCmd 
      Caption         =   "&Edit"
      Begin VB.Menu CopyValueCmd 
         Caption         =   "Copy &ValueBar"
      End
      Begin VB.Menu CopyGraphCmd 
         Caption         =   "Copy &Graph"
      End
      Begin VB.Menu e1 
         Caption         =   "-"
      End
      Begin VB.Menu EditorCmd 
         Caption         =   "&Editor"
      End
   End
   Begin VB.Menu DataM 
      Caption         =   "&Data"
      Begin VB.Menu VarNamesCmd 
         Caption         =   "&Name Variables..."
      End
      Begin VB.Menu SetNodataFlagCmd 
         Caption         =   "Set N&o-Data Flag"
      End
      Begin VB.Menu zd1 
         Caption         =   "-"
      End
      Begin VB.Menu SelDatesCmd 
         Caption         =   "&Select Dates..."
      End
      Begin VB.Menu ViewDatesCmd 
         Caption         =   "&View Dates"
      End
      Begin VB.Menu zd2 
         Caption         =   "-"
      End
      Begin VB.Menu ModValuesCmd 
         Caption         =   "&Modify Variable..."
      End
      Begin VB.Menu CreateVariableCmd 
         Caption         =   "&Create New Variable..."
      End
      Begin VB.Menu CreateModelProjectionCmd 
         Caption         =   "Create Model &Projection..."
         Visible         =   0   'False
      End
      Begin VB.Menu DeleteVariableCmd 
         Caption         =   "De&lete Variable..."
      End
      Begin VB.Menu zd3 
         Caption         =   "-"
      End
      Begin VB.Menu GroupByDateCmd 
         Caption         =   "Group by &Date..."
      End
      Begin VB.Menu GroupByDOWCmd 
         Caption         =   "Group by DO&W"
      End
      Begin VB.Menu zd4 
         Caption         =   "-"
      End
      Begin VB.Menu TimeScalesCmd 
         Caption         =   "&Time Scales"
         Begin VB.Menu ToHourlyCmd 
            Caption         =   "To &Hourly"
            Enabled         =   0   'False
         End
         Begin VB.Menu ToDailyCmd 
            Caption         =   "To &Daily"
            Enabled         =   0   'False
         End
         Begin VB.Menu ToMonthlyCmd 
            Caption         =   "To &Monthly"
            Enabled         =   0   'False
         End
         Begin VB.Menu ToYearlyCmd 
            Caption         =   "To &Yearly"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu zd5 
         Caption         =   "-"
      End
      Begin VB.Menu MergeCmd 
         Caption         =   "Merge With Previous Data &File"
      End
      Begin VB.Menu AvgIntDataCmd 
         Caption         =   "&Average Intermittent Data..."
      End
   End
   Begin VB.Menu GraphBut 
      Caption         =   "&Graphics"
      Begin VB.Menu GridCmd 
         Caption         =   "G&rid"
      End
      Begin VB.Menu SetScalesCmd 
         Caption         =   "&Set Scales..."
      End
      Begin VB.Menu zg1 
         Caption         =   "-"
      End
      Begin VB.Menu PointCmd 
         Caption         =   "&Point"
         Checked         =   -1  'True
      End
      Begin VB.Menu ZoomCmd 
         Caption         =   "Zoom &In"
         Enabled         =   0   'False
      End
      Begin VB.Menu ZoomTimeOutCmd 
         Caption         =   "Zoom &Out"
         Enabled         =   0   'False
      End
      Begin VB.Menu zg2 
         Caption         =   "-"
      End
      Begin VB.Menu PredIntCmd 
         Caption         =   "&Prediction Intervals"
      End
      Begin VB.Menu zg3 
         Caption         =   "-"
      End
      Begin VB.Menu BinFitCmd 
         Caption         =   "&Bin Fit w/ BoxWiskers"
      End
      Begin VB.Menu zg4 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu HourlyProfilesRotCmd 
         Caption         =   "&Hourly Profiles Rotation..."
         Visible         =   0   'False
      End
   End
   Begin VB.Menu SavingsGrpCmd 
      Caption         =   "&Analysis"
      Begin VB.Menu CalcTotalCmd 
         Caption         =   "Calculate &Total"
      End
      Begin VB.Menu CalcSavingsCmd 
         Caption         =   "Calculate &Savings"
      End
      Begin VB.Menu CalcLEACmd 
         Caption         =   "Calculate &Lean Energy Breakdown"
      End
      Begin VB.Menu zs1 
         Caption         =   "-"
      End
      Begin VB.Menu DefineMeasurementErrorCmd 
         Caption         =   "&Define Measurement Error..."
      End
   End
   Begin VB.Menu SimulationCmd 
      Caption         =   "Si&mulation"
      Visible         =   0   'False
      Begin VB.Menu DefineLoadCmd 
         Caption         =   "Define &Load..."
      End
      Begin VB.Menu DefineSystemCmd 
         Caption         =   "Define &System..."
      End
      Begin VB.Menu DefinePlantCmd 
         Caption         =   "Define &Plant..."
      End
      Begin VB.Menu zm1 
         Caption         =   "-"
      End
      Begin VB.Menu RunSimulationCmd 
         Caption         =   "&Run Simulation"
      End
   End
   Begin VB.Menu HelpCmd 
      Caption         =   "&Help"
      Begin VB.Menu AboutCmd 
         Caption         =   "&About"
         Visible         =   0   'False
      End
      Begin VB.Menu ShowHelpCmd 
         Caption         =   "&Show Help"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AboutCmd_Click()

AboutDB.Show

End Sub

Private Sub AvgIntDataCmd_Click()
Dim i, j, sdate, edate
On Error GoTo avgintdatace

'if data are intermittent, as in utility billing data with an unequal number
'of days in each period, this routine will replace the total energy use per
'period with the average energy use per day in each period

'show time scale options
Screen.MousePointer = 0
TSOptsDB.NDOpt.Visible = False
TSOptsDB.AvailOpt.Visible = False
TSOptsDB.Label1.Caption = "Don't Change"
TSOptsDB.Caption = "Averge Intermittent Data Options"
TSOptsDB.Show MODAL
TSOptsDB.NDOpt.Visible = True
TSOptsDB.AvailOpt.Visible = True
TSOptsDB.Label1.Caption = "Total"
TSOptsDB.Caption = "Time Scale Options"
If tsoptscancel = True Then GoTo avgintdatace
Screen.MousePointer = 11

'show perdone
PerDone.Show 0
DoEvents
totregs = numrecs
increment = 10

For i = 1 To numrecs - 1
   'finds average of other records
   For j = 1 To numflds
      If j <> mofld And j <> dyfld And j <> yrfld And j <> hrfld And totvars(j) = 0 Then 'avg intermittent data
         sdate = CDate(DateOnly("d", i))
         edate = CDate(DateOnly("d", i + 1))
         If timeint = 3 Then 'daily data
            diff = edate - sdate
         ElseIf timeint = 4 Then 'hourly data
            diff = (edate - sdate) * 24 + (d(i + 1, hrfld) - d(i, hrfld))
         Else
            MsgBox "Average Intermittent Data is not appropriate for the current timescale."
            GoTo avgintdatace
         End If
                            
         If d(i + 1, j) <> ndflag And diff <> 0 Then
            d(i + 1, j) = d(i + 1, j) / diff
         End If
      End If
   Next j
   
   'sets first record to ndflag because there is no starting date for first record (ie first rec is start date for second period)
   If i = 1 Then
      For j = 1 To numflds
         If j <> mofld And j <> dyfld And j <> yrfld And j <> hrfld Then d(1, j) = ndflag
      Next j
   End If
  
   'update perdone
   regnum = regnum + 1
   Call UpdatePerDone(regnum, totregs, increment)
Next i

avgintdatace:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume avgintdataes
End If
avgintdataes:
End Sub

Private Sub BinFitBut_Click()
Dim yc As Integer, i As Integer, j As Integer
Dim numpasses As Integer, count As Integer
Dim intsize, xlow, xhigh, noprint As Integer
Dim ssm, std
On Error GoTo bfce

'call binfit dia box to get num intervals
BinFitDB.Show MODAL
If binfitcancel = True Then GoTo bfes:
Screen.MousePointer = 11

'call SetAll to fill x() and y()
Call SetAll
ReDim NumObs(numints) As Integer, ytot(numints), ybar(numints), xmid(numints)

'prints errors
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo bfes:
ElseIf G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined.", , "Error"
   GoTo bfes:
ElseIf xfld(1) = 0 Then
   MsgBox "No X variable has been defined.", , "Error"
   GoTo bfes:
End If

modtype$ = "BinFit"

'sets pass
If G1G2But.Value = True And gmodnum = 2 Then pass = 2 Else pass = 1

'refills x() and y() if appropriate
If G1G2But.Value = True Then Call SetGrp
 
If setminmax = True Then
   xmax = xmaxset
   xmin = xminset
End If

'on each pass, resets counting variables
For j = 1 To numints
   NumObs(j) = 0
   ytot(j) = 0
Next j

'determines the interval size
intsize = (xmax - xmin) / numints

'bins, counts, and totals data
For i = 1 To n
   For j = 1 To numints
      xlow = xmin + (j - 1) * intsize
      xhigh = xmin + j * intsize
      xmid(j) = (xlow + xhigh) / 2
      Select Case j
         Case 1
            If x(i, 2) >= xlow And x(i, 2) <= xhigh Then
               NumObs(j) = NumObs(j) + 1
               ytot(j) = ytot(j) + y(i, 1)
               j = 1
               GoTo bfni:
            End If
         Case Else
            If x(i, 2) > xlow And x(i, 2) <= xhigh Then
               NumObs(j) = NumObs(j) + 1
               ytot(j) = ytot(j) + y(i, 1)
               j = 1
               GoTo bfni:
            End If
         End Select
   Next j
bfni: Next i

If BinFitCmd.Checked = True Then
    'finds max numobs and dims bindata()
    Dim maxobs
    For i = 1 To numints
       If i = 1 Then
          maxobs = NumObs(i)
       ElseIf NumObs(i) > maxobs Then
          maxobs = NumObs(i)
       End If
       'initialize numobs counter
       NumObs(i) = 0
    Next i
    ReDim bindata(numints, maxobs)
    
    'loads bindata for percentile for box and whiskers
    For i = 1 To n
       For j = 1 To numints
          xlow = xmin + (j - 1) * intsize
          xhigh = xmin + j * intsize
          xmid(j) = (xlow + xhigh) / 2
          Select Case j
              Case 1
                  If x(i, 2) >= xlow And x(i, 2) <= xhigh Then
                     NumObs(j) = NumObs(j) + 1
                     bindata(j, NumObs(j)) = y(i, 1)
                     j = 1
                     GoTo bfni2:
                  End If
              Case Else
                  If x(i, 2) > xlow And x(i, 2) <= xhigh Then
                     NumObs(j) = NumObs(j) + 1
                     bindata(j, NumObs(j)) = y(i, 1)
                     j = 1
                     GoTo bfni2:
                  End If
              End Select
       Next j
bfni2: Next i

   'finds percentiles
   ReDim p10(numints), p25(numints), p50(numints), p75(numints), p90(numints)
   For i = 1 To numints
      Call Percentile(i, NumObs(i), bindata(), p10(i), p25(i), p50(i), p75(i), p90(i))
   Next i
End If 'end of b&w block

'finds mean of each interval
For i = 1 To numints
   If NumObs(i) > 0 Then
      ybar(i) = ytot(i) / NumObs(i)
   Else
      ybar(i) = ndflag
   End If
Next i

'calcs std
ssm = 0
For i = 1 To n
   For j = 1 To numints
      xlow = xmin + (j - 1) * intsize
      xhigh = xmin + j * intsize
      xmid(j) = (xlow + xhigh) / 2
      Select Case j
         Case 1
            If x(i, 2) >= xlow And x(i, 2) <= xhigh Then
               'NumObs(j) = NumObs(j) + 1
               ssm = ssm + (y(i, 1) - ybar(j)) ^ 2
               'ytot(j) = ytot(j) + y(i, 1)
               j = 1
               GoTo bfnii:
            End If
         Case Else
            If x(i, 2) > xlow And x(i, 2) <= xhigh Then
               'NumObs(j) = NumObs(j) + 1
               'ytot(j) = ytot(j) + y(i, 1)
               ssm = ssm + (y(i, 1) - ybar(j)) ^ 2
               j = 1
               GoTo bfnii:
            End If
         End Select
   Next j
bfnii: Next i
std = (ssm / (n - j)) ^ 0.5
'cvstd = std / ymean * 100

'For BinFit Savings Calc
If pass = 1 Then
   ReDim ybarbf(numints)
   xmaxbf = xmax
   xminbf = xmin
   For i = 1 To numints
      ybarbf(i) = ybar(i)
   Next i
   rmsepre = std
   premodtype$ = modtype$
   npre = n
   modelstats = " N = " + Format$(n, "####") + "   Num Intervals = " + Format$(numints, "####") + "   RMSE = " + Format$(std, "0.0")
   modeleqn = vn$(yfld) + " = mean of each interval"
End If
 
 'prints results to valuebar
 sf$ = "###0.0#"
If G1G2But.Value = False Or (G1G2But.Value = True And pass = 1) Then ValueBar.Cls
Call PrintHeaderLine(gmodnum, modtype$)
ValueBar.Print " N = " + Format$(n, "####") + "   Num Intervals = " + Format$(numints, "####") + "   RMSE = " + Format$(std, "0.0")
For i = 1 To numints
   xlow = xmin + (i - 1) * intsize
   xhigh = xmin + i * intsize
   'ValueBar.Print " (" + Format$(xmid(i), sf$) + "," + Format$(ybar(i), sf$) + ")";
   ValueBar.Print " (" + Format$(xlow, sf$) + "<=x<" + Format$(xhigh, sf$) + ", " + Format$(ybar(i), sf$) + ")";
   If i = 5 Or i = 10 Or i = 15 Or i = 20 Then ValueBar.Print
Next i
If i <> 6 And i <> 11 And i <> 16 And i <> 21 Then ValueBar.Print

'draws line of connected means
predints = 0
noprint = False
For i = 2 To numints
   x1 = xmid(i - 1)
   y1 = ybar(i - 1)
   x2 = xmid(i)
   y2 = ybar(i)
   If y1 = ndflag Then
       GoTo bni:
   ElseIf y2 = ndflag Then
       noprint = True
   Else
       If noprint = True Then
           If gx1 = 0 And gy1 = 0 Then
              x1 = xmid(1)
              y1 = ybar(1)
           Else
              x1 = gx1
              y1 = gy1
           End If
       End If
       gx1 = x2
       gy1 = y2
       Call DrawModelLine(predint, pass, x1, y1, x2, y2)
       noprint = False
   End If
bni: Next i

'b&w: draws box & whiskers markers
If BinFitCmd.Checked = True Then
   For i = 1 To numints
      'kkk
      If p90(i) <> ndflag Then
         x1 = xmid(i)
         'top vert line
         Call DrawModelLine(predint, 3, x1, p90(i), x1, p75(i))
         'box left
         Call DrawModelLine(predint, 3, x1 - 0.5, p75(i), x1 - 0.5, p25(i))
         'box right
         Call DrawModelLine(predint, 3, x1 + 0.5, p75(i), x1 + 0.5, p25(i))
         'box top
         Call DrawModelLine(predint, 3, x1 - 0.5, p75(i), x1 + 0.5, p75(i))
         'box median
         Call DrawModelLine(predint, 3, x1 - 0.5, p50(i), x1 + 0.5, p50(i))
         'box bottom
         Call DrawModelLine(predint, 3, x1 - 0.5, p25(i), x1 + 0.5, p25(i))
         'bot vert line
         Call DrawModelLine(predint, 3, x1, p25(i), x1, p10(i))
      End If
   Next i
End If

'resets gmodnum
If G1G2But.Value = True Then
   If gmodnum = 1 Then
      gmodnum = 2
   Else 'gmodnum = 2
      gmodnum = 1
   End If
End If

bfce:
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume bfes
End If
bfes:
Screen.MousePointer = 0
GraphBox.MousePointer = 0
binfitcancel = False
End Sub

Private Sub BinFitCmd_Click()
   If BinFitCmd.Checked = False Then BinFitCmd.Checked = True Else BinFitCmd.Checked = False
End Sub

Private Sub CalcLEACmd_Click()
Dim i, j, k, ybar1, ybar2
Dim badrec, ivar, ivar1, ivar2
Dim npost, pred, meas, savings
Dim PredErr, MeasErr, SavErr

On Error GoTo cscce

'init counting variables
npost = 0
pred = 0
meas = 0
savings = 0

'select model type and calc savings
Select Case premodtype$
   
   Case "Mean"
      
      MsgBox "Lean Energy breakdown is available only with 3PH-MVR or 3PC-MVR models.", , "Error"
      GoTo cses:

      For i = 1 To numrecs
          If grpd(i) = 2 And d(i, yfld) <> ndflag Then
             npost = npost + 1
             pred = pred + ymeanpre
             meas = meas + d(i, yfld)
             'store predicted value in group 2 residual field for timeseries plot of predicted and actual use
             d(i, numflds) = ymeanpre
          Else
             d(i, numflds) = -99
          End If
      Next i
      savings = pred - meas
   
   Case "Median"
      
      MsgBox "Lean Energy breakdown is available only with 3PH-MVR or 3PC-MVR models.", , "Error"
      GoTo cses:
      
      For i = 1 To numrecs
          If grpd(i) = 2 And d(i, yfld) <> ndflag Then
             npost = npost + 1
             pred = pred + ymedianpre
             meas = meas + d(i, yfld)
             'store predicted value in group 2 residual field for timeseries plot of predicted and actual use
             d(i, numflds) = ymedianpre
          Else
             d(i, numflds) = -99
          End If
      Next i
      savings = pred - meas
   
   Case "2P"
      
      MsgBox "Lean Energy breakdown is available only with 3PH-MVR or 3PC-MVR models.", , "Error"
      GoTo cses:
      
      For i = 1 To numrecs
          If grpd(i) = 2 And d(i, xfld(1)) <> ndflag And d(i, yfld) <> ndflag Then
             npost = npost + 1
             pred = pred + (betapre(1) + betapre(2) * d(i, xfld(1)))
             meas = meas + d(i, yfld)
             'store predicted value in group 2 residual field for timeseries plot of predicted and actual use
             d(i, numflds) = betapre(1) + betapre(2) * d(i, xfld(1))
          Else
             d(i, numflds) = -99
          End If
      Next i
      savings = pred - meas
   
   Case "3P"
      
      'redimensions d() with 10 more fields for holding predicted values for timeseries display
      ReDim dtemp(numrecs, numflds + 10)
      For i = 1 To numrecs
         For j = 1 To numflds
            dtemp(i, j) = d(i, j)
         Next j
      Next i
      ReDim d(numrecs, numflds + 10)
      For i = 1 To numrecs
         For j = 1 To numflds
            d(i, j) = dtemp(i, j)
         Next j
      Next i
      
      'initialize counting variables
      npost = 0
      pred = 0
      efac = 0
      ewea = 0
      eprod = 0

      For i = 1 To numrecs
         'scan xvars for ndflag
         badrec = False
         For j = 1 To numxvarspre
            If d(i, xfld(j)) = ndflag Then badrec = True
         Next j
          
         'If grpd(i) = 1 And badrec = False And d(i, yfld) <> ndflag Then
         If badrec = False Then

            'count num obs in post
            npost = npost + 1
            
            'calc total ypred, efacility, ewea-dep and eprod-dep
            If indexpre = 0 Then '3pc
               If d(i, xfld(1)) <= xcppre Then ivar = 0 Else ivar = 1
            Else '3ph
               If d(i, xfld(1)) <= xcppre Then ivar = 1 Else ivar = 0
            End If
            'pred = pred + betapre(1) + betapre(2) * (d(i, xfld(1)) - xcppre) * ivar
            pred = pred + ycppre + slopepre * (d(i, xfld(1)) - xcppre) * ivar
            efac = efac + ycppre
            ewea = ewea + slopepre * (d(i, xfld(1)) - xcppre) * ivar
            For j = 2 To numxvarspre
               pred = pred + ivcoefspre(j) * d(i, xfld(j))
               eprod = eprod + ivcoefspre(j) * d(i, xfld(j))
            Next j
                       
            'stores predicted whole plant energy use in numflds + 1 for timeseries plot
            d(i, numflds + 1) = ycppre + slopepre * (d(i, xfld(1)) - xcppre) * ivar
            For j = 2 To numxvarspre
               d(i, numflds + 1) = d(i, numflds + 1) + ivcoefspre(j) * d(i, xfld(j))
            Next j
            
            'stores facility energy use in numflds + 2 for timeseries plot
            d(i, numflds + 2) = ycppre
            
            'stores weather-dependent energy use in numflds + 3 for timeseries plot
            d(i, numflds + 3) = slopepre * (d(i, xfld(1)) - xcppre) * ivar
            
            'stores production-dependent energy use in numflds + 4 for timeseries plot
            For j = 2 To numxvarspre
               d(i, numflds + 4) = d(i, numflds + 4) + ivcoefspre(j) * d(i, xfld(j))
            Next j
            

         Else
            d(i, numflds) = -99
         End If
      Next i
      
      'calculate average values
      If npost > 0 Then
         pred = pred / npost
         efac = efac / npost
         ewea = ewea / npost
         eprod = eprod / npost
      Else
         pred = 0
         efac = 0
         ewea = 0
         eprod = 0
      End If
   
   Case "4P"
      
      MsgBox "Lean Energy breakdown is available only with 3PH-MVR or 3PC-MVR models.", , "Error"
      GoTo cses:
      
      
      For i = 1 To numrecs
          
         'scan xvars for ndflag
         badrec = False
         For j = 1 To numxvarspre
            If d(i, xfld(j)) = ndflag Then badrec = True
         Next j
                    
         If grpd(i) = 2 And badrec = False And d(i, yfld) <> ndflag Then
            npost = npost + 1
            If d(i, xfld(1)) <= xcppre Then ivar = 0 Else ivar = 1
            pred = pred + betapre(1) + betapre(2) * d(i, xfld(1)) + betapre(3) * (d(i, xfld(1)) - xcppre) * ivar
            For j = 2 To numxvarspre
               pred = pred + ivcoefspre(j) * d(i, xfld(j))
            Next j
            meas = meas + d(i, yfld)
            
            'store predicted value in group 2 residual field for timeseries plot of predicted and actual use
            d(i, numflds) = betapre(1) + betapre(2) * d(i, xfld(1)) + betapre(3) * (d(i, xfld(1)) - xcppre) * ivar
            For j = 2 To numxvarspre
               d(i, numflds) = d(i, numflds) + ivcoefspre(j) * d(i, xfld(j))
            Next j

         Else
            d(i, numflds) = -99
         
         End If
      Next i
      savings = pred - meas
            
   Case "5P"
      
      MsgBox "Lean Energy breakdown is available only with 3PH-MVR or 3PC-MVR models.", , "Error"
      GoTo cses:
      
      For i = 1 To numrecs
         
         'scan xvars for ndflag
         badrec = False
         For j = 1 To numxvarspre
            If d(i, xfld(j)) = ndflag Then badrec = True
         Next j
                    
         If grpd(i) = 2 And badrec = False And d(i, yfld) <> ndflag Then
            'calc post
            npost = npost + 1
            
            'calc pred
            If d(i, xfld(1)) <= xcp1pre Then
               ivar1 = 1
            Else
               ivar1 = 0
            End If
            If d(i, xfld(1)) <= xcp2pre Then
               ivar2 = 0
            Else
               ivar2 = 1
            End If
            pred = pred + ycppre + lspre * ivar1 * (d(i, xfld(1)) - xcp1pre) + rspre * ivar2 * (d(i, xfld(1)) - xcp2pre)
            For j = 2 To numxvarspre
               pred = pred + ivcoefspre(j) * d(i, xfld(j))
            Next j
            
            'calc meas
            meas = meas + d(i, yfld)
         
            'store predicted value in group 2 residual field for timeseries plot of predicted and actual use
            d(i, numflds) = ycppre + lspre * ivar1 * (d(i, xfld(1)) - xcp1pre) + rspre * ivar2 * (d(i, xfld(1)) - xcp2pre)
            For j = 2 To numxvarspre
               d(i, numflds) = d(i, numflds) + ivcoefspre(j) * d(i, xfld(j))
            Next j

         Else
            d(i, numflds) = -99
         
         End If
      Next i
      savings = pred - meas
   
   Case "MLR"
      
      MsgBox "Lean Energy breakdown is available only with 3PH-MVR or 3PC-MVR models.", , "Error"
      GoTo cses:
      
      For i = 1 To numrecs
         If grpd(i) = 2 And d(i, yfld) <> ndflag Then
            
            badrec = False
            For j = 1 To numxvarspre
               If d(i, xfld(j)) = ndflag Then badrec = True
            Next j
            
            If badrec = False Then
               npost = npost + 1
               ypred = betapre(1)
               For j = 1 To numxvarspre
                  ypred = ypred + betapre(j + 1) * d(i, xfld(j))
               Next j
               pred = pred + ypred
               meas = meas + d(i, yfld)
            
               'store predicted value in group 2 residual field for timeseries plot of predicted and actual use
               d(i, numflds) = betapre(1)
               For j = 1 To numxvarspre
                  d(i, numflds) = d(i, numflds) + betapre(j + 1) * d(i, xfld(j))
               Next j
               
   
            Else
               d(i, numflds) = -99
            
            
            End If
         End If
      Next i
      savings = pred - meas
   
   Case "BinFit"
      
      MsgBox "Lean Energy breakdown is available only with 3PH-MVR or 3PC-MVR models.", , "Error"
      GoTo cses:
      
      Dim xlow, xhigh, intsize
      
      'determines the interval size
      intsize = (xmaxbf - xminbf) / numints
   
      For i = 1 To numrecs
         If grpd(i) = 2 And d(i, xfld(1)) <> ndflag And d(i, yfld) <> ndflag Then
            'runs through the j intervals looking for the proper bin
            For j = 1 To numints
               xlow = xminbf + (j - 1) * intsize
               xhigh = xminbf + j * intsize
               'for first interval use >= on low bound
               If j = 1 Then
                 If d(i, xfld(1)) >= xlow And d(i, xfld(1)) <= xhigh Then
                    npost = npost + 1
                    'if ybarbf = ndflag then search for nearest non-zero ybarbf
                    If ybarbf(j) = ndflag Then
                       For k = j + 1 To numints
                          If ybarbf(k) <> ndflag Then
                             'calc pred using this ybarbf
                             pred = pred + ybarbf(k)
                             Exit For
                          End If
                       Next k
                    Else
                       pred = pred + ybarbf(j)
                    End If
                    meas = meas + d(i, yfld)
                    j = 1
                    Exit For
                 End If
               Else  'use only > on low bound
                 If d(i, xfld(1)) > xlow And d(i, xfld(1)) <= xhigh Then
                    npost = npost + 1
                    If ybarbf(j) = ndflag Then
                       ybar1 = ndflag
                       ybar2 = ndflag
                       'find nearest preceding ybarbf()
                       For k = j - 1 To 1 Step -1
                          If ybarbf(k) <> ndflag Then
                             'calc pred using this ybarbf
                             ybar1 = ybarbf(k)
                             Exit For
                          End If
                       Next k
                       'find nearest subsequent ybarbf()
                       For k = j + 1 To numints
                          If ybarbf(k) <> ndflag Then
                             'calc pred using this ybarbf
                             ybar2 = ybarbf(k)
                             Exit For
                          End If
                       Next k
                       'calc pred
                       If ybar1 <> ndflag And ybar2 <> ndflag Then
                          pred = pred + (ybar1 + ybar2) / 2
                       ElseIf ybar1 = ndflag Then
                          pred = pred + ybar2
                       ElseIf ybar2 = ndflag Then
                          pred = pred + ybar1
                       End If
                    Else
                       pred = pred + ybarbf(j)
                    End If
                    meas = meas + d(i, yfld)
                    j = 1
                    Exit For
                 End If
               End If
            Next j
         End If
      Next i
      savings = pred - meas
   
   Case Else
      
      MsgBox "Lean Energy breakdown is available only with 3PH-MVR or 3PC-MVR models.", , "Error"
      GoTo cses:

End Select

'prints results
ValueBar.Cls
ValueBar.Print " Baseline Model Stats: Type = " + premodtype$ + "  " + modelstats
ValueBar.Print " Baseline Model Eqn: " + modeleqn
ValueBar.Print " Facility = " + Format(efac, "#,##0.0") + " (" + Format(efac / pred * 100, "0.0") + " %)    Wea-dep = " + Format(ewea, "#,##0.0") + " (" + Format(ewea / pred * 100, "0.0") + " %)    Prod-dep = " + Format(eprod, "#,##0.0") + " (" + Format(eprod / pred * 100, "0.0") + " %)    Total = " + Format(pred, "#,##0.0") + " (100%)"
ValueBar.Print

'display timeseries of actual, facility, wea-dep and prod-dep use
'turns on calcleadisplay to override normal timeseries 2 operation
calcleadisplay = True
'displays actual use
Call TimeSeriesBut_Click
'displays facility energy use
AdditionalVar = 2
Call TimeSeries2But_Click
'displays weather-dependent energy use
AdditionalVar = 3
Call TimeSeries2But_Click
'displays production-dependent energy use
AdditionalVar = 4
Call TimeSeries2But_Click
'turns off calcleadisplay so timeseries2 works as usual
calcleadisplay = False
AdditionalVar = 0

cscce:
Screen.MousePointer = 0
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error"
   Resume cses
End If
cses:

End Sub

Private Sub CalcSavingsCmd_Click()
Dim i, j, k, ybar1, ybar2
Dim badrec, ivar, ivar1, ivar2
Dim npost, pred, meas, savings
Dim PredErr, MeasErr, SavErr

'On Error GoTo cscce

'init counting variables
npost = 0
pred = 0
meas = 0
savings = 0

'select model type and calc savings
Select Case premodtype$
   
   Case "Mean"
      For i = 1 To numrecs
          If grpd(i) = 2 And d(i, yfld) <> ndflag Then
             npost = npost + 1
             pred = pred + ymeanpre
             meas = meas + d(i, yfld)
             'store predicted value in group 2 residual field for timeseries plot of predicted and actual use
             d(i, numflds) = ymeanpre
          Else
             d(i, numflds) = -99
          End If
      Next i
      savings = pred - meas
   
   Case "Median"
      For i = 1 To numrecs
          If grpd(i) = 2 And d(i, yfld) <> ndflag Then
             npost = npost + 1
             pred = pred + ymedianpre
             meas = meas + d(i, yfld)
             'store predicted value in group 2 residual field for timeseries plot of predicted and actual use
             d(i, numflds) = ymedianpre
          Else
             d(i, numflds) = -99
          End If
      Next i
      savings = pred - meas
   
   Case "2P"
      For i = 1 To numrecs
          If grpd(i) = 2 And d(i, xfld(1)) <> ndflag And d(i, yfld) <> ndflag Then
             npost = npost + 1
             pred = pred + (betapre(1) + betapre(2) * d(i, xfld(1)))
             meas = meas + d(i, yfld)
             'store predicted value in group 2 residual field for timeseries plot of predicted and actual use
             d(i, numflds) = betapre(1) + betapre(2) * d(i, xfld(1))
          Else
             d(i, numflds) = -99
          End If
      Next i
      savings = pred - meas
   
   Case "3P"
      
      For i = 1 To numrecs
         'scan xvars for ndflag
         badrec = False
         For j = 1 To numxvarspre
            If d(i, xfld(j)) = ndflag Then badrec = True
         Next j
          
         If grpd(i) = 2 And badrec = False And d(i, yfld) <> ndflag Then
            'count num obs in post
            npost = npost + 1
            
            'calc ypred
            If indexpre = 0 Then '3pc
               If d(i, xfld(1)) <= xcppre Then ivar = 0 Else ivar = 1
            Else '3ph
               If d(i, xfld(1)) <= xcppre Then ivar = 1 Else ivar = 0
            End If
            'pred = pred + betapre(1) + betapre(2) * (d(i, xfld(1)) - xcppre) * ivar
            pred = pred + ycppre + slopepre * (d(i, xfld(1)) - xcppre) * ivar
            For j = 2 To numxvarspre
               pred = pred + ivcoefspre(j) * d(i, xfld(j))
            Next j
                       
            'calc ypost
            meas = meas + d(i, yfld)
            
            'store predicted value in group 2 residual field for timeseries plot of predicted and actual use
            d(i, numflds) = ycppre + slopepre * (d(i, xfld(1)) - xcppre) * ivar
            For j = 2 To numxvarspre
               d(i, numflds) = d(i, numflds) + ivcoefspre(j) * d(i, xfld(j))
            Next j

         Else
            d(i, numflds) = -99
         
         End If
      Next i
      savings = pred - meas
   
   Case "4P"
      For i = 1 To numrecs
          
         'scan xvars for ndflag
         badrec = False
         For j = 1 To numxvarspre
            If d(i, xfld(j)) = ndflag Then badrec = True
         Next j
                    
         If grpd(i) = 2 And badrec = False And d(i, yfld) <> ndflag Then
            npost = npost + 1
            If d(i, xfld(1)) <= xcppre Then ivar = 0 Else ivar = 1
            pred = pred + betapre(1) + betapre(2) * d(i, xfld(1)) + betapre(3) * (d(i, xfld(1)) - xcppre) * ivar
            For j = 2 To numxvarspre
               pred = pred + ivcoefspre(j) * d(i, xfld(j))
            Next j
            meas = meas + d(i, yfld)
            
            'store predicted value in group 2 residual field for timeseries plot of predicted and actual use
            d(i, numflds) = betapre(1) + betapre(2) * d(i, xfld(1)) + betapre(3) * (d(i, xfld(1)) - xcppre) * ivar
            For j = 2 To numxvarspre
               d(i, numflds) = d(i, numflds) + ivcoefspre(j) * d(i, xfld(j))
            Next j

         Else
            d(i, numflds) = -99
         
         End If
      Next i
      savings = pred - meas
            
   Case "5P"
      For i = 1 To numrecs
         
         'scan xvars for ndflag
         badrec = False
         For j = 1 To numxvarspre
            If d(i, xfld(j)) = ndflag Then badrec = True
         Next j
                    
         If grpd(i) = 2 And badrec = False And d(i, yfld) <> ndflag Then
            'calc post
            npost = npost + 1
            
            'calc pred
            If d(i, xfld(1)) <= xcp1pre Then
               ivar1 = 1
            Else
               ivar1 = 0
            End If
            If d(i, xfld(1)) <= xcp2pre Then
               ivar2 = 0
            Else
               ivar2 = 1
            End If
            pred = pred + ycppre + lspre * ivar1 * (d(i, xfld(1)) - xcp1pre) + rspre * ivar2 * (d(i, xfld(1)) - xcp2pre)
            For j = 2 To numxvarspre
               pred = pred + ivcoefspre(j) * d(i, xfld(j))
            Next j
            
            'calc meas
            meas = meas + d(i, yfld)
         
            'store predicted value in group 2 residual field for timeseries plot of predicted and actual use
            d(i, numflds) = ycppre + lspre * ivar1 * (d(i, xfld(1)) - xcp1pre) + rspre * ivar2 * (d(i, xfld(1)) - xcp2pre)
            For j = 2 To numxvarspre
               d(i, numflds) = d(i, numflds) + ivcoefspre(j) * d(i, xfld(j))
            Next j

         Else
            d(i, numflds) = -99
         
         End If
      Next i
      savings = pred - meas
   
   Case "MLR"
      For i = 1 To numrecs
         If grpd(i) = 2 And d(i, yfld) <> ndflag Then
            
            badrec = False
            For j = 1 To numxvarspre
               If d(i, xfld(j)) = ndflag Then badrec = True
            Next j
            
            If badrec = False Then
               npost = npost + 1
               ypred = betapre(1)
               For j = 1 To numxvarspre
                  ypred = ypred + betapre(j + 1) * d(i, xfld(j))
               Next j
               pred = pred + ypred
               meas = meas + d(i, yfld)
            
               'store predicted value in group 2 residual field for timeseries plot of predicted and actual use
               d(i, numflds) = betapre(1)
               For j = 1 To numxvarspre
                  d(i, numflds) = d(i, numflds) + betapre(j + 1) * d(i, xfld(j))
               Next j
               
   
            Else
               d(i, numflds) = -99
            
            
            End If
         End If
      Next i
      savings = pred - meas
   
   Case "BinFit"
      Dim xlow, xhigh, intsize
      
      'determines the interval size
      intsize = (xmaxbf - xminbf) / numints
   
      For i = 1 To numrecs
         If grpd(i) = 2 And d(i, xfld(1)) <> ndflag And d(i, yfld) <> ndflag Then
            'runs through the j intervals looking for the proper bin
            For j = 1 To numints
               xlow = xminbf + (j - 1) * intsize
               xhigh = xminbf + j * intsize
               'for first interval use >= on low bound
               If j = 1 Then
                 If d(i, xfld(1)) >= xlow And d(i, xfld(1)) <= xhigh Then
                    npost = npost + 1
                    'if ybarbf = ndflag then search for nearest non-zero ybarbf
                    If ybarbf(j) = ndflag Then
                       For k = j + 1 To numints
                          If ybarbf(k) <> ndflag Then
                             'calc pred using this ybarbf
                             pred = pred + ybarbf(k)
                             Exit For
                          End If
                       Next k
                    Else
                       pred = pred + ybarbf(j)
                    End If
                    meas = meas + d(i, yfld)
                    j = 1
                    Exit For
                 End If
               Else  'use only > on low bound
                 If d(i, xfld(1)) > xlow And d(i, xfld(1)) <= xhigh Then
                    npost = npost + 1
                    If ybarbf(j) = ndflag Then
                       ybar1 = ndflag
                       ybar2 = ndflag
                       'find nearest preceding ybarbf()
                       For k = j - 1 To 1 Step -1
                          If ybarbf(k) <> ndflag Then
                             'calc pred using this ybarbf
                             ybar1 = ybarbf(k)
                             Exit For
                          End If
                       Next k
                       'find nearest subsequent ybarbf()
                       For k = j + 1 To numints
                          If ybarbf(k) <> ndflag Then
                             'calc pred using this ybarbf
                             ybar2 = ybarbf(k)
                             Exit For
                          End If
                       Next k
                       'calc pred
                       If ybar1 <> ndflag And ybar2 <> ndflag Then
                          pred = pred + (ybar1 + ybar2) / 2
                       ElseIf ybar1 = ndflag Then
                          pred = pred + ybar2
                       ElseIf ybar2 = ndflag Then
                          pred = pred + ybar1
                       End If
                    Else
                       pred = pred + ybarbf(j)
                    End If
                    meas = meas + d(i, yfld)
                    j = 1
                    Exit For
                 End If
               End If
            Next j
         End If
      Next i
      savings = pred - meas
   
   Case Else
      MsgBox "No model of pre retrofit data has been selected.", , "Error"
      GoTo cses:
End Select


If npre > 0 And savings <> 0 And npost > 0 Then

   'calc uncertainty
   PredErr = 1.96 * rmsepre * ((1 + 2 / npre) ^ 0.5) * (npost ^ 0.5)
   MeasErr = meas * (PerMeasErr / 100)
   SavErr = (PredErr ^ 2 + MeasErr ^ 2) ^ 0.5
   
   'prints results
   ValueBar.Cls
   'ValueBar.Print " " + UCase$(filename$)
   ValueBar.Print " Baseline Model Stats: Type = " + premodtype$ + "  " + modelstats
   ValueBar.Print " Baseline Model Eqn: " + modeleqn
   'ValueBar.Print " " + vn(yfld); " savings using a " + premodtype$ + " pre-retrofit model, during " + Format(npost, "#,###") + " post observations:"
   ValueBar.Print " Observations During Post-retrofit Period: " + Format$(npost, "#,###")
   'ValueBar.Print " Predicted Pre Use   = " + Format(pred, "#,##0.0") + "  ±" + Format(PredErr, "#,##0.0") + "  ±" + Format(PredErr / pred * 100, "0.0") + "%"
   'ValueBar.Print " Measured Post Use   = " + Format(meas, "#,##0.0") + "  ±" + Format(MeasErr, "#,##0.0") + "  ±" + Format(MeasErr / meas * 100, "0.0") + "%"
   ValueBar.Print " Adjusted Baseline Use During Post-retrofit Period   = " + Format(pred, "#,##0.0") + "  ±" + Format(PredErr, "#,##0.0") + "  ±" + Format(PredErr / pred * 100, "0.0") + "%"
   ValueBar.Print " Actual Use During Post-retrofit Period   = " + Format(meas, "#,##0.0") + "  ±" + Format(MeasErr, "#,##0.0") + "  ±" + Format(MeasErr / meas * 100, "0.0") + "%"
   ValueBar.Print " Total Savings During Post-retrofit Period = " + Format(savings, "#,##0.0") + "  ±" + Format(SavErr, "#,##0.0") + "  ±" + Format(SavErr / savings * 100, "0.0") + "%"
   ValueBar.Print " Average Savings Per Observation During Post-retrofit Period = " + Format(savings / npost, "#,##0.0") + "  ±" + Format(SavErr / npost, "#,##0.0") + "  ±" + Format(SavErr / savings * 100, "0.0") + "%"

Else

   MsgBox "Unable to calculate savings.", , "Error"

End If

'display timeseries of acutal and predicted use
Call TimeSeriesBut_Click
calcsavedisplay = True
Call TimeSeries2But_Click
calcsavedisplay = False

cscce:
Screen.MousePointer = 0
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error"
   Resume cses
End If
cses:
End Sub

Private Sub CalcSavingsPECmd_Click()
Dim i, j, k, Msg$, no As Integer, numpasses As Integer
Dim pnum, pdenom, prevres, P, ytot, ymean
Dim sse, mse, rmse, cvrmse, savings, postuse
Dim det, r2, adjr2, trace, rss, sig2, var
Dim c11, c22, c23, c12, wyz, t
Dim ModelErr, MeasErr, PredErr, preduse 'global measerr
Dim xr As Integer, xc As Integer, yr As Integer, yc As Integer
Dim xtr As Integer, xtc As Integer, xtxr As Integer, xtxc As Integer
Dim qr As Integer, qc As Integer, betar As Integer, betac As Integer
Dim yhatpstr  As Integer, yhatpstc As Integer
Dim ar As Integer, ac As Integer, xbqr As Integer, xbqc As Integer
Dim xrpst As Integer, xbqtr As Integer, xbqtc As Integer
Dim mtr As Integer, mtc As Integer, mvr As Integer, mvc As Integer
Dim qtr As Integer, qtc As Integer, xbqvr  As Integer, xbqvc As Integer
Dim skiprec, ncount
Dim ybar1, ybar2
Dim indvar1, indvar2
On Error GoTo scpece

'checks for pre and post
'If g1name$ <> "Pre" Or g2name$ <> "Post" Then
'   Msg$ = "Savings calculation is available only when groups have been named 'Pre' and 'Post'."
'   MsgBox Msg$, , "Error"
'   GoTo scpece
'End If

'clears screen
Screen.MousePointer = 11


'inits vars for 'old' calc save routines
pred = 0
meas = 0
saved = 0
numobspost = 0

If modtype$ = "Mean" Then

   'fills x() and y(), and calcs n for the chosen variable
   Call SetAll
   
   'fills xpst(), ypst()
   pass = 2
   Call SetGrp
   xrpst = n
   xcpst = 2
   ReDim xpst(xrpst, xcpst), ypst(xrpst, 1)
   For i = 1 To xrpst
      ypst(i, 1) = y(i, 1)
      For j = 1 To xcpst
         xpst(i, j) = x(i, j)
      Next j
   Next i
   
   'fills x(), y() with pre data
   pass = 1
   Call SetGrp
   xr = n
   yr = n
   yc = 1
   xc = 1

   'checks for pre and post data before continuing
   If xr = 0 Or xrpst = 0 Then
      Msg$ = "No data available for savings calculation."
      MsgBox Msg$, , "Error"
      GoTo scpece
   End If
   
   'dim arrays
   ReDim xt(xc, xr), xtx(xc, xc), xtxi(xc, xc), q(xc, xr)
   ReDim xtotal(xc), yhatpst(xrpst, 1)
   ReDim xtotalp(xc), xmeanp(xc), xb(1, xc)
   ReDim v(xr, 1)
   ReDim A(1, xr), m(1, xr), xn(1, xc)
   ReDim mt(xr, 1), mv(1, 1)
   ReDim qt(xr, xc), xdum(xr, 1), qdum(xr, 1)
   
   'calcs ymean
   For i = 1 To xr
      ytot = ytot + y(i, 1)
   Next i
   ymean = ytot / xr
      
   'calcs params
   ReDim mres(xr)
   sse = 0: pnum = 0: pdenom = 0
   For i = 1 To xr
      mres(i) = y(i, 1) - ymean
      sse = sse + mres(i) ^ 2
      If i = 1 Then
         prevres = mres(1)
      Else
         pnum = pnum + mres(i) * prevres
         pdenom = pdenom + mres(i) ^ 2
         prevres = mres(i)
      End If
   Next i
   mse = sse / (n - xc)
   rmse = mse ^ 0.5
   cvrmse = rmse / ymean * 100
   P = pnum / pdenom
   
   'calcs savings
   savings = 0: postuse = 0
   For i = 1 To xrpst
      postuse = postuse + ypst(i, 1)
      savings = savings + ymean - ypst(i, 1)
   Next i
      
ElseIf modtype$ = "Median" Then
       
       For i = 1 To numrecs
           If grpd(i) = 2 And d(i, yfld) <> ndflag Then
              numobspost = numobspost + 1
              pred = pred + ymedianpre
              meas = meas + d(i, yfld)
           End If
       Next i
       saved = pred - meas


ElseIf modtype = "2P" Then

   'fills x() and y(), and calcs n for the chosen variable
   Call SetAll
   
   'fills xpst(), ypst()
   pass = 2
   Call SetGrp
   xrpst = n
   xcpst = 2
   ReDim xpst(xrpst, xcpst), ypst(xrpst, 1)
   For i = 1 To xrpst
      ypst(i, 1) = y(i, 1)
      For j = 1 To xcpst
         xpst(i, j) = x(i, j)
      Next j
   Next i
   
   'fills x(), y() with pre data
   pass = 1
   Call SetGrp
   xr = n
   yr = n
   yc = 1
   xc = 2
   numxvars = 1
   
   'checks for pre and post data before continuing
   If xr = 0 Or xrpst = 0 Then
      Msg$ = "No data available for savings calculation."
      MsgBox Msg$, , "Error"
      GoTo scpece
   End If
   
   'dim arrays
   ReDim xt(xc, xr), xtx(xc, xc), xtxi(xc, xc), q(xc, xr)
   ReDim xtotal(xc), yhatpst(xrpst, 1)
   ReDim xtotalp(xc), xmeanp(xc), xb(1, xc)
   ReDim v(xr, 1)
   ReDim A(1, xr), m(1, xr), xn(1, xc)
   ReDim mt(xr, 1), mv(1, 1)
   ReDim qt(xr, xc), xdum(xr, 1), qdum(xr, 1)

   'finds Regression parameters
   Call Trans(x(), xr, xc, xt(), xtr, xtc)
   Call Mult(xt(), xtr, xtc, x(), xr, xc, xtx(), xtxr, xtxc)
   Call Invert(xtx(), xtxr, xtxi(), det, sec())
   Call Mult(xtxi(), xtxr, xtxr, xt(), xtr, xtc, q(), qr, qc)
   Call Mult(q(), qr, qc, y(), yr, yc, beta(), betar, betac)
   Call Inf(x(), xr, xc, y(), yr, yc, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), P)

   'calcs savings
   Call Mult(xpst(), xrpst, xc, beta(), betar, betac, yhatpst(), yhatpstr, yhatpstc)
   savings = 0: postuse = 0
   For i = 1 To xrpst
      postuse = postuse + ypst(i, 1)
      savings = savings + yhatpst(i, 1) - ypst(i, 1)
   Next i

ElseIf modtype$ = "3P" Then
       
       For i = 1 To numrecs
           If grpd(i) = 2 And d(i, xfld(1)) <> ndflag And d(i, yfld) <> ndflag Then
              'post
              numobspost = numobspost + 1
              If hcindex = 0 Then
                 If d(i, xfld(1)) <= xcppre Then ivar = 0 Else ivar = 1
              Else
                 If d(i, xfld(1)) <= xcppre Then ivar = 1 Else ivar = 0
              End If
              pred = pred + betapre(1) + betapre(2) * (d(i, xfld(1)) - xcppre) * ivar
              meas = meas + d(i, yfld)
           End If
       Next i
       saved = pred - meas

ElseIf modtype$ = "4P" Then
       
   For i = 1 To numrecs
       If grpd(i) = 2 And d(i, xfld(1)) <> ndflag And d(i, yfld) <> ndflag Then
          'post
          numobspost = numobspost + 1
          If d(i, xfld(1)) <= xcppre Then ivar = 0 Else ivar = 1
          pred = pred + betapre(1) + betapre(2) * d(i, xfld(1)) + betapre(3) * (d(i, xfld(1)) - xcppre) * ivar
          meas = meas + d(i, yfld)
       End If
   Next i
   saved = pred - meas


ElseIf modtype$ = "5P" Then

   For i = 1 To numrecs
      If grpd(i) = 2 And d(i, xfld(1)) <> ndflag And d(i, yfld) <> ndflag Then
         'post
         numobspost = numobspost + 1
         If d(i, xfld(1)) <= cp1pre Then
           indvar1 = 1
         Else
            indvar1 = 0
         End If
                 
        If d(i, xfld(1)) <= cp2pre Then
           indvar2 = 0
        Else
           indvar2 = 1
        End If
   
         pred = pred + ycppre + lspre * indvar1 * (d(i, xfld(1)) - cp1pre) + rspre * indvar2 * (d(i, xfld(1)) - cp2pre)
         meas = meas + d(i, yfld)
      End If
   Next i
   saved = pred - meas

ElseIf modtype$ = "MLR" Then

   'sets numvars from global variable
   numxvars = mlrxvars
   xc = mlrxvars + 1
   xcpst = mlrxvars + 1

   'redimensions X(), Y(), and IndexNum()
   ReDim x(numrecs, xc), y(numrecs, 1)
   ReDim xpst(numrecs, xcpst), ypst(numrecs, 1)

   'fills x() and xpst()
   xr = 0
   xrpst = 0
   For j = 1 To numrecs
      
      If grpd(j) = 1 And d(j, yfld) <> ndflag Then
         'pre....

         'checks for no data flag
         skiprec = False
         For k = 1 To numxvars
            If d(j, xfld(k)) = ndflag Then skiprec = True
         Next k
         
         If skiprec = False Then 'fills x() and y()
            xr = xr + 1
            x(xr, 1) = 1
            For k = 1 To numxvars
               x(xr, k + 1) = d(j, xfld(k))
            Next k
            y(xr, 1) = d(j, yfld)
         End If

      
      ElseIf grpd(j) = 2 And d(j, yfld) <> ndflag Then
         'post....

         'checks for no data flag
         skiprec = False
         For k = 1 To numxvars
            If d(j, xfld(k)) = ndflag Then skiprec = True
         Next k
         
         If skiprec = False Then 'fills xpst() and ypst()
            xrpst = xrpst + 1
            xpst(xrpst, 1) = 1
            For k = 1 To numxvars
               xpst(xrpst, k + 1) = d(j, xfld(k))
            Next k
            ypst(xrpst, 1) = d(j, yfld)
         End If
      
      End If
   Next j
   yr = xr
   yc = 1

   'prints error if n = 0
   If xr = 0 Then
      MsgBox "No data is available to model.", , "Error"
      GoTo scpece
   End If
   
   'dim arrays
   ReDim xt(xc, xr), xtx(xc, xc), xtxi(xc, xc), q(xc, xr)
   ReDim xtotal(xc), yhatpst(xrpst, 1)
   ReDim xtotalp(xc), xmeanp(xc), xb(1, xc)
   ReDim v(xr, 1)
   ReDim A(1, xr), m(1, xr), xn(1, xc)
   ReDim mt(xr, 1), mv(1, 1)
   ReDim qt(xr, xc), xdum(xr, 1), qdum(xr, 1)

   'finds Regression parameters
   Call Trans(x(), xr, xc, xt(), xtr, xtc)
   Call Mult(xt(), xtr, xtc, x(), xr, xc, xtx(), xtxr, xtxc)
   Call Invert(xtx(), xtxr, xtxi(), det, sec())
   Call Mult(xtxi(), xtxr, xtxr, xt(), xtr, xtc, q(), qr, qc)
   Call Mult(q(), qr, qc, y(), yr, yc, beta(), betar, betac)
   Call Inf(x(), xr, xc, y(), yr, yc, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), P)

   'calcs savings
   Call Mult(xpst(), xrpst, xc, beta(), betar, betac, yhatpst(), yhatpstr, yhatpstc)
   savings = 0: postuse = 0
   For i = 1 To xrpst
      postuse = postuse + ypst(i, 1)
      savings = savings + yhatpst(i, 1) - ypst(i, 1)
   Next i


ElseIf modtype$ = "BinFit" Then

       If setminmax = False Then
          MsgBox "User must define Xmin and Xmax for Bin Fit model before savings can be calculated.", , "Error"
          GoTo scpece
       End If

       Dim xlow, xhigh, intsize
       'determines the interval size
       intsize = (xmaxbf - xminbf) / numints

       For i = 1 To numrecs
          If grpd(i) = 2 And d(i, xfld(1)) <> ndflag And d(i, yfld) <> ndflag Then
             'runs through the j intervals looking for the proper bin
             For j = 1 To numints
                xlow = xminbf + (j - 1) * intsize
                xhigh = xminbf + j * intsize
                'for first interval use >= on low bound
                If j = 1 Then
                  If d(i, xfld(1)) >= xlow And d(i, xfld(1)) <= xhigh Then
                     numobspost = numobspost + 1
                     'if ybarbf = ndflag then search for nearest non-zero ybarbf
                     If ybarbf(j) = ndflag Then
                        For k = j + 1 To numints
                           If ybarbf(k) <> ndflag Then
                              'calc pred using this ybarbf
                              pred = pred + ybarbf(k)
                              Exit For
                           End If
                        Next k
                     Else
                        pred = pred + ybarbf(j)
                     End If
                     meas = meas + d(i, yfld)
                     j = 1
                     Exit For
                  End If
                Else  'use only > on low bound
                  If d(i, xfld(1)) > xlow And d(i, xfld(1)) <= xhigh Then
                     numobspost = numobspost + 1
                     If ybarbf(j) = ndflag Then
                        ybar1 = ndflag
                        ybar2 = ndflag
                        'find nearest preceding ybarbf()
                        For k = j - 1 To 1 Step -1
                           If ybarbf(k) <> ndflag Then
                              'calc pred using this ybarbf
                              ybar1 = ybarbf(k)
                              Exit For
                           End If
                        Next k
                        'find nearest subsequent ybarbf()
                        For k = j + 1 To numints
                           If ybarbf(k) <> ndflag Then
                              'calc pred using this ybarbf
                              ybar2 = ybarbf(k)
                              Exit For
                           End If
                        Next k
                        'calc pred
                        If ybar1 <> ndflag And ybar2 <> ndflag Then
                           pred = pred + (ybar1 + ybar2) / 2
                        ElseIf ybar1 = ndflag Then
                           pred = pred + ybar2
                        ElseIf ybar2 = ndflag Then
                           pred = pred + ybar1
                        End If
                     Else
                        pred = pred + ybarbf(j)
                     End If
                     meas = meas + d(i, yfld)
                     j = 1
                     Exit For
                  End If
                End If
             Next j
          End If
       Next i
       saved = pred - meas

Else
   
   Msg$ = "No model of pre retrofit data has been selected."
   MsgBox Msg$, , "Error"
   GoTo scpece

End If


If modtype$ = "Mean" Or modtype$ = "2P" Or modtype$ = "MLR" Then
   'calculate uncertainty using Ruch et al. Hybrid Method
   
   'Step 0
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 0"
   For i = 2 To xc
      For j = 1 To xr
         xtotal(i) = xtotal(i) + x(j, i)
      Next j
      xmean(i) = xtotal(i) / xr
   Next i
   For i = 2 To xc
      For j = 1 To xr
         x(j, i) = x(j, i) - xmean(i)
      Next j
   Next i
   For i = 2 To xc
      For j = 1 To xrpst
         xpst(j, i) = xpst(j, i) - xmean(i)
      Next j
   Next i
   
   'Step 1
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 1"
   Call Trans(x(), xr, xc, xt(), xtr, xtc)
   Call Mult(xt(), xtr, xtc, x(), xr, xc, xtx(), xtxr, xtxc)
   Call Invert(xtx(), xtxr, xtxi(), det, sec())
   Call Mult(xtxi(), xtxr, xtxr, xt(), xtr, xtc, q(), qr, qc)
   
   'Step 2
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 2"
   For i = 2 To xc
      For j = 1 To xrpst
         xtotalp(i) = xtotalp(i) + xpst(j, i)
      Next j
      xmeanp(i) = xtotalp(i) / xrpst
   Next i
   xb(1, 1) = 1
   For i = 2 To xc
     xb(1, i) = xmeanp(i)
   Next i
   
   'Step 3
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 3"
   k = P * (1 - P ^ xrpst) / (1 - P)
   For i = 1 To xr
      v(i, 1) = k * P ^ (xr - i)
   Next i
   
   'Step 4
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 4"
   xn(1, 1) = 1
   For i = 2 To xc
      xn(1, i) = x(xr, i)
   Next i
   Call Mult(xn(), 1, xc, q(), qr, qc, A(), ar, ac)
   A(1, xr) = A(1, xr) - 1
   For i = 1 To ac
      m(1, i) = -k * A(1, i)
   Next i
   
   'Step 5
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 5"
   'define subprogram WYZCalc()
   
   'Step 6
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 6"
   ReDim xbq(1, qc), xbqt(qc, 1)
   Call Mult(xb(), 1, xc, q(), qr, qc, xbq(), xbqr, xbqc)
   Call Trans(xbq(), xbqr, xbqc, xbqt(), xbqtr, xbqtc)
   Call WYZCalc(xbqt(), xbqt(), P, xr, wyz)
   c11 = wyz * xrpst ^ 2
   
   'Step 7
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 7"
   Call Trans(m(), 1, xr, mt(), mtr, mtc)
   Call WYZCalc(mt(), mt(), P, xr, wyz)
   Call Mult(m(), 1, xr, v(), xr, 1, mv(), mvr, mvc)
   c23 = -2 * mv(1, 1) + wyz
   
   'Step 8
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 8"
   ReDim xbqv(xbqr, 1)
   Call WYZCalc(xbqt(), mt(), P, xr, wyz)
   Call Mult(xbq(), xbqr, xbqc, v(), xr, 1, xbqv(), xbqvr, xbqvc)
   c12 = 2 * xrpst * (wyz - xbqv(1, 1))
   
   'Step 9
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 9"
   c22 = (2 * P ^ (xrpst + 1) - xrpst * (P ^ 2 - 1) - 2 * P) / (1 - P) ^ 2
   
   'Step 10
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 10"
   Call Trans(q(), qr, qc, qt(), qtr, qtc)
   For i = 1 To xc
      For j = 1 To xr
         xdum(j, 1) = x(j, i)
         qdum(j, 1) = qt(j, i)
      Next j
      Call WYZCalc(xdum(), qdum(), P, xr, wyz)
      trace = trace + wyz
   Next i
   rss = (rmse ^ 2) * (xr - xc)
   sig2 = rss / (xr - trace)
   var = sig2 * (c11 + c12 + c22 + c23)
   t = 1.96 + 2.7 / (xr - 2)
   ModelErr = t * var ^ 0.5
   MeasErr = postuse * posterr / 100
   PredErr = (ModelErr ^ 2 + MeasErr ^ 2) ^ 0.5
   preduse = savings + postuse
End If

'print results
ValueBar.Cls
If modtype$ = "2P" Or modtype = "MLR" Then
   
   ValueBar.Print "Savings using a " + modtype$ + " model, during " + Format(xrpst, "#,###") + " post observations:"
   ValueBar.Print "Predicted Pre Use   = " + Format(preduse, "#,###,##0.0") + "  ±" + Format(ModelErr, "###,##0.0") + "  ±" + Format(ModelErr / preduse * 100, "#####0.0") + "%"
   ValueBar.Print "Measured Post Use   = " + Format(postuse, "#,###,##0.0") + "  ±" + Format(MeasErr, "###,##0.0") + "  ±" + Format(MeasErr / postuse * 100, "#####0.0") + "%"
   ValueBar.Print "Savings During Post = " + Format(savings, "#,###,##0.0") + "  ±" + Format(PredErr, "###,##0.0") + "  ±" + Format(PredErr / savings * 100, "#####0.0") + "%"
   ValueBar.Print "Avg Savings Per Obs = " + Format(savings / xrpst, "#,###,##0.0") + "  ±" + Format(PredErr / xrpst, "###,##0.0") + "  ±" + Format(PredErr / savings * 100, "#####0.0") + "%"

ElseIf modtype$ = "Mean" Then
   
   ValueBar.Print "Savings using a " + modtype$ + " model, during " + Format(xrpst, "#,###") + " post observations:"
   ValueBar.Print "Predicted Pre Use   = " + Format(preduse, "#,###,##0.0") + "  ±" + Format(ModelErr, "###,##0.0") + "  ±" + Format(ModelErr / preduse * 100, "#####0.0") + "%"
   ValueBar.Print "Measured Post Use   = " + Format(postuse, "#,###,##0.0") + "  ±" + Format(MeasErr, "###,##0.0") + "  ±" + Format(MeasErr / postuse * 100, "#####0.0") + "%"
   ValueBar.Print "Savings During Post = " + Format(savings, "#,###,##0.0") + "  ±" + Format(PredErr, "###,##0.0") + "  ±" + Format(PredErr / savings * 100, "#####0.0") + "%"
   ValueBar.Print "Avg Savings Per Obs = " + Format(savings / xrpst, "#,###,##0.0") + "  ±" + Format(PredErr / xrpst, "###,##0.0") + "  ±" + Format(PredErr / savings * 100, "#####0.0") + "%"

Else
   
   ValueBar.Print "Savings using a " + modtype$ + " model, during " + Format(numobspost, "#,###") + " post observations:"
   ValueBar.Print "Predicted Pre Use   = " + Format(pred, "#,###,##0.0")
   ValueBar.Print "Measured Post Use   = " + Format(meas, "#,###,##0.0")
   ValueBar.Print "Savings During Post = " + Format(saved, "#,###,##0.0")
   ValueBar.Print "Avg Savings Per Obs = " + Format(saved / numobspost, "#,###,##0.0")

End If

scpece:
Screen.MousePointer = 0
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error"
   Resume scpees
End If
scpees:
End Sub

Private Sub CalcTotalCmd_Click()
Dim i As Integer, count As Integer, numpasses As Integer
Dim ytot
Dim no As Integer
On Error GoTo tbce

Screen.MousePointer = 11
Call SetAll
ReDim IndexNumO(n) As Integer, WDWEO(n) As Integer, XO(n), YO(n)

'prints error if n = 0 or wdwe but no groupfld
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo tbce:
ElseIf G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined.", , "Error"
   GoTo tbce:
End If

modtype$ = "T"

'set numpasses
If G1G2But.Value = False Then
   numpasses = 1
Else
   numpasses = 2
End If

For pass = 1 To numpasses
   
   'refills x() and y() if appropriate
   If numpasses = 2 Then Call SetGrp
    
    'calcs y
    ytot = 0
    For i = 1 To n
       ytot = ytot + y(i, 1)
    Next i

    'prints results
    sf$ = "###0.0000"
    If numpasses = 1 Then ValueBar.Cls
    If numpasses = 2 And pass = 1 Then ValueBar.Cls
    ValueBar.Print " " + UCase$(filename$);
    If numpasses = 1 Then
       ValueBar.Print "   Single Group " + vn$(yfld) + " Model"
    ElseIf numpasses = 2 And pass = 2 Then
       ValueBar.Print "   " + g2name$ + "  " + vn$(yfld) + "  Model"
    ElseIf numpasses = 2 And pass = 1 Then
       ValueBar.Print "   " + g1name$ + "  " + vn$(yfld) + "  Model"
    End If
    ValueBar.Print " N = " + Format$(n, "####") + "   Total = " + Format$(ytot, "###0.00")
    'prints results to log file
    If SaveLogCmd.Checked = True Then
        Print #3, " " + UCase$(filename$);
        If numpasses = 1 Then
           Print #3, "   Single Group " + vn$(yfld) + " Model"
        ElseIf numpasses = 2 And pass = 2 Then
           Print #3, "   " + g2name$ + "  " + vn$(yfld) + "  Model"
        ElseIf numpasses = 2 And pass = 1 Then
           Print #3, "   " + g1name$ + "  " + vn$(yfld) + "  Model"
        End If
        Print #3, " N = " + Format$(n, "####") + "   Total = " + Format$(ytot, "###0.00")
        Print #3,: Print #3,
    End If
Next pass

tbce:
Screen.MousePointer = 0
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume tbes
End If
tbes:
End Sub

Private Sub ChangeTimeScalesBut_Click()

TimeScaleDB.Show MODAL

End Sub

Private Sub ClearGraphBut_Click()
   GraphBox.Cls
End Sub

Private Sub ClearValueBut_Click()
   ValueBar.Cls
End Sub

Private Sub CopyAllCmd_Click()
    Clipboard.Clear
    Clipboard.SetData ValueBar.Image
    Clipboard.SetData GraphBox.Image
End Sub

Private Sub CopyGraphCmd_Click()
     Clipboard.Clear
     Clipboard.SetData GraphBox.Image
End Sub

Private Sub CopyValueCmd_Click()
     Clipboard.Clear
     Clipboard.SetData ValueBar.Image
End Sub


Private Sub CreateModelProjectionCmd_Click()
Dim i, j, k, ybar1, ybar2
Dim xlow, xhigh, intsize
ReDim modvar(numrecs), dtemp(numrecs, numflds), vntemp$(numflds)
On Error GoTo createmodvarce

'show perdone
Dim regnum, totregs, increment
increment = 10
totregs = 2 * numrecs
Screen.MousePointer = 11
PerDone.Caption = "Processing Data..."
PerDone.Show 0
DoEvents

Select Case modtype$
   Case "Mean"
      For i = 1 To numrecs
         modvar(i) = ymeanpre
         
         'fills dtemp()
         For j = 1 To numflds
            dtemp(i, j) = d(i, j)
         Next j
         
         'update perdone
         regnum = regnum + 1
         Call UpdatePerDone(regnum, totregs, increment)
      Next i
   Case "Median"
      For i = 1 To numrecs
         modvar(i) = ymeanpre
         
         'fills dtemp()
         For j = 1 To numflds
            dtemp(i, j) = d(i, j)
         Next j
         
         'update perdone
         regnum = regnum + 1
         Call UpdatePerDone(regnum, totregs, increment)
      Next i
   Case "2P"
      For i = 1 To numrecs
         If d(i, xfld(1)) <> ndflag Then
            modvar(i) = betapre(1) + betapre(2) * d(i, xfld(1))
         Else
            modvar(i) = ndflag
         End If
         
         'fills dtemp()
         For j = 1 To numflds
            dtemp(i, j) = d(i, j)
         Next j
         
         'update perdone
         regnum = regnum + 1
         Call UpdatePerDone(regnum, totregs, increment)
      Next i
   Case "4P"
      For i = 1 To numrecs
         If d(i, xfld(1)) <> ndflag Then
            If d(i, xfld(1)) <= xcppre Then ivar = 0 Else ivar = 1
            modvar(i) = betapre(1) + betapre(2) * d(i, xfld(1)) + betapre(3) * (d(i, xfld(1)) - xcppre) * ivar
         Else
            modvar(i) = ndflag
         End If
         
         'fills dtemp()
         For j = 1 To numflds
            dtemp(i, j) = d(i, j)
         Next j
         
         'update perdone
         regnum = regnum + 1
         Call UpdatePerDone(regnum, totregs, increment)
      Next i
   Case "3P"
      For i = 1 To numrecs
          If d(i, xfld(1)) <> ndflag Then
             If hcindex = 0 Then
                If d(i, xfld(1)) <= xcppre Then ivar = 0 Else ivar = 1
             Else
                If d(i, xfld(1)) <= xcppre Then ivar = 1 Else ivar = 0
             End If
             modvar(i) = betapre(1) + betapre(2) * (d(i, xfld(1)) - xcppre) * ivar
          Else
             modvar(i) = ndflag
          End If
         
         'fills dtemp()
         For j = 1 To numflds
            dtemp(i, j) = d(i, j)
         Next j
         
         'update perdone
         regnum = regnum + 1
         Call UpdatePerDone(regnum, totregs, increment)
      Next i
   Case "MLR"
      For i = 1 To numrecs
         nd = False
         For j = 1 To mlrxvars
            If d(i, xfld(j)) = ndflag Then nd = True
         Next j
         If nd = False Then
            For j = 1 To mlrxvars
               modvar(i) = modvar(i) + betapre(j + 1) * d(i, xfld(j))
            Next j
            modvar(i) = modvar(i) + betapre(1)
         Else
            modvar(i) = ndflag
         End If
                  
         'fills dtemp()
         For j = 1 To numflds
            dtemp(i, j) = d(i, j)
         Next j
         
         'update perdone
         regnum = regnum + 1
         Call UpdatePerDone(regnum, totregs, increment)
      Next i
   Case "BinFit"
      'determines the interval size
      intsize = (xmaxbf - xminbf) / numints
   
      For i = 1 To numrecs
         If d(i, xfld(1)) <> ndflag Then
            'runs through the j intervals looking for the proper bin
            For j = 1 To numints
               xlow = xminbf + (j - 1) * intsize
               xhigh = xminbf + j * intsize
               If j = 1 Then 'use >= on low bound
                 If d(i, xfld(1)) >= xlow And d(i, xfld(1)) <= xhigh Then
                    numobspost = numobspost + 1
                    If ybarbf(j) = ndflag Then 'search for nearest non-zero ybarbf
                       For k = j + 1 To numints
                          If ybarbf(k) <> ndflag Then
                             'calc pred using this ybarbf
                             modvar(i) = ybarbf(k)
                             Exit For
                          End If
                       Next k
                    Else
                       modvar(i) = ybarbf(j)
                    End If
                    j = 1
                    Exit For
                 End If
               Else  'use only > on low bound
                 If d(i, xfld(1)) > xlow And d(i, xfld(1)) <= xhigh Then
                    numobspost = numobspost + 1
                    If ybarbf(j) = ndflag Then 'find nearest preceding ybarbf()
                       ybar1 = ndflag
                       ybar2 = ndflag
                       For k = j - 1 To 1 Step -1
                          If ybarbf(k) <> ndflag Then
                             'calc pred using this ybarbf
                             ybar1 = ybarbf(k)
                             Exit For
                          End If
                       Next k
                       'find nearest subsequent ybarbf()
                       For k = j + 1 To numints
                          If ybarbf(k) <> ndflag Then
                             'calc pred using this ybarbf
                             ybar2 = ybarbf(k)
                             Exit For
                          End If
                       Next k
                       'calc pred
                       If ybar1 <> ndflag And ybar2 <> ndflag Then
                          modvar(i) = (ybar1 + ybar2) / 2
                       ElseIf ybar1 = ndflag Then
                          modvar(i) = ybar2
                       ElseIf ybar2 = ndflag Then
                          modvar(i) = ybar1
                       End If
                    Else
                       modvar(i) = ybarbf(j)
                    End If
                    j = 1
                    Exit For
                 End If
               End If
            Next j
         Else
            modvar(i) = ndflag
         End If
         'temporarily stores d() in dtemp() so d() can be redimensioned
         For j = 1 To numflds
            dtemp(i, j) = d(i, j)
         Next j
         
         'fills dtemp()
         For j = 1 To numflds
            dtemp(i, j) = d(i, j)
         Next j
         
         'update perdone
         regnum = regnum + 1
         Call UpdatePerDone(regnum, totregs, increment)
      Next i
   Case Else
      MsgBox "No model has been selected.", , "Error"
      GoTo createmodvarce:
End Select

'redims and fills d() [moving 3 resid vars to last]
ReDim d(numrecs, numflds + 1)
For i = 1 To numrecs
   For j = 1 To numflds - 3
      d(i, j) = dtemp(i, j)
   Next j
   d(i, numflds - 2) = modvar(i)
   For j = numflds - 1 To numflds + 1
      d(i, j) = dtemp(i, j - 1)
   Next j
   
   'update perdone
   regnum = regnum + 1
   Call UpdatePerDone(regnum, totregs, increment)
Next i

'fills vn$() [moving 3 resid vars to last]
nvn$ = modtype$ + vn$(yfld)
For i = 1 To numflds
   vntemp$(i) = vn$(i)
Next i
ReDim vn$(numflds + 1)
For i = 1 To numflds - 3
   vn$(i) = vntemp$(i)
Next i
vn$(numflds - 2) = nvn$
For i = numflds - 1 To numflds + 1
   vn$(i) = vntemp$(i - 1)
Next i

'resets numflds
numflds = numflds + 1
numfldsd = numfldsd + 1

'initializes variable boxes
Do While Main.XVarBox.ListCount > 0
   Main.XVarBox.RemoveItem 0
   Main.YVarBox.RemoveItem 0
Loop

'fills variable boxes with new names
For i = 1 To numflds
   Main.XVarBox.AddItem vn$(i)
   Main.YVarBox.AddItem vn$(i)
Next i

'reinitializes timescale variables
timeinthr = False
timeintdy = False
timeintmo = False
timeintyr = False

createmodvarce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error in Calc Savings Procedure"
   Resume createmodvares
End If
createmodvares:

End Sub


Private Sub CreateVariableCmd_Click()
   NewVarDB.Show MODAL
End Sub

Private Sub d1_Click()

End Sub

Private Sub d2_Click()

End Sub


Private Sub d3_Click()

End Sub


Private Sub d4_Click()

End Sub


Private Sub d5_Click()

End Sub


Private Sub DefineLoadCmd_Click()

LoadDB.Show 1

End Sub

Private Sub DefineMeasurementErrorCmd_Click()

MeasErrDB.Show 1

End Sub

Private Sub DefineSystemCmd_Click()

SystemDB.Show 1

End Sub

Private Sub DeleteVariableCmd_Click()

DeleteDB.Show 1

End Sub

Private Sub EditorCmd_Click()

ESEditor.Show 1

End Sub

Private Sub ExitCmd_Click()

Close
If timeinthr = True Then Kill "dhr.tmp"
If timeintdy = True Then Kill "ddy.tmp"
If timeintmo = True Then Kill "dmo.tmp"
If timeintyr = True Then Kill "dyr.tmp"
End

End Sub

Private Sub FivePBut_Click()
Dim i, j, k 'pass is global
Dim numxvars, xcp1, sexcp1, xcp2, sexcp2, ycp, seycp, ls, sels, rs, sers, rmse, cvrmse, r2
ReDim A(numrecs, 3), ivcoefs(7), seivcoefs(7)
On Error GoTo p5ce

'show perdone
Screen.MousePointer = 11
PerDone.Caption = "Processing Data..."
PerDone.Show 0
DoEvents
 
'sets modtype
modtype$ = "5P"

'calls SetAll to fill x(), y(), mins maxs etc.
Call SetAll
ReDim IndexNumO(n) As Integer, WDWEO(n) As Integer, XO(n), YO(n)

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo p5ce:
ElseIf G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined.", , "Error"
   GoTo p5ce:
ElseIf xfld(1) = 0 Then
   MsgBox "No X variable has been defined.", , "Error"
   GoTo p5ce:
End If

'sets pass
If G1G2But.Value = True And gmodnum = 2 Then pass = 2 Else pass = 1

'refills x() and y() if appropriate
If G1G2But.Value = True Then Call SetGrp

'calls FivePMVR
numxvars = 1
Call FivePMVR(x(), y(), n, numxvars, xcp1, sexcp1, xcp2, sexcp2, ycp, seycp, ls, sels, rs, sers, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)

'creates global variables for savings calcuations
If pass = 1 Then
    premodtype$ = modtype$
    rmsepre = rmse
    npre = n
    numxvarspre = 1
    xcp1pre = xcp1
    xcp2pre = xcp2
    ycppre = ycp
    lspre = ls
    rspre = rs
   modelstats = " N = " + Format$(n, "####") + "  R2 = " + Format$(r2, "0.00") + "  RMSE = " + Format$(rmse, "#,##0.00") + "  CV-RMSE = " + Format$(cvrmse, "##0.00") + "%"
   modeleqn = vn$(yfld) + " = " + Format(ycp, "#,##0.00") + "  -  " + Format(ls, "#,##0.00") + " (" + Format(xcp1, "#,##0.00") + " - " + vn$(xfld(1)) + ")+  +  " + Format(rs, "#,##0.00") + " (" + vn$(xfld(1)) + " - " + Format(xcp2, "#,##0.00") + ")+"
End If

'prints results
sf$ = "###0.0000"
If G1G2But.Value = False Or (G1G2But.Value = True And pass = 1) Then ValueBar.Cls
Call PrintHeaderLine(gmodnum, modtype$)
ValueBar.Print " N = " + Format$(n, "####") + "  R2 = " + Format$(r2, "0.00") + "  RMSE = " + Format$(rmse, sf$) + "  CV-RMSE = " + Format$(cvrmse, "##0.0\%")
ValueBar.Print " Ycp = " + Format$(ycp, sf$) + " (" + Format$(seycp, sf$) + ")";
ValueBar.Print "  Xcp1 = " + Format$(xcp1, sf$) + " (" + Format$(sexcp1, sf$) + ")";
ValueBar.Print "  Xcp2 = " + Format$(xcp2, sf$) + " (" + Format$(sexcp2, sf$) + ")";
ValueBar.Print " LS = " + Format$(ls, sf$) + " (" + Format$(sels, sf$) + ")";
ValueBar.Print "  RS = " + Format$(rs, sf$) + " (" + Format$(sers, sf$) + ")"
'ValueBar.Print "Model: "; "Y"; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp1, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(xcp2, "#,##0.00"); ")+" ' + "; Format(ivcoefs(2), "#,##0.00"); " IV1  +  "; Format(ivcoefs(3), "#,##0.00"); " IV2"
'print model equation
ValueBar.Print "Model: "; vn$(yfld); " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp1, "#,##0.00"); " - "; vn$(xfld(1)); ")+  +  "; Format(rs, "#,##0.00"); " ("; vn$(xfld(1)); " - "; Format(xcp2, "#,##0.00"); ")+";
For i = 2 To numxvars
   ValueBar.Print " + "; Format(ivcoefs(i), "#,##0.00"); " "; vn$(xfld(i)); '" X" + Format(i, "0"); '
Next i
ValueBar.Print
ValueBar.Print

'draws model lines
GraphBox.DrawWidth = 1
'middle model line
Call DrawModelLine(predint, pass, xcp1, ycp, xcp2, ycp)
'right & left model lines
For rightleft = 1 To 2
   Select Case rightleft
      Case 1
         'left model line
         slope = ls
         xstart = xminr
         xcp = xcp1
      Case 2
         'right model line
         slope = rs
         xstart = xmaxr
         xcp = xcp2
   End Select
         
   x1 = xstart
   y1 = slope * (x1 - xcp) + ycp
   x2 = xcp
   y2 = ycp
   
   If y1 < yminr Then
      y1 = yminr
      x1 = (y1 - ycp) / slope + xcp
   ElseIf y1 > ymaxr Then
      y1 = ymaxr
      x1 = (y1 - ycp) / slope + xcp
   End If
   Call DrawModelLine(predint, pass, x1, y1, x2, y2)
Next rightleft

'resets gmodnum
If G1G2But.Value = True Then
   If gmodnum = 1 Then
      gmodnum = 2
   Else 'gmodnum = 2
      gmodnum = 1
   End If
End If

p5ce:
Screen.MousePointer = 0
Unload PerDone
If Err = 6 Or Err = 11 Then
   MsgBox "Try using the Data, Modify Values menu items to rescale variables with large values so that the magnitude of those values is decreased.", , "Error"
   Resume p5es
ElseIf Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume p5es
End If
p5es:

End Sub

Private Sub FivePMVRBut_Click()
Dim i, j
Dim numxvars, xcp1, sexcp1, xcp2, sexcp2, ycp, seycp, ls, sels, rs, sers, rmse, cvrmse, r2
ReDim A(numrecs, 2), ivcoefs(7), seivcoefs(7)
Dim xpnt, ypnt, ivar1, ivar2
'pass, xr are global
On Error GoTo psmce

modtype$ = "5P"

'calls Mult Reg dialog box
MultRgDB.Show 1
If cancelmlr = True Then GoTo psmce:
'counts numxvars for multrgdb
numxvars = 1
Do While xfld(numxvars) <> 0
   numxvars = numxvars + 1
Loop
numxvars = numxvars - 1
'displays error message if no vars are selected
If numxvars = 0 Or yfld = 0 Then
   MsgBox "No variables have been selected.  Remember to press the 'Set' button after each variable is selected.", , "Error"
   GoTo psmce:
End If
Screen.MousePointer = 11

'sets pass
If G1G2But.Value = True And gmodnum = 2 Then pass = 2 Else pass = 1
   
'fills X(), Y(), and sets n
ReDim x(numrecs, numxvars + 1), y(numrecs, 1)
Call FillXYGrp(pass, numxvars, x(), y(), n)

'sets xrows and columns
xr = n

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo psmce:
ElseIf G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined.", , "Error"
   GoTo psmce:
ElseIf xfld(1) = 0 Then
   MsgBox "No X variable has been defined.", , "Error"
   GoTo psmce:
End If

'calls FivePMVR
Call FivePMVR(x(), y(), n, numxvars, xcp1, sexcp1, xcp2, sexcp2, ycp, seycp, ls, sels, rs, sers, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)

'creates global variables for savings calcuations
If pass = 1 Then
   premodtype$ = modtype$
   rmsepre = rmse
   npre = n
   numxvarspre = numxvars
   xcp1pre = xcp1
   xcp2pre = xcp2
   ycppre = ycp
   lspre = ls
   rspre = rs
   ReDim ivcoefspre(7)
   For i = 2 To numxvars
      ivcoefspre(i) = ivcoefs(i)
   Next i
   modelstats = " N = " + Format$(n, "####") + "  R2 = " + Format$(r2, "0.00") + "  RMSE = " + Format$(rmse, "#,##0.00") + "  CV-RMSE = " + Format$(cvrmse, "##0.00") + "%"
   modeleqn = vn$(yfld) + " = " + Format(ycp, "#,##0.00") + "  -  " + Format(ls, "#,##0.00") + " (" + Format(xcp1, "#,##0.00") + " - " + vn$(xfld(1)) + ")+  +  " + Format(rs, "#,##0.00") + " (" + vn$(xfld(1)) + " - " + Format(xcp2, "#,##0.00") + ")+"
   For i = 2 To numxvars
      modeleqn = modeleqn + " + " + Format(ivcoefs(i), "#,##0.00") + " " + vn$(xfld(i))
   Next i
End If

'prints results
sf$ = "###0.0000"
If G1G2But.Value = False Or (G1G2But.Value = True And pass = 1) Then ValueBar.Cls
Call PrintHeaderLine(gmodnum, modtype$)
ValueBar.Print " N = " + Format$(n, "####") + "  R2 = " + Format$(r2, "0.00") + "  RMSE = " + Format$(rmse, "#,##0.00") + "  CV-RMSE = " + Format$(cvrmse, "##0.00") + "%"
ValueBar.Print " Ycp = " + Format$(ycp, sf$) + " (" + Format$(seycp, sf$) + ")";
ValueBar.Print "  Xcp1 = " + Format$(xcp1, sf$) + " (" + Format$(sexcp1, sf$) + ")";
ValueBar.Print "  Xcp2 = " + Format$(xcp2, sf$) + " (" + Format$(sexcp2, sf$) + ")";
ValueBar.Print " LS = " + Format$(ls, sf$) + " (" + Format$(sels, sf$) + ")";
ValueBar.Print "  RS = " + Format$(rs, sf$) + " (" + Format$(sers, sf$) + ")"
If numxvars >= 2 Then
   For i = 2 To numxvars
      'If numxvars = 4 Then ValueBar.Print
      ValueBar.Print "  X" + Format(i, "0") + " = "; Format(ivcoefs(i), "0.0000"); " ("; Format(seivcoefs(i), "0.0000"); ")  ";
   Next i
   ValueBar.Print
End If
'print model equation
ValueBar.Print "Model: "; vn$(yfld); " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp1, "#,##0.00"); " - "; vn$(xfld(1)); ")+  +  "; Format(rs, "#,##0.00"); " ("; vn$(xfld(1)); " - "; Format(xcp2, "#,##0.00"); ")+";
For i = 2 To numxvars
   ValueBar.Print " + "; Format(ivcoefs(i), "#,##0.00"); " "; vn$(xfld(i)); '" X" + Format(i, "0"); '
Next i
ValueBar.Print

'calcs prediction interval
t = 1.96 + 2.7 / (n - 3)
predint = t * rmse * (1 + 1 / n) ^ 0.5

'mlr graphing
'calcs ypred and saves values of xpred and ypred before they get changed by the call to xybut
ReDim ypred(n), xpred(n)
For i = 1 To xr
   'If y(i, 1) = 21.8 Then Stop
   'calc ypred and res
   If x(i, 2) <= xcp1 Then ivar1 = 1 Else ivar1 = 0
   If x(i, 2) <= xcp2 Then ivar2 = 0 Else ivar2 = 1
   ypred(i) = ycp + ls * (x(i, 2) - xcp1) * ivar1 + rs * (x(i, 2) - xcp2) * ivar2
   For j = 2 To numxvars
      ypred(i) = ypred(i) + ivcoefs(j) * x(i, j + 1)
   Next j
   xpred(i) = x(i, 2)
Next i

'replot xy graph of y and first x var in mult regression
If G1G2But.Value = False Or (G1G2But.Value = True And pass = 1) Or (G1G2But.Value = True And pass = 2 And Main.XVarBox.ListIndex + 1 <> xfld(1)) Then
   'selects the first x variable in mult reg as the x variable for xy plot
   Main.XVarBox.ListIndex = xfld(1) - 1
   'selects the first y variable in mult reg as the y variable for xy plot
   Main.YVarBox.ListIndex = yfld - 1
   'calls xy plot
   Call XYBut_Click
End If

'draws mlr points
GraphBox.DrawWidth = 3
For i = 1 To xr
   ypnt = (ypred(i) - yminr) / yranger * 100
   xpnt = (xpred(i) - xminr) / xranger * 100
   
   If xpnt >= 0 And xpnt <= 100 And ypnt >= 0 And ypnt <= 100 Then
      If G1G2But.Value = False Then
         GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), BLACK, B
      Else
         If pass = 1 Then
            GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), BLUE, B
         Else
            GraphBox.Circle (xpnt, ypnt), 0.8, RED
         End If
      End If
   End If
Next i
ReDim ypred(1), xpred(1)
GraphBox.DrawWidth = 1

'If numpasses = 2 Then Call SetAll
If G1G2But.Value = True And pass = 2 Then Call SetAll

'resets gmodnum
If G1G2But.Value = True Then
  If gmodnum = 1 Then
     gmodnum = 2
  Else 'gmodnum = 2
     gmodnum = 1
  End If
End If

psmce:
Screen.MousePointer = 0
Unload PerDone
If Err = 6 Or Err = 11 Then
  MsgBox "Try using the Data, Modify Values menu items to rescale variables with large values so that the magnitude of those values is decreased.", , "Error"
  Resume psmes
ElseIf Err Then
  MsgBox """" + Error(Err) + """", , "Error"
  Resume psmes
End If
psmes:

End Sub

Private Sub Form_Load()

Main.Move 0.1 * Screen.Width, 0.1 * Screen.Height, 0.8 * Screen.Width, 0.8 * Screen.Height


SelDatesCmd.Enabled = False
ViewDatesCmd.Enabled = False
ZoomCmd.Enabled = False
VarNamesCmd.Enabled = False
CreateVariableCmd.Enabled = False
DeleteVariableCmd.Enabled = False
ModValuesCmd.Enabled = False
Main.AvgIntDataCmd.Enabled = False
SaveDataCmd.Enabled = False
Main.SaveNamesCmd.Enabled = False
GroupByDOWCmd.Enabled = False
GroupByDateCmd.Enabled = False
SaveLogCmd.Enabled = False
CreateModelProjectionCmd.Enabled = False
TimeScalesCmd.Enabled = False


'set default for 3d
thetad = 45

'set default for group modeling
gmodnum = 1

'open the file listed as a command line
If Command <> "" Then ' If no command line.
   OpenDataCmd_Click
End If

'set number of files the first time emodel opens
numfiles = 0


'*** Code added by HelpWriter ***
'*** Subroutine added by HelpWriter ***
    'SetAppHelp Me.hWnd
    Call SetAppHelp(Me.hWnd)
'***********************************

'=====================================================================
'To use these subroutines to access WINHELP, you need to add
'at least this one subroutine call to your code
'     o  In the Form_Load event of your main Form enter:
'        Call SetAppHelp(Me.hWnd) 'To setup helpfile variables
'         (If you are not interested in keyword searching or context
'         sensitive help, this is the only call you need to make!)
'=====================================================================


End Sub

Private Sub Form_Resize()
Dim sw, sh

sw = Main.ScaleWidth / 9930
sh = Main.ScaleHeight / 6630

ValueBar.Move 0, 0, 7695 * sw, 1320 * sh
ClearValueBut.Move 7695 * sw - 245, 0
GraphBox.Move 0, 1320 * sh, 7695 * sw, 4575 * sh
ClearGraphBut.Move 7695 * sw - 245, 1320 * sh

'YVarLab.Move 7905 * sw, 120 * sh
'YVarBox.Move 7905 * sw, 360 * sh
'XVarLab.Move 7905 * sw, 1920 * sh
'XVarBox.Move 7905 * sw, 2160 * sh
YVarLab.Move 7800 * sw, 120 * sh
YVarBox.Move 7800 * sw, 360 * sh
XVarLab.Move 7800 * sw, 1920 * sh
XVarBox.Move 7800 * sw, 2160 * sh

TimeSeriesBut.Move 7800 * sw, 3720 * sh
'TimeSeries2But.Move 7800 * sw + 1080, 3720 * sh
TimeSeries2But.Move 7800 * sw + 1.88 * 1080, 3720 * sh
XYBut.Move 7800 * sw, 4200 * sh
XYAnimateBut.Move 7800 * sw, 4680 * sh
HistBut.Move 7800 * sw, 5160 * sh
HourlyProfilesBut.Move 7800 * sw, 5640 * sh

SingleGroupBut.Move 6720 * sw, 6195 * sh
G1G2But.Move 8280 * sw, 6240 * sh
G1Lab.Move 8520 * sw, 6240 * sh, 360
G2Lab.Move 8520 * sw + 360, 6240 * sh, 360

If 1 = 2 Then
MeanBut.Move 200 * sw, 6030 * sh
MedianBut.Move 200 * sw, 6030 * sh + 250
TwoPBut.Move 1100 * sw, 6030 * sh
MLRBut.Move 1100 * sw + 750, 6030 * sh
ThreePBut(0).Move 2000 * sw, 6030 * sh
ThreePMVRBut(0).Move 2000 * sw + 750, 6030 * sh
ThreePBut(1).Move 2000 * sw, 6030 * sh + 250
ThreePMVRBut(1).Move 2000 * sw + 750, 6030 * sh + 250
FourPBut.Move 2900 * sw, 6030 * sh
FourPMVRBut.Move 2900 * sw + 750, 6030 * sh
FivePBut.Move 3800 * sw, 6030 * sh
FivePMVRBut.Move 3800 * sw + 750, 6030 * sh
'MLRBut.Move 4700 * sw, 6030 * sh
BinFitBut.Move 5600 * sw, 6030 * sh
End If

MeanBut.Move 200 * sw, 6030 * sh
MedianBut.Move 200 * sw, 6030 * sh + 250
TwoPBut.Move 1300 * sw, 6030 * sh
MLRBut.Move 1300 * sw + 750, 6030 * sh
ThreePBut(0).Move 2400 * sw, 6030 * sh
ThreePMVRBut(0).Move 2400 * sw + 750, 6030 * sh
ThreePBut(1).Move 2400 * sw, 6030 * sh + 250
ThreePMVRBut(1).Move 2400 * sw + 750, 6030 * sh + 250
FourPBut.Move 3500 * sw, 6030 * sh
FourPMVRBut.Move 3500 * sw + 750, 6030 * sh
FivePBut.Move 4600 * sw, 6030 * sh
FivePMVRBut.Move 4600 * sw + 750, 6030 * sh
BinFitBut.Move 5700 * sw, 6030 * sh


End Sub

Private Sub Form_Unload(Cancel As Integer)
Close
If timeinthr = True Then Kill "dhr.tmp"
If timeintdy = True Then Kill "ddy.tmp"
If timeintmo = True Then Kill "dmo.tmp"
If timeintyr = True Then Kill "dyr.tmp"


'*** Code added by HelpWriter ***
'*** Subroutine added by HelpWriter ***
    QuitHelp
'***********************************

End

DoEvents

End Sub

Private Sub FourPBut_Click()
ReDim A(numrecs, 3), rmsei(100)
Dim numints As Integer, itration As Integer
Dim i As Integer, j As Integer, n1 As Integer, n2 As Integer
Dim numpasses As Integer, count As Integer
Dim totregs, numregs
Dim xmin4p, seycp
On Error GoTo cpce

'show perdone
Screen.MousePointer = 11
PerDone.Caption = "Processing Data..."
PerDone.Show 0
DoEvents

'call SetAll to fill x(),y(), min, max etc
Call SetAll
ReDim IndexNumO(n) As Integer, WDWEO(n) As Integer, XO(n), YO(n)

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo cpce:
ElseIf G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined.", , "Error"
   GoTo cpce:
ElseIf xfld(1) = 0 Then
   MsgBox "No X variable has been defined.", , "Error"
   GoTo cpce:
End If

modtype = "4P"

'sets pass
If G1G2But.Value = True And gmodnum = 2 Then pass = 2 Else pass = 1

    
'refills x() and y() if appropriate
If G1G2But.Value = True Then Call SetGrp

'sets xrows and x columns
xr = n
xc = 2
    
'Calcs initial size of search grid
numints = 10
inc = (xmax - xmin) / numints

'find total regressions to be performed for perdone
totregs = 2 * (numints - 1) + 1 'for fine search


'FINDS CP BY USING COURSE AND THEN FINER SEARCHES, ITRATION = NUM OF SEARCHES
For itration = 1 To 2
   
   'set xmin4p and inc depending on course or fine grid search
   If itration = 1 Then
      xmin4p = xmin
   Else 'itration = 2 for fine grid...
      xmin4p = bestcp - inc
      inc = 2 * inc / numints
   End If

   For j = 1 To numints - 1
      rmsei(j) = 0
   Next j
   
   For i = 1 To numints - 1   'FOR EACH INCREMENT....
      'INITIALIZE COUNTERS AND ARRAYS
      n1 = 0: n2 = 0

      'FILL A()
      cp = xmin4p + i * inc
      For j = 1 To n
         If x(j, 2) <= cp Then
            n1 = n1 + 1
            IndVar = 0
         Else
            n2 = n2 + 1
            IndVar = 1
         End If
         A(j, 1) = 1
         A(j, 2) = x(j, 2)
         A(j, 3) = IndVar * (x(j, 2) - cp)
      Next j

     'prints error if n = 0
     If n = 0 Then
       MsgBox "No data is available to model.", , "Error"
       Screen.MousePointer = 0
       GoTo cpce:
     End If

      Call Reg(A(), n, 3, y(), n, 1, beta(), betar, betac, sec())
      Call Inf(A(), n, 3, y(), n, 1, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), P)
     
      'SELECTS CP AT MINIMUM RMSE
      rmsei(i) = rmse
      If i = 1 Then
         rmsemin = rmsei(1)
         bestcp = cp
      End If
      If rmsei(i) < rmsemin Then
         rmsemin = rmsei(i)
         bestcp = cp
      End If
      
      'update perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, 1)
   
   Next i
Next itration

'USES BEST CP TO FIND FINAL STATS
n1 = 0: n2 = 0
cp = bestcp
For j = 1 To n
   If x(j, 2) <= cp Then
      n1 = n1 + 1
      IndVar = 0
   Else
      n2 = n2 + 1
      IndVar = 1
   End If
   A(j, 1) = 1
   A(j, 2) = x(j, 2)
   A(j, 3) = IndVar * (x(j, 2) - cp)
Next j
Call Reg(A(), n, 3, y(), n, 1, beta(), betar, betac, sec())
Call Inf(A(), n, 3, y(), n, 1, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), P)

'calcs cps and slopes from beta()
xcp = bestcp
ycp = beta(1, 1) + beta(2, 1) * bestcp
ls = beta(2, 1)
rs = beta(2, 1) + beta(3, 1)
If Abs(beta(2, 1)) > 0 And Abs(xcp) > 0 Then
   seycp = (sigma(1) ^ 2 + (beta(2, 1) * xcp) ^ 2 * ((sigma(2) / beta(2, 1)) ^ 2 + (sexcp / xcp) ^ 2)) ^ 0.5
Else
   seycp = inc
End If

'creates global variables for savings calcuations
If pass = 1 Then
    premodtype$ = modtype$
    rmsepre = rmse
    npre = n
    numxvarspre = numxvars
    xcppre = xcp
    betapre(1) = beta(1, 1)
    betapre(2) = beta(2, 1)
    betapre(3) = beta(3, 1)
    modelstats = " N = " + Format$(n, "####") + "  R2 = " + Format$(r2, "0.00") + "  RMSE = " + Format$(rmse, "#,##0.00") + "  CV-RMSE = " + Format$(cvrmse, "##0.0") + "%"
    modeleqn = vn$(yfld) + " = " + Format$(ycp, "#,##0.00") + "  -  " + Format$(ls, "#,##0.00") + " (" + Format$(xcp, "#,##0.00") + " - " + vn$(xfld(1)) + ")+  +  " + Format$(rs, "#,##0.00") + " (" + vn$(xfld(1)) + " - " + Format$(xcp, "#,##0.00") + ")+"
End If

'prints results
Call DWTest(3, dwr$)
sf$ = "###0.0000"
If G1G2But.Value = False Or (G1G2But.Value = True And pass = 1) Then ValueBar.Cls
Call PrintHeaderLine(gmodnum, modtype$)
'ValueBar.Print " N = " + Format$(n, "####") + "  N1 = " + Format$(n1, "####") + "  N2 = " + Format$(n2, "####")
ValueBar.Print " N = " + Format$(n, "####") + "  R2 = " + Format$(r2, "0.00") + "  RMSE = " + Format$(rmse, sf$) + "  CV-RMSE = " + Format$(cvrmse, "##0.0\%")
ValueBar.Print " Ycp = " + Format$(ycp, sf$) + " (" + Format$(seycp, sf$) + ")";
ValueBar.Print "  Xcp = " + Format$(xcp, sf$) + " (" + Format$(inc, sf$) + ")";
ValueBar.Print "  LS = " + Format$(ls, sf$) + " (" + Format$(sigma(2), sf$) + ")";
ValueBar.Print "  RS = " + Format$(rs, sf$) + " (" + Format$((sigma(2) ^ 2 + sigma(3) ^ 2) ^ 0.5, sf$) + ")"
'ValueBar.Print "  p = " + Format$(autocor, "0.00") + "  DW = " + Format$(dw, "0.00") + " (" + dwr$ + ")"
'print model
ValueBar.Print "Model: "; vn$(yfld); " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp, "#,##0.00"); " - "; vn$(xfld(1)); ")+  +  "; Format(rs, "#,##0.00"); " ("; vn$(xfld(1)); " - "; Format(xcp, "#,##0.00"); ")+";
'For i = 2 To numxvars
   'ValueBar.Print " + "; Format(ivcoefs(i), "#,##0.00"); " "; vn$(xfld(i)); '" X" + Format(i, "0"); '
'Next i
ValueBar.Print
ValueBar.Print

'ValueBar.Print "Model: "; "Y"; " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp, "#,##0.00"); " - T)+  +  "; Format(rs, "#,##0.00"); " (T - "; Format(xcp, "#,##0.00"); ")+"
'ValueBar.Print

'calcs prediction interval
t = 1.96 + 2.7 / (n - 3)
predint = t * rmse * (1 + 1 / n) ^ 0.5

'draws left and right model line
For rightleft = 1 To 2
    'calcs model line coordinates
    Select Case rightleft
        Case 1
            'left model line
            slope = ls
            xstart = xminr
        Case 2
            'right model line
            slope = rs
            xstart = xmaxr
        End Select
            
    x1 = xstart
    y1 = slope * (x1 - xcp) + ycp
    
    x2 = xcp
    y2 = ycp
    
    If y1 < yminr Then
        y1 = yminr
        x1 = (y1 - ycp) / slope + xcp
    ElseIf y1 > ymaxr Then
        y1 = ymaxr
        x1 = (y1 - ycp) / slope + xcp
    End If
    
    Call DrawModelLine(predint, pass, x1, y1, x2, y2)
Next rightleft

'resets gmodnum
If G1G2But.Value = True Then
   If gmodnum = 1 Then
      gmodnum = 2
   Else 'gmodnum = 2
      gmodnum = 1
   End If
End If

cpce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   MsgBox """" + Error(Err) + """" + Chr(13) + Chr(10) + "Try using the Data, Modify Values menu items to rescale variables with large values so that their magnitudes are decreased.", , "Error"
   Resume cpes
End If
cpes:
End Sub

Private Sub FourPMVRBut_Click()
Dim i, j
Dim numxvars, xcp, sexcp, ycp, seycp, ls, sels, rs, sers, rmse, cvrmse, r2
ReDim ivcoefs(7), seivcoefs(7)
Dim ivar, b1, b2, b3, xpnt, ypnt
'pass, xr are global
On Error GoTo psmce

'sets modtype
modtype$ = "4P"

'calls mult reg dialog box
MultRgDB.Show 1
If cancelmlr = True Then GoTo psmce:
'counts numxvars for multrgdb
numxvars = 1
Do While xfld(numxvars) <> 0
   numxvars = numxvars + 1
Loop
numxvars = numxvars - 1
'displays error message if no vars are selected
If numxvars = 0 Or yfld = 0 Then
   MsgBox "No variables have been selected.  Remember to press the 'Set' button after each variable is selected.", , "Error"
   GoTo psmce:
End If
Screen.MousePointer = 11

'sets pass
If G1G2But.Value = True And gmodnum = 2 Then pass = 2 Else pass = 1
   
'fills X(), Y(), and sets n after call to mult reg db
ReDim x(numrecs, numxvars + 1), y(numrecs, 1)
Call FillXYGrp(pass, numxvars, x(), y(), n)

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo psmce:
ElseIf G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined.", , "Error"
   GoTo psmce:
ElseIf xfld(1) = 0 Then
   MsgBox "No X variable has been defined.", , "Error"
   GoTo psmce:
End If

'calls FourPMVR
Call FourPMVR(x(), y(), n, numxvars, xcp, sexcp, ycp, seycp, ls, sels, rs, sers, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)
b2 = ls
b1 = ycp - b2 * xcp
b3 = rs - b2

'creates global variables for savings calcuations
If pass = 1 Then
   premodtype$ = modtype$
   rmsepre = rmse
   npre = n
   numxvarspre = numxvars
   xcppre = xcp
   betapre(1) = b1
   betapre(2) = b2
   betapre(3) = b3
   ReDim ivcoefspre(7)
   For i = 2 To numxvars
      ivcoefspre(i) = ivcoefs(i)
   Next i
   modelstats = " N = " + Format$(n, "####") + "  R2 = " + Format$(r2, "0.00") + "  RMSE = " + Format$(rmse, sf$) + "  CV-RMSE = " + Format$(cvrmse, "##0.00") + "%"
   modeleqn = vn$(yfld) + " = " + Format(ycp, "#,##0.00") + "  -  " + Format(ls, "#,##0.00") + " (" + Format(xcp, "#,##0.00") + " - " + vn$(xfld(1)) + ")+  +  " + Format(rs, "#,##0.00") + " (" + vn$(xfld(1)) + " - " + Format(xcp, "#,##0.00") + ")+"
   For i = 2 To numxvars
      modeleqn = modeleqn + " + " + Format(ivcoefs(i), "#,##0.00") + " " + vn$(xfld(i))
   Next i
ValueBar.Print

End If

'prints results
'Call DWTest(3, dwr$)
sf$ = "###0.0000"
If G1G2But.Value = False Or (G1G2But.Value = True And pass = 1) Then ValueBar.Cls
Call PrintHeaderLine(gmodnum, modtype$)
ValueBar.Print " N = " + Format$(n, "####") + "  R2 = " + Format$(r2, "0.00") + "  RMSE = " + Format$(rmse, sf$) + "  CV-RMSE = " + Format$(cvrmse, "##0.00") + "%"
ValueBar.Print " Ycp = " + Format$(ycp, sf$) + " (" + Format$(seycp, sf$) + ")";
ValueBar.Print "  Xcp = " + Format$(xcp, sf$) + " (" + Format$(sexcp, sf$) + ")";
ValueBar.Print "  LS = " + Format$(ls, sf$) + " (" + Format$(sels, sf$) + ")";
ValueBar.Print "  RS = " + Format$(rs, sf$) + " (" + Format$(sers, sf$) + ")"
'ValueBar.Print "  p = " + Format$(autocor, "0.00") + "  DW = " + Format$(dw, "0.00") + " (" + dwr$ + ")"
If numxvars >= 2 Then
   For i = 2 To numxvars
      ValueBar.Print " X" + Format(i, "0") + " = "; Format(ivcoefs(i), "0.0000"); " ("; Format(seivcoefs(i), "0.0000"); ")  ";
   Next i
   ValueBar.Print
End If
'print model
ValueBar.Print "Model: "; vn$(yfld); " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp, "#,##0.00"); " - "; vn$(xfld(1)); ")+  +  "; Format(rs, "#,##0.00"); " ("; vn$(xfld(1)); " - "; Format(xcp, "#,##0.00"); ")+";
For i = 2 To numxvars
   ValueBar.Print " + "; Format(ivcoefs(i), "#,##0.00"); " "; vn$(xfld(i)); '" X" + Format(i, "0"); '
Next i
ValueBar.Print


'calcs prediction interval
t = 1.96 + 2.7 / (n - 3)
predint = t * rmse * (1 + 1 / n) ^ 0.5

'mlr graphing
xr = n
ReDim ypred(xr), xpred(xr)
'calcs ypred and saves values of xpred and ypred before they get changed by the call to xybut
For i = 1 To xr
   If x(i, 2) <= xcp Then ivar = 0 Else ivar = 1
   ypred(i) = b1 + b2 * x(i, 2) + b3 * (x(i, 2) - xcp) * ivar
   For j = 2 To numxvars
      ypred(i) = ypred(i) + ivcoefs(j) * x(i, j + 1)
   Next j
   xpred(i) = x(i, 2)
Next i

'replot xy graph of y and first x var in mult regression
If G1G2But.Value = False Or (G1G2But.Value = True And pass = 1) Or (G1G2But.Value = True And pass = 2 And Main.XVarBox.ListIndex + 1 <> xfld(1)) Then
   'selects the first x variable in mult reg as the x variable for xy plot
   Main.XVarBox.ListIndex = xfld(1) - 1
   'selects the first y variable in mult reg as the y variable for xy plot
   Main.YVarBox.ListIndex = yfld - 1
   'calls xy plot
   Call XYBut_Click
End If

'draws mlr points
GraphBox.DrawWidth = 3
For i = 1 To xr
   ypnt = (ypred(i) - yminr) / yranger * 100
   xpnt = (xpred(i) - xminr) / xranger * 100
   
   If xpnt >= 0 And xpnt <= 100 And ypnt >= 0 And ypnt <= 100 Then
      If G1G2But.Value = False Then
         GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), BLACK, B
      Else
         If pass = 1 Then
            GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), BLUE, B
         Else
            GraphBox.Circle (xpnt, ypnt), 0.8, RED
         End If
      End If
   End If
Next i
ReDim ypred(1), xpred(1)
GraphBox.DrawWidth = 1

'If numpasses = 2 Then Call SetAll
If G1G2But.Value = True And pass = 2 Then Call SetAll

'resets gmodnum
If G1G2But.Value = True Then
  If gmodnum = 1 Then
     gmodnum = 2
  Else 'gmodnum = 2
     gmodnum = 1
  End If
End If

psmce:
Screen.MousePointer = 0
Unload PerDone
If Err = 6 Or Err = 11 Then
  MsgBox "Try using the Data, Modify Values menu items to rescale variables with large values so that the magnitude of those values is decreased.", , "Error"
  Resume psmes
ElseIf Err Then
  MsgBox """" + Error(Err) + """", , "Error"
  Resume psmes
End If
psmes:

End Sub

Private Sub G1G2But_Click()
gmodnum = 1
End Sub

Private Sub GraphBox_DblClick()
Dim hr, dY
Dim i, m As Integer
Dim x1, y1, z1, x2, y2, z2
Dim xs1 As Integer, ys1, xs2 As Integer, ys2
Dim yvar
Dim diff, lastdiff, theta

If gtype$ = "HP" Then
   
   'sets scale and rotation angle
   GraphBox.Scale (-0.1 * res, 1.1 * res)-(1.1 * res, -0.1 * res)
   theta = thetab / 360 * 2 * 3.1416
   sint = Sin(theta)
   cost = Cos(theta)
   
   'find closest xs1 to xclick
   For dY = 1 To numdays
      'define the first base point in each daily load trace
      x1 = dY / numdays * res: y1 = 0: z1 = res
      xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
      ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
      
      diff = Abs(xclick - xs1)
      If dY = 1 Then
         lastdiff = Abs(xclick - xs1)
      Else
         If diff > lastdiff Then
            Exit For
         Else
            lastdiff = diff
         End If
      End If
   Next dY
   dy1 = dY - 1
   'kkk
   Call SingleProfile
   Screen.MousePointer = 0
End If




End Sub

Private Sub GraphBox_KeyDown(keycode As Integer, Shift As Integer)
   
'kkk
If gtype$ = "HP" Then
   
   'increments dy1
   If keycode = KEY_RIGHT Then
         dy1 = dy1 + 1
   ElseIf keycode = KEY_LEFT Then
         dy1 = dy1 - 1
   End If

   'keeps cursor in bounds
   If dy1 < 0 Then dy1 = 0
   If dy1 > numdays - 1 Then dy1 = numdays - 1
   
   Call SingleProfile
   Screen.MousePointer = 0
   Exit Sub
End If

'increments cursor one time interval
If keycode = KEY_RIGHT Then
      currec = currec + 1
ElseIf keycode = KEY_LEFT Then
      currec = currec - 1
End If

'keeps cursor in bounds
If currec < 1 Then currec = 1
If currec > numrecs Then currec = numrecs

Call Cursor

'displays new value bar
If gtype$ = "TS" Then
   Call CalendarDate(currec, FoundDate$)
   ValueBar.Cls
   ValueBar.Print FoundDate$ + "    " + YVarBox.Text + " = " + Str$(d(currec, yfld))
ElseIf gtype$ = "XY" Then
   Call CalendarDate(currec, FoundDate$)
   ValueBar.Cls
   'ValueBar.Print " " + UCase$(filename$)
   ValueBar.Print FoundDate$ + "    " + vn$(xfld(1)) + " = " + Str$(d(currec, xfld(1))) + "    " + YVarBox.Text + " = " + Str$(d(currec, yfld))
End If

End Sub

Private Sub GraphBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim tol, xtol, ytol, pnt$
Dim i As Integer

'pass location of pass to double click command 'kk
'for hourly profiles graph
xclick = x                                     'kk
yclick = y                                     'kk

If ZoomCmd.Checked = True And gtype$ = "TS" Then
   
   zoom = True
   'make global vars x1,y1 to mark start place of zoom
   x1 = x
   y1 = y
   
   'make global vars lastx,lasty to erase first intermediate box in mousemove procedure
   lastx = x
   lasty = y

Else
    'sets search tolerance, must click within tolerance of actual point
    pnt$ = "searching"
    tol = 0.01
    
    Select Case gtype$
    
    Case "XY"
       
       'sets tolerance for how close mouse must be to click
       xtol = tol * xranger
       ytol = tol * yranger
       
       'makes sure click is in bounds
       If x >= 0 And x <= 100 And y >= 0 And y <= 100 Then
           
           'clear valuebar
           ValueBar.Cls
           'ValueBar.Print " " + UCase$(filename$)
           
           'change to graph coordinates
           yval = y / 100 * yranger + yminr
           xval = x / 100 * xranger + xminr
        
           'search for point close to click
           For i = 1 To numrecs
              If d(i, xfld(1)) > (xval - xtol) And d(i, xfld(1)) < (xval + xtol) Then
                   If d(i, yfld) > (yval - ytol) And d(i, yfld) < (yval + ytol) Then
                      'found point
                      Call CalendarDate(i, FoundDate$)
                      ValueBar.Print FoundDate$ + "    " + vn$(xfld(1)) + " = " + Str$(d(i, xfld(1))) + "    " + YVarBox.Text + " = " + Str$(d(i, yfld))
                      pnt$ = "found"
                      currec = i
                   End If
              End If
           Next i
       End If
    
       If pnt$ = "found" Then
           Call Cursor
       Else
           ValueBar.Cls
           ValueBar.Print "Point Not Found"
       End If
    
    Case "TS"
       
       If zoommode = False Then
         i1 = 1
         i2 = numrecs
       End If

       xtol = tol * (i2 - i1 + 1)
       ytol = tol * yranger
    
       If x >= 0 And x <= 100 And y >= 0 And y <= 100 Then
           ValueBar.Cls
           'ValueBar.Print " " + UCase$(filename$)
           
           xval = x / 100 * (i2 - i1 + 1) + i1
           yval = y / 100 * yranger + yminr
           
           lowi = xval - xtol
           highi = xval + xtol
           If lowi < 1 Then lowi = 1
           If highi > numrecs Then highi = numrecs
            
           For i = lowi To highi
              If i > (xval - xtol) And i < (xval + xtol) Then
                   If d(i, yfld) > (yval - ytol) And d(i, yfld) < (yval + ytol) Then
                      Call CalendarDate(i, FoundDate$)
                      ValueBar.Print FoundDate$ + "    " + YVarBox.Text + " = " + Str$(d(i, yfld))
                      pnt$ = "found"
                      currec = i
                   End If
              End If
           Next i
       End If
    
       If pnt$ = "found" Then
          Call Cursor
       Else
          ValueBar.Cls
          ValueBar.Print "Point Not Found"
       End If
    
    End Select

End If

gbes:
End Sub

Private Sub GraphBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If gtype$ = "TS" And ZoomCmd.Checked = True And zoom = True Then
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
   'erase last line
   GraphBox.Line (x1, y1)-(lastx, lasty), BackColor, B
   'draw intermediate lines
   GraphBox.Line (x1, y1)-(x, y), QBColor(9), B
   'save coordinates of last line
   lastx = x
   lasty = y
End If
End Sub

Private Sub GraphBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

If gtype$ = "TS" And ZoomCmd.Checked = True Then
   zoom = False
   
   'draw final box
   GraphBox.Line (x1, x2)-(x, y), QBColor(9), B
   
   x2 = x
   y2 = y
   
   'makes x1 < x2
   If x2 < x1 Then
      x1temp = x1
      x1 = x2
      x2 = x1temp
   End If
   
   'keeps x1 and x2 within legal bounds
   If x1 > 100 Or x2 < 0 Then
      MsgBox "Limit size of zoom area to within the X axis bounds.", , "Error"
      GoTo mues:
   End If
   If x1 < 0 Then x1 = 0
   If x2 > 100 Then x2 = 100
   
   'converts x1 and x2 into i values (i1=1, i2=numrecs)
   xscale = i2 - i1 + 1
   i1o = i1
   i1 = i1o + x1 * xscale / 100
   i2 = i1o + x2 * xscale / 100
   
   'calls timeseries button
   zoommode = True
   ZoomTimeOutCmd.Enabled = True
   TimeSeriesBut.Value = True
End If

mues:
End Sub

Private Sub GridCmd_Click()

GridCmd.Checked = Not GridCmd.Checked
If gtype$ = "TS" Then
   TimeSeriesBut.Value = True
ElseIf gtype$ = "XY" Then
   XYBut.Value = True
End If

End Sub

Private Sub GroupByDateCmd_Click()
Dim sdate, i As Integer
Dim errflag As Integer
ReDim grpd(numrecs)

AddGrpDB.Show MODAL

If agcancelflag = True Then GoTo ages:
Screen.MousePointer = 11

'finds g1 and g2 names
If AddGrpDB.Group1Box.Text = "" Then g1name$ = "G1" Else g1name$ = AddGrpDB.Group1Box.Text
If AddGrpDB.Group2Box.Text = "" Then g2name$ = "G2" Else g2name$ = AddGrpDB.Group2Box.Text

'finds start date from sdatebox
sd$ = AddGrpDB.StartDateBox.Text
If sd$ = "" Then
   g1start = 0
Else
   'Call TxtToSer(sd$, g1start, errflag)
   g1start = SerialDate(sd$, errflag)
End If

'finds end date from edatebox
ed$ = AddGrpDB.EndDateBox.Text
If ed$ = "" Then
   g1end = 9999999999#
Else
   'Call TxtToSer(ed$, g1end, errflag)
   g1end = SerialDate(ed$, errflag)
End If

If errflag = True Then
   MsgBox "Enter dates exactly as shown in Time Interval text box.", , "Error"
   GoTo ages:
End If

For i = 1 To numrecs
   Call SerDate(i, sdate)
   If sdate >= g1start And sdate <= g1end Then
      grpd(i) = 1
   Else
      grpd(i) = 2
   End If
Next i

grpdefined = True
grptype = "date"

'defines resid variables and fills variable boxes
vn$(numfldsd + 2) = g1name$ + "_Resids"
vn$(numfldsd + 3) = g2name$ + "_Resids"
For i = 1 To 2
   XVarBox.RemoveItem numfldsd + 1
   YVarBox.RemoveItem numfldsd + 1
Next i
For i = 1 To 2
    XVarBox.AddItem vn$(numfldsd + i + 1), numfldsd + i
    YVarBox.AddItem vn$(numfldsd + i + 1), numfldsd + i
Next i

agce:
Screen.MousePointer = 0
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume ages
End If
ages:
End Sub

Private Sub GroupByDOWCmd_Click()
Dim i As Integer, j As Integer, intyr As Integer
Dim yr As Integer, mo As Integer, dY As Integer, numdays As Integer
Dim dow As Integer
On Error GoTo adddowce

If d(1, yrfld) < 1989 Then
   MsgBox "Day of week procedure is only available for dates after January 1, 1989", , "Error"
   GoTo adddowes:
End If

Screen.MousePointer = 11

For i = 1 To numrecs
    yr = d(i, yrfld)
    mo = d(i, mofld)
    dY = d(i, dyfld)
    dow = Weekday(Str(mo) + "/" + Str(dY) + "/" + Str(yr))
        
    If dow = 1 Or dow = 7 Then
       grpd(i) = 2
    Else
       grpd(i) = 1
    End If
Next i

grpdefined = True
grptype = "dow"
g1name$ = "WD"
g2name$ = "WE"

'defines resid variables and fills variable boxes
vn$(numfldsd + 2) = "WD_Resids"
vn$(numfldsd + 3) = "WE_Resids"
For i = 1 To 2
   XVarBox.RemoveItem numfldsd + 1
   YVarBox.RemoveItem numfldsd + 1
Next i
For i = 1 To 2
    XVarBox.AddItem vn$(numfldsd + i + 1), numfldsd + i
    YVarBox.AddItem vn$(numfldsd + i + 1), numfldsd + i
Next i


adddowce:
Screen.MousePointer = 0
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume adddowes
End If
adddowes:
End Sub

Private Sub HistBut_Click()
Dim i As Integer, j As Integer
Dim numpasses As Integer, nmax As Integer, n1, n2
Dim emax$, rootmax, powermax
Dim xinc, xscale, xloc, yinc, yscale, yloc

'error checking
On Error GoTo hbce

'call hist dia box to get num intervals
histcancel = False
HistDB.Show MODAL
If histcancel = True Then GoTo hbce:
Screen.MousePointer = 11
'dims based on hnumints
ReDim NumObs(2, hnumints) As Integer, prob(2, hnumints)

'call SetAll to fill x() and y()
Call SetAll

'sets xmaxr, xminr, xranger
If hsetminmax = True Then 'user sets min and max
   xmaxr = hxmaxset
   xminr = hxminset
Else 'uses overall mins and maxs from setall
   xmaxr = ymax
   xminr = ymin
End If
xranger = xmaxr - xminr

'prints errors
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo hbce:
ElseIf G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined.", , "Error"
   GoTo hbce:
End If

'sets graph type for point find procedure
gtype$ = "HIST"
ZoomCmd.Enabled = False
GraphBox.Cls
GraphBox.Scale (-15, 110)-(115, -20)
GraphBox.FontName = "Helv"
GraphBox.FontSize = 8

'set numpasses
If G1G2But.Value = False Then
   numpasses = 1
Else
   numpasses = 2
End If

For pass = 1 To numpasses
   
   'refills x() and y() if appropriate
   If numpasses = 2 Then
      Call SetGrp
      If pass = 1 Then n1 = n Else n2 = n
   End If
    
   'on each pass, resets arrays which total values
   For j = 1 To hnumints
      NumObs(pass, j) = 0
   Next j
   
   'determines the interval size
   intsize = (xmaxr - xminr) / hnumints
   
   'finds numobs(pass) per bin
   For i = 1 To n
      For j = 1 To hnumints
         xlow = xminr + (j - 1) * intsize
         xhigh = xminr + j * intsize
         Select Case j
            Case 1
               If y(i, 1) >= xlow And y(i, 1) <= xhigh Then
                  NumObs(pass, j) = NumObs(pass, j) + 1
                  Exit For
               End If
            Case Else
               If y(i, 1) > xlow And y(i, 1) <= xhigh Then
                  NumObs(pass, j) = NumObs(pass, j) + 1
                  Exit For
            End If
         End Select
      Next j
   Next i

   'finds prob() of occurrance and overall max obs
   For i = 1 To hnumints
      'finds prob()
      prob(pass, i) = NumObs(pass, i) / n

      'finds overall max obs
      If pass = 1 And i = 1 Then
         nmax = NumObs(pass, i)
      ElseIf NumObs(pass, i) > nmax Then
         nmax = NumObs(pass, i)
      End If
   Next i
Next pass



'rounds nmax up
emax$ = Format$(nmax, "0.0E+00")
rootmax = 1 + Val(Left$(emax$, 1))
powermax = Val(Right(emax$, 3))
nmax = rootmax * 10 ^ powermax

'kkk
If G1G2But.Value = False Then 'single group
   
   'prints results to valuebar
   sf$ = "0.0"
   ValueBar.Cls
   ValueBar.Print "Histogram: N = " + Format$(n, "####") + "   Num Intervals = " + Format$(hnumints, "####")
   For i = 1 To hnumints
      ValueBar.Print "("; i; ","; NumObs(1, i); ","; Format(NumObs(1, i) / n * 100, sf$); "%)  ";
      If i = Int(i / 3) * 3 Then ValueBar.Print
   Next i
      
   'Draws first y axis
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
   GraphBox.Line (0, 0)-(0, 100)
   GraphBox.CurrentX = -13
   GraphBox.CurrentY = 108
   GraphBox.Print "Frequency"
      
   'Draws second y axis
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
   GraphBox.Line (100, 0)-(100, 100)
   GraphBox.CurrentX = 92
   GraphBox.CurrentY = 108
   GraphBox.Print "Probability"
      
   'sets params required to print y scales
   numincs = 5
   fy$ = "###0"
   ymin = 0
   ymax = nmax
   yrange = ymax - ymin
   yinc = yrange / numincs
   
   fy2$ = "0.00"
   y2max = nmax / n
   y2range = y2max - ymin
   y2inc = y2range / numincs
      
   For i = 0 To numincs
      'prints first y scale labels
      yscale = ymin + i * yinc
      yloc = (yscale - ymin) / yrange * 100
      GraphBox.CurrentX = -10
      GraphBox.CurrentY = yloc + 3
      GraphBox.Print Format$(yscale, fy$)
      GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
      GraphBox.Line (0, yloc)-(1, yloc)
      If GridCmd.Checked = -1 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
         GraphBox.Line (0, yloc)-(100, yloc)
      End If
   
      'prints second y scale labels
      y2scale = ymin + i * y2inc
      yloc = (y2scale - ymin) / y2range * 100
      GraphBox.CurrentX = 100.5
      GraphBox.CurrentY = yloc + 3
      GraphBox.Print Format$(y2scale, fy2$)
      GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
      GraphBox.Line (99, yloc)-(100, yloc)
   Next i
      
   'Draws x axis
   GraphBox.Line (0, 0)-(100, 0)
   GraphBox.CurrentX = 90
   GraphBox.CurrentY = -9.5
   GraphBox.Print YVarBox.Text
      
   'finds format of x scale labels
   If Abs(xminr) <= Abs(xmaxr) Then xm = Abs(xmaxr) Else xm = Abs(xminr)
   If xm < 10 Then
      fx$ = "0.00"
   ElseIf xm < 50 Then
      fx$ = "#0.0"
   Else
      fx$ = "###0"
   End If
      
   'prints x scale labels
   For i = 0 To hnumints Step (hnumints / 4)
      xscale = xminr + i * intsize
      xloc = (xscale - xminr) / xranger * 100
      GraphBox.CurrentX = xloc
      GraphBox.CurrentY = -5
      GraphBox.Print Format$(xscale, fx$)
      GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
      GraphBox.Line (xloc, 0)-(xloc, 1)
   Next i
   
   'draws histogram
   predint = 0
   pass = 1
   yminr = 0
   ymaxr = nmax
   yranger = ymaxr - yminr
   
   For i = 1 To hnumints
         'draws left vertical line
         x1 = xminr + (i - 1) * intsize
         y1 = 0
         x2 = x1
         y2 = NumObs(pass, i)
         Call DrawModelLine(predint, pass, x1, y1, x2, y2)
         
         'draw horizontal line
         x1 = xminr + (i - 1) * intsize
         y1 = NumObs(pass, i)
         x2 = xminr + i * intsize
         y2 = NumObs(pass, i)
         Call DrawModelLine(predint, pass, x1, y1, x2, y2)

         'draw right vertical line
         x1 = xminr + i * intsize
         y1 = 0
         x2 = x1
         y2 = NumObs(pass, i)
         Call DrawModelLine(predint, pass, x1, y1, x2, y2)
   Next i

Else 'G1G2 procedure

  'prints results to valuebar
   sf$ = "0.0"
   ValueBar.Cls
   ValueBar.Print "Histogram: N1 = " + Format$(n1, "####") + "   N2 = " + Format$(n2, "####") + "   Num Intervals = " + Format$(hnumints, "####")
   ReDim nn(2)
   nn(1) = n1
   nn(2) = n2
   For i = 1 To hnumints
      For pass = 1 To 2
         ValueBar.Print "("; pass; ","; i; ","; NumObs(pass, i); ","; Format(NumObs(pass, i) / nn(pass) * 100, sf$); "%)  ";
         If pass = 2 Then ValueBar.Print
      Next pass
   Next i

   'Draws first y axis
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
   GraphBox.Line (0, 0)-(0, 100)
   GraphBox.CurrentX = -13
   GraphBox.CurrentY = 108
   GraphBox.Print "Frequency"
   
   'Draws second y axis
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
   GraphBox.Line (100, 0)-(100, 100)
   GraphBox.CurrentX = 92
   GraphBox.CurrentY = 108
   GraphBox.Print "Probability"
   
   'Draws center dividing line and g1/g2 labels
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
   GraphBox.Line (50, 0)-(50, 100)
   GraphBox.CurrentX = 20
   GraphBox.CurrentY = 108
   GraphBox.Print g1name$
   GraphBox.CurrentX = 70
   GraphBox.CurrentY = 108
   GraphBox.Print g2name$
   
   'sets params for subs.g1g2histlines()
   ymaxr = nmax
   yminr = 0
   yranger = ymaxr - yminr

   'sets params required to print y scales
   numincs = 5
   fy$ = "###0"
   ymin = 0
   ymax = nmax
   yrange = ymax - ymin
   yinc = yrange / numincs
   
   fy2$ = ".00"
   y21max = nmax / n1
   y21range = y21max - ymin
   y21inc = y21range / numincs
   y22max = nmax / n2
   y22range = y22max - ymin
   y22inc = y22range / numincs
   
   For i = 0 To numincs
      'prints first y scale labels
      yscale = ymin + i * yinc
      yloc = (yscale - ymin) / yrange * 100
      GraphBox.CurrentX = -10
      GraphBox.CurrentY = yloc + 3
      GraphBox.Print Format$(yscale, fy$)
      GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
      GraphBox.Line (0, yloc)-(1, yloc)
      If GridCmd.Checked = -1 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
         GraphBox.Line (0, yloc)-(100, yloc)
      End If
   
      'prints second y scale labels
      y21scale = ymin + i * y21inc
      y22scale = ymin + i * y22inc
      yloc = (y21scale - ymin) / y21range * 100
      GraphBox.CurrentX = 100.5
      GraphBox.CurrentY = yloc + 3
      GraphBox.Print Format$(y21scale, fy2$) + "/" + Format$(y22scale, fy2$)
      GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
      GraphBox.Line (99, yloc)-(100, yloc)
   Next i
   
   'Draws x axis
   GraphBox.Line (0, 0)-(100, 0)
   GraphBox.CurrentX = 90
   GraphBox.CurrentY = -9.5
   GraphBox.Print YVarBox.Text
   
   'finds format of x scale labels
   If Abs(xminr) <= Abs(xmaxr) Then xm = Abs(xmaxr) Else xm = Abs(xminr)
   If xm < 10 Then
      fx$ = "0.00"
   ElseIf xm < 50 Then
      fx$ = "#0.0"
   Else
      fx$ = "###0"
   End If
   
   For pass = 1 To 2
      'prints x scale labels
      For i = 0 To hnumints Step (hnumints / 4)
         If pass = 1 And i = hnumints Then Exit For
         xscale = xminr + i * intsize
         xloc = (xscale - xminr) / xranger * 50
         If pass = 2 Then xloc = xloc + 50
         GraphBox.CurrentX = xloc
         GraphBox.CurrentY = -5
         GraphBox.Print Format$(xscale, fx$)
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
         GraphBox.Line (xloc, 0)-(xloc, 1)
      Next i
         
      'draws histograms
      For i = 1 To hnumints
         'draws left vertical line
         x1 = xminr + (i - 1) * intsize
         y1 = 0
         x2 = x1
         y2 = NumObs(pass, i)
         Call G1G2HistLines(pass, x1, y1, x2, y2)
         
         'draw horizontal line
         x1 = xminr + (i - 1) * intsize
         y1 = NumObs(pass, i)
         x2 = xminr + i * intsize
         y2 = NumObs(pass, i)
         Call G1G2HistLines(pass, x1, y1, x2, y2)
   
         'draw right vertical line
         x1 = xminr + i * intsize
         y1 = 0
         x2 = x1
         y2 = NumObs(pass, i)
         Call G1G2HistLines(pass, x1, y1, x2, y2)
      Next i
   Next pass
End If

hbce:
Screen.MousePointer = 0
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error"
   Resume hbes:
End If
hbes:
End Sub

Private Sub HourlyProfilesBut_Click()
Dim hr, dY
Dim theta ', sint, cost
Dim i
Dim x1, y1, z1, x2, y2, z2
'assumes only valid x points are the xscreen values 0 to res
Dim xs1 As Integer, ys1, xs2 As Integer, ys2
'new dims for emodel
Dim yvar ', numdays
Dim ckdate$, m As Integer
Dim rootmax, powermax
Dim totregs As Double
On Error GoTo hpce

'sets graphtype
gtype$ = "HP"

'clears screen and makes mouse pointer into an hourglass
GraphBox.Cls
Screen.MousePointer = 11

'show perdone
Screen.MousePointer = 11
PerDone.Caption = "Processing Data..."
PerDone.Show 0
totregs = numrecs
regnum = 0
DoEvents

'finds ymax
yvar = Main.YVarBox.ListIndex + 1
For i = 1 To numrecs
   If i = 1 And d(i, yvar) <> ndflag Then
      ymax = d(i, yvar)
   ElseIf d(i, yvar) > ymax And d(i, yvar) <> ndflag Then
      ymax = d(i, yvar)
   End If
   
  'update perdone
   regnum = regnum + 1
   If Int(regnum / 2) * 2 = regnum Then Call UpdatePerDone(regnum / 2, totregs, 1)
Next i

'rounds ymaxr up to even number
emax$ = Format$(ymax, "0.0E+00")
rootmax = 1 + Val(Left$(emax$, 1))
powermax = Val(Right(emax$, 3))
ymaxr = rootmax * 10 ^ powermax

'reads axis angle from text box
theta = thetad / 360 * 2 * 3.1416
sint = Sin(theta)
cost = Cos(theta)

'reads x-res from text box and uses it to set scale
res = 400 'Int(Val(resTB.Text))
GraphBox.Scale (-0.1 * res, 1.1 * res)-(1.1 * res, -0.1 * res)

'initializes maskmin and maskmax
ReDim maskmax(0 To res), maskmin(0 To res), maskmaxnew(0 To res), maskminnew(0 To res)
For i = 0 To res
   maskminnew(i) = res
   maskmaxnew(i) = 0
Next i

'finds total number of days
If numrecs = Int(numrecs / 24) * 24 Then
   numdays = numrecs / 24
Else
   numdays = Int(numrecs / 24) + 1
End If

'processes data
dY = numdays
For i = 1 To numrecs
   If d(i, hrfld) = 0 Then
      dY = dY - 1
      
      'define the first base point in each daily load trace
      x1 = dY / numdays * res: y1 = 0: z1 = res
      xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
      ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
      
      'draw data lines
      For hr = 0 To 23
         If d(i + hr, yvar) = ndflag Then
            y2 = 0
         Else
            y2 = d(i + hr, yvar) / ymaxr * res
         End If
         x2 = dY / numdays * res
         z2 = res - (hr / 23 * res)
         xs2 = (x2 - (z2 * cost) + (res * cost)) / (cost + 1)
         ys2 = (y2 - (z2 * sint) + (res * sint)) / (sint + 1)
         'graphbox.Line (xs1, ys1)-(xs2, ys2)
         Call D3Line(xs1, ys1, xs2, ys2, x1)
         xs1 = xs2
         ys1 = ys2
      Next hr
   
      'define the second base point in each daily load trace
      x2 = dY / numdays * res
      y2 = 0
      z2 = 0
      xs2 = (x2 - (z2 * cost) + (res * cost)) / (cost + 1)
      ys2 = (y2 - (z2 * sint) + (res * sint)) / (sint + 1)
      'graphbox.Line (xs1, ys1)-(xs2, ys2)
      Call D3Line(xs1, ys1, xs2, ys2, x1)
      
      'connect base points
      xs1 = xs2
      ys1 = ys2
      x2 = dY / numdays * res
      y2 = 0
      z2 = res
      xs2 = (x2 - (z2 * cost) + (res * cost)) / (cost + 1)
      ys2 = (y2 - (z2 * sint) + (res * sint)) / (sint + 1)
      'graphbox.Line (xs1, ys1)-(xs2, ys2)
      Call D3Line(xs1, ys1, xs2, ys2, x1)


      'draw vertical hour lines
      'currently disabled, can add as option
      If 1 = 2 Then
         For hr = 0 To 23 Step 6
            'define base point
            x1 = dY / numdays * res: y1 = 0: z1 = res * (1 - hr / 23)
            xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
            ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
            
            If d(i + hr, yvar) = ndflag Then
               y2 = 0
            Else
               y2 = d(i + hr, yvar) / ymaxr * res
            End If
            x2 = dY / numdays * res
            z2 = res - (hr / 23 * res)
            xs2 = (x2 - (z2 * cost) + (res * cost)) / (cost + 1)
            ys2 = (y2 - (z2 * sint) + (res * sint)) / (sint + 1)
            'GraphBox.Line (xs1, ys1)-(xs2, ys2)
            Call D3Line(xs1, ys1, xs2, ys2, x1)
         Next hr
      End If


   End If
  
  'update perdone
   regnum = regnum + 1
   If Int(regnum / 2) * 2 = regnum Then Call UpdatePerDone(regnum / 2, totregs, 1)
   
Next i

'draws cursor
dY = numdays
For i = 1 To numrecs
   If d(i, hrfld) = 0 Then
      dY = dY - 1
      
      currec = dY

      'define the first base point in each daily load trace
      x1 = dY / numdays * res: y1 = 0: z1 = res
      xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
      ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)

      GraphBox.Line (xs1 - 10, ys1 + 10)-(xs1 + 10, ys1 - 10), BLUE
      GraphBox.Line (xs1 - 10, ys1 - 10)-(xs1 + 10, ys1 + 10), BLUE
      Exit For
   End If
   
Next i
dy1 = dY

'draws x axis
x1 = 0: y1 = 0: z1 = 0: x2 = res: y2 = 0: z2 = 0
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
xs2 = (x2 - (z2 * cost) + (res * cost)) / (cost + 1)
ys2 = (y2 - (z2 * sint) + (res * sint)) / (sint + 1)
'graphbox.Line (xs1, ys1)-(xs2, ys2)
Call D3Line(xs1, ys1, xs2, ys2, x1)
'draws y axis
x1 = 0: y1 = 0: z1 = 0: x2 = 0: y2 = res: z2 = 0
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
xs2 = (x2 - (z2 * cost) + (res * cost)) / (cost + 1)
ys2 = (y2 - (z2 * sint) + (res * sint)) / (sint + 1)
'graphbox.Line (xs1, ys1)-(xs2, ys2)
Call D3Line(xs1, ys1, xs2, ys2, x1)
'draws z axis
x1 = 0: y1 = 0: z1 = 0: x2 = 0: y2 = 0: z2 = res
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
xs2 = (x2 - (z2 * cost) + (res * cost)) / (cost + 1)
ys2 = (y2 - (z2 * sint) + (res * sint)) / (sint + 1)
'graphbox.Line (xs1, ys1)-(xs2, ys2)
Call D3Line(xs1, ys1, xs2, ys2, x1)

'draws 23 hour label
x1 = res + 3: y1 = 0: z1 = 0
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
GraphBox.CurrentX = xs1: GraphBox.CurrentY = ys1
GraphBox.Print "23"
'draws 18 hour label
x1 = res + 3: y1 = 0: z1 = res / 4
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
GraphBox.CurrentX = xs1: GraphBox.CurrentY = ys1
GraphBox.Print "18"
x1 = (numdays - 1) / numdays * res: y1 = 0: z1 = res * (1 - (18 / 23))
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
x2 = ((numdays - 1) / numdays * res) + 15: y1 = 0: z1 = res * (1 - (18 / 23))
xs2 = (x2 - (z1 * cost) + (res * cost)) / (cost + 1)
GraphBox.Line (xs1, ys1)-(xs2, ys1)
'draws 12 hour label
x1 = res + 3: y1 = 0: z1 = res / 2
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
GraphBox.CurrentX = xs1: GraphBox.CurrentY = ys1
GraphBox.Print "12"
x1 = (numdays - 1) / numdays * res: y1 = 0: z1 = res * (1 - (12 / 23))
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
x2 = ((numdays - 1) / numdays * res) + 15: y1 = 0: z1 = res * (1 - (12 / 23))
xs2 = (x2 - (z1 * cost) + (res * cost)) / (cost + 1)
GraphBox.Line (xs1, ys1)-(xs2, ys1)
'draws 6 hour label
x1 = res + 3: y1 = 0: z1 = res * 0.75
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
GraphBox.CurrentX = xs1: GraphBox.CurrentY = ys1
GraphBox.Print "6"
x1 = (numdays - 1) / numdays * res: y1 = 0: z1 = res * (1 - (6 / 23))
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
x2 = ((numdays - 1) / numdays * res) + 15: y1 = 0: z1 = res * (1 - (6 / 23))
xs2 = (x2 - (z1 * cost) + (res * cost)) / (cost + 1)
GraphBox.Line (xs1, ys1)-(xs2, ys1)
'draw ymax line
x1 = 0: y1 = res: z1 = 0
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
x1 = 15: y1 = res: z1 = 0
xs2 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys2 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
GraphBox.Line (xs1, ys1)-(xs2, ys2)
'draws variable label
x1 = 2: y1 = res + 20: z1 = 0
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
GraphBox.CurrentX = xs1: GraphBox.CurrentY = ys1
GraphBox.Print YVarBox.Text + " = " + Format(ymaxr)

'draws startdate label
x1 = -4: y1 = -10: z1 = res
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
GraphBox.CurrentX = xs1: GraphBox.CurrentY = ys1
For i = 1 To numrecs
   If d(i, hrfld) = 0 Then m = i
Next i
timeint = 3
Call CalendarDate(m, ckdate$)
timeint = 4
GraphBox.Print ckdate$
'draws enddate label
x1 = res: y1 = -10: z1 = res
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
GraphBox.CurrentX = xs1: GraphBox.CurrentY = ys1
For i = 1 To numrecs
   If d(i, hrfld) = 0 Then
      m = i
      Exit For
   End If
Next i
timeint = 3
Call CalendarDate(m, ckdate$)
timeint = 4
GraphBox.Print ckdate$

'set focus so moving cursor with keyboard will work
GraphBox.SetFocus

hpce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error"
   Resume hpes
End If
hpes:
End Sub

Private Sub HourlyProfilesRotCmd_Click()

RotateDB.Show 0

End Sub

Private Sub MeanBut_Click()
Dim i As Integer, count As Integer, numpasses As Integer
Dim ytot, ssm, std, cvstd, t
On Error GoTo mbce

Screen.MousePointer = 11
Call SetAll

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo mbce:
ElseIf G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined.", , "Error"
   GoTo mbce:
End If

modtype$ = "Mean"

'sets pass
If G1G2But.Value = True And gmodnum = 2 Then pass = 2 Else pass = 1
    
'refills x() and y() if appropriate
If G1G2But.Value = True Then Call SetGrp

'calcs y
ytot = 0
For i = 1 To n
   ytot = ytot + y(i, 1)
Next i
ymean = ytot / n

'calcs std and cvstd
ssm = 0
For i = 1 To n
   ssm = ssm + (y(i, 1) - ymean) ^ 2
Next i
std = (ssm / (n - 1)) ^ 0.5
cvstd = std / ymean * 100

If pass = 1 Then
   ymeanpre = ymean
   premodtype$ = modtype$
   rmsepre = std
   npre = n
   modelstats = " N = " + Format$(n, "####") + "   Std Dev = " + Format$(std, "#,##0.00") + "   CV-Std Dev = " + Format$(cvstd, "##0.00") + "%"
   modeleqn = vn$(yfld) + " = " + Format$(ymean, "#,##0.00")
End If


'calcs prediction interval bounds
t = 1.96 + 2.7 / (n - 1)
predint = t * std

'prints results
sf$ = "###0.0000"
If G1G2But.Value = False Or (G1G2But.Value = True And pass = 1) Then ValueBar.Cls
Call PrintHeaderLine(gmodnum, modtype$)
ValueBar.Print " N = " + Format$(n, "####") + "   Ymean = " + Format$(ymean, "#,##0.00") + "   Std Dev = " + Format$(std, sf$) + "   CV-Std Dev = " + Format$(cvstd, "##0.00") + "%"
ValueBar.Print

'calcs lines coordinates
x1 = xminr: y1 = ymean
x2 = xmaxr: y2 = ymean

Call DrawModelLine(predint, pass, x1, y1, x2, y2)
    
If G1G2But.Value = True And pass = 2 Then Call SetAll

'resets gmodnum
If G1G2But.Value = True Then
   If gmodnum = 1 Then
      gmodnum = 2
   Else 'gmodnum = 2
      gmodnum = 1
   End If
End If
   
mbce:
Screen.MousePointer = 0
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume mbes
End If
mbes:

End Sub

Private Sub MedianBut_Click()
Dim t, i As Integer, numpasses As Integer
Dim seymedian, cvymedian
On Error GoTo medbce

Screen.MousePointer = 11
Call SetAll

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo medbce:
ElseIf G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined.", , "Error"
   GoTo medbce:
End If

modtype$ = "Median"
    
'sets pass
If G1G2But.Value = True And gmodnum = 2 Then pass = 2 Else pass = 1


'refills x() and y() if appropriate
If G1G2But.Value = True Then Call SetGrp

'calcs median, semedian, cvymedian
Call Median(n, y(), ymedian, seymedian)
cvymedian = seymedian / ymedian * 100
If pass = 1 Then
   premodtype$ = modtype$
   rmsepre = seymedian
   npre = n
   ymedianpre = ymedian
   modelstats = " N = " + Format$(n, "####") + "   Std Err = " + Format$(seymedian, "#,##0.00") + "   CV-Std Err = " + Format$(cvymedian, "##0.00") + "%"
   modeleqn = vn$(yfld) + " = " + Format$(ymedian, "#,##0.00")
End If

'calcs prediction interval bounds
t = 1.96 + 2.7 / (n - 1)
predint = t * seymedian

'prints results
sf$ = "###0.0000"
If G1G2But.Value = False Or (G1G2But.Value = True And pass = 1) Then ValueBar.Cls
Call PrintHeaderLine(gmodnum, modtype$)
ValueBar.Print " N = " + Format$(n, "####") + "   Ymedian = " + Format$(ymedian, "###0.00") + "   Std Err = " + Format$(seymedian, sf$) + "   CV-Std Err = " + Format$(cvymedian, "##0.00") + "%"
ValueBar.Print

'calcs lines coordinates
x1 = xminr: y1 = ymedian
x2 = xmaxr: y2 = ymedian

Call DrawModelLine(predint, pass, x1, y1, x2, y2)
    
If G1G2But.Value = True Then
   If gmodnum = 1 Then
      gmodnum = 2
   Else 'gmodnum = 2
      gmodnum = 1
   End If
End If


medbce:
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error"
   Resume medbes
End If
medbes:
Screen.MousePointer = 0
End Sub

Private Sub MergeCmd_Click()
Dim i, j, startrec, startdate, enddate, dydate, numnewflds
Dim startdate1, enddate1, timeperrec1, startdate2, enddate2, timeperrec2
Dim tnpdy$, tfndy$, numrecsdy, numfldsdy, numfldsddy, filenamedy$, ndflagdy, yrflddy, moflddy, dyflddy, hrflddy, timeintdy
Dim tnpmo$, tfnmo$, numrecsmo, numfldsmo, numfldsdmo, filenamemo$, ndflagmo, yrfldmo, mofldmo, dyfldmo, hrfldmo, timeintmo
On Error GoTo mergece
'this subroutine integrates two data sets into one

'in this subroutine we identify variables and arrays as "monthly"
'or "daily", but this is simply a nuemonic for clarity.  the subroutine can
'merge data with any timescales.

'show perdone
Dim regnum, totregs, increment
increment = 10
If numrecs > numrecs2 Then
   totregs = 2 * numrecs + 3 * numrecs2
Else
   totregs = 3 * numrecs + 2 * numrecs2
End If
Screen.MousePointer = 11
PerDone.Caption = "Processing Data..."
PerDone.Show 0
DoEvents

'because merge always merges high freq data such as "daily" data into
'low freq data such as "monthly" data, this finds
'which of d() or d2()is the daily and which is the monthly matrix
startdate1 = DateOnly("d", 1)
enddate1 = DateOnly("d", numrecs)
timeperrec1 = (CDate(enddate1) - CDate(startdate1)) / numrecs
startdate2 = DateOnly("d2", 1)
enddate2 = DateOnly("d2", numrecs2)
timeperrec2 = (CDate(enddate2) - CDate(startdate2)) / numrecs2

'fills ddy(), dmo() and sets dy and mo parameters
If timeperrec2 >= timeperrec1 Then
   
   ReDim ddy(numrecs, numfldsd), grpddy(numrecs), vndy$(numfldsd)
   For i = 1 To numrecs
      For j = 1 To numfldsd
         ddy(i, j) = d(i, j)
      Next j
      grpddy(i) = grpd(i)
      
      'update perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, increment)
   Next i
   For i = 1 To numfldsd
      vndy$(i) = vn$(i)
   Next i
   
   ReDim dmo(numrecs2, numfldsd2), grpdmo(numrecs2), vnmo$(numfldsd2)
   For i = 1 To numrecs2
      For j = 1 To numfldsd2
         dmo(i, j) = d2(i, j)
      Next j
      grpdmo(i) = grpd2(i)
   
      'update perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, increment)
   Next i
   For i = 1 To numfldsd2
      vnmo$(i) = vn2$(i)
   Next i
   
   tnpdy$ = tnp$
   tfndy$ = tfn$
   numrecsdy = numrecs
   numfldsdy = numflds
   numfldsddy = numfldsd
   filenamedy$ = filename$
   ndflagdy = ndflag
   yrflddy = yrfld
   moflddy = mofld
   dyflddy = dyfld
   hrflddy = hrfld
   timeintdy = timeint
   
   tnpmo$ = tnp2$
   tfnmo$ = tfn2$
   numrecsmo = numrecs2
   numfldsmo = numflds2
   numfldsdmo = numfldsd2
   filenamemo$ = filename2$
   ndflagmo = ndflag2
   yrfldmo = yrfld2
   mofldmo = mofld2
   dyfldmo = dyfld2
   hrfldmo = hrfld2
   timeintmo = timeint2

Else
   
   ReDim ddy(numrecs2, numfldsd2), grpddy(numrecs2), vndy$(numfldsd2)
   For i = 1 To numrecs2
      For j = 1 To numfldsd2
         ddy(i, j) = d2(i, j)
      Next j
      grpddy(i) = grpd2(i)
      
     'update perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, increment)
   Next i
   For i = 1 To numfldsd2
      vndy$(i) = vn2$(i)
   Next i
         
   ReDim dmo(numrecs, numfldsd), grpdmo(numrecs), vnmo$(numfldsd)
   For i = 1 To numrecs
      For j = 1 To numfldsd
         dmo(i, j) = d(i, j)
      Next j
      grpdmo(i) = grpd(i)
      
      'update perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, increment)
   Next i
   For i = 1 To numfldsd
      vnmo$(i) = vn$(i)
   Next i
   
   tnpdy$ = tnp2$
   tfndy$ = tfn2$
   numrecsdy = numrecs2
   numfldsdy = numflds2
   numfldsddy = numfldsd2
   filenamedy$ = filename2$
   ndflagdy = ndflag2
   yrflddy = yrfld2
   moflddy = mofld2
   dyflddy = dyfld2
   hrflddy = hrfld2
   timeintdy = timeint2
   
   tnpmo$ = tnp$
   tfnmo$ = tfn$
   numrecsmo = numrecs
   numfldsmo = numflds
   numfldsdmo = numfldsd
   filenamemo$ = filename$
   ndflagmo = ndflag
   yrfldmo = yrfld
   mofldmo = mofld
   dyfldmo = dyfld
   hrfldmo = hrfld
   timeintmo = timeint

End If

'dims temp arrays
ReDim numdata(numfldsddy), davg(numrecsmo, numfldsddy)

'fills davg() with ndflags
For i = 1 To numrecsmo
   For j = 1 To numfldsddy
      davg(i, j) = ndflagmo
   Next j
   
  'update perdone
   regnum = regnum + 1
   Call UpdatePerDone(regnum, totregs, increment)
Next i

startrec = 1
'integrates data in d() into data in dmo()
For i = 2 To numrecsmo
   'finds start and end date for each period in monthly file
   startdate = CDate(Format(dmo(i - 1, mofldmo)) + "/" + Format(dmo(i - 1, dyfldmo)) + "/" + Format(dmo(i - 1, yrfldmo)))
   enddate = CDate(Format(dmo(i, mofldmo)) + "/" + Format(dmo(i, dyfldmo)) + "/" + Format(dmo(i, yrfldmo)))
    
   'inititializes numdata(k)
   For k = 1 To numfldsddy
      numdata(k) = 0
   Next k
   
   'finds average temperature in energy period
   For j = startrec To numrecsdy
      dydate = CDate(Format(ddy(j, moflddy)) + "/" + Format(ddy(j, dyflddy)) + "/" + Format(ddy(j, yrflddy)))
      
      If dydate > startdate And dydate <= enddate Then
         
         'finds average
         For k = 1 To numfldsddy
            If k <> moflddy And k <> dyflddy And k <> hrflddy And k <> yrflddy And ddy(j, k) <> ndflagdy Then
               numdata(k) = numdata(k) + 1
               davg(i, k) = (ddy(j, k) + davg(i, k) * (numdata(k) - 1)) / numdata(k)
            End If
         Next k
         
         'set new startrec
         startrec = j
      End If
      
      'update perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, increment)
      
      If dydate > enddate Then Exit For
   Next j
   
   
Next i
   
'refills d()
ReDim d(numrecsmo, numfldsdmo + numfldsddy - timeintdy + 3)
For i = 1 To numrecsmo
   'fills d() with contents of dmo()
   For j = 1 To numfldsdmo
      d(i, j) = dmo(i, j)
   Next j
   'fills d() with averaged fields from d()
   numnewflds = 0
   For k = 1 To numfldsddy
      If k <> moflddy And k <> dyflddy And k <> hrflddy And k <> yrflddy Then
         numnewflds = numnewflds + 1
         d(i, numfldsdmo + numnewflds) = davg(i, k)
      End If
   Next k
   'fills grouping array
   grpd(i) = grpdmo(i)
   
   'update perdone
   regnum = regnum + 1
   Call UpdatePerDone(regnum, totregs, increment)
Next i

'refills vn$()
ReDim vnnew$(numfldsddy)
'saves non time stamp var names
numnewflds = 0
For k = 1 To numfldsddy
   If k <> moflddy And k <> dyflddy And k <> hrflddy And k <> yrflddy Then
      numnewflds = numnewflds + 1
      vnnew$(numnewflds) = vndy$(k)
   End If
Next k
ReDim vn$(numfldsdmo + numfldsddy - timeintdy + 3)
'adds var names from dmo to vn$()
For k = 1 To numfldsdmo
   vn$(k) = vnmo$(k)
Next k
'adds var names from d() to vn$()
For k = 1 To numnewflds
   vn$(k + numfldsdmo) = vnnew$(k)
Next k
vn$(numfldsdmo + numfldsddy - timeintdy + 1) = "SG_Resids"
vn$(numfldsdmo + numfldsddy - timeintdy + 2) = g1name$ + "_Resids"
vn$(numfldsdmo + numfldsddy - timeintdy + 3) = g2name$ + "_Resids"

'resets all global variables
tnp$ = tnpmo$
tfn$ = tfnmo$
numrecs = numrecsmo
numflds = numfldsdmo + numfldsddy - timeintdy + 3
numfldsd = numfldsdmo + numfldsddy - timeintdy
filename$ = filenamemo$
ndflag = ndflagmo
yrfld = yrfldmo
mofld = mofldmo
dyfld = dyfldmo
hrfld = hrfldmo
timeint = timeintmo

'clears arrays to save memory
ReDim ddy(1, 1), grpddy(1), vndy$(1), dmo(1, 1), grpdmo(1), vnmo$(1)
ReDim d2(1, 1), grpd2(1), vn2$(1), numdata(1), davg(1, 1), vnnew$(1)

'refills variable name boxes
getvnsfrommerge = True
Call OpenVar("")
getvnsfrommerge = False

'disables merge command until another data file is opened
Main.MergeCmd.Enabled = False

mergece:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume mergees
End If
mergees:
End Sub

Private Sub MLRBut_Click()
Dim i, j As Integer
Dim k, numpasses, count As Integer
Dim xpnt, ypnt
On Error GoTo mlrce

If Main.G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined."
   GoTo mlrce:
End If

'shows Mult Reg dialog box
MultRgDB.Show 1
If cancelmlr = True Then GoTo mlrce:
Screen.MousePointer = 11
'counts numxvars
numxvars = 1
Do While xfld(numxvars) <> 0
   numxvars = numxvars + 1
Loop
numxvars = numxvars - 1
'displays error message if no vars are selected
If numxvars = 0 Or yfld = 0 Then
   MsgBox "No variables have been selected.  Remember to press the 'Set' button after each variable is selected.", , "Error"
   GoTo mlrce:
End If

modtype$ = "MLR"

'sets pass
If G1G2But.Value = True And gmodnum = 2 Then pass = 2 Else pass = 1

'fills X(), Y(), and sets n
ReDim x(numrecs, numxvars + 1), y(numrecs, 1)
Call FillXYGrp(pass, numxvars, x(), y(), n)

'sets xrows and columns
xr = n
xc = numxvars + 1

'prints error if n = 0
If n = 0 Then
    MsgBox "No data is available to model.", , "Error"
    GoTo mlrce:
End If
 
'calls Regression routine
Call Reg(x(), xr, xc, y(), n, 1, beta(), betar, betac, sec())
Call Inf(x(), xr, xc, y(), n, 1, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), P)
 
'creates global variables for savings calculations
If pass = 1 Then
   premodtype$ = modtype$
   rmsepre = rmse
   npre = n
   numxvarspre = numxvars
   For i = 1 To numxvarspre + 1
      betapre(i) = beta(i, 1)
   Next i
   modelstats = " N = " + Format$(n, "####") + "  R2 = " + Format$(r2, "0.00") + "  RMSE = " + Format$(rmse, "#,##0.00") + "  CV-RMSE = " + Format$(cvrmse, "##0.00") + "%"
   modeleqn = vn$(yfld) + " = " + Format(beta(1, 1), "#,##0.00") + "  +  " + Format(beta(2, 1), "#,##0.00") + " " + vn$(xfld(1))
   For i = 2 To numxvars
      modeleqn = modeleqn + " + " + Format(beta(i + 1, 1), "#,##0.00") + " " + vn$(xfld(i))
   Next i
End If

'prints results
'Call DWTest(3, dwr$)
sf$ = "###0.0000"
If G1G2But.Value = False Or (G1G2But.Value = True And pass = 1) Then ValueBar.Cls
Call PrintHeaderLine(gmodnum, modtype$)
Main.ValueBar.Print " N = " + Format$(n, "####") + "     R2 = " + Format$(r2, "0.00") + "    RMSE = " + Format$(rmse, "###0.00") + "    CV-RMSE = " + Format$(cvrmse, "##0.0\%")
'prints coefficients
Main.ValueBar.Print " Yint = " + Format$(beta(1, 1), sf$) + " (" + Format$(sigma(1), sf$) + ")";
Main.ValueBar.Print "    " + vn$(xfld(1)) + " = " + Format$(beta(2, 1), sf$) + " (" + Format$(sigma(1 + 1), sf$) + ")";
For i = 2 To numxvars
    Main.ValueBar.Print "   " + vn$(xfld(i)) + " = " + Format$(beta(i + 1, 1), sf$) + " (" + Format$(sigma(i + 1), sf$) + ")";
    If numxvars > 3 And i = 3 Then Main.ValueBar.Print
    'iodd% = (i - 1) / 2
    'If iodd% = (i - 1) / 2 Then Main.ValueBar.Print
Next i
Main.ValueBar.Print
'ieven% = numxvars / 2
'If ieven% = numxvars / 2 Then Main.ValueBar.Print
'Main.ValueBar.Print " p = " + Format$(autocor, "0.00") + "   DW = " + Format$(dw, "0.00") + " (" + dwr$ + ")"
'print model
ValueBar.Print "Model: "; vn$(yfld); " = "; Format(beta(1, 1), "#,##0.00"); "  +  "; Format(beta(2, 1), "#,##0.00"); " "; vn$(xfld(1));
For i = 2 To numxvars
   ValueBar.Print " + "; Format(beta(i + 1, 1), "#,##0.00"); " "; vn$(xfld(i)); '" X" + Format(i, "0"); '
Next i
ValueBar.Print
If numxvars < 4 Then ValueBar.Print


'mlr graphing
ReDim ypred(xr), xpred(xr)
'saves values of xpred and ypred before they get changed by the call to xybut
For i = 1 To xr
   ypred(i) = yhat(i, 1)
   xpred(i) = x(i, 2)
Next i

'replot xy graph of y and first x var in mult regression
If G1G2But.Value = False Or (G1G2But.Value = True And pass = 1) Or (G1G2But.Value = True And pass = 2 And Main.XVarBox.ListIndex + 1 <> xfld(1)) Then
   'selects the first x variable in mult reg as the x variable for xy plot
   Main.XVarBox.ListIndex = xfld(1) - 1
   'selects the first y variable in mult reg as the y variable for xy plot
   Main.YVarBox.ListIndex = yfld - 1
   'calls xy plot
   Call XYBut_Click
End If

'draws mlr points
GraphBox.DrawWidth = 3
For i = 1 To xr
   ypnt = (ypred(i) - yminr) / yranger * 100
   xpnt = (xpred(i) - xminr) / xranger * 100
   
   If xpnt >= 0 And xpnt <= 100 And ypnt >= 0 And ypnt <= 100 Then
      If G1G2But.Value = False Then
         GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), BLACK, B
      Else
         If pass = 1 Then
            GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), BLUE, B
         Else
            GraphBox.Circle (xpnt, ypnt), 0.8, RED
         End If
      End If
   End If
Next i
ReDim ypred(1), xpred(1)
GraphBox.DrawWidth = 1

'If numpasses = 2 Then Call SetAll
If G1G2But.Value = True And pass = 2 Then Call SetAll

'resets gmodnum
If G1G2But.Value = True Then
   If gmodnum = 1 Then
      gmodnum = 2
   Else 'gmodnum = 2
      gmodnum = 1
   End If
End If

mlrce:
Screen.MousePointer = 0
If Err = 6 Or Err = 9 Or Err = 11 Then
   MsgBox "Try reducing the value of a variable by scaling it using the Data, Modify Values menu items.", , "Overflow Error"
   Resume mlres
ElseIf Err Then
   MsgBox """" + Error(Err) + """", , "Error in MLR Procedure"
   Resume mlres
End If

mlres:
End Sub

Private Sub ModValuesCmd_Click()

ModValDB.Show 1

End Sub

Private Sub OpenCmd_Click()
Dim row As Integer, col As Integer, i As Long

fopen = 1

'display dialog box and get info
FileOpen.Show MODAL

Screen.MousePointer = 11

'gets filename$ from dialog box
filename$ = tnp$ + tfn$
If filename$ = "" Then GoTo oces:


'initialize sdd() and edd()
numbdate = 0
For i = 1 To 20
   SDD(i) = 0
   EDD(i) = 0
Next i

'finds numflds
'chr$(32) is a space
numflds = 0
Open filename$ For Input As #1
Line Input #1, l$
l$ = LTrim$(RTrim$(l$))
numchars = Len(l$)
For i = 1 To numchars
   Char$ = Mid$(l$, i, 1)
   Select Case Char$
      Case Chr$(32)
        If lc$ = "C" Then
           numflds = numflds + 1
           lc$ = "S"
        End If
      Case Else
         lc$ = "C"
   End Select
Next i
numfldsd = numflds + 1
numflds = numfldsd + 3

'finds numrecs and dimensions D()
numrecs = 1
While Not EOF(1)
    numrecs = numrecs + 1
    Line Input #1, l$
Wend
Close #1
ReDim d(numrecs, numflds)

'fills D()
Open filename$ For Input As #1
While Not EOF(1)
    row = row + 1
    For col = 1 To numfldsd
        Input #1, d(row, col)
    Next col
Wend
Close #1

'displays file name and numrecs in listbox
filename$ = tfn$
ValueBar.Cls
ValueBar.Print " " + UCase$(filename$) + "     Number Observations = " + Format$(numrecs, "####") + "     Number Columns = " + Format$(numfldsd, "###")

'calls variable name cmd
Screen.MousePointer = 0
ReDim vn$(numflds)
Call VarNamesCmd_Click
VarNamesCmd.Enabled = True
SaveLogCmd.Enabled = True

oces:
Screen.MousePointer = 0
End Sub


Private Sub OpenDataOldCmd_Click()
Dim row As Integer, col As Integer, i As Long
On Error GoTo odatcce

fopen = 1

'display dialog box and get info
opencode = "*.dat"
FileOpen.Show MODAL

'gets filename$ from dialog box
filename$ = tnp$ + tfn$
If filename$ = "" Then GoTo odatcce
Screen.MousePointer = 11

'sets default no data flag
ndflag = -99

'initialize sdd() and edd()
numbdate = 0
For i = 1 To 20
   SDD(i) = 0
   EDD(i) = 0
Next i

'finds numflds
'chr$(32) is a space
numflds = 0
Open filename$ For Input As #1
Line Input #1, l$
l$ = LTrim$(RTrim$(l$))
numchars = Len(l$)
For i = 1 To numchars
   Char$ = Mid$(l$, i, 1)
   Select Case Char$
      Case Chr$(32)
        If lc$ = "C" Then
           numflds = numflds + 1
           lc$ = "S"
        End If
      Case Else
         lc$ = "C"
   End Select
Next i
numfldsd = numflds + 1
numflds = numfldsd + 3

'finds numrecs and dimensions D()
numrecs = 1
While Not EOF(1)
    numrecs = numrecs + 1
    Line Input #1, l$
Wend
Close #1
ReDim d(numrecs, numflds)
ReDim grpd(numrecs)

'fills D()
Open filename$ For Input As #1
While Not EOF(1)
    row = row + 1
    For col = 1 To numfldsd
        Input #1, d(row, col)
    Next col
Wend
Close #1

'displays file name and numrecs in listbox
filename$ = tfn$
ValueBar.Cls
ValueBar.Print " " + UCase$(filename$) + "     Number Observations = " + Format$(numrecs, "####") + "     Number Columns = " + Format$(numfldsd, "###")

'calls variable name cmd
Screen.MousePointer = 0
ReDim vn$(numflds)
'Call VarNamesCmd_Click

For i = 1 To numfldsd
   vn$(i) = "V" + Format$(i, "00")
Next i
Do While Main.XVarBox.ListCount > 0
   XVarBox.RemoveItem 0
   YVarBox.RemoveItem 0
Loop
For i = 1 To numfldsd
   'fills variable boxes with new names
   XVarBox.AddItem vn$(i)
   YVarBox.AddItem vn$(i)
Next i

'defines resid variables
vn$(numfldsd + 1) = "SG_Resids"
vn$(numfldsd + 2) = g1name$ + "_Resids"
vn$(numfldsd + 3) = g2name$ + "_Resids"
For i = 1 To 3
    XVarBox.AddItem vn$(numfldsd + i)
    YVarBox.AddItem vn$(numfldsd + i)
Next i
'fills resid variable with ndflags
For i = 1 To numrecs
   For j = numfldsd + 1 To numfldsd + 3
      d(i, j) = ndflag
   Next j
Next i

'reads *.var if it can be found
tfn = Trim$(tfn$)
For i = 1 To Len(tfn$)
   If Mid$(tfn$, i, 1) = "." Then Exit For
Next i
For j = 0 To FileOpen.File1.ListCount
   If FileOpen.File1.List(j) = Left$(tfn$, i - 1) + ".var" Then
      Call OpenVar(tnp$ + Left$(tfn$, i - 1) + ".var")
   End If
Next j

VarNamesCmd.Enabled = True
SaveLogCmd.Enabled = True


odatcce:
Screen.MousePointer = 0
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error"
   Resume odatces
End If
odatces:
End Sub

Private Sub OpenDailyCmd_Click()

End Sub

Private Sub OpenDataCmd_Click()
Dim row As Integer, col As Integer, i As Long
Dim varfilefound
On Error GoTo od2ce

fopen = 1

If numfiles > 0 Then Call SaveFileParams

If Command = "" Then 'normal opening procedure
   'display dialog box and get info
   opencode = "*.*"
   FileOpen.Show MODAL
   tnp$ = LCase$(tnp$)
   tfn$ = LCase$(tfn$)
   
   'gets filename$ from dialog box
   filename$ = tnp$ + tfn$
   If filename$ = "" Then GoTo od2ce

Else 'opening via a command line string
   filename$ = Command
   tnp$ = ""
   For i = 1 To Len(filename$)
      If Mid$(filename$, i, 1) = "\" Then
         'c:\emodel17\p5test.dat
         tnp$ = Left$(filename$, i)
         tfn$ = Right$(filename$, Len(filename$) - i)
      End If
   Next i
   If tnp$ = "" Then
      tfn$ = filename$
   End If
End If

Screen.MousePointer = 11

'sets default no data flag
ndflag = -99

'initialize sdd() and edd()
numbdate = 0
For i = 1 To 20
   SDD(i) = 0
   EDD(i) = 0
Next i

'finds delimiter type, numflds, numrecs
numrecs = 0
numflds = 0
Open filename$ For Input As #1
While Not EOF(1)
   'reads a line
   Line Input #1, l$
   l$ = Trim$(l$)
   If Len(l$) > 0 Then
      numrecs = numrecs + 1
   
      If numrecs = 1 Then
         
         'determines if comma, tab (chr(9)), or space (chr(32)) delimited
         delim = ""
         For charnum = 1 To Len(l$)
            If Mid$(l$, charnum, 1) = "," Then
               delim = ","
               Exit For
            End If
         Next charnum
         If delim <> "," Then
            For charnum = 1 To Len(l$)
               If Mid$(l$, charnum, 1) = Chr(9) Then
                  delim = Chr(9)
                  Exit For
               End If
            Next charnum
         End If
         If delim = "" Then delim = Chr(32)
   
         'determines numflds and dims field()
         numflds = 0
         If delim = "," Or delim = Chr(9) Then
            For charnum = 1 To Len(l$)
               If Mid$(l$, charnum, 1) = delim Or charnum = Len(l$) Then
                  numflds = numflds + 1
               End If
            Next charnum
         Else 'delim = " "
            For charnum = 1 To Len(l$)
               If Mid$(l$, charnum, 1) = delim Then
                  If Mid$(l$, charnum - 1, 1) <> delim Then numflds = numflds + 1
               ElseIf charnum = Len(l$) Then
                  numflds = numflds + 1
               End If
            Next charnum
         End If
      End If
   End If
Wend
Close
numfldsd = numflds
numflds = numflds + 3
ReDim d(numrecs, numflds)
ReDim grpd(numrecs)

'fills d(row,col)
'If delim = Chr(32) Then
'   'fills d(row,col) if space delimited.
'   'Even though the next procedure works with space delimited data,
'   'I put this in to increase reading speed for space delimited data.
'   Open filename$ For Input As #1
'   For i = 1 To numrecs
'      For j = 1 To numfldsd
'         Input #1, d(i, j)
'      Next j
'   Next i
'Else
   'fills d(row,col) if tab, comma or space delim
   'use this method to fill d() since this method filters empty header and footer lines and the 'space only' method commented out above does not.
   row = 0
   Open filename$ For Input As #1
   While Not EOF(1)
      'reads a line
      Line Input #1, l$
      l$ = Trim$(l$)
      If Len(l$) > 0 Then
         row = row + 1
         'parses the line into fields
         scharnum = 1
         fnum = 0
         For charnum = 1 To Len(l$)
            If Mid$(l$, charnum, 1) = delim Then
               If Mid$(l$, charnum - 1, 1) <> delim Then
                  fnum = fnum + 1
                  If fnum > numflds Then
                     Msg$ = "Incorrect number of columns in line " + Str$(row) + " of " + UCase$(filename$) + "."
                     MsgBox Msg$, , "Error"
                     Close
                     Screen.MousePointer = 0
                     Exit Sub
                  End If
                  echarnum = charnum - 1
                  flength = echarnum - scharnum + 1
                  d(row, fnum) = Val(Trim(Mid$(l$, scharnum, flength)))
                  scharnum = echarnum + 2
               End If
            ElseIf charnum = Len(l$) Then
               fnum = fnum + 1
               echarnum = charnum
               flength = echarnum - scharnum + 1
               d(row, fnum) = Val(Trim(Mid$(l$, scharnum, flength)))
            End If
         Next charnum
      End If
   Wend
'End If
Close

'displays file name and numrecs in listbox
filename$ = tfn$
ValueBar.Cls
ValueBar.Print " " + UCase$(filename$) + "     Number Observations = " + Format$(numrecs, "####") + "     Number Columns = " + Format$(numfldsd, "###")

'reinits variable names
Screen.MousePointer = 0
ReDim vn$(numflds)

For i = 1 To numfldsd
   vn$(i) = "V" + Format$(i, "00")
Next i
Do While Main.XVarBox.ListCount > 0
   XVarBox.RemoveItem 0
   YVarBox.RemoveItem 0
Loop
For i = 1 To numfldsd
   'fills variable boxes with new names
   XVarBox.AddItem vn$(i)
   YVarBox.AddItem vn$(i)
Next i

'defines resid variables
vn$(numfldsd + 1) = "SG_Resids"
vn$(numfldsd + 2) = g1name$ + "_Resids"
vn$(numfldsd + 3) = g2name$ + "_Resids"
For i = 1 To 3
    XVarBox.AddItem vn$(numfldsd + i)
    YVarBox.AddItem vn$(numfldsd + i)
Next i
'fills resid variable with ndflags
For i = 1 To numrecs
   For j = numfldsd + 1 To numfldsd + 3
      d(i, j) = ndflag
   Next j
Next i

'reads *.var if it can be found, else calls db to name variables
varfilefound = False
If Command = "" Then 'normal open: search for *.var
   tfn$ = Trim$(tfn$)
   For i = 1 To Len(tfn$)
      If Mid$(tfn$, i, 1) = "." Then Exit For
   Next i
   For j = 0 To FileOpen.File1.ListCount
      If LCase$(FileOpen.File1.List(j)) = Left$(tfn$, i - 1) + ".var" Then
         Call OpenVar(tnp$ + Left$(tfn$, i - 1) + ".var")
         varfilefound = True
      ElseIf LCase$(FileOpen.File1.List(j)) = Left$(tfn$, i - 1) + ".dvn" Then
         Call OpenVar(tnp$ + Left$(tfn$, i - 1) + ".dvn")
         varfilefound = True
      End If
   Next j
Else 'command line open: open *.var
   For i = 1 To Len(tfn$)
      If Mid$(tfn$, i, 1) = "." Then Exit For
   Next i
   Call OpenVar(tnp$ + Left$(tfn$, i - 1) + ".var")
   varfilefound = True
End If
VarNamesCmd.Enabled = True
SaveLogCmd.Enabled = True

If varfilefound = False Then DefineDB.Show 1

'counts number of files open
numfiles = numfiles + 1

'resets tsoptsshown so that TSOptsDB will be shown
'the first time someone changes time scales
tsoptsshown = False

od2ce:
Screen.MousePointer = 0
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error"
   Resume od2es
End If
od2es:
End Sub

Private Sub OpenNamesCmd_Click()
On Error GoTo oncce

'show fileopen db
opencode = "*.var;*.dvn"
FileOpen.Show MODAL
If tnp$ + tfn$ = "" Then Exit Sub
Screen.MousePointer = 11

'read in variable names
Call OpenVar(tnp$ + tfn$)

oncce:
Screen.MousePointer = 0
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error"
   Resume onces
End If
onces:
End Sub

Private Sub PointCmd_Click()

PointCmd.Checked = True
ZoomCmd.Checked = False
GraphBox.MousePointer = 0

End Sub

Private Sub PredIntCmd_Click()
   PredIntCmd.Checked = Not PredIntCmd.Checked
End Sub

Private Sub PrintBestCmd_Click()

Call DrawPic(Main.GraphBox)

End Sub

Private Sub PrintCmd_Click()
   
If 1 = 2 Then
   MeanBut.Visible = False
   MedianBut.Visible = False
   TwoPBut.Visible = False
   ThreePBut(0).Visible = False
   ThreePBut(1).Visible = False
   FourPBut.Visible = False
   FivePBut.Visible = False
   MLRBut.Visible = False
   YVarBox.Visible = False
   XVarBox.Visible = False
   TimeSeriesBut.Visible = False
   TimeSeries2But.Visible = False
   XYBut.Visible = False
   XYAnimateBut.Visible = False
   HourlyProfilesBut.Visible = False
   G1G2But.Visible = False
   SingleGroupBut.Visible = False
   G1Lab.Visible = False
   G2Lab.Visible = False
   YVarLab.Visible = False
   XVarLab.Visible = False
   ClearValueBut.Visible = False
   ClearGraphBut.Visible = False
   HistBut.Visible = False
   BinFitBut.Visible = False
   
   
   Main.PrintForm
   
   MeanBut.Visible = True
   MedianBut.Visible = True
   TwoPBut.Visible = True
   ThreePBut(0).Visible = True
   ThreePBut(1).Visible = True
   FourPBut.Visible = True
   FivePBut.Visible = True
   MLRBut.Visible = True
   YVarBox.Visible = True
   XVarBox.Visible = True
   TimeSeriesBut.Visible = True
   TimeSeries2But.Visible = True
   XYBut.Visible = True
   XYAnimateBut.Visible = True
   HourlyProfilesBut.Visible = True
   G1G2But.Visible = True
   SingleGroupBut.Visible = True
   G1Lab.Visible = True
   G2Lab.Visible = True
   XVarLab.Visible = True
   YVarLab.Visible = True
   ClearValueBut.Visible = True
   ClearGraphBut.Visible = True
   HistBut.Visible = True
   BinFitBut.Visible = True
   
End If

   Main.PrintForm
End Sub



Private Sub ProjectSavingsCmd_Click()
Dim i, j, k, Msg$, no As Integer, numpasses As Integer
Dim pnum, pdenom, prevres, P, ytot, ymean
Dim sse, mse, rmse, cvrmse, savings, postuse
Dim det, r2, adjr2, trace, rss, sig2, var
Dim c11, c22, c23, c12, wyz, t
Dim ModelErr, MeasErr, PredErr, preduse 'global measerr
Dim xr As Integer, xc As Integer, yr As Integer, yc As Integer
Dim xtr As Integer, xtc As Integer, xtxr As Integer, xtxc As Integer
Dim qr As Integer, qc As Integer, betar As Integer, betac As Integer
Dim yhatpstr  As Integer, yhatpstc As Integer
Dim ar As Integer, ac As Integer, xbqr As Integer, xbqc As Integer
Dim xrpst As Integer, xbqtr As Integer, xbqtc As Integer
Dim mtr As Integer, mtc As Integer, mvr As Integer, mvc As Integer
Dim qtr As Integer, qtc As Integer, xbqvr  As Integer, xbqvc As Integer
Dim skiprec, ncount
Dim ybar1, ybar2
Dim indvar1, indvar2
On Error GoTo projectsavingsce

'checks for pre and post
If g1name$ <> "Pre" Or g2name$ <> "Post" Then
   Msg$ = "Savings calculation is available only when groups have been named 'Pre' and 'Post'."
   MsgBox Msg$, , "Error"
   GoTo projectsavingsce
End If

'clears screen
Screen.MousePointer = 11


'inits vars for 'old' calc save routines
pred = 0
meas = 0
saved = 0
numobspost = 0

If modtype$ = "Mean" Then

   'fills x() and y(), and calcs n for the chosen variable
   Call SetAll
   
   'fills xpst(), ypst()
   pass = 2
   Call SetGrp
   xrpst = n
   xcpst = 2
   ReDim xpst(xrpst, xcpst), ypst(xrpst, 1)
   For i = 1 To xrpst
      ypst(i, 1) = y(i, 1)
      For j = 1 To xcpst
         xpst(i, j) = x(i, j)
      Next j
   Next i
   
   'fills x(), y() with pre data
   pass = 1
   Call SetGrp
   xr = n
   yr = n
   yc = 1
   xc = 1

   'checks for pre and post data before continuing
   If xr = 0 Or xrpst = 0 Then
      Msg$ = "No data available for savings calculation."
      MsgBox Msg$, , "Error"
      GoTo projectsavingsce
   End If
   
   'dim arrays
   ReDim xt(xc, xr), xtx(xc, xc), xtxi(xc, xc), q(xc, xr)
   ReDim xtotal(xc), yhatpst(xrpst, 1)
   ReDim xtotalp(xc), xmeanp(xc), xb(1, xc)
   ReDim v(xr, 1)
   ReDim A(1, xr), m(1, xr), xn(1, xc)
   ReDim mt(xr, 1), mv(1, 1)
   ReDim qt(xr, xc), xdum(xr, 1), qdum(xr, 1)
   
   'calcs ymean
   For i = 1 To xr
      ytot = ytot + y(i, 1)
   Next i
   ymean = ytot / xr
      
   'calcs params
   ReDim mres(xr)
   sse = 0: pnum = 0: pdenom = 0
   For i = 1 To xr
      mres(i) = y(i, 1) - ymean
      sse = sse + mres(i) ^ 2
      If i = 1 Then
         prevres = mres(1)
      Else
         pnum = pnum + mres(i) * prevres
         pdenom = pdenom + mres(i) ^ 2
         prevres = mres(i)
      End If
   Next i
   mse = sse / (n - xc)
   rmse = mse ^ 0.5
   cvrmse = rmse / ymean * 100
   P = pnum / pdenom
   
   'calcs savings
   savings = 0: postuse = 0
   For i = 1 To xrpst
      postuse = postuse + ypst(i, 1)
      savings = savings + ymean - ypst(i, 1)
   Next i
      
ElseIf modtype$ = "Median" Then
       
       For i = 1 To numrecs
           If grpd(i) = 2 And d(i, yfld) <> ndflag Then
              numobspost = numobspost + 1
              pred = pred + ymedianpre
              meas = meas + d(i, yfld)
           End If
       Next i
       saved = pred - meas


ElseIf modtype = "2P" Then

   'fills x() and y(), and calcs n for the chosen variable
   Call SetAll
   
   'fills xpst(), ypst()
   pass = 2
   Call SetGrp
   xrpst = n
   xcpst = 2
   ReDim xpst(xrpst, xcpst), ypst(xrpst, 1)
   For i = 1 To xrpst
      ypst(i, 1) = y(i, 1)
      For j = 1 To xcpst
         xpst(i, j) = x(i, j)
      Next j
   Next i
   
   'fills x(), y() with pre data
   pass = 1
   Call SetGrp
   xr = n
   yr = n
   yc = 1
   xc = 2
   numxvars = 1
   
   'checks for pre and post data before continuing
   If xr = 0 Or xrpst = 0 Then
      Msg$ = "No data available for savings calculation."
      MsgBox Msg$, , "Error"
      GoTo projectsavingsce
   End If
   
   'dim arrays
   ReDim xt(xc, xr), xtx(xc, xc), xtxi(xc, xc), q(xc, xr)
   ReDim xtotal(xc), yhatpst(xrpst, 1)
   ReDim xtotalp(xc), xmeanp(xc), xb(1, xc)
   ReDim v(xr, 1)
   ReDim A(1, xr), m(1, xr), xn(1, xc)
   ReDim mt(xr, 1), mv(1, 1)
   ReDim qt(xr, xc), xdum(xr, 1), qdum(xr, 1)

   'finds Regression parameters
   Call Trans(x(), xr, xc, xt(), xtr, xtc)
   Call Mult(xt(), xtr, xtc, x(), xr, xc, xtx(), xtxr, xtxc)
   Call Invert(xtx(), xtxr, xtxi(), det, sec())
   Call Mult(xtxi(), xtxr, xtxr, xt(), xtr, xtc, q(), qr, qc)
   Call Mult(q(), qr, qc, y(), yr, yc, beta(), betar, betac)
   Call Inf(x(), xr, xc, y(), yr, yc, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), P)

   'calcs savings
   Call Mult(xpst(), xrpst, xc, beta(), betar, betac, yhatpst(), yhatpstr, yhatpstc)
   savings = 0: postuse = 0
   For i = 1 To xrpst
      postuse = postuse + ypst(i, 1)
      savings = savings + yhatpst(i, 1) - ypst(i, 1)
   Next i

ElseIf modtype$ = "3P" Then
       
       For i = 1 To numrecs
           If grpd(i) = 2 And d(i, xfld(1)) <> ndflag And d(i, yfld) <> ndflag Then
              'post
              numobspost = numobspost + 1
              If hcindex = 0 Then
                 If d(i, xfld(1)) <= xcppre Then ivar = 0 Else ivar = 1
              Else
                 If d(i, xfld(1)) <= xcppre Then ivar = 1 Else ivar = 0
              End If
              pred = pred + betapre(1) + betapre(2) * (d(i, xfld(1)) - xcppre) * ivar
              meas = meas + d(i, yfld)
           End If
       Next i
       saved = pred - meas

ElseIf modtype$ = "4P" Then
       
   For i = 1 To numrecs
       If grpd(i) = 2 And d(i, xfld(1)) <> ndflag And d(i, yfld) <> ndflag Then
          'post
          numobspost = numobspost + 1
          If d(i, xfld(1)) <= xcppre Then ivar = 0 Else ivar = 1
          pred = pred + betapre(1) + betapre(2) * d(i, xfld(1)) + betapre(3) * (d(i, xfld(1)) - xcppre) * ivar
          meas = meas + d(i, yfld)
       End If
   Next i
   saved = pred - meas


ElseIf modtype$ = "5P" Then

   For i = 1 To numrecs
      If grpd(i) = 2 And d(i, xfld(1)) <> ndflag And d(i, yfld) <> ndflag Then
         'post
         numobspost = numobspost + 1
         If d(i, xfld(1)) <= cp1pre Then
           indvar1 = 1
         Else
            indvar1 = 0
         End If
                 
        If d(i, xfld(1)) <= cp2pre Then
           indvar2 = 0
        Else
           indvar2 = 1
        End If
   
         pred = pred + ycppre + lspre * indvar1 * (d(i, xfld(1)) - cp1pre) + rspre * indvar2 * (d(i, xfld(1)) - cp2pre)
         meas = meas + d(i, yfld)
      End If
   Next i
   saved = pred - meas

ElseIf modtype$ = "MLR" Then

   'sets numvars from global variable
   numxvars = mlrxvars
   xc = mlrxvars + 1
   xcpst = mlrxvars + 1

   'redimensions X(), Y(), and IndexNum()
   ReDim x(numrecs, xc), y(numrecs, 1)
   ReDim xpst(numrecs, xcpst), ypst(numrecs, 1)

   'fills x() and xpst()
   xr = 0
   xrpst = 0
   For j = 1 To numrecs
      
      If grpd(j) = 1 And d(j, yfld) <> ndflag Then
         'pre....

         'checks for no data flag
         skiprec = False
         For k = 1 To numxvars
            If d(j, xfld(k)) = ndflag Then skiprec = True
         Next k
         
         If skiprec = False Then 'fills x() and y()
            xr = xr + 1
            x(xr, 1) = 1
            For k = 1 To numxvars
               x(xr, k + 1) = d(j, xfld(k))
            Next k
            y(xr, 1) = d(j, yfld)
         End If

      
      ElseIf grpd(j) = 2 And d(j, yfld) <> ndflag Then
         'post....

         'checks for no data flag
         skiprec = False
         For k = 1 To numxvars
            If d(j, xfld(k)) = ndflag Then skiprec = True
         Next k
         
         If skiprec = False Then 'fills xpst() and ypst()
            xrpst = xrpst + 1
            xpst(xrpst, 1) = 1
            For k = 1 To numxvars
               xpst(xrpst, k + 1) = d(j, xfld(k))
            Next k
            ypst(xrpst, 1) = d(j, yfld)
         End If
      
      End If
   Next j
   yr = xr
   yc = 1

   'prints error if n = 0
   If xr = 0 Then
      MsgBox "No data is available to model.", , "Error"
      GoTo projectsavingsce
   End If
   
   'dim arrays
   ReDim xt(xc, xr), xtx(xc, xc), xtxi(xc, xc), q(xc, xr)
   ReDim xtotal(xc), yhatpst(xrpst, 1)
   ReDim xtotalp(xc), xmeanp(xc), xb(1, xc)
   ReDim v(xr, 1)
   ReDim A(1, xr), m(1, xr), xn(1, xc)
   ReDim mt(xr, 1), mv(1, 1)
   ReDim qt(xr, xc), xdum(xr, 1), qdum(xr, 1)

   'finds Regression parameters
   Call Trans(x(), xr, xc, xt(), xtr, xtc)
   Call Mult(xt(), xtr, xtc, x(), xr, xc, xtx(), xtxr, xtxc)
   Call Invert(xtx(), xtxr, xtxi(), det, sec())
   Call Mult(xtxi(), xtxr, xtxr, xt(), xtr, xtc, q(), qr, qc)
   Call Mult(q(), qr, qc, y(), yr, yc, beta(), betar, betac)
   Call Inf(x(), xr, xc, y(), yr, yc, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), P)

   'calcs savings
   Call Mult(xpst(), xrpst, xc, beta(), betar, betac, yhatpst(), yhatpstr, yhatpstc)
   savings = 0: postuse = 0
   For i = 1 To xrpst
      postuse = postuse + ypst(i, 1)
      savings = savings + yhatpst(i, 1) - ypst(i, 1)
   Next i


ElseIf modtype$ = "BinFit" Then

       If setminmax = False Then
          MsgBox "User must define Xmin and Xmax for Bin Fit model before savings can be calculated.", , "Error"
          GoTo projectsavingsce
       End If

       Dim xlow, xhigh, intsize
       'determines the interval size
       intsize = (xmaxbf - xminbf) / numints

       For i = 1 To numrecs
          If grpd(i) = 2 And d(i, xfld(1)) <> ndflag And d(i, yfld) <> ndflag Then
             'runs through the j intervals looking for the proper bin
             For j = 1 To numints
                xlow = xminbf + (j - 1) * intsize
                xhigh = xminbf + j * intsize
                'for first interval use >= on low bound
                If j = 1 Then
                  If d(i, xfld(1)) >= xlow And d(i, xfld(1)) <= xhigh Then
                     numobspost = numobspost + 1
                     'if ybarbf = ndflag then search for nearest non-zero ybarbf
                     If ybarbf(j) = ndflag Then
                        For k = j + 1 To numints
                           If ybarbf(k) <> ndflag Then
                              'calc pred using this ybarbf
                              pred = pred + ybarbf(k)
                              Exit For
                           End If
                        Next k
                     Else
                        pred = pred + ybarbf(j)
                     End If
                     meas = meas + d(i, yfld)
                     j = 1
                     Exit For
                  End If
                Else  'use only > on low bound
                  If d(i, xfld(1)) > xlow And d(i, xfld(1)) <= xhigh Then
                     numobspost = numobspost + 1
                     If ybarbf(j) = ndflag Then
                        ybar1 = ndflag
                        ybar2 = ndflag
                        'find nearest preceding ybarbf()
                        For k = j - 1 To 1 Step -1
                           If ybarbf(k) <> ndflag Then
                              'calc pred using this ybarbf
                              ybar1 = ybarbf(k)
                              Exit For
                           End If
                        Next k
                        'find nearest subsequent ybarbf()
                        For k = j + 1 To numints
                           If ybarbf(k) <> ndflag Then
                              'calc pred using this ybarbf
                              ybar2 = ybarbf(k)
                              Exit For
                           End If
                        Next k
                        'calc pred
                        If ybar1 <> ndflag And ybar2 <> ndflag Then
                           pred = pred + (ybar1 + ybar2) / 2
                        ElseIf ybar1 = ndflag Then
                           pred = pred + ybar2
                        ElseIf ybar2 = ndflag Then
                           pred = pred + ybar1
                        End If
                     Else
                        pred = pred + ybarbf(j)
                     End If
                     meas = meas + d(i, yfld)
                     j = 1
                     Exit For
                  End If
                End If
             Next j
          End If
       Next i
       saved = pred - meas

Else
   
   Msg$ = "No model of pre retrofit data has been selected."
   MsgBox Msg$, , "Error"
   GoTo projectsavingsce

End If


If modtype$ = "Mean" Or modtype$ = "2P" Or modtype$ = "MLR" Then
   'calculate uncertainty using Ruch et al. Hybrid Method
   
   'Step 0
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 0"
   For i = 2 To xc
      For j = 1 To xr
         xtotal(i) = xtotal(i) + x(j, i)
      Next j
      xmean(i) = xtotal(i) / xr
   Next i
   For i = 2 To xc
      For j = 1 To xr
         x(j, i) = x(j, i) - xmean(i)
      Next j
   Next i
   For i = 2 To xc
      For j = 1 To xrpst
         xpst(j, i) = xpst(j, i) - xmean(i)
      Next j
   Next i
   
   'Step 1
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 1"
   Call Trans(x(), xr, xc, xt(), xtr, xtc)
   Call Mult(xt(), xtr, xtc, x(), xr, xc, xtx(), xtxr, xtxc)
   Call Invert(xtx(), xtxr, xtxi(), det, sec())
   Call Mult(xtxi(), xtxr, xtxr, xt(), xtr, xtc, q(), qr, qc)
   
   'Step 2
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 2"
   For i = 2 To xc
      For j = 1 To xrpst
         xtotalp(i) = xtotalp(i) + xpst(j, i)
      Next j
      xmeanp(i) = xtotalp(i) / xrpst
   Next i
   xb(1, 1) = 1
   For i = 2 To xc
     xb(1, i) = xmeanp(i)
   Next i
   
   'Step 3
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 3"
   k = P * (1 - P ^ xrpst) / (1 - P)
   For i = 1 To xr
      v(i, 1) = k * P ^ (xr - i)
   Next i
   
   'Step 4
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 4"
   xn(1, 1) = 1
   For i = 2 To xc
      xn(1, i) = x(xr, i)
   Next i
   Call Mult(xn(), 1, xc, q(), qr, qc, A(), ar, ac)
   A(1, xr) = A(1, xr) - 1
   For i = 1 To ac
      m(1, i) = -k * A(1, i)
   Next i
   
   'Step 5
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 5"
   'define subprogram WYZCalc()
   
   'Step 6
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 6"
   ReDim xbq(1, qc), xbqt(qc, 1)
   Call Mult(xb(), 1, xc, q(), qr, qc, xbq(), xbqr, xbqc)
   Call Trans(xbq(), xbqr, xbqc, xbqt(), xbqtr, xbqtc)
   Call WYZCalc(xbqt(), xbqt(), P, xr, wyz)
   c11 = wyz * xrpst ^ 2
   
   'Step 7
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 7"
   Call Trans(m(), 1, xr, mt(), mtr, mtc)
   Call WYZCalc(mt(), mt(), P, xr, wyz)
   Call Mult(m(), 1, xr, v(), xr, 1, mv(), mvr, mvc)
   c23 = -2 * mv(1, 1) + wyz
   
   'Step 8
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 8"
   ReDim xbqv(xbqr, 1)
   Call WYZCalc(xbqt(), mt(), P, xr, wyz)
   Call Mult(xbq(), xbqr, xbqc, v(), xr, 1, xbqv(), xbqvr, xbqvc)
   c12 = 2 * xrpst * (wyz - xbqv(1, 1))
   
   'Step 9
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 9"
   c22 = (2 * P ^ (xrpst + 1) - xrpst * (P ^ 2 - 1) - 2 * P) / (1 - P) ^ 2
   
   'Step 10
   'GraphBox.CurrentX = -10
   'GraphBox.Print "Step 10"
   Call Trans(q(), qr, qc, qt(), qtr, qtc)
   For i = 1 To xc
      For j = 1 To xr
         xdum(j, 1) = x(j, i)
         qdum(j, 1) = qt(j, i)
      Next j
      Call WYZCalc(xdum(), qdum(), P, xr, wyz)
      trace = trace + wyz
   Next i
   rss = (rmse ^ 2) * (xr - xc)
   sig2 = rss / (xr - trace)
   var = sig2 * (c11 + c12 + c22 + c23)
   t = 1.96 + 2.7 / (xr - 2)
   ModelErr = t * var ^ 0.5
   MeasErr = postuse * MeasErr / 100
   PredErr = (ModelErr ^ 2 + MeasErr ^ 2) ^ 0.5
   preduse = savings + postuse
End If

'print results
ValueBar.Cls
If modtype$ = "2P" Or modtype = "MLR" Then
   
   ValueBar.Print "Savings using a " + modtype$ + " model, during " + Format(xrpst, "#,###") + " post observations:"
   ValueBar.Print "Predicted Pre Use   = " + Format(preduse, "#,###,##0.0") + "  ±" + Format(ModelErr, "###,##0.0") + "  ±" + Format(ModelErr / preduse * 100, "#####0.0") + "%"
   ValueBar.Print "Measured Post Use   = " + Format(postuse, "#,###,##0.0") + "  ±" + Format(MeasErr, "###,##0.0") + "  ±" + Format(MeasErr / postuse * 100, "#####0.0") + "%"
   ValueBar.Print "Savings During Post = " + Format(savings, "#,###,##0.0") + "  ±" + Format(PredErr, "###,##0.0") + "  ±" + Format(PredErr / savings * 100, "#####0.0") + "%"
   ValueBar.Print "Avg Savings Per Obs = " + Format(savings / xrpst, "#,###,##0.0") + "  ±" + Format(PredErr / xrpst, "###,##0.0") + "  ±" + Format(PredErr / savings * 100, "#####0.0") + "%"

ElseIf modtype$ = "Mean" Then
   
   ValueBar.Print "Savings using a " + modtype$ + " model, during " + Format(xrpst, "#,###") + " post observations:"
   ValueBar.Print "Predicted Pre Use   = " + Format(preduse, "#,###,##0.0") + "  ±" + Format(ModelErr, "###,##0.0") + "  ±" + Format(ModelErr / preduse * 100, "#####0.0") + "%"
   ValueBar.Print "Measured Post Use   = " + Format(postuse, "#,###,##0.0") + "  ±" + Format(MeasErr, "###,##0.0") + "  ±" + Format(MeasErr / postuse * 100, "#####0.0") + "%"
   ValueBar.Print "Savings During Post = " + Format(savings, "#,###,##0.0") + "  ±" + Format(PredErr, "###,##0.0") + "  ±" + Format(PredErr / savings * 100, "#####0.0") + "%"
   ValueBar.Print "Avg Savings Per Obs = " + Format(savings / xrpst, "#,###,##0.0") + "  ±" + Format(PredErr / xrpst, "###,##0.0") + "  ±" + Format(PredErr / savings * 100, "#####0.0") + "%"

Else
   
   ValueBar.Print "Savings using a " + modtype$ + " model, during " + Format(numobspost, "#,###") + " post observations:"
   ValueBar.Print "Predicted Pre Use   = " + Format(pred, "#,###,##0.0")
   ValueBar.Print "Measured Post Use   = " + Format(meas, "#,###,##0.0")
   ValueBar.Print "Savings During Post = " + Format(saved, "#,###,##0.0")
   ValueBar.Print "Avg Savings Per Obs = " + Format(saved / numobspost, "#,###,##0.0")

End If

projectsavingsce:
Screen.MousePointer = 0
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error"
   Resume projectsavingses
End If
projectsavingses:
End Sub

Private Sub RunSimulationCmd_Click()
'Program listing from ESIM.BAS by Kelly Kissock
'ESIM.BAS is a simplified systems model of energy use in a two zone
'commercial building. It can simulate CAV, VAV systems with economizer cycles,
'hot deck reset schedules and latent cooling loads.
ReDim QCCM(numrecs), QHCM(numrecs), EDEM(numrecs)
'Dim tr, tcc, thc, fi, foa, ua, vr, vmini, vmine, pcp, phfg, s1, s2
Dim vr, vmini, vmine, pcp, phfg, s1, s2
Dim i
'Dim cavon, vavon, hdreseton, econon, laton, solaron
Dim solaron
Dim wbe, toa, woa, hsol
Dim qi, qli, qle
Dim qlv, qlp, qcl
Dim vri, vre, vci, vce, vhi, vhe, vcc, vhc
Dim tma, qcc, qhc, EDE
Dim ecode

On Error GoTo rsce
Screen.MousePointer = 11
ecode = False

'input parameters
'vavflow, foa, thc, tr, tcc are global vars set in system dialog box
'tr = 72                                         '(F)
'tcc = 55                                        '(F)
'thc = 120                                       '(F)
'fi = .65                                        'fi = Ai / (Ai + Ae)
'foa = .1                                        'foa = voa / vr
'ua = 20000                                      '(Btu/hrF)
'vr = 1.666666 * floorarea * 60             '(ft3/hr) = cfm/ft2 * ft2 * 60(min/hr) for CV
vr = cavflow * floorarea * 60             '(ft3/hr) = cfm/ft2 * ft2 * 60(min/hr) for CV
'vmini = .33 * floorarea * 60 * fi  '(ft3/hr) = cfm/ft2 * ft2 * 60(min/hr) for VAV
'vmine = .33 * floorarea * 60 * (1 - fi)    '(ft3/hr) = cfm/ft2 * ft2 * 60(min/hr) for VAV
vmini = vavflow * floorarea * 60 * fi  '(ft3/hr) = cfm/ft2 * ft2 * 60(min/hr) for VAV
vmine = vavflow * floorarea * 60 * (1 - fi)    '(ft3/hr) = cfm/ft2 * ft2 * 60(min/hr) for VAV

'define constants
pcp = 0.018                                     '(Btu/ft3F)
phfg = 80                                       '(Btu/ft3) air density * latent heat of vaporization
s1 = 10000                                      '(Btu/hr) solar load constant
s2 = 400                                        '(Btu/hrF) solar load slope

'turns on and off system and control options, 1=on, 0=off
'cavon = 0
'vavon = 1
'hdreseton = 0
'econon = 0
'laton = 1
solaron = 0

'begin simulation
For i = 1 To numrecs
   
   'fills d2() with nodata flags
   QCCM(i) = ndflag
   QHCM(i) = ndflag
   EDEM(i) = ndflag
   
   'collect input data
   'this is for hourly.dat .var
   wbe = d(i, 9)        '(kWh/hr)
   toa = d(i, 5)        '(F)
   woa = d(i, 6)        '(lbw/lba)
   hsol = d(i, 7)       '(Btu/ft2hr)
   
   'skips record if no-data flag for input
   If wbe = -99 Or toa = -99 Or (laton = 1 And woa = -99) Or (solaron = 1 And hsol = -99) Then
      GoTo nxti:
   End If
   
   'calcs zone loads
   qi = (wbe * 3412) + (numoccs * 255)          'kWh/hr to btu/hr, 255 btu(senible)/person hr
   qli = fi * qi                                                'total loads to interior zone
   If solaron = 1 Then
      qle = (1 - fi) * qi + ua * (toa - tr) + (s1 + s2 * toa)      'total loads to exterior zone
   Else
      qle = (1 - fi) * qi + ua * (toa - tr)       'total loads to exterior zone
   End If
  
   'hot deck reset
   If hdreseton = True Then 'use standard hd reset
      If toa <= 50 Then
         thc = 120
      ElseIf toa > 50 And toa < 80 Then
         thc = 120 - (120 - 80) / (80 - 50) * (toa - 50)
      Else
         thc = 80
      End If
   ElseIf ophdreseton = True Then 'use optimized hd reset
      thc1 = (foa * toa) + ((1 - foa) * tr) + 1
      thc2 = tr - (qle / ((1 - fi) * vr * pcp)) + 1
      If thc1 >= thc2 Then thc = thc1 Else thc = thc2
   Else 'use fixed heat coil temp
      thc = thcfixed
   End If

   'economizer
   If econon = True Then
      If toa <= tcc Then
         foa = (tcc - tr) / (toa - tr)
      ElseIf toa > tcc And toa <= tr Then
         foa = 1
      Else
         foa = 0.1 'minimum outside air flow rate
      End If
   Else 'use fixed foa
      foa = foafixed
   End If

   'latent loads
   If laton = True Then
      If woa > 0.0092 Then
         qlv = vr * foa * phfg * (woa - 0.0092)    'latent load due to ventilation
      Else
         qlv = 0
      End If
      qlp = numoccs * 255          '255 btu(latent)/ occupant hour
      qcl = (qlv + qlp) / 1000000  'total latent Btu to MBtu
   Else
      qcl = 0
   End If
  
   'system equations
   If cavon = True Then 'cav system
      'zone total mass flow rates
      vri = fi * vr
      vre = (1 - fi) * vr
      'int zone ebalance, finds vci
      vci = (vri * pcp * (tr - thc) - qli) / (pcp * (tcc - thc))
      'ext zone ebalance, finds vce
      vce = (vre * pcp * (tr - thc) - qle) / (pcp * (tcc - thc))
      'total air flow mass balance, finds vcc, vhc
      vcc = vci + vce
      vhc = vr - vcc
   Else 'vavon = 1 Then 'vav system
      'int zone ebalance, finds vci, vhi
      If qli > 0 Then
         vci = qli / (pcp * (tr - tcc))
         vhi = 0
         If vci < vmini Then       'cv formula for volume = vmin
             vci = (vmini * pcp * (tr - thc) - qli) / (pcp * (tcc - thc))
             vhi = vmini - vci
         End If
      Else
         vhi = qli / (pcp * (tr - thc))
         vci = 0
         If vhi < vmini Then
             vci = (vmini * pcp * (tr - thc) - qli) / (pcp * (tcc - thc))
             vhi = vmini - vci
         End If
      End If
      'ext zone ebalance, finds vce, vhe
      If qle > 0 Then
         vce = qle / (pcp * (tr - tcc))
         vhe = 0
         If vce < vmine Then        'cv formula for volume = vmin
             vce = (vmine * pcp * (tr - thc) - qle) / (pcp * (tcc - thc))
             vhe = vmine - vce
         End If
      Else
         vhe = qle / (pcp * (tr - thc))
         vce = 0
         If vhe < vmine Then
             vce = (vmine * pcp * (tr - thc) - qle) / (pcp * (tcc - thc))
             vhe = vmine - vce
         End If
      End If
      'total air flow mass balance, finds vcc, vhc
      vcc = vci + vce
      vhc = vhi + vhe
   End If

   'oamixing box ebalance, finds tma
   tma = foa * toa + (1 - foa) * tr

   'cooling coil ebalance, finds qcc
   qcc = (vcc * pcp * (tma - tcc)) / 1000000    'Btu/hr to MBtu/hr
   qcc = qcc + qcl   'qcl = latent load

   'heating coil ebalance, finds qhc
   qhc = vhc * pcp * (thc - tma) / 1000000      'Btu/hr to MBtu/hr
   
   'checks for errors
   If vci < 0 Or vce < 0 Or vhc < 0 Or qcc < 0 Or qhc < 0 Then ecode = True
   
   'calcs EDE
   If qcc + qhc > 0 Then
      EDE = (qcc - qhc) / (qcc + qhc) * 100
   Else
      EDE = ndflag
   End If

   'fills data matrix with simulation output
   QCCM(i) = qcc
   QHCM(i) = qhc
   EDEM(i) = EDE
nxti:
Next i

'add new variables
numsims = numsims + 1
Call AddVar(QCCM(), "CC" + Format(numsims) + "_(MBtu/hr)")
Call AddVar(QHCM(), "HC" + Format(numsims) + "_(MBtu/hr)")
Call AddVar(EDEM(), "EDE" + Format(numsims) + "_(%)")

'prints error message
ValueBar.Cls
If ecode = True Then
   Msg$ = "The HVAC system could not meet the zone loads using these settings, creating negative values for parameters that cannot be negative."
   MsgBox Msg$, , "Error"
End If

'prints output message
ValueBar.Print "Cooling Coil (CC) and Heating Coil (HC) energy use and"
ValueBar.Print "EDE (EDE) from simulation #" + Format(numsims) + " have been added to the variable list."

rsce:
Screen.MousePointer = 0
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error"
   Resume rses
End If
rses:
End Sub

Private Sub SaveDataCmd_Click()
Dim i As Integer, j As Integer

tnp$ = ""
tfn$ = ""

FileSave.Show MODAL
Screen.MousePointer = 11

If tfn$ = "*.*" Or tfn$ = "" Then GoTo pdes:

'gets filename$ from dialog box
outfile$ = tnp$ + tfn$
If outfile$ = "" Then GoTo pdes:

Open outfile$ For Output As #1
For i = 1 To numrecs
   For j = 1 To numfldsd
      Print #1, d(i, j);
   Next j
   For j = numfldsd + 4 To numflds
      Print #1, d(i, j);
   Next j
   Print #1,
Next i

Close #1

pdes:
Screen.MousePointer = 0

End Sub

Private Sub SaveLogCmd_Click()
Static alreadyopen As Integer

SaveLogCmd.Checked = Not SaveLogCmd.Checked

If SaveLogCmd.Checked = True And alreadyopen = False Then
     alreadyopen = True
     filename$ = tnp$ + filename$
     lengthfn = Len(filename$)
     logname$ = Left$(filename$, lengthfn - 4) + ".log"

     Open logname$ For Output As #3
End If

End Sub

Private Sub SaveNamesCmd_Click()
Dim i As Integer
On Error GoTo sncce

tnp$ = ""
tfn$ = ""
opencode = "*.var"
FileSave.Show MODAL

If tfn$ = "*.*" Or tfn$ = "" Then GoTo sncce:
Open tnp$ + tfn$ For Output As #1

For i = 1 To numfldsd
    Print #1, vn$(i)
Next i
For i = numfldsd + 4 To numflds
    Print #1, vn$(i)
Next i

sncce:
Close #1
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume snces
End If
snces:

End Sub

Private Sub SelDatesCmd_Click()

SelDatesDB.Show 1

End Sub

Private Sub SetNodataFlagCmd_Click()
NoDataDB.Show MODAL
End Sub

Private Sub SetScalesCmd_Click()
   SetScalesDB.Show 1
End Sub


Private Sub ShowHelpCmd_Click()

ShowHelpContents

End Sub

Private Sub SingleGroupBut_Click()
gmodnum = 1
End Sub

Private Sub ThreePMVRBut_Click(Index As Integer)
Dim i, j
Dim numxvars, xcp, sexcp, ycp, seycp, slope, seslope, rmse, cvrmse, r2
Dim ls, sels, rs, sers
ReDim ivcoefs(7), seivcoefs(7)
Dim ivar, xpnt, ypnt
'pass, xr are global
On Error GoTo psmce

'sets modtype
modtype$ = "3P"

'calls Mult Reg dialog box
MultRgDB.Show 1
If cancelmlr = True Then GoTo psmce:

'counts numxvars for multrgdb
numxvars = 1
Do While xfld(numxvars) <> 0
   numxvars = numxvars + 1
Loop
numxvars = numxvars - 1

'displays error message if no vars are selected
If numxvars = 0 Or yfld = 0 Then
   MsgBox "No variables have been selected.  Remember to press the 'Set' button after each variable is selected.", , "Error"
   GoTo psmce:
End If

'unloads multrgdb
Unload MultRgDB
Screen.MousePointer = 11

'sets pass
If G1G2But.Value = True And gmodnum = 2 Then pass = 2 Else pass = 1
   
'fills X(), Y(), and sets n
ReDim x(numrecs, numxvars + 1), y(numrecs, 1)
Call FillXYGrp(pass, numxvars, x(), y(), n)

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo psmce:
ElseIf G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined.", , "Error"
   GoTo psmce:
ElseIf xfld(1) = 0 Then
   MsgBox "No X variable has been defined.", , "Error"
   GoTo psmce:
End If

'calls threepmvr() index = 0 for 3PC and 1 for 3ph
Call ThreePMVR(x(), y(), n, numxvars, Index, xcp, sexcp, ycp, seycp, slope, seslope, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)


'sets right and left slopes
If Index = 0 Then '3pc
   ls = 0: sels = 0
   rs = slope: sers = seslope
Else '3ph
   ls = slope: sels = seslope
   rs = 0: sers = 0
End If

'creates global variables for savings calcuations
If pass = 1 Then
   premodtype$ = modtype$
   rmsepre = rmse
   npre = n
   numxvarspre = numxvars
   xcppre = xcp
   ycppre = ycp
   slopepre = slope
   indexpre = Index
   ReDim ivcoefspre(7)
   For i = 2 To numxvars
      ivcoefspre(i) = ivcoefs(i)
   Next i
   modelstats = " N = " + Format$(n, "####") + "  R2 = " + Format$(r2, "0.00") + "  RMSE = " + Format$(rmse, "#,##0.00") + "  CV-RMSE = " + Format$(cvrmse, "##0.00") + "%"
   modeleqn = vn$(yfld) + " = " + Format(ycp, "#,##0.00") + "  -  " + Format(ls, "#,##0.00") + " (" + Format(xcp, "#,##0.00") + " - " + vn$(xfld(1)) + ")+  +  " + Format(rs, "#,##0.00") + " (" + vn$(xfld(1)) + " - " + Format(xcp, "#,##0.00") + ")+"
   For i = 2 To numxvars
      modeleqn = modeleqn + " + " + Format(ivcoefs(i), "#,##0.00") + " " + vn$(xfld(i))
   Next i
End If


'prints results
'Call DWTest(3, dwr$)
sf$ = "###0.0000"
If G1G2But.Value = False Or (G1G2But.Value = True And pass = 1) Then ValueBar.Cls
Call PrintHeaderLine(gmodnum, modtype$)
ValueBar.Print " N = " + Format$(n, "####") + "  R2 = " + Format$(r2, "0.00") + "  RMSE = " + Format$(rmse, sf$) + "  CV-RMSE = " + Format$(cvrmse, "##0.0\%")
ValueBar.Print " Ycp = " + Format$(ycp, sf$) + " (" + Format$(seycp, sf$) + ")";
ValueBar.Print "  Xcp = " + Format$(xcp, sf$) + " (" + Format$(sexcp, sf$) + ")";
ValueBar.Print "  LS = " + Format$(ls, sf$) + " (" + Format$(sels, sf$) + ")";
ValueBar.Print "  RS = " + Format$(rs, sf$) + " (" + Format$(sers, sf$) + ")"
'ValueBar.Print "  p = " + Format$(autocor, "0.00") + "  DW = " + Format$(dw, "0.00") + " (" + dwr$ + ")"
If numxvars >= 2 Then
   For i = 2 To numxvars
      ValueBar.Print " X" + Format(i, "0") + " = "; Format(ivcoefs(i), "0.0000"); " ("; Format(seivcoefs(i), "0.0000"); ")  ";
   Next i
   ValueBar.Print
End If
'print model
ValueBar.Print "Model: "; vn$(yfld); " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp, "#,##0.00"); " - "; vn$(xfld(1)); ")+  +  "; Format(rs, "#,##0.00"); " ("; vn$(xfld(1)); " - "; Format(xcp, "#,##0.00"); ")+";
For i = 2 To numxvars
   ValueBar.Print " + "; Format(ivcoefs(i), "#,##0.00"); " "; vn$(xfld(i)); '" X" + Format(i, "0"); '
Next i
ValueBar.Print

'calcs prediction interval
t = 1.96 + 2.7 / (n - 3)
predint = t * rmse * (1 + 1 / n) ^ 0.5

'mlr graphing
xr = n
ReDim ypred(xr), xpred(xr)
'calcs ypred and saves values of xpred and ypred before they get changed by the call to xybut
For i = 1 To xr
   If Index = 0 Then '3pc
      If x(i, 2) <= xcp Then ivar = 0 Else ivar = 1
   Else '3ph
      If x(i, 2) <= xcp Then ivar = 1 Else ivar = 0
   End If
   ypred(i) = ycp + slope * (x(i, 2) - xcp) * ivar
   For j = 2 To numxvars
      ypred(i) = ypred(i) + ivcoefs(j) * x(i, j + 1)
   Next j
   xpred(i) = x(i, 2)
Next i

'replot xy graph of y and first x var in mult regression
If G1G2But.Value = False Or (G1G2But.Value = True And pass = 1) Or (G1G2But.Value = True And pass = 2 And Main.XVarBox.ListIndex + 1 <> xfld(1)) Then
   'selects the first x variable in mult reg as the x variable for xy plot
   Main.XVarBox.ListIndex = xfld(1) - 1
   'selects the first y variable in mult reg as the y variable for xy plot
   Main.YVarBox.ListIndex = yfld - 1
   'calls xy plot
   Call XYBut_Click
End If

'draws mlr points
GraphBox.DrawWidth = 3
For i = 1 To xr
   ypnt = (ypred(i) - yminr) / yranger * 100
   xpnt = (xpred(i) - xminr) / xranger * 100
   
   If xpnt >= 0 And xpnt <= 100 And ypnt >= 0 And ypnt <= 100 Then
      If G1G2But.Value = False Then
         GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), BLACK, B
      Else
         If pass = 1 Then
            GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), BLUE, B
         Else
            GraphBox.Circle (xpnt, ypnt), 0.8, RED
         End If
      End If
   End If
Next i
ReDim ypred(1), xpred(1)
GraphBox.DrawWidth = 1

'If numpasses = 2 Then Call SetAll
If G1G2But.Value = True And pass = 2 Then Call SetAll

'resets gmodnum
If G1G2But.Value = True Then
  If gmodnum = 1 Then
     gmodnum = 2
  Else 'gmodnum = 2
     gmodnum = 1
  End If
End If

psmce:
Screen.MousePointer = 0
Unload PerDone
If Err = 6 Or Err = 11 Then
  MsgBox "Try using the Data, Modify Values menu items to rescale variables with large values so that the magnitude of those values is decreased.", , "Error"
  Resume psmes
ElseIf Err Then
  MsgBox """" + Error(Err) + """", , "Error"
  Resume psmes
End If
psmes:

End Sub

Private Sub TwoPBut_Click()
Dim i, j
Dim b1, seb1, b2, seb2, rmse, cvrmse, r2
Dim t, predint, x1, y1, x2, y2
On Error GoTo slrce

'sets model type
modtype$ = "2P"
Screen.MousePointer = 11

'fills x() and y()
Call SetAll
If G1G2But.Value = True And gmodnum = 2 Then pass = 2 Else pass = 1
If G1G2But.Value = True Then Call SetGrp
 
'prints errors
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo slrce:
ElseIf G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined.", , "Error"
   GoTo slrce:
ElseIf xfld(1) = 0 Then
   MsgBox "No X variable has been defined.", , "Error"
   GoTo slrce:
End If

'calls regression
'Call Reg(x(), xr, xc, y(), n, 1, beta(), betar, betac, sec())
'Call Inf(x(), xr, xc, y(), n, 1, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), P)
Call TwoP(x(), y(), n, b1, seb1, b2, seb2, rmse, cvrmse, r2)

'creates global variables for savings calculations
If pass = 1 Then
   premodtype$ = modtype$
   rmsepre = rmse
   npre = n
   betapre(1) = b1
   betapre(2) = b2
   modelstats = " N = " + Format$(n, "####") + "  R2 = " + Format$(r2, "0.00") + "  RMSE = " + Format$(rmse, "#,##0.00") + "  CV-RMSE = " + Format$(cvrmse, "##0.00") + "%"
   modeleqn = vn$(yfld) + " = " + Format(beta(1, 1), "#,##0.00") + "  +  " + Format(beta(2, 1), "#,##0.00") + " " + vn$(xfld(1))
End If

'prints results
'Call DWTest(2, dwr$)
sf$ = "###0.0000"
If G1G2But.Value = False Or (G1G2But.Value = True And pass = 1) Then ValueBar.Cls
Call PrintHeaderLine(gmodnum, modtype$)
ValueBar.Print " N = " + Format$(n, "####") + "   R2 = " + Format$(r2, "0.00") + "   RMSE = " + Format$(rmse, sf$) + "   CV-RMSE = " + Format$(cvrmse, "##0.0\%")
ValueBar.Print " Yint = " + Format$(b1, sf$) + " (" + Format$(seb1, sf$) + ")     Slope = " + Format$(b2, sf$) + " (" + Format$(seb2, sf$) + ")"
'ValueBar.Print " p = " + Format$(autocor, "0.00") + "   DW = " + Format$(dw, "0.00") + " (" + dwr$ + ")"
'print model
ValueBar.Print "Model: "; vn$(yfld); " = "; Format(beta(1, 1), "#,##0.00"); "  +  "; Format(beta(2, 1), "#,##0.00"); " "; vn$(xfld(1))
ValueBar.Print

'calcs prediction interval
If n <> 2 Then t = 1.96 + 2.7 / (n - 2)
predint = t * rmse * (1 + 1 / n) ^ 0.5

'calcs coordinates of model line and draws line
x1 = xminr
'y1 = beta(1, 1) + beta(2, 1) * x1
y1 = b1 + b2 * x1
x2 = xmaxr
'y2 = beta(1, 1) + beta(2, 1) * x2
y2 = b1 + b2 * x2
If y1 < yminr Then
   y1 = yminr
   'x1 = (y1 - beta(1, 1)) / beta(2, 1)
   x1 = (y1 - b1) / b2
ElseIf y1 > ymaxr Then
   y1 = ymaxr
   'x1 = (y1 - beta(1, 1)) / beta(2, 1)
   x1 = (y1 - b1) / b2
End If
If y2 < yminr Then
   y2 = yminr
   'x2 = (y2 - beta(1, 1)) / beta(2, 1)
   x2 = (y2 - b1) / b2
ElseIf y2 > ymaxr Then
   y2 = ymaxr
   'x2 = (y2 - beta(1, 1)) / beta(2, 1)
   x2 = (y2 - b1) / b2
End If
Call DrawModelLine(predint, pass, x1, y1, x2, y2)

'resets gmodnum
If G1G2But.Value = True Then
   If gmodnum = 1 Then
      gmodnum = 2
   Else 'gmodnum = 2
      gmodnum = 1
   End If
End If

'error handling
slrce:
Screen.MousePointer = 0
If Err = 6 Or Err = 11 Then
   MsgBox "Try using the Data, Modify Values menu items to rescale variables with large values so that the magnitude of those values is decreased.", , "Error"
   Resume slres
ElseIf Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume slres
End If
slres:
End Sub

Private Sub SLRBut_Click()

End Sub

Private Sub ThreePBut_Click(Index As Integer)
ReDim A(numrecs, 2), rmsei(100)
Dim numints As Integer, itration As Integer
Dim i As Integer, j As Integer
Dim n1 As Integer, n2 As Integer, count As Integer
Dim numpasses As Integer
Dim totregs, numregs
ReDim ivcoefs(7), seivcoefs(7)
On Error GoTo psmce

'set modtype$
modtype$ = "3P"
'Index = 0 for 3PC and index = 1 for 3ph

'show perdone
Screen.MousePointer = 11
PerDone.Caption = "Processing Data..."
PerDone.Show 0
DoEvents

'Fills x(), y(), mins maxs etc.
Call SetAll
If G1G2But.Value = True And gmodnum = 2 Then pass = 2 Else pass = 1
If G1G2But.Value = True Then Call SetGrp

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo psmce:
ElseIf G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined.", , "Error"
   GoTo psmce:
ElseIf xfld(1) = 0 Then
   MsgBox "No X variable has been defined.", , "Error"
   GoTo psmce:
End If

'calls 3p+mvr but with no additional independent variables
numxvars = 1
Call ThreePMVR(x(), y(), n, numxvars, Index, xcp, sexcp, ycp, seycp, slope, seslope, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)

'sets right and left slopes
If Index = 0 Then '3pc
   ls = 0: lsse = 0
   rs = slope: rsse = seslope
Else '3ph
   ls = slope: lsse = seslope
   rs = 0: rsse = 0
End If

'creates global variables for savings calcuations
If pass = 1 Then
   premodtype$ = modtype$
   rmsepre = rmse
   npre = n
   numxvarspre = numxvars
   xcppre = xcp
   ycppre = ycp
   slopepre = slope
   indexpre = Index
   ReDim ivcoefspre(7)
   For i = 2 To numxvars
      ivcoefspre(i) = ivcoefs(i)
   Next i
   modelstats = " N = " + Format$(n, "####") + "  R2 = " + Format$(r2, "0.00") + "  RMSE = " + Format$(rmse, "#,##0.00") + "  CV-RMSE = " + Format$(cvrmse, "##0.00") + "%"
   modeleqn = vn$(yfld) + " = " + Format(ycp, "#,##0.00") + "  -  " + Format(ls, "#,##0.00") + " (" + Format(xcp, "#,##0.00") + " - " + vn$(xfld(1)) + ")+  +  " + Format(rs, "#,##0.00") + " (" + vn$(xfld(1)) + " - " + Format(xcp, "#,##0.00") + ")+"
End If


'prints results
'Call DWTest(3, dwr$)
sf$ = "###0.0000"
If G1G2But.Value = False Or (G1G2But.Value = True And pass = 1) Then ValueBar.Cls
Call PrintHeaderLine(gmodnum, modtype$)
ValueBar.Print " N = " + Format$(n, "####") + "  R2 = " + Format$(r2, "0.00") + "  RMSE = " + Format$(rmse, sf$) + "  CV-RMSE = " + Format$(cvrmse, "##0.0\%")
ValueBar.Print " Ycp = " + Format$(ycp, sf$) + " (" + Format$(seycp, sf$) + ")";
ValueBar.Print "  Xcp = " + Format$(xcp, sf$) + " (" + Format$(sexcp, sf$) + ")";
ValueBar.Print "  LS = " + Format$(ls, sf$) + " (" + Format$(lsse, sf$) + ")";
ValueBar.Print "  RS = " + Format$(rs, sf$) + " (" + Format$(rsse, sf$) + ")"
'ValueBar.Print " N = " + Format$(n, "####") + "  N1 = " + Format$(n1, "####") + "  N2 = " + Format$(n2, "####");
'ValueBar.Print "  p = " + Format$(autocor, "0.00") + "  DW = " + Format$(dw, "0.00") + " (" + dwr$ + ")"
'ValueBar.Print
'print model
ValueBar.Print "Model: "; vn$(yfld); " = "; Format(ycp, "#,##0.00"); "  -  "; Format(ls, "#,##0.00"); " ("; Format(xcp, "#,##0.00"); " - "; vn$(xfld(1)); ")+  +  "; Format(rs, "#,##0.00"); " ("; vn$(xfld(1)); " - "; Format(xcp, "#,##0.00"); ")+";
'For i = 2 To numxvars
'   ValueBar.Print " + "; Format(ivcoefs(i), "#,##0.00"); " "; vn$(xfld(i)); '" X" + Format(i, "0"); '
'Next i
ValueBar.Print
ValueBar.Print

'calcs prediction interval
t = 1.96 + 2.7 / (n - 3)
predint = t * rmse * (1 + 1 / n) ^ 0.5

'draws left and right model line
GraphBox.DrawWidth = 1
For rightleft = 1 To 2
    Select Case rightleft
       Case 1
          'left model line
          slope = ls
          xstart = xminr
       Case 2
          'right model line
          slope = rs
          xstart = xmaxr
     End Select
          
    x1 = xstart
    y1 = slope * (x1 - xcp) + ycp
    
    x2 = xcp
    y2 = ycp
    
    If y1 < yminr Then
       y1 = yminr
       x1 = (y1 - ycp) / slope + xcp
    ElseIf y1 > ymaxr Then
       y1 = ymaxr
       x1 = (y1 - ycp) / slope + xcp
    End If
    
Call DrawModelLine(predint, pass, x1, y1, x2, y2)
Next rightleft

'resets gmodnum
If G1G2But.Value = True Then
  If gmodnum = 1 Then
     gmodnum = 2
  Else 'gmodnum = 2
     gmodnum = 1
  End If
End If

psmce:
Screen.MousePointer = 0
Unload PerDone
If Err = 6 Or Err = 11 Then
  MsgBox "Try using the Data, Modify Values menu items to rescale variables with large values so that the magnitude of those values is decreased.", , "Error"
  Resume psmes
ElseIf Err Then
  MsgBox """" + Error(Err) + """", , "Error"
  Resume psmes
End If
psmes:
End Sub

Private Sub TimeScalesCmd_Click()

'enable longer time-scales than current data
If timeint <= 1 Then 'yearly data
   'ToDailyCmd.enabled = True
   'ToWeeklyCmd.enabled = True
   'ToMonthlyCmd.enabled = True
   'ToYearlyCmd.enabled = True
ElseIf timeint = 2 Then 'monthly data
   'ToDailyCmd.enabled = True
   'ToWeeklyCmd.enabled = True
   'ToMonthlyCmd.enabled = True
   ToYearlyCmd.Enabled = True
ElseIf timeint = 3 Then 'daily data
   'ToDailyCmd.enabled = True
   'ToWeeklyCmd.Enabled = True
   ToMonthlyCmd.Enabled = True
   ToYearlyCmd.Enabled = True
ElseIf timeint = 4 Then
   ToDailyCmd.Enabled = True
   'ToWeeklyCmd.Enabled = True
   ToMonthlyCmd.Enabled = True
   ToYearlyCmd.Enabled = True
End If

'enable previous forms of data
If timeinthr = True Then ToHourlyCmd.Enabled = True
If timeintdy = True Then ToDailyCmd.Enabled = True
If timeintmo = True Then ToMonthlyCmd.Enabled = True
If timeintyr = True Then ToYearlyCmd.Enabled = True

End Sub

Private Sub TimeSeries2But_Click()
Dim i, j As Integer
Dim xinc, xscale, xloc, yinc, yscale, yloc
Dim xpnt, ypnt
Dim ckdate$, t$, Msg$
Dim grpindex

On Error GoTo ts2ce

'gives error if n = 0 or no time stamp variables or no dow fld has been named
If n = 0 Then
   'gives error if no y variable is selected
   MsgBox "No data is available to plot.", , "Error"
   GoTo ts2es:
ElseIf timeint = 0 Then
   MsgBox "No time stamp variable names (Yr, Mo, Dy or Hr) have been defined.", , "Error"
   GoTo ts2es:
ElseIf G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined.", , "Error"
   GoTo ts2es:
End If

'sets graph type for point find procedure
Screen.MousePointer = 11
gtype$ = "TS"
ZoomCmd.Enabled = True
GraphBox.Scale (-15, 110)-(115, -20)
GraphBox.FontName = "Helv"
GraphBox.FontSize = 8

'sets yfld
'calcsaveddisplay and calcleadisplay are global variables set in the CalcSavingsCmd and CalcLEACmd modules
'When set to "True", they override the y-value listbox indicator of which variable to graph
'and specify a new yfld where predicted use has been stored
If calcsavedisplay = True Then
   yfld = numflds
ElseIf calcleadisplay = True Then
   yfld = numflds + AdditionalVar
Else 'get yfld from standard list box
   yfld = YVarBox.ListIndex + 1
End If

'if not zoommode then use all data
If zoommode = False Then
  i1 = 1: i2 = numrecs
End If

'sets y-scales
If SetScalesDB.NumYLabelsTB.Text <> "" Then
   'use manual scaling
   
   'clears right axis and labels
   GraphBox.Line (82, 109)-(115, 101), BackColor, BF
   GraphBox.Line (98, 101)-(115, -1), BackColor, BF
   
   'draws new right axis
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
   GraphBox.Line (100, 0)-(100, 100)
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
   GraphBox.Line (80, 105)-(83, 105), BLACK
   GraphBox.CurrentX = 85
   GraphBox.CurrentY = 108
   GraphBox.Print YVarBox.Text
   
   'gets num y labels, yminr, ymaxr from text box
   numylabels = Val(SetScalesDB.NumYLabelsTB.Text)
   ymaxr = Val(SetScalesDB.YmaxTB.Text)
   yminr = Val(SetScalesDB.YminTB.Text)
   yranger = ymaxr - yminr

   'set y increment
   yinc = yranger / numylabels
   
   'finds format of x and y scale labels
   If Abs(yminr) <= Abs(ymaxr) Then ym = Abs(ymaxr) Else ym = Abs(yminr)
   If ym < 1 Then
       fy$ = "0.0000"
   ElseIf ym < 10 Then
       fy$ = "0.00"
   ElseIf ym < 50 Then
       fy$ = "#0.0"
   Else
       fy$ = "#,##0"
   End If
   
   'prints y axis scale labels
   For i = 0 To numylabels
      yscale = yminr + i * yinc
      yloc = (yscale - yminr) / yranger * 100
      GraphBox.CurrentX = 101
      GraphBox.CurrentY = yloc + 3
      GraphBox.Print Format$(yscale, fy$)
      GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
      GraphBox.Line (99, yloc)-(100, yloc)
      If GridCmd.Checked = -1 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
         GraphBox.Line (0, yloc)-(100, yloc)
      End If
   Next i

ElseIf SetScalesDB.SingleScalesCB.Value = 1 Or calcsavedisplay = True Or calcleadisplay = True Then
   'use single scales
   
   'finds ymin1 and ymax1 for left(first) axis
   For i = i1 To i2
      If d(i, yfldo) <> ndflag Then
          j = j + 1
          If j = 1 Then
             ymin1 = d(i, yfldo)
             ymax1 = d(i, yfldo)
          End If
          If d(i, yfldo) < ymin1 Then ymin1 = d(i, yfldo)
          If d(i, yfldo) > ymax1 Then ymax1 = d(i, yfldo)
      End If
   Next i
   
   'finds new yminr and ymaxr for right axis
   For i = i1 To i2
      If d(i, yfld) <> ndflag Then
          j = j + 1
          If j = 1 Then
             ymin2 = d(i, yfld)
             ymax2 = d(i, yfld)
          End If
          If d(i, yfld) < ymin2 Then ymin2 = d(i, yfld)
          If d(i, yfld) > ymax2 Then ymax2 = d(i, yfld)
      End If
   Next i

   'use greater of left or right mins and maxs
   If ymin1 <= ymin2 Then ymin = ymin1 Else ymin = ymin2
   If ymax1 >= ymax2 Then ymax = ymax1 Else ymax = ymax2
   
   'round ymax up and ymin down if necessary
   emax$ = Format$(ymax, "0.0E+00")
   rootmax = 1 + Val(Left$(emax$, 1))
   powermax = Val(Right(emax$, 3))
   ymaxr = rootmax * 10 ^ powermax
   
   'finds yminr, yroot and ypower
   If ymin < 0 Then
   
      'rounds ymin down to even number
      emax$ = Format$(ymin, "0.0E+00")
      rootmin = -1 + Val(Left$(emax$, 2))
      powermin = Val(Right(emax$, 3))
      yminr = rootmin * 10 ^ powermin
      
      'sets power and root
      If powermax > powermin Then
         ypower = powermax
         yroot = rootmax + 1
         yminr = -1 * 10 ^ ypower
      ElseIf powermin > powermax Then
         ypower = powermin
         yroot = Abs(rootmin) + 1
         ymaxr = 1 * 10 ^ ypower
      ElseIf powermin = powermax Then
         ypower = powermax
         yroot = rootmax - rootmin
      End If
   Else 'ymin > 0
      
      'sets ymin, power and root
      yminr = 0
      yroot = rootmax
      ypower = powermax
   End If
   yranger = ymaxr - yminr

   'selects interval to determine how many labels
   If yroot >= 5 Then
      yintindex = 1
   Else
      yintindex = 0.5
   End If
   
   'selects yfmt
   If ymaxr - yminr > 5 Then
      yfmt = "#,###,##0"
   ElseIf ymaxr - yminr > 1 Then
      yfmt = "0.00"
   Else
      yfmt = "0.000"
   End If


   'fill yminss, ymaxss for pass to timeseries1
   yminss = yminr
   ymaxss = ymaxr

   If calcleadisplay = True Then
      'do nothing
   Else
      'call timeseries1 to plot first series (with ss=true)
      ss = True
      TimeSeriesBut.Value = True
      ss = False
      yfld = YVarBox.ListIndex + 1
   End If
   
   'sets yfld
   If calcsavedisplay = True Then
      yfld = numflds
   ElseIf calcleadisplay = True Then
      yfld = numflds + AdditionalVar
   Else
      yfld = YVarBox.ListIndex + 1
   End If
   
   'draws new right axis
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
   GraphBox.Line (100, 0)-(100, 100)
   
   'draws label(s) for new trace(s)
   If calcsavedisplay = True Then
      GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
      GraphBox.Line (80, 105)-(83, 105), MAGENTA
      GraphBox.CurrentX = 85
      GraphBox.CurrentY = 108
      GraphBox.ForeColor = MAGENTA
      GraphBox.Print "Adjusted Baseline"
      GraphBox.ForeColor = BLACK
   ElseIf calcleadisplay = True Then
      If AdditionalVar = 1 Then
         'GraphBox.Print "Predicted"
      ElseIf AdditionalVar = 2 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
         GraphBox.Line (23, 105)-(29, 105), RED
         GraphBox.CurrentX = 30
         GraphBox.CurrentY = 108
         GraphBox.ForeColor = RED
         GraphBox.Print "Fac Use"
         GraphBox.ForeColor = BLACK
      ElseIf AdditionalVar = 3 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 3
         GraphBox.Line (53, 105)-(59, 105), BLUE
         GraphBox.CurrentX = 60
         GraphBox.CurrentY = 108
         GraphBox.ForeColor = BLUE
         GraphBox.Print "Wea-dep Use"
         GraphBox.ForeColor = BLACK
      ElseIf AdditionalVar = 4 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 4
         GraphBox.Line (83, 105)-(89, 105), GREEN
         GraphBox.CurrentX = 90
         GraphBox.CurrentY = 108
         GraphBox.ForeColor = GREEN
         GraphBox.Print "Prod-dep Use"
         GraphBox.ForeColor = BLACK
      End If

      
      If 1 = 2 Then 'old stacked labels
      If AdditionalVar = 1 Then
         'GraphBox.Print "Predicted"
      ElseIf AdditionalVar = 2 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
         GraphBox.Line (73, 106)-(78, 106), RED
         GraphBox.CurrentX = 80
         GraphBox.CurrentY = 109
         GraphBox.ForeColor = RED
         GraphBox.Print "Fac Use"
         GraphBox.ForeColor = BLACK
      ElseIf AdditionalVar = 3 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 3
         GraphBox.Line (73, 102)-(78, 102), BLUE
         GraphBox.CurrentX = 80
         GraphBox.CurrentY = 105
         GraphBox.ForeColor = BLUE
         GraphBox.Print "Wea-Dep Use"
         GraphBox.ForeColor = BLACK
      ElseIf AdditionalVar = 4 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 4
         GraphBox.Line (73, 98)-(78, 98), GREEN
         GraphBox.CurrentX = 80
         GraphBox.CurrentY = 101
         GraphBox.ForeColor = GREEN
         GraphBox.Print "Prod-Dep Use"
         GraphBox.ForeColor = BLACK
      End If
      End If
      
      GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 1
   Else 'normal timeseries 2
      GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
      GraphBox.Line (80, 105)-(83, 105), BLACK
      GraphBox.CurrentX = 85
      GraphBox.CurrentY = 108
      GraphBox.Print YVarBox.Text
      GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 1
   End If
   

   'plots right axis labels
   For i = 0 To yroot Step yintindex
      GraphBox.CurrentX = 102
      ylabel = yminr + (i * 10 ^ ypower)
      ylabelpos = (i / yroot * 100)
      GraphBox.CurrentY = ylabelpos + 1
      GraphBox.Print Format(ylabel, yfmt)
      GraphBox.Line (99, ylabelpos)-(100, ylabelpos)
      If GridCmd.Checked = -1 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
         GraphBox.Line (0, ylabelpos)-(100, ylabelpos)
      End If
   Next i
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0


Else 'use auto-scaling for y2 axis
   
   Call SetAll

   'clears right axis and labels
   GraphBox.Line (82, 109)-(115, 101), BackColor, BF
   GraphBox.Line (98, 101)-(115, -1), BackColor, BF
   
   'draws new right axis
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
   GraphBox.Line (100, 0)-(100, 100)
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
   GraphBox.Line (80, 105)-(83, 105), BLACK
   GraphBox.CurrentX = 85
   GraphBox.CurrentY = 108
   GraphBox.Print YVarBox.Text

   'plots right axis labels
   For i = 0 To yroot Step yintindex
      GraphBox.CurrentX = 102
      ylabel = yminr + (i * 10 ^ ypower)
      ylabelpos = (i / yroot * 100)
      GraphBox.CurrentY = ylabelpos + 1
      GraphBox.Print Format(ylabel, yfmt)
      GraphBox.Line (99, ylabelpos)-(100, ylabelpos)
      If GridCmd.Checked = -1 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
         GraphBox.Line (0, ylabelpos)-(100, ylabelpos)
      End If
   Next i
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
End If

'Draws second time-series line
GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
For i = i1 To i2
   xpnt = (i - i1) / (i2 - i1 + 1) * 100
   ypnt = (d(i, yfld) - yminr) / yranger * 100
   
   If i = 1 Or ld = ndflag Or d(i, yfld) = ndflag Then
      
      'no data so skip
      lxpnt = xpnt: lypnt = ypnt: ld = d(i, yfld)
   
   Else
      
      'draw line
      If G1G2But.Value = True Then 'grouping is on
        
         'grpindex makes a line 'leaving' a point the same group color as the point
         'and makes the cursor the right color
         If i = 1 Then grpindex = 1 Else grpindex = i - 1
         If grpd(grpindex) = 1 Then
            GraphBox.Line (lxpnt, lypnt)-(xpnt, ypnt), CYAN
         Else
            GraphBox.Line (lxpnt, lypnt)-(xpnt, ypnt), MAGENTA
         End If
      
      Else 'grouping is off: draw black line
            If calcleadisplay = True Then
               If AdditionalVar = 2 Then
                  GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
                  GraphBox.Line (lxpnt, lypnt)-(xpnt, ypnt), RED
               ElseIf AdditionalVar = 3 Then
               GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 3
                  GraphBox.Line (lxpnt, lypnt)-(xpnt, ypnt), BLUE
               ElseIf AdditionalVar = 4 Then
                  GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 4
                  GraphBox.Line (lxpnt, lypnt)-(xpnt, ypnt), GREEN
               End If
               GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 1
            ElseIf calcsavedisplay = True Then
               GraphBox.Line (lxpnt, lypnt)-(xpnt, ypnt), BLACK
            Else 'normal timeseries 2
               GraphBox.Line (lxpnt, lypnt)-(xpnt, ypnt), BLACK
            End If
      End If
      lxpnt = xpnt: lypnt = ypnt: ld = d(i, yfld)

   End If
Next i

ts2ce:
If Err = 6 Or Err = 11 Then
   MsgBox "No data available to plot.", , "Error"
   Resume ts2es
ElseIf Err = 9 Then
   MsgBox "Draw zoom-rectangle within the x-scale and y-scale limits.  Select Graphics, Zoom Out and try again.", , "Error"
   Resume ts2es
ElseIf Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume ts2es
End If

ts2es:
If zoommode = True Then
   Screen.MousePointer = 0
   GraphBox.MousePointer = 2
Else
   Screen.MousePointer = 0
End If

End Sub

Private Sub TimeSeriesBut_Click()
Dim i, m As Integer
Dim xinc, xscale, xloc, yinc, yscale, yloc
Dim xpnt, ypnt
Dim ckdate$, t$, Msg$
Dim grpindex
On Error GoTo tsce

Screen.MousePointer = 11

If ss = False Then
   'calls SetAll to set x(),y(),n,minr,maxr,root,power,fmt,intindex
   Call SetAll
End If

'exits if no data or no timestamp data
If n = 0 Then
   'gives error if no y variable is selected
   MsgBox "No data is available to plot.", , "Error"
   GoTo tses:
ElseIf timeint = 0 Then
   'gives error if no time stamp variables are defined
   MsgBox "No time stamp variable names (Yr, Mo, Dy or Hr) have been defined.", , "Error"
   GoTo tses:
ElseIf G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined.", , "Error"
   GoTo tses:
End If

'sets graph type for point find procedure
gtype$ = "TS"
ZoomCmd.Enabled = True

'clears graphbox and selects font
GraphBox.Cls
GraphBox.Scale (-15, 110)-(115, -20)
GraphBox.FontName = "Helv"
GraphBox.FontSize = 8

'Draw axis and labels
GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
GraphBox.Line (0, 0)-(0, 100)
GraphBox.Line (0, 0)-(100, 0)
GraphBox.CurrentX = 90
GraphBox.CurrentY = -9.5
GraphBox.Print "Date-Time"

'prints y-axis title
GraphBox.CurrentX = -13
GraphBox.CurrentY = 108
If ss = False Then
   GraphBox.Print YVarBox.Text
Else 'ss = true
   'reprints orig yvarname for ts2-single scales procedure
   GraphBox.Print ynameo$
End If

'resets i1 & 12 to all data if not zoom-mode
If zoommode = False Then
  i1 = 1: i2 = numrecs
End If

'prints x-axis scales
xinc = (i2 - i1 + 1) / 4
t$ = "##"
For k = i1 To i2 Step xinc
   m = k
   'finds the calandar date and print it at each x label
   Call CalendarDate(m, ckdate$)
   xloc = (m - i1) / (i2 - i1 + 1) * 100
   GraphBox.CurrentX = xloc - 6
   GraphBox.CurrentY = -5
   GraphBox.Print ckdate$
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
   GraphBox.Line (xloc, 0)-(xloc, 1)
   If GridCmd.Checked = -1 Then
      GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
      GraphBox.Line (xloc, 0)-(xloc, 100)
   End If
Next k
'adds last x label
Call CalendarDate(i2, ckdate$)
xloc = (i2 - i1) / (i2 - i1 + 1) * 100
GraphBox.CurrentX = xloc - 6
GraphBox.CurrentY = -5
GraphBox.Print ckdate$
GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
GraphBox.Line (xloc, 0)-(xloc, 1)
If GridCmd.Checked = True Then
    GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
    GraphBox.Line (xloc, 0)-(xloc, 100)
End If

'sets y-scales
If ss = True Then 'replots orig yvar for ts2-single scales procedure

   yminr = yminss
   ymaxr = ymaxss
   yranger = ymaxr - yminr
   yfld = yfldo

   'plots y labels
   For i = 0 To yroot Step yintindex
      GraphBox.CurrentX = -14
      ylabel = yminr + (i * 10 ^ ypower)
      ylabelpos = (i / yroot * 100)
      GraphBox.CurrentY = ylabelpos + 1
      GraphBox.Print Format(ylabel, yfmt)
      GraphBox.Line (0, ylabelpos)-(1, ylabelpos)
      If GridCmd.Checked = -1 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
         GraphBox.Line (0, ylabelpos)-(100, ylabelpos)
      End If
   Next i
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0

ElseIf SetScalesDB.NumYLabelsTB.Text <> "" Then 'use manually-set y scales

   'sets num y labels (yminr & ymaxr set in SetAllBut)
   numylabels = Val(SetScalesDB.NumYLabelsTB.Text)
   
   'set increment for labels
   yinc = yranger / numylabels
   
   'prints y axis scale labels
   For i = 0 To numylabels
      yscale = yminr + i * yinc
      yloc = (yscale - yminr) / yranger * 100
      GraphBox.CurrentX = -10
      GraphBox.CurrentY = yloc + 3
      GraphBox.Print Format$(yscale, yfmt)
      GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
      GraphBox.Line (0, yloc)-(1, yloc)
      If GridCmd.Checked = -1 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
         GraphBox.Line (0, yloc)-(100, yloc)
      End If
   Next i

Else 'use auto-scaling for y axis

   'plots y labels
   For i = 0 To yroot Step yintindex
      GraphBox.CurrentX = -14
      ylabel = yminr + (i * 10 ^ ypower)
      ylabelpos = (i / yroot * 100)
      GraphBox.CurrentY = ylabelpos + 1
      GraphBox.Print Format(ylabel, yfmt)
      GraphBox.Line (0, ylabelpos)-(1, ylabelpos)
      If GridCmd.Checked = -1 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
         GraphBox.Line (0, ylabelpos)-(100, ylabelpos)
      End If
   Next i
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
End If

'prints title and draws 0 line for residuals
If yfld > numfldsd And yfld <= numfldsd + 3 Then
    GraphBox.CurrentX = 39
    GraphBox.CurrentY = 108
    GraphBox.Print "Residuals (Y - Ypredicted)"
    yloc = (0 - yminr) / yranger * 100
    GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
    GraphBox.Line (0, yloc)-(100, yloc)
    GraphBox.CurrentX = 101
    GraphBox.CurrentY = yloc + 3
    GraphBox.Print "0"
End If

'draws timeseries line
GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
For i = i1 To i2
   
   'find location
   xpnt = (i - i1) / (i2 - i1 + 1) * 100
   ypnt = (d(i, yfld) - yminr) / yranger * 100
   
   'either save point or drawline
   If i = i1 Or ld = ndflag Or d(i, yfld) = ndflag Then 'save point
      'record last good point for nodata points
      lxpnt = xpnt: lypnt = ypnt: ld = d(i, yfld)
   Else 'draw line
      If G1G2But.Value = True Then
         'grpindex makes a line 'leaving' a point the same group color as the point
         'and makes the cursor the right color
         If i = 1 Then grpindex = 1 Else grpindex = i - 1
         If grpd(grpindex) = 1 Then
            GraphBox.Line (lxpnt, lypnt)-(xpnt, ypnt), BLUE
         Else
            'group 2 so draw blue line
            GraphBox.Line (lxpnt, lypnt)-(xpnt, ypnt), RED
         End If
      Else
         'no grouping so draw black line
         GraphBox.Line (lxpnt, lypnt)-(xpnt, ypnt), BLACK
      End If
      
      'record last point
      lxpnt = xpnt: lypnt = ypnt: ld = d(i, yfld)
   End If
Next i

tsce:
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume tses
End If
tses:
yfldo = yfld
ynameo$ = YVarBox.Text
Screen.MousePointer = 0
If zoommode = True Then GraphBox.MousePointer = 2
End Sub

Private Sub ToDailyCmd_Click()
Dim i, j, curdy, mustavg
ReDim ndf(numflds)
On Error GoTo todailyce

'prepare perdone
Dim regnum, totregs, increment
increment = 10
PerDone.Caption = "Processing Data..."
Screen.MousePointer = 11

'hourly to daily

If timeintdy = True Then
   
   'show perdone
   PerDone.Show 0
   DoEvents
   totregs = numrecsdy
   
   Open "ddy.tmp" For Input As #1
   numrecs = numrecsdy
   ReDim d(numrecs, numflds)
   For i = 1 To numrecs
      For j = 1 To numflds
         Input #1, d(i, j)
      Next j
      
      'update perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, increment)
   Next i
   Close

Else

   'show time scale options
   Screen.MousePointer = 0
   If tsoptsshown <> True Then TSOptsDB.Show MODAL
   If tsoptscancel = True Then GoTo todailyce
   Screen.MousePointer = 11
   
   'show perdone
   PerDone.Show 0
   DoEvents
   If timeinthr <> True Then totregs = 2 * numrecs Else totregs = numrecs
   
   If timeinthr <> True Then 'save orig hourly data to file
      Open "dhr.tmp" For Output As #1
      For i = 1 To numrecs
         For j = 1 To numflds
            Print #1, d(i, j);
         Next j
         Print #1,
         
         'update perdone
         regnum = regnum + 1
         Call UpdatePerDone(regnum, totregs, increment)
      Next i
      Close
      timeinthr = True
      ToHourlyCmd.Enabled = True
      numrecshr = numrecs
   End If
   
   Call TimeScaleChange("todaily", 24, dyfld, "ddy.tmp", numrecsdy, regnum, totregs, increment)
   
   timeintdy = True
   ToDailyCmd.Enabled = True
End If

timeint = 3
todailyce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume todailyes
End If
todailyes:
End Sub

Private Sub ToHourlyCmd_Click()
Dim i, j
On Error GoTo tohourlyce

'prepare perdone
Dim regnum, totregs, increment
increment = 10
PerDone.Caption = "Processing Data..."

Screen.MousePointer = 11

If timeinthr = True Then
   'show perdone
   PerDone.Show 0
   DoEvents
   totregs = numrecshr

   Open "dhr.tmp" For Input As #1
   numrecs = numrecshr
   ReDim d(numrecs, numflds)
   For i = 1 To numrecs
      For j = 1 To numflds
         Input #1, d(i, j)
      Next j
      
      'update perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, increment)
   Next i
   Close
   timeint = 4

End If

tohourlyce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume tohourlyes
End If
tohourlyes:
End Sub
Private Sub ToMonthlyCmd_Click()
Dim i, j, curmo, periods, mustavg
ReDim ndf(numflds)
On Error GoTo tomonthlyce

'prepare perdone
Dim regnum, totregs, increment
increment = 10
PerDone.Caption = "Processing Data..."
Screen.MousePointer = 11

If timeintmo = True Then
   'show perdone
   PerDone.Show 0
   DoEvents
   totregs = numrecsmo
   
   Open "dmo.tmp" For Input As #1
   numrecs = numrecsmo
   ReDim d(numrecs, numflds)
   For i = 1 To numrecs
      For j = 1 To numflds
         Input #1, d(i, j)
      Next j
     
      'update perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, increment)
   Next i
   Close

Else

   'show time scale options
   Screen.MousePointer = 0
   If tsoptsshown <> True Then TSOptsDB.Show MODAL
   If tsoptscancel = True Then GoTo tomonthlyce
   Screen.MousePointer = 11

   'show perdone
   PerDone.Show 0
   DoEvents
   totregs = 2 * numrecs
   
   If timeint = 4 Then
      'hourly to monthly
      periods = 22 * 27
      If timeinthr <> True Then 'print orig hourly data to file
         Open "dhr.tmp" For Output As #1
         For i = 1 To numrecs
            For j = 1 To numflds
               Print #1, d(i, j);
            Next j
            Print #1,
            'update perdone
            regnum = regnum + 1
            Call UpdatePerDone(regnum, totregs, increment)
         Next i
         Close
         timeinthr = True
         ToHourlyCmd.Enabled = True
         numrecshr = numrecs
      End If
   
   ElseIf timeint = 3 Then
      'daily to monthly
      periods = 28
      
      If timeintdy <> True Then 'print orig hourly data to file
         Open "ddy.tmp" For Output As #1
         For i = 1 To numrecs
            For j = 1 To numflds
               Print #1, d(i, j);
            Next j
            Print #1,
            'update perdone
            regnum = regnum + 1
            Call UpdatePerDone(regnum, totregs, increment)
         Next i
         Close
         timeintdy = True
         ToDailyCmd.Enabled = True
         numrecsdy = numrecs
      End If
   End If
   
   Call TimeScaleChange("tomonthly", periods, mofld, "dmo.tmp", numrecsmo, regnum, totregs, increment)
   
   timeintmo = True
   ToMonthlyCmd.Enabled = True
End If

timeint = 2
tomonthlyce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume tomonthlyes
End If
tomonthlyes:

End Sub

Private Sub ToYearlyCmd_Click()
Dim i, j, curyr, periods, mustavg
ReDim ndf(numflds)
On Error GoTo toyearlyce

'prepare perdone
Dim regnum, totregs, increment
increment = 10
PerDone.Caption = "Processing Data..."
Screen.MousePointer = 11

If timeintyr = True Then
   
   'show perdone
   PerDone.Show 0
   DoEvents
   totregs = numrecsyr

   Open "dyr.tmp" For Input As #1
   numrecs = numrecsyr
   ReDim d(numrecs, numflds)
   For i = 1 To numrecs
      For j = 1 To numflds
         Input #1, d(i, j)
      Next j
      
      'update perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, increment)
   Next i
   Close

Else

   'show time scale options
   Screen.MousePointer = 0
   If tsoptsshown <> True Then TSOptsDB.Show MODAL
   If tsoptscancel = True Then GoTo toyearlyce
   Screen.MousePointer = 11
   
   'show perdone
   PerDone.Show 0
   DoEvents
   totregs = 2 * numrecs

   If timeint = 4 Then
      'hourly to yearly
      periods = 8760
      
      If timeinthr <> True Then 'print orig hourly data to file
         Open "dhr.tmp" For Output As #1
         For i = 1 To numrecs
            For j = 1 To numflds
               Print #1, d(i, j);
            Next j
            Print #1,
            
            'update perdone
            regnum = regnum + 1
            Call UpdatePerDone(regnum, totregs, increment)
         Next i
         Close
         timeinthr = True
         ToHourlyCmd.Enabled = True
         numrecshr = numrecs
      End If
   
   ElseIf timeint = 3 Then
      'daily to yearly
      periods = 365
      
      If timeintdy <> True Then  'print orig hourly data to file
         Open "ddy.tmp" For Output As #1
         For i = 1 To numrecs
            For j = 1 To numflds
               Print #1, d(i, j);
            Next j
            Print #1,
            
            'update perdone
            regnum = regnum + 1
            Call UpdatePerDone(regnum, totregs, increment)
         Next i
         Close
         timeintdy = True
         ToDailyCmd.Enabled = True
         numrecsdy = numrecs
      End If
   
   ElseIf timeint = 2 Then
      
      'monthly to yearly
      periods = 12
      
      If timeintmo <> True Then 'print orig hourly data to file
         Open "dmo.tmp" For Output As #1
         For i = 1 To numrecs
            For j = 1 To numflds
               Print #1, d(i, j);
            Next j
            Print #1,
            
            'update perdone
            regnum = regnum + 1
            Call UpdatePerDone(regnum, totregs, increment)
         Next i
         Close
         timeintmo = True
         ToMonthlyCmd.Enabled = True
         numrecsmo = numrecs
      End If
   
   End If
   
   Call TimeScaleChange("toyearly", periods, yrfld, "dyr.tmp", numrecsyr, regnum, totregs, increment)
   
   timeintyr = True
   ToYearlyCmd.Enabled = True
End If

timeint = 1
toyearlyce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume toyearlyes
End If
toyearlyes:

End Sub

Private Sub VarNamesCmd_Click()

DefineDB.Show 1

End Sub

Private Sub ViewDatesCmd_Click()
Dim t$, sdate$, edate$, sddt$, eddt$, g$

t$ = "##"

'prints output to valuebar
ValueBar.Cls
ValueBar.Print " " + UCase$(filename$);
'finds and prints start and end dates
Call CalendarDate(1, sdate$)
Call CalendarDate(numrecs, edate$)
ValueBar.Print "    Start = " + sdate$ + "     End= " + edate$
'finds and prints deleted days
For i = 1 To numbdate
    If i = 1 Then
        ValueBar.Print " Deleted:  ";
    ElseIf i = 4 Or i = 7 Or i = 10 Or i = 13 Or i = 16 Then
        ValueBar.Print "               ";
    End If

    Call SerToCal(SDD(i), sddt$)
    Call SerToCal(EDD(i), eddt$)
    ValueBar.Print sddt$ + "-" + eddt$ + "     ";
    If i = 3 Or i = 6 Or i = 9 Or i = 12 Or i = 15 Then ValueBar.Print
Next i
If numbdate > 0 Then ValueBar.Print
'finds and prints groups
If grpdefined = True Then
   If grptype = "dow" Then
      ValueBar.Print " Weekday / Weekend Grouping"
   Else 'grptype = "date"
      'print start date
      If g1start = 0 Then
         Call CalendarDate(1, ckdate$)
      Else
         Call SerToCal(g1start, ckdate$)
      End If
      ValueBar.Print " " + g1name$ + " Start Date: " + ckdate$;
      
      'print end date
      If g1end = 9999999999# Then
         Call CalendarDate(numrecs, ckdate$)
      Else
         Call SerToCal(g1end, ckdate$)
      End If
      ValueBar.Print "     " + g1name$ + " End Date: " + ckdate$
   End If
End If


'prints output to logfile
If SaveLogCmd.Checked = True Then
    Print #3, " " + UCase$(filename$);
    Print #3, "    Start = " + sdate$ + "     End= " + edate$
    'finds and prints deleted days
    For i = 1 To numbdate
        If i = 1 Then
           Print #3, " Deleted:  ";
        ElseIf i = 4 Or i = 7 Or i = 10 Or i = 13 Or i = 16 Then
           Print #3, "               ";
        End If
    
        Call SerToCal(SDD(i), sddt$)
        Call SerToCal(EDD(i), eddt$)
        Print #3, sddt$ + "-" + eddt$ + "     ";
        If i = 3 Or i = 6 Or i = 9 Or i = 12 Or i = 15 Then Print #3,
    Next i
    If numbdate > 0 Then Print #3,
   If grpdefined = True Then
      If grptype = "dow" Then
         Print #3, " Weekday / Weekend Grouping"
      Else 'grptype = "date"
          Call SerToCal(g1start, ckdate$)
          Print #3, " " + g1name$ + " Start Date: " + ckdate$;
          Call SerToCal(g1end, ckdate$)
          Print #3, "     " + g1name$ + " End Date: " + ckdate$
       End If
    End If
    Print #3,: Print #3,
End If

End Sub

Private Sub XYAnimateBut_Click()

PointCmd.Checked = True
ZoomCmd.Checked = False
GraphBox.MousePointer = 0
ZoomTimeOutCmd.Enabled = False
zoommode = False


ANCNTLDB.Show 0

End Sub

Private Sub XYBut_Click()
Dim i
Dim xinc, xscale, xloc, yinc, yscale, yloc
Dim xpnt, ypnt
Dim ylabel, ylabelpos, xlabel, xlabelpos

Screen.MousePointer = 11
Call SetAll

PointCmd.Checked = True
ZoomCmd.Checked = False
GraphBox.MousePointer = 0
ZoomTimeOutCmd.Enabled = False
zoommode = False

If n = 0 Or xranger = 0 Then
   MsgBox "No data is available to plot.", , "Error"
   GoTo xyes:
ElseIf G1G2But.Value = True And grpdefined = False Then
   MsgBox "No grouping has been defined.", , "Error"
   GoTo xyes:
End If

'sets graph type for point find procedure
gtype$ = "XY"
ZoomCmd.Enabled = False
xcur = 0: ycur = 0: newxygraph = True

GraphBox.Cls
GraphBox.Scale (-15, 110)-(115, -20)
GraphBox.FontName = "Helv"
GraphBox.FontSize = 8

'Draw axis and labels
GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
GraphBox.Line (0, 0)-(0, 100)
GraphBox.Line (0, 0)-(100, 0)
GraphBox.CurrentX = -13
GraphBox.CurrentY = 108
GraphBox.Print YVarBox.Text
GraphBox.CurrentX = 90
GraphBox.CurrentY = -9.5
GraphBox.Print XVarBox.Text
If GridCmd.Checked = -1 Then
   GraphBox.Line (100, 0)-(100, 100)
   GraphBox.Line (0, 100)-(100, 100)
End If

'sets x-scales
If SetScalesDB.NumXLabelsTB.Text <> "" Then 'user defined scaling
   
   numxlabels = Val(SetScalesDB.NumXLabelsTB.Text)
   xinc = xranger / numxlabels
   For i = 0 To numxlabels
      'prints x axis scale labels
      xscale = xminr + i * xinc
      xloc = (xscale - xminr) / xranger * 100
      GraphBox.CurrentX = xloc
      GraphBox.CurrentY = -5
      GraphBox.Print Format$(xscale, xfmt)
      GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
      GraphBox.Line (xloc, 0)-(xloc, 1)
      If GridCmd.Checked = -1 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
         GraphBox.Line (xloc, 0)-(xloc, 100)
      End If
   Next i

Else 'use auto-scaling

   'prints x-labels
   For i = 0 To xroot Step xintindex
      GraphBox.CurrentY = -0.5
      xlabel = xminr + i * 10 ^ xpower
      xlabelpos = i / xroot * 100
      GraphBox.CurrentX = xlabelpos - 2
      GraphBox.Print Format(xlabel, xfmt)
      GraphBox.Line (xlabelpos, 1)-(xlabelpos, 0)
      If GridCmd.Checked = -1 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
         GraphBox.Line (xlabelpos, 0)-(xlabelpos, 100)
      End If
   Next i
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0

End If

'sets y-scales
If SetScalesDB.NumYLabelsTB.Text <> "" Then 'use manual scaling
   
   numylabels = Val(SetScalesDB.NumYLabelsTB.Text)
   yinc = yranger / numylabels
   'prints y axis scale labels
   For i = 0 To numylabels
      yscale = yminr + i * yinc
      yloc = (yscale - yminr) / yranger * 100
      GraphBox.CurrentX = -10
      GraphBox.CurrentY = yloc + 3
      GraphBox.Print Format$(yscale, yfmt)
      GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
      GraphBox.Line (0, yloc)-(1, yloc)
      If GridCmd.Checked = -1 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
         GraphBox.Line (0, yloc)-(100, yloc)
      End If
   Next i

Else 'use auto-scaling for y axis

   'plots y labels
   For i = 0 To yroot Step yintindex
      GraphBox.CurrentX = -14
      ylabel = yminr + (i * 10 ^ ypower)
      ylabelpos = (i / yroot * 100)
      GraphBox.CurrentY = ylabelpos + 1
      GraphBox.Print Format(ylabel, yfmt)
      GraphBox.Line (0, ylabelpos)-(1, ylabelpos)
      If GridCmd.Checked = -1 Then
         GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 2
         GraphBox.Line (0, ylabelpos)-(100, ylabelpos)
      End If
   Next i
   GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
End If

'if graph is residuals then prints title and draws 0 line
If yfld > numfldsd And yfld <= numfldsd + 3 Then
    GraphBox.CurrentX = 39
    GraphBox.CurrentY = 108
    GraphBox.Print "Residuals (Y - Ypredicted)"
    yloc = (0 - yminr) / yranger * 100
    GraphBox.DrawWidth = 1: GraphBox.DrawStyle = 0
    GraphBox.Line (0, yloc)-(100, yloc)
    GraphBox.CurrentX = 101
    GraphBox.CurrentY = yloc + 3
    GraphBox.Print "0"
End If

'Draw points
GraphBox.DrawWidth = 1
For i = 1 To n
    xpnt = (x(i, 2) - xminr) / xranger * 100
    ypnt = (y(i, 1) - yminr) / yranger * 100

    If G1G2But.Value = False Then
       GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), BLACK, B
    Else
       If grpxy(i) = 1 Then
          GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), BLUE, B
       Else
          GraphBox.Circle (xpnt, ypnt), 0.8, RED
       End If
    End If
Next i

xyes:
Screen.MousePointer = 0
GraphBox.MousePointer = 0
End Sub

Private Sub ZoomCmd_Click()

ZoomCmd.Checked = True
PointCmd.Checked = False
GraphBox.MousePointer = 2

'   If ZoomCmd.Checked = True Then
'      ZoomCmd.Checked = False
'   ElseIf ZoomCmd.Checked = False Then
'      ZoomCmd.Checked = True
'   End If

'If ZoomCmd.Checked = True Then
'    GraphBox.MousePointer = 2
'Else
'    GraphBox.MousePointer = 0
'End If

End Sub

Private Sub ZoomTimeOutCmd_Click()
'   ZoomCmd.Checked = False
'   GraphBox.MousePointer = 0
   zoommode = False
   TimeSeriesBut.Value = True
   ZoomTimeOutCmd.Enabled = False
End Sub

