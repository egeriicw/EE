VERSION 5.00
Begin VB.Form ANCNTLDB 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Animation Control Box"
   ClientHeight    =   915
   ClientLeft      =   210
   ClientTop       =   900
   ClientWidth     =   7515
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
   ScaleHeight     =   915
   ScaleWidth      =   7515
   Begin VB.CheckBox ClearScreenBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "ClearScreen"
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   75
      TabIndex        =   8
      Top             =   225
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton BeginBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Begin"
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton QuitBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Quit"
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton ContinueBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Continue"
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   480
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton PauseBut 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Pause"
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox NumPntsBox 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   3480
      TabIndex        =   2
      Text            =   "7"
      Top             =   165
      Width           =   1020
   End
   Begin VB.HScrollBar SpeedBar 
      Height          =   255
      LargeChange     =   500
      Left            =   2520
      Max             =   1
      Min             =   10000
      SmallChange     =   200
      TabIndex        =   0
      Top             =   600
      Value           =   1
      Width           =   1995
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Speed"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1755
      TabIndex        =   3
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Points To Display"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1770
      TabIndex        =   1
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "ANCNTLDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BeginBut_Click()
Dim i, j
Dim k, kk, l
Dim xinc, xscale, xloc, yinc, yscale, yloc
Dim xpnt, ypnt

'sets mousepointer to hourglass
Screen.MousePointer = 11

'fills x() and y() and sets mins and maxs for graphing
Call SetAll

If n = 0 Or xranger = 0 Then
   MsgBox "No data is available to plot.", , "Error"
   GoTo xyaes:
ElseIf Main.G1G2But.Value = True And grpdefined = False Then
    MsgBox "No grouping has been defined.", , "Error"
    GoTo xyaes:
End If

'sets graph type for point find procedure and other params
gtype$ = "XY"
Main.ZoomCmd.Enabled = False
xcur = 0: ycur = 0
Main.GraphBox.FontName = "Helv"
Main.GraphBox.FontSize = 8

'clears screen
Main.GraphBox.Scale (-15, 110)-(115, -20)
If ClearScreenBut.Value = 1 Then Main.GraphBox.Cls
Main.GraphBox.AutoRedraw = False

'Draw axis and labels
Main.GraphBox.DrawWidth = 1: Main.GraphBox.DrawStyle = 0
Main.GraphBox.Line (0, 0)-(0, 100)
Main.GraphBox.Line (0, 0)-(100, 0)
Main.GraphBox.CurrentX = -13
Main.GraphBox.CurrentY = 108
Main.GraphBox.Print Main.YVarBox.Text
Main.GraphBox.CurrentX = 90
Main.GraphBox.CurrentY = -9
Main.GraphBox.Print Main.XVarBox.Text

'sets x-scales
If SetScalesDB.NumXLabelsTB.Text <> "" Then 'use manual scaling
   
   numxlabels = Val(SetScalesDB.NumXLabelsTB.Text)
   xinc = xranger / numxlabels
   For i = 0 To numxlabels
      'prints x axis scale labels
      xscale = xminr + i * xinc
      xloc = (xscale - xminr) / xranger * 100
      Main.GraphBox.CurrentX = xloc
      Main.GraphBox.CurrentY = -5
      Main.GraphBox.Print Format$(xscale, xfmt)
      Main.GraphBox.DrawWidth = 1: Main.GraphBox.DrawStyle = 0
      Main.GraphBox.Line (xloc, 0)-(xloc, 1)
      If Main.GridCmd.Checked = -1 Then
         Main.GraphBox.DrawWidth = 1: Main.GraphBox.DrawStyle = 2
         Main.GraphBox.Line (xloc, 0)-(xloc, 100)
      End If
   Next i

Else 'use auto-scaling

   'prints x-labels
   For i = 0 To xroot Step xintindex
      Main.GraphBox.CurrentY = -0.5
      xlabel = xminr + i * 10 ^ xpower
      xlabelpos = i / xroot * 100
      Main.GraphBox.CurrentX = xlabelpos - 2
      Main.GraphBox.Print Format(xlabel, xfmt)
      Main.GraphBox.Line (xlabelpos, 1)-(xlabelpos, 0)
      If Main.GridCmd.Checked = -1 Then
         Main.GraphBox.DrawWidth = 1: Main.GraphBox.DrawStyle = 2
         Main.GraphBox.Line (xlabelpos, 0)-(xlabelpos, 100)
      End If
   Next i
   Main.GraphBox.DrawWidth = 1: Main.GraphBox.DrawStyle = 0
   
End If

'sets y-scales
If SetScalesDB.NumYLabelsTB.Text <> "" Then 'use manual scaling
   
   numylabels = Val(SetScalesDB.NumYLabelsTB.Text)
   yinc = yranger / numylabels
   'prints y axis scale labels
   For i = 0 To numylabels
      yscale = yminr + i * yinc
      yloc = (yscale - yminr) / yranger * 100
      Main.GraphBox.CurrentX = -10
      Main.GraphBox.CurrentY = yloc + 3
      Main.GraphBox.Print Format$(yscale, yfmt)
      Main.GraphBox.DrawWidth = 1: Main.GraphBox.DrawStyle = 0
      Main.GraphBox.Line (0, yloc)-(1, yloc)
      If Main.GridCmd.Checked = -1 Then
         Main.GraphBox.DrawWidth = 1: Main.GraphBox.DrawStyle = 2
         Main.GraphBox.Line (0, yloc)-(100, yloc)
      End If
   Next i

Else 'use auto-scaling for y axis

   'plots y labels
   For i = 0 To yroot Step yintindex
      Main.GraphBox.CurrentX = -14
      ylabel = yminr + (i * 10 ^ ypower)
      ylabelpos = (i / yroot * 100)
      Main.GraphBox.CurrentY = ylabelpos + 1
      Main.GraphBox.Print Format(ylabel, yfmt)
      Main.GraphBox.Line (0, ylabelpos)-(1, ylabelpos)
      If Main.GridCmd.Checked = -1 Then
         Main.GraphBox.DrawWidth = 1: Main.GraphBox.DrawStyle = 2
         Main.GraphBox.Line (0, ylabelpos)-(100, ylabelpos)
      End If
   Next i
   Main.GraphBox.DrawWidth = 1: Main.GraphBox.DrawStyle = 0
End If


Main.GraphBox.DrawWidth = 1
For i = 1 To n

    'erases old points
    j = i - numpnts
    If j > 0 Then
       xpnt = (x(j, 2) - xminr) / xranger * 100
       ypnt = (y(j, 1) - yminr) / yranger * 100
       If Main.G1G2But.Value = False Then
          Main.GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), Main.GraphBox.BackColor, B
       Else
          If grpxy(j) = 1 Then
             Main.GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), Main.GraphBox.BackColor, B
          Else
             Main.GraphBox.Circle (xpnt, ypnt), 0.8, Main.GraphBox.BackColor
          End If
       End If
        
       'delay for speed control
       For k = 1 To delay
          For kk = 1 To 3
             l = k / 3
          Next kk
       Next k
    End If
    
    'draws new points
    xpnt = (x(i, 2) - xminr) / xranger * 100
    ypnt = (y(i, 1) - yminr) / yranger * 100
    If Main.G1G2But.Value = False Then
       Main.GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), BLACK, B
    Else
       If grpxy(i) = 1 Then
          Main.GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), BLUE, B
       Else
          Main.GraphBox.Circle (xpnt, ypnt), 0.8, RED
       End If
    End If

Next i

'erases old markers
For i = n + 1 To n + numpnts
    j = i - numpnts
    If j > 0 Then
       xpnt = (x(j, 2) - xminr) / xranger * 100
       ypnt = (y(j, 1) - yminr) / yranger * 100
       If Main.G1G2But.Value = False Then
          Main.GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), Main.GraphBox.BackColor, B
       Else
          If grpxy(j) = 1 Then
             Main.GraphBox.Line (xpnt - 0.7, ypnt + 1)-(xpnt + 0.7, ypnt - 1), Main.GraphBox.BackColor, B
          Else
             Main.GraphBox.Circle (xpnt, ypnt), 0.8, Main.GraphBox.BackColor
          End If
       End If
        
        'delays program for speed control
        For k = 1 To delay
            For kk = 1 To 3
               l = k / 3
            Next kk
        Next k
    End If
Next i

xyaes:
Screen.MousePointer = 0
Main.GraphBox.AutoRedraw = True

End Sub

Private Sub ContinueBut_Click()
   pause = False
End Sub

Private Sub Form_Load()

'center control db approximately over data box
ANCNTLDB.Move Main.Left + (Main.Width - ANCNTLDB.Width) / 4, Main.Top + ANCNTLDB.Height / 2

'set default numpoints and delay
numpnts = 7
delay = 1
SpeedBar.Value = delay
NumPntsBox.Text = Format$(numpnts, "####")
PauseBut.Enabled = False
ContinueBut.Enabled = False
pause = False
End Sub

Private Sub GoBut_Click()
   Main.XYAnimateBut.Value = True
End Sub

Private Sub NumPntsBox_Change()
   numpnts = Val(NumPntsBox.Text)
End Sub

Private Sub PauseBut_Click()
   pause = True
End Sub

Private Sub QuitBut_Click()
  Unload ANCNTLDB
End Sub

Private Sub SpeedBar_Change()
   SpeedBar_Scroll
End Sub

Private Sub SpeedBar_Scroll()
   delay = SpeedBar.Value
   If delay = 0 Then delay = 1
End Sub

