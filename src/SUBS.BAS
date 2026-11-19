Attribute VB_Name = "SUBS"

Sub AddVar(NewVar(), NewVarName$)
ReDim dtemp(numrecs, numflds + 1)
On Error GoTo avce

Screen.MousePointer = 11

'Transfers D() to DTemp()
For i = 1 To numrecs
   For j = 1 To numflds
      dtemp(i, j) = d(i, j)
   Next j
Next i

'adds new variable to DTemp()
For i = 1 To numrecs
   dtemp(i, numflds + 1) = NewVar(i)
Next i

'redims D() with one more field
numflds = numflds + 1
ReDim d(numrecs, numflds)

'refills D() with DTemp()
For i = 1 To numrecs
   For j = 1 To numflds
      d(i, j) = dtemp(i, j)
   Next j
Next i

'fills VN$()
ReDim vntemp$(numflds)
For i = 1 To numflds - 1
   vntemp$(i) = vn$(i)
Next i
vntemp$(i) = NewVarName$
ReDim vn$(numflds)
For i = 1 To numflds
   vn$(i) = vntemp$(i)
Next i

'adds new variable to y & x var boxes
Main.YVarBox.AddItem NewVarName$
Main.XVarBox.AddItem NewVarName$

avce:
Screen.MousePointer = 0
ReDim dtemp(1, 1)
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error"
   Resume aves:
End If
aves:
End Sub

Sub CalendarDate(recnum, ckdate$)
'recnum As Integer, ckdate$
Dim t$

t$ = "#0"

Select Case timeint
    Case 1
        ckdate$ = Format$(d(recnum, yrfld), t$)
    Case 2
        ckdate$ = Format$(d(recnum, mofld), t$) + "/" + Format$(d(recnum, yrfld), t$)
    Case 3
        ckdate$ = Format$(d(recnum, mofld), t$) + "/" + Format$(d(recnum, dyfld), t$) + "/" + Format$(d(recnum, yrfld), t$)
    Case 4
        ckdate$ = Format$(d(recnum, mofld), t$) + "/" + Format$(d(recnum, dyfld), t$) + "/" + Format$(d(recnum, yrfld), t$) + "/" + Format$(d(recnum, hrfld), t$)
End Select

'nht
If timeint = 4 And mnfld > 0 Then
   ckdate$ = ckdate$ + ":" + Format$(d(recnum, mnfld), t$)
End If

End Sub

Sub Cursor()
Static xcur, ycur, prevrec, prevgtype$

'sets legal bounds for current record number
If currec > numrecs Then currec = numrecs
If currec < 1 Then currec = 1

If gtype$ = "TS" Then

   'finds original color
   If zoommode = False Then
      i1 = 1
      i2 = numrecs
   End If
   PC2& = Main.GraphBox.Point((currec - i1) / (i2 - i1 + 1) * 100, (d(currec, yfld) - yminr) / yranger * 100)

   'erases old cursor if in bounds
   If xcur > 0 And xcur <= 100 And ycur > 0 And ycur <= 100 And prevgtype$ = gtype$ Then
      
      'saves previous color
      PC& = Main.GraphBox.Point(xcur, ycur)
      
      'erases old cursor
      Main.GraphBox.DrawWidth = 4
      Main.GraphBox.PSet (xcur, ycur), Main.GraphBox.BackColor
      
      'replaces last point
      Main.GraphBox.DrawWidth = 1
      Main.GraphBox.PSet (xcur, ycur), PC&
   End If

   'this is here and in the main.opencmd to get the alt+  menu to work
   If fopen = 0 Then ReDim d(1, 1)
   
   'skip if no data
   If d(currec, xfld(1)) = ndflag Or d(currec, yfld) = ndflag Then Exit Sub

   'find graph coordinates
   If zoommode = False Then
      i1 = 1
      i2 = numrecs
   End If
   xcur = (currec - i1) / (i2 - i1 + 1) * 100
   ycur = (d(currec, yfld) - yminr) / yranger * 100
   prevgtype$ = gtype$

   'displays current cursor
   If xcur >= 0 And xcur <= 100 And ycur >= 0 And ycur <= 100 Then
       Main.GraphBox.DrawWidth = 4
       Main.GraphBox.PSet (xcur, ycur), PC2&
   End If

ElseIf gtype$ = "XY" Then

   If newxygraph = True Then
      xcur = 0
      ycur = 0
   End If
   
   'erases old cursor if in bounds
   If xcur > 0 And xcur <= 100 And ycur > 0 And ycur <= 100 And prevgtype$ = gtype$ Then
      
      If Main.G1G2But.Value = False Then
         Main.GraphBox.Line (xcur - 0.7, ycur + 1)-(xcur + 0.7, ycur - 1), Main.GraphBox.BackColor, BF
         Main.GraphBox.Line (xcur - 0.7, ycur + 1)-(xcur + 0.7, ycur - 1), BLACK, B
      Else
         If grpd(prevrec) = 1 Then
            Main.GraphBox.Line (xcur - 0.7, ycur + 1)-(xcur + 0.7, ycur - 1), Main.GraphBox.BackColor, BF
            Main.GraphBox.Line (xcur - 0.7, ycur + 1)-(xcur + 0.7, ycur - 1), BLUE, B
         Else
            Main.GraphBox.FillStyle = 0
            Main.GraphBox.FillColor = Main.GraphBox.BackColor
            Main.GraphBox.Circle (xcur, ycur), 0.8, Main.GraphBox.BackColor
            Main.GraphBox.FillStyle = 1
            Main.GraphBox.Circle (xcur, ycur), 0.8, RED
         End If
      End If
   End If

   'this is here and in the main.opencmd to get the alt+  menu to work
   If fopen = 0 Then ReDim d(1, 1)
   
   'skip if no data
   If d(currec, xfld(1)) = ndflag Or d(currec, yfld) = ndflag Then Exit Sub

   'find graph coordinates for current cursor
   xcur = (d(currec, xfld(1)) - xminr) / xranger * 100
   ycur = (d(currec, yfld) - yminr) / yranger * 100
   prevrec = currec
   prevgtype$ = gtype$
   newxygraph = False
   

   'displays current cursor
   If xcur >= 0 And xcur <= 100 And ycur >= 0 And ycur <= 100 Then
      
      If Main.G1G2But.Value = False Then
         Main.GraphBox.Line (xcur - 0.7, ycur + 1)-(xcur + 0.7, ycur - 1), BLACK, BF
      Else
         'If grpxy(currec) = 1 Then
         If grpd(currec) = 1 Then
            Main.GraphBox.Line (xcur - 0.7, ycur + 1)-(xcur + 0.7, ycur - 1), BLUE, BF
         Else
            Main.GraphBox.FillStyle = 0
            Main.GraphBox.FillColor = RED
            Main.GraphBox.Circle (xcur, ycur), 0.8, RED
            Main.GraphBox.FillStyle = 1
         End If
      End If
   End If
End If


End Sub

Sub D3Line(x1m, y1m, x2m, y2m, xnew)
Dim x1, y1, x2, y2
Dim x1t, y1t
Dim mask1, mask2
Dim n, seg, ytop, ybot, color
Dim i, dydx
Dim xa, ya, xb, yb
Static xold

'pass variables to x1 so x1m is not changed
x1 = x1m: y1 = y1m: x2 = x2m: y2 = y2m

'sets color
'If y1 < .25 * ymaxr Then
'   color = blue
'ElseIf y1 < .5 * ymaxr Then
'   color = green
'ElseIf y1 < .75 * ymaxr Then
'   color = yellow
'ElseIf y1 < ymaxr Then
'   color = red
'Else
   color = BLACK
'End If

'order (x1,y1) and (x2,y2) so x1 > x2
If x1 > x2 Then
   x1t = x1: y1t = y1
   x1 = x2: y1 = y2
   x2 = x1t: y2 = y1t
End If

'refill maskmax and maskmin when appropriate
If xnew <> xold Then
   For i = 1 To res
      maskmin(i) = maskminnew(i)
      maskmax(i) = maskmaxnew(i)
   Next i
End If

If x1 = x2 Then 'line is vertical
   
   'order (x1,y1) and (x2,y2) so that y1 > y2
   If y2 > y1 Then
      x1t = x1: y1t = y1
      x1 = x2: y1 = y2
      x2 = x1t: y2 = y1t
   End If

   'break line into n segments and process each segment
   n = 50
   seg = (y1 - y2) / n
   For i = 1 To n
      'define top and bottom of line segment
      ytop = y1 - seg * (i - 1)
      ybot = y1 - seg * i
      
      'find if either top or bottom segments are clear
      If ytop > maskmin(x1) And ytop < maskmax(x1) Then mask1 = True Else mask1 = False
      If ybot > maskmin(x1) And ybot < maskmax(x1) Then mask2 = True Else mask2 = False
   
      'draw line segment if either end is clear
      If mask1 = False Or mask2 = False Then
         Main.GraphBox.Line (x1, ytop)-(x1, ybot), color
      End If
   Next i
   
   'update mask
   If y1 > maskmaxnew(x1) Then maskmaxnew(x1) = y1
   If y2 < maskminnew(x1) Then maskminnew(x1) = y2
   
Else 'line goes between atleast two diff x's
   
   'calc slope of line
   dydx = (y2 - y1) / (x2 - x1)

   For i = 1 To (x2 - x1)
      
      'calc postion of each segment (xa,ya)-(xb,yb)
      xa = x1 + (i - 1): ya = y1 + dydx * (i - 1)
      xb = x1 + i: yb = y1 + dydx * i

      'find if either top or bottom segments are clear
      If ya > maskmin(xa) And ya < maskmax(xa) Then mask1 = True Else mask1 = False
      If yb > maskmin(xb) And yb < maskmax(xb) Then mask2 = True Else mask2 = False
   
      
      'draw line segment if either end is clear
      If mask1 = False Or mask2 = False Then Main.GraphBox.Line (xa, ya)-(xb, yb), color
      
      'update mask
      If ya > maskmaxnew(xa) Then maskmaxnew(xa) = ya
      If ya < maskminnew(xa) Then maskminnew(xa) = ya
      If yb > maskmaxnew(xb) Then maskmaxnew(xb) = yb
      If yb < maskminnew(xb) Then maskminnew(xb) = yb
   Next i

End If
      
xold = xnew


End Sub

Sub DrawModelLine(predint, pass%, x1p, y1p, x2p, y2p)
Dim x1, y1, x2, y2

'Transfer x1pass to x1 so x1pass will not be changed
x1 = x1p
y1 = y1p
x2 = x2p
y2 = y2p

'calcs coordinates from 0-100 in x and y
If modtype$ = "Mean" Or modtype$ = "Median" Then
   x1 = 0
   x2 = 100
Else
   x1 = (x1 - xminr) / xranger * 100
   x2 = (x2 - xminr) / xranger * 100
End If
y1 = (y1 - yminr) / yranger * 100
y2 = (y2 - yminr) / yranger * 100
predint = predint / yranger * 100

'determines color of model line
If Main.G1G2But.Value = True And pass% = 1 Then
    color& = BLUE
    Main.GraphBox.DrawStyle = 0
    Main.GraphBox.DrawWidth = 2
ElseIf Main.G1G2But.Value = True And pass% = 2 Then
    color& = RED
    Main.GraphBox.DrawStyle = 2
    Main.GraphBox.DrawWidth = 2
Else
    color& = BLACK
    Main.GraphBox.DrawStyle = 0
    Main.GraphBox.DrawWidth = 2
End If

'draws model line
Main.GraphBox.Line (x1, y1)-(x2, y2), color&

If Main.PredIntCmd.Checked = True Then
    slope = (y2 - y1) / (x2 - x1)

    For topbot = 1 To 2
        x1p = x1
        x2p = x2
        If topbot = 1 Then
            'top
            y1p = y1 + predint
            y2p = y2 + predint
        Else
            'bottom
            y1p = y1 - predint
            y2p = y2 - predint
        End If
        
        Select Case slope
           Case 0
              If y1p < 0 Or y1p > 100 Then
                 y1p = -50
                 y2p = -50
              End If
           Case Else
            If y1p < 0 Then
                x1p = (0 - y1p) / slope + x1p
                y1p = 0
            ElseIf y1p > 100 Then
                x1p = (100 - y1p) / slope + x1p
                y1p = 100
            End If
            If y2p < 0 Then
                x2p = (0 - y2p) / slope + x2p
                y2p = 0
            ElseIf y2p > 100 Then
                x2p = (100 - y2p) / slope + x2p
                y2p = 100
            End If
        End Select
        
        Main.GraphBox.DrawWidth = 1: Main.GraphBox.DrawStyle = 2
        Main.GraphBox.Line (x1p, y1p)-(x2p, y2p), color&
    Next topbot
End If

Main.GraphBox.DrawWidth = 1

End Sub

Sub DrawPic(pic As Control)
'from page 393 of the VB 2.0 Programers Guide
'this subprogram will print the image in GraphBox to the
'screen at the resolution of the printer

scrx = Screen.TwipsPerPixelX
scry = Screen.TwipsPerPixelY
pic.ScaleMode = 3
For yrd = 0 To (pic.ScaleHeight - 1)
    pely = pic.Top + (yrd * scry)
    For xrd = 0 To (pic.ScaleWidth - 1)
        pelx = pic.Left + (xrd * scrx)
        pelc = pic.Point(xrd, yrd)
        If pelc <> QBColor(7) Then
            Printer.Line (pelx, pely)-Step(scrx, scry), pelc, BF
        End If
    Next xrd
    DoEvents
Next yrd
'SetClientPrintOrigin Card

Printer.EndDoc
End Sub

Sub DWTest(P, dwr$)

Select Case P
   Case 2
      dl = 0.31624 + 0.29655 * Log(n)
      du = 0.88254 + 0.17875 * Log(n)
   Case 3
      dl = 0.046637 + 0.35361 * Log(n)
      du = 1.24391 + 1.00448 * Log(n)
End Select

If dw < dl Then
   dwr$ = "p>0"
ElseIf dw > du Then
   dwr$ = "p=0"
Else
   dwr$ = "I"
End If

End Sub

Sub G1G2HistLines(pass%, x1, y1, x2, y2)
'calcs coordinates from 0-100 in x and y

x1 = (x1 - xminr) / xranger * 50
x2 = (x2 - xminr) / xranger * 50
If pass% = 2 Then
    x1 = x1 + 50
    x2 = x2 + 50
End If
y1 = (y1 - yminr) / yranger * 100
y2 = (y2 - yminr) / yranger * 100

'determines color of model line
If Main.G1G2But.Value = True And pass% = 1 Then
    color& = BLUE
ElseIf Main.G1G2But.Value = True And pass% = 2 Then
    color& = RED
Else
    color& = BLACK
End If

'draws model line
Main.GraphBox.DrawWidth = 1
Main.GraphBox.DrawStyle = 0
Main.GraphBox.Line (x1, y1)-(x2, y2), color&

End Sub

Sub Inf(x(), xr As Integer, xc As Integer, y(), yr As Integer, yc As Integer, beta(), betar As Integer, betac As Integer, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), roe)
'CALULATES INFERENCE PARAMETERS
Dim yhatr As Integer, yhatc As Integer
Dim i As Integer, j As Integer
Dim pnum, pdenom
ReDim xtot(8), yhat(numrecs, 1), resx(numrecs)

'DEFINITIONS
n = xr
numvars = xc - 1
numcoefs = xc

'CALCS X AND Y TOTALS AND MEANS
ytot = 0: SSY = 0
For j = 1 To numvars
   xtot(j) = 0
Next j
For i = 1 To n
   ytot = ytot + y(i, 1)
   SSY = SSY + y(i, 1) ^ 2
   For j = 1 To numvars
      xtot(j) = xtot(j) + x(i, j + 1)
   Next j
Next i
ymean = ytot / n
For i = 1 To numvars
   xmean(i) = xtot(i) / n
Next i

'CALCS SSE, MSE, RMSE, CV-RMSE AND R2
Call Mult(x(), xr, xc, beta(), betar, betac, yhat(), yhatr, yhatc)

ssm = 0: sse = 0
acnum = 0: acden = 0: dwnum = 0: dwden = 0
For i = 1 To n
   'fills resx() which corresponds to the x matrix
   'If i = 204 Then Stop
   resx(i) = y(i, 1) - yhat(i, 1)
   ssm = ssm + (y(i, 1) - ymean) ^ 2
   sse = sse + resx(i) ^ 2

   'calcs autocorr And durbinwatson
   Select Case i
   Case 1
       prevres = resx(i)
   Case Else
       acnum = acnum + prevres * resx(i)
       acden = acden + prevres ^ 2
       dwnum = dwnum + (resx(i) - prevres) ^ 2
       pnum = pnum + resx(i) * prevres
       pdenom = pdenom + resx(i) ^ 2
       prevres = resx(i)
   End Select
   dwden = dwden + resx(i) ^ 2
Next i
If acden <> 0 Then autocor = acnum / acden 'if statement enables regression where n = p
If dwden <> 0 Then dw = dwnum / dwden 'if statement enables regression where n = p

'fills d() with resids
If Main.G1G2But.Value = True Then
   If pass = 1 Then restype = 2 Else restype = 3
Else
   restype = 1
End If
idi = 1
For i = 1 To numrecs
   If i = IndexNum(idi) Then
      d(i, numfldsd + restype) = resx(idi)
      idi = idi + 1
   Else
      d(i, numfldsd + restype) = ndflag
   End If
Next i

If n - xc = 0 Then 'same number of data points as reg coef gives perfect fit
   mse = 0
Else
   mse = sse / (n - xc)
End If
rmse = mse ^ 0.5
cvrmse = rmse / ymean * 100
r2 = 1 - sse / ssm
If n - xc = 0 Then 'same number of data points as reg coef gives perfect fit
   adjr2 = 1
Else
   adjr2 = 1 - (n - 1) / (n - xc) * sse / ssm
End If
If pdemnom <> 0 Then roe = pnum / pdenom 'if statement enables regression where n = p

'CALCS THE STANDARD ERROR OF THE COEFFICIENTS
For i = 1 To numcoefs
   sigma(i) = rmse * sec(i) ^ 0.5
Next i

End Sub

Sub Invert(A(), s As Integer, B(), det, sec())
' THIS PROCEDURE WILL Invert UP TO AN 8 BY 8 MATRIX
' CHANGE THE DIMINSIONS IN THE "DIM INDEX()" STATEMENT FOR BIGGER MATRICES
ReDim Index(8, 3)

numrows = s
numcols = s

'TransFER A() TO B(), AND DO ALL FUTURE WORK ON B()
errflag = 0
For i = 1 To numrows
  For j = 1 To numcols
     B(i, j) = A(i, j)
  Next j
  Index(i, 3) = 0
Next i

det = 1
For i = 1 To numrows
   ' SEARCH FOR BIGGEST PIVOT ELEMENT
   big = 0
   For j = 1 To numcols
      If Index(j, 3) = 1 Then GoTo 5390
      For k = 1 To numcols
         If Index(k, 3) > 1 Then Main.ValueBar.Print "ERROR: MATRIX IS SINGULAR": End
         If Index(k, 3) = 1 Then GoTo 5380
         If big >= Abs(B(j, k)) Then GoTo 5380
         irow = j
         icol = k
         big = Abs(B(j, k))
5380  Next k
5390  Next j
   Index(icol, 3) = Index(icol, 3) + 1
   Index(i, 1) = irow
   Index(i, 2) = icol
  
   'INTERCHANGE ROWS TO PUT PIVOT ON THE DIAGONAL
   If irow = icol Then GoTo 5580
   det = -det
   For l = 1 To numcols
      hold = B(irow, l)
      B(irow, l) = B(icol, l)
      B(icol, l) = hold
   Next l
  
5580 'DIVIDE PIVOT ROW BY PIVOT ELEMENT
    Pivot = B(icol, icol)
    det = det * Pivot
    B(icol, icol) = 1
    For l = 1 To numcols
       B(icol, l) = B(icol, l) / Pivot
    Next l
  
    'REDUCE NON-PIVOT ROWS
    For l = 1 To numcols
       If l = icol Then GoTo 5810
       t = B(l, icol)
       B(l, icol) = 0
       For m = 1 To numcols
          B(l, m) = B(l, m) - B(icol, m) * t
       Next m
5810 Next l
Next i

'INTERCHANGE COLUMNS
For i = 1 To numcols
   l = numcols - i + 1
   If Index(l, 1) = Index(l, 2) Then GoTo 5960
   irow = Index(l, 1)
   icol = Index(l, 2)
   For k = 1 To numcols
      hold = B(k, irow)
      B(k, irow) = B(k, icol)
      B(k, irow) = hold
   Next k
5960 Next i

For k = 1 To numcols
   If Index(k, 3) <> 1 Then
      Main.ValueBar.Print "ERROR:  MATRIX IS SINGULAR":
   End If
Next k

For i = 1 To numcols
  sec(i) = B(i, i)
Next i

inves:
End Sub

Sub Median(yobs, ydata(), med, semed)
Dim i, tmp, swit, n
Dim P, keyyobs, intpart, fracpart
Dim p25s, p50s
ReDim odata(yobs)

'passes ydata() to odata() [ordered data] so original matrix is not changed
For i = 1 To yobs
   odata(i) = ydata(i, 1)
Next i
n = yobs

'sorts odata()
swit = True
While swit = True
   swit = False
   For i = 1 To n - 1
      If odata(i + 1) < odata(i) Then
         tmp = odata(i)
         odata(i) = odata(i + 1)
         odata(i + 1) = tmp
         swit = True
      End If
   Next i
Wend

'finds median
P = 0.5
keyyobs = (n + 1) * P
intpart = Int(keyyobs)
fracpart = keyyobs - Int(keyyobs)
If intpart = 0 Then
   med = (1 - fracpart) * odata(1) + fracpart * odata(1)
ElseIf intpart = n Then
   med = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart)
Else
   med = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart + 1)
End If

'finds 25th percentile
P = 0.25
keyyobs = (n + 1) * P
intpart = Int(keyyobs)
fracpart = keyyobs - Int(keyyobs)
If intpart = 0 Then
   p25s = (1 - fracpart) * odata(1) + fracpart * odata(1)
ElseIf intpart = n Then
   p25s = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart)
Else
   p25s = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart + 1)
End If

'finds 75th percentile
P = 0.75
keyyobs = (n + 1) * P
intpart = Int(keyyobs)
fracpart = keyyobs - Int(keyyobs)
If intpart = 0 Then
   p75s = (1 - fracpart) * odata(1) + fracpart * odata(1)
ElseIf intpart = n Then
   p75s = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart)
Else
   p75s = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart + 1)
End If

'calcs semed
semed = (p75s - p25s) / (n ^ 0.5)


End Sub

Sub Mult(ma(), mar As Integer, mac As Integer, mb(), mbr As Integer, mbc As Integer, mc(), mcr As Integer, mcc As Integer)

If mac <> mbr Then
   Msg$ = "NUM X COL <> NUM Y ROWS"
   MsgBox Msg$, , "Error in Subroutine Mult"
   Exit Sub
End If

mcr = mar
mcc = mbc

For R = 1 To mcr
   For c = 1 To mcc
         mc(R, c) = 0
   Next c
Next R

For R = 1 To mcr
   For c = 1 To mcc
      For i = 1 To mac
        mc(R, c) = mc(R, c) + ma(R, i) * mb(i, c)
      Next i
   Next c
Next R

End Sub

Sub OpenVar(varname$)
Dim i As Integer
Dim yrflag As Integer, moflag As Integer, dyflag As Integer, hrflag As Integer

If getvnsfrommerge = False Then
   'read in variable names
   Open varname$ For Input As #1
   i = 0
   Do While Not EOF(1)
       i = i + 1
       Input #1, vn$(i)
       If i = numfldsd Then Exit Do
   Loop
   Close #1

   'check for correct number
   If i <> numfldsd Then
       Msg$ = Str$(numfldsd) + "fields in the data file and " + Str$(i) + " variable names in the variable name file."
       MsgBox Msg$, , "Error"
       GoTo ovce
   End If
End If

'initializes variable boxes
Do While Main.XVarBox.ListCount > 0
   Main.XVarBox.RemoveItem 0
   Main.YVarBox.RemoveItem 0
Loop

mnfld = 0 'nht
For i = 1 To numfldsd
   'fills variable boxes with new names
   Main.XVarBox.AddItem vn$(i)
   Main.YVarBox.AddItem vn$(i)

   'finds fields of time stamp
   If LCase$(vn$(i)) = "yr" Then
      yrfld = i: yrflag = 1
   ElseIf LCase$(vn$(i)) = "mo" Then
      mofld = i: moflag = 1
   ElseIf LCase$(vn$(i)) = "dy" Then
      dyfld = i: dyflag = 1
   ElseIf LCase$(vn$(i)) = "hr" Then
      hrfld = i: hrflag = 1
   ElseIf LCase$(vn$(i)) = "mn" Then 'nht
      mnfld = i
   End If
Next i

'defines resid variables
vn$(numfldsd + 1) = "SG_Resids"
vn$(numfldsd + 2) = g1name$ + "_Resids"
vn$(numfldsd + 3) = g2name$ + "_Resids"
For i = 1 To 3
    Main.XVarBox.AddItem vn$(numfldsd + i)
    Main.YVarBox.AddItem vn$(numfldsd + i)
Next i
For i = 1 To numrecs
   For j = numfldsd + 1 To numfldsd + 3
      d(i, j) = ndflag
   Next j
Next i

'finds timeint
timeint = yrflag + moflag + dyflag + hrflag 'nht
If timeint > 0 Then
    Main.SelDatesCmd.Enabled = True
    Main.ViewDatesCmd.Enabled = True
End If
   
'reinitializes timescale variables
timeinthr = False
timeintdy = False
timeintmo = False
timeintyr = False
Unload TSOptsDB

'enable longer time-scales than current data
Main.TimeScalesCmd.Enabled = True
If timeint = 2 Then 'monthly data
   Main.ToYearlyCmd.Enabled = True
ElseIf timeint = 3 Then 'daily data
   Main.ToMonthlyCmd.Enabled = True
   Main.ToYearlyCmd.Enabled = True
ElseIf timeint = 4 Then  'hourly data
   Main.ToDailyCmd.Enabled = True
   Main.ToMonthlyCmd.Enabled = True
   Main.ToYearlyCmd.Enabled = True
End If

'enables hourly profiles cmd
If timeint = 4 Then
   Main.HourlyProfilesBut.Enabled = True
Else
   Main.HourlyProfilesBut.Enabled = False
End If

'enables mod and create var commands
Main.CreateVariableCmd.Enabled = True
Main.DeleteVariableCmd.Enabled = True
Main.CreateModelProjectionCmd.Enabled = True
Main.ModValuesCmd.Enabled = True
Main.SaveDataCmd.Enabled = True
Main.SaveNamesCmd.Enabled = True
Main.GroupByDateCmd.Enabled = True
Main.AvgIntDataCmd.Enabled = True

If yrflag + moflag + dyflag = 3 Then Main.GroupByDOWCmd.Enabled = True

'refills the year field in d() with 4 digit date if it has two digit date
'If yrflag = 1 And d(1, yrfld) < 100 Then
For i = 1 To numrecs
   If yrflag = 1 And d(i, yrfld) >= 50 And d(i, yrfld) <= 99 Then
      d(i, yrfld) = d(i, yrfld) + 1900
   ElseIf yrflag = 1 And d(i, yrfld) >= 0 And d(i, yrfld) <= 49 Then
      d(i, yrfld) = d(i, yrfld) + 2000
   End If
Next i

ovce:
Screen.MousePointer = 0
If Err Then
   Msg$ = """" + Error(Err) + """"
   MsgBox Msg$, , "Error"
   Resume oves
End If
oves:

End Sub

Sub Percentile(VarNum, yobs, ydata(), p10s, p25s, p50s, p75s, p90s)
Dim i, tmp, swit, n
Dim P, keyyobs, intpart, fracpart
ReDim odata(yobs)

'scans for too few data points
If yobs <= 0 Then
   p10s = ndflag
   p25s = ndflag
   p50s = ndflag
   p75s = ndflag
   p90s = ndflag
   Exit Sub
End If

'if not ndflag, then passes ydata() to odata()
'[ordered data] so original matrix is not changed
n = 0
For i = 1 To yobs
   If ydata(VarNum, i) <> ndflag Then
      n = n + 1
      odata(n) = ydata(VarNum, i)
   End If
Next i

'sorts odata()
swit = True
While swit = True
   swit = False
   For i = 1 To n - 1
      If odata(i + 1) < odata(i) Then
         tmp = odata(i)
         odata(i) = odata(i + 1)
         odata(i + 1) = tmp
         swit = True
      End If
   Next i
Wend

'finds 10th percentile
P = 0.1
keyyobs = (n + 1) * P
intpart = Int(keyyobs)
fracpart = keyyobs - Int(keyyobs)
If intpart = 0 Then
   p10s = (1 - fracpart) * odata(1) + fracpart * odata(1)
ElseIf intpart = n Then
   p10s = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart)
Else
   p10s = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart + 1)
End If
'finds 25th percentile
P = 0.25
keyyobs = (n + 1) * P
intpart = Int(keyyobs)
fracpart = keyyobs - Int(keyyobs)
If intpart = 0 Then
   p25s = (1 - fracpart) * odata(1) + fracpart * odata(1)
ElseIf intpart = n Then
   p25s = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart)
Else
   p25s = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart + 1)
End If
'finds 50th percentile (median)
P = 0.5
keyyobs = (n + 1) * P
intpart = Int(keyyobs)
fracpart = keyyobs - Int(keyyobs)
If intpart = 0 Then
   p50s = (1 - fracpart) * odata(1) + fracpart * odata(1)
ElseIf intpart = n Then
   p50s = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart)
Else
   p50s = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart + 1)
End If
'finds 75th percentile
P = 0.75
keyyobs = (n + 1) * P
intpart = Int(keyyobs)
fracpart = keyyobs - Int(keyyobs)
If intpart = 0 Then
   p75s = (1 - fracpart) * odata(1) + fracpart * odata(1)
ElseIf intpart = n Then
   p75s = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart)
Else
   p75s = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart + 1)
End If
'finds 90th percentile
P = 0.9
keyyobs = (n + 1) * P
intpart = Int(keyyobs)
fracpart = keyyobs - Int(keyyobs)
If intpart = 0 Then
   p90s = (1 - fracpart) * odata(1) + fracpart * odata(1)
ElseIf intpart = n Then
   p90s = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart)
Else
   p90s = (1 - fracpart) * odata(intpart) + fracpart * odata(intpart + 1)
End If

End Sub

Sub Reg(x(), xr As Integer, xc As Integer, y(), yr As Integer, yc As Integer, beta(), betar As Integer, betac As Integer, sec())
ReDim xt(8, numrecs), xtx(8, 8), xtxi(8, 8), xtotal(8, numrecs)
Dim xtr As Integer, xtc As Integer
Dim xtxr As Integer, xtxc As Integer
Dim xtotalr As Integer, xtotalc As Integer
Dim nv As Integer, i As Integer, j As Integer

'MATRIX RegRESSION ROUTINE

' THESE ARE MY SUB-PROGRAM TOOLS
' CALL Invert(A(), AR, AINV(), det, SEC())
' CALL Mult(A(), AR, AC, B(), BR, BC, AB(), ABR, ABC)
' CALL Trans(A(), AR, AC, AT(), ATR, ATC)
' WHERE  AR=NUM ROWS IN A()  AND  AC=NUM COLS IN A()

'CALCS X MEANS AND NORMALIZES X VARIABLES
nv = xc - 1
ReDim xt1(nv), xm(nv)
For j = 1 To nv
   xt1(j) = 0
Next
For i = 1 To n
   For j = 1 To nv
     xt1(j) = xt1(j) + x(i, j + 1)
   Next
Next
For j = 1 To nv
   xm(j) = xt1(j) / n
   If xm(j) = 0 Then xm(j) = 1
  For i = 1 To n
      x(i, j + 1) = x(i, j + 1) / xm(j)
   Next
Next

Call Trans(x(), xr, xc, xt(), xtr, xtc)
Call Mult(xt(), xtr, xtc, x(), xr, xc, xtx(), xtxr, xtxc)
Call Invert(xtx(), xtxr, xtxi(), det, sec())
Call Mult(xtxi(), xtxr, xtxc, xt(), xtr, xtc, xtotal(), xtotalr, xtotalc)
Call Mult(xtotal(), xtotalr, xtotalc, y(), yr, yc, beta(), betar, betac)

'RETURNS TO ORIGINAL X VALUES AND COEFFICIENTS
For i = 1 To n
   For j = 1 To nv
      x(i, j + 1) = x(i, j + 1) * xm(j)
   Next
Next
For j = 1 To nv
   beta(j + 1, 1) = beta(j + 1, 1) / xm(j)
   sec(j + 1) = sec(j + 1) / (xm(j) ^ 2)
Next

End Sub

Sub SerDate(recnum As Integer, sdate)

Select Case timeint
    Case 1
        sdate = d(recnum, yrfld) * 10 ^ 6
    Case 2
        sdate = d(recnum, yrfld) * 10 ^ 6 + d(recnum, mofld) * 10 ^ 4
    Case 3
        sdate = d(recnum, yrfld) * 10 ^ 6 + d(recnum, mofld) * 10 ^ 4 + d(recnum, dyfld) * 10 ^ 2
    Case 4
        sdate = d(recnum, yrfld) * 10 ^ 6 + d(recnum, mofld) * 10 ^ 4 + d(recnum, dyfld) * 10 ^ 2 + d(recnum, hrfld)
End Select

End Sub

Sub SerToCal(sdate, ckdate)
Dim dt$, yr$, mo$, dY$, hr$

dt$ = LTrim$(RTrim$(Str$(sdate)))
yr$ = Left$(dt$, 4)
mo$ = Mid$(dt$, 5, 2)
dY$ = Mid$(dt$, 7, 2)
hr$ = Right$(dt$, 2)

Select Case timeint
   Case 1
      ckdate = yr$
   Case 2
      ckdate = yr$ + "/" + mo$
   Case 3
      ckdate = mo$ + "/" + dY$ + "/" + yr$
   Case 4
      ckdate = mo$ + "/" + dY$ + "/" + yr$ + "/" + hr$
End Select

If sdate = 0 Then
    ckdate = "Start of Data File"
End If
If edate = 9999999999# Then
    ckdate = "End of Data File"
End If

End Sub

Sub SetAll()
Dim i As Integer
ReDim x(numrecs, 2), y(numrecs, 1)
ReDim IndexNum(numrecs) As Integer
ReDim grpxy(numrecs)

'NOTE xmax,xmin,ymin,ymax,xrange,yrange refer to true mins and maxs
'xmaxr,xminr,yminr,ymaxr,xranger,yranger refer to rounded mins and maxs

'sets xfld and yfld
xfld(1) = Main.XVarBox.ListIndex + 1
If ss = True Then
   yfld = yfldo
Else
   yfld = Main.YVarBox.ListIndex + 1
End If

'fills x() and y()
n = 0
For i = 1 To numrecs
   If d(i, xfld(1)) <> ndflag And d(i, yfld) <> ndflag Then
        
      'counts num obs in x() and y()
      n = n + 1
      
      'fills x() and y()
      x(n, 1) = 1
      x(n, 2) = d(i, xfld(1))
      y(n, 1) = d(i, yfld)
      IndexNum(n) = i
   
      'finds max and mins
      If n = 1 Then
         xmin = x(1, 2)
         xmax = x(1, 2)
         ymin = y(1, 1)
         ymax = y(1, 1)
      End If
      If x(n, 2) < xmin Then xmin = x(n, 2)
      If x(n, 2) > xmax Then xmax = x(n, 2)
      If y(n, 1) < ymin Then ymin = y(n, 1)
      If y(n, 1) > ymax Then ymax = y(n, 1)
   
      If grpd(i) = 1 Then grpxy(n) = 1
      If grpd(i) = 2 Then grpxy(n) = 2
   End If
Next i

'sets ranges for actual mins and maxs
xrange = xmax - xmin
yrange = ymax - ymin

'overrides yminr & ymaxr if ss from timeseries2
If ss = True Then
   ymaxr = ymaxss - 1
   yminr = yminss + 1
End If

'rounds ymaxr up to even number
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

'selects interval to determine how many labels
If yroot >= 5 Then
   yintindex = 1
Else
   yintindex = 0.5
End If

'rounds xmax up to even number
emax$ = Format$(xmax, "0.0E+00")
rootmax = 1 + Val(Left$(emax$, 1))
powermax = Val(Right(emax$, 3))
xmaxr = rootmax * 10 ^ powermax

'finds xmin, xroot and xpower
If xmin < 0 Then

   'rounds xmin down to even number
   emax$ = Format$(xmin, "0.0E+00")
   rootmin = -1 + Val(Left$(emax$, 2))
   powermin = Val(Right(emax$, 3))
   xminr = rootmin * 10 ^ powermin
   
   'sets power and root
   If powermax > powermin Then
      xpower = powermax
      xroot = rootmax + 1
      xminr = -1 * 10 ^ xpower
   ElseIf powermin > powermax Then
      xpower = powermin
      xroot = Abs(rootmin) + 1
      xmaxr = 1 * 10 ^ xpower
   ElseIf powermin = powermax Then
      xpower = powermax
      xroot = rootmax - rootmin
   End If

Else 'xmin > 0
   
   'sets xmin, power and root
   xminr = 0
   xroot = rootmax
   xpower = powermax

End If

'selects interval to determine how many labels
If xroot >= 5 Then
   xintindex = 1
Else
   xintindex = 0.5
End If

' changes xmax and xmin if set scales db is in use
If SetScalesDB.XMinTB.Text <> "" Then xminr = Val(SetScalesDB.XMinTB.Text)
If SetScalesDB.XMaxTB.Text <> "" Then xmaxr = Val(SetScalesDB.XMaxTB.Text)
If SetScalesDB.YminTB.Text <> "" Then yminr = Val(SetScalesDB.YminTB.Text)
If SetScalesDB.YmaxTB.Text <> "" Then ymaxr = Val(SetScalesDB.YmaxTB.Text)

'sets rounded ranges
xranger = xmaxr - xminr
yranger = ymaxr - yminr

'selects xfmt
If xmaxr - xminr > 5 Then
   xfmt = "#,###,##0"
ElseIf xmaxr - xminr > 1 Then
   xfmt = "0.00"
Else
   xfmt = "0.000"
End If

'selects yfmt
If ymaxr - yminr >= 5 Then
   yfmt = "#,###,##0"
ElseIf ymaxr - yminr > 1 Then
   yfmt = "0.00"
Else
   yfmt = "0.000"
End If

End Sub

Sub SetGrp()
Dim i As Integer
ReDim x(numrecs, 2), y(numrecs, 1)
ReDim IndexNum(numrecs) As Integer
ReDim grpxy(numrecs)
'refills x() and y() with only the appropriate group

n = 0
For i = 1 To numrecs
   If d(i, xfld(1)) <> ndflag And d(i, yfld) <> ndflag And grpd(i) = pass Then
        
      'counts num obs in x() and y()
      n = n + 1
      
      'fills x() and y()
      x(n, 1) = 1
      x(n, 2) = d(i, xfld(1))
      y(n, 1) = d(i, yfld)
      IndexNum(n) = i
      
      'finds new max and mins if necessary
      If modtype$ = "5P" Or modtype$ = "4P" Or modtype$ = "3P" Or modtype$ = "BinFit" Then
         If n = 1 Then
            xmin = x(1, 2)
            xmax = x(1, 2)
            ymin = y(1, 1)
            ymax = y(1, 1)
         End If
         If x(n, 2) < xmin Then xmin = x(n, 2)
         If x(n, 2) > xmax Then xmax = x(n, 2)
         If y(n, 1) < ymin Then ymin = y(n, 1)
         If y(n, 1) > ymax Then ymax = y(n, 1)
      End If

   End If
Next i

'sets ranges for actual mins and maxs if necessary
If modtype$ = "5P" Or modtype$ = "4P" Or modtype$ = "3P" Or modtype$ = "BinFit" Then
   xrange = xmax - xmin
   yrange = ymax - ymin
End If

End Sub

Sub SingleProfile()
Dim hr, dY
Dim i, m As Integer
Dim x1, y1, z1, x2, y2, z2
Dim xs1 As Integer, ys1, xs2 As Integer, ys2
Dim yvar ', numdays
Dim diff, lastdiff, theta

'sets scale and rotation angle
Main.GraphBox.Scale (-0.1 * res, 1.1 * res)-(1.1 * res, -0.1 * res)
'theta = main.RotAngSB.Value / 360 * 2 * 3.1416
theta = thetad / 360 * 2 * 3.1416
sint = Sin(theta)
cost = Cos(theta)

'note: to make this procedure run faster,
'remove this init mask block,and change every
'Call D3Line to Main.GraphBox.Line
'initializes maskmin and maskmax
'ReDim maskmax(0 To res), maskmin(0 To res), maskmaxnew(0 To res), maskminnew(0 To res)
'For i = 0 To res
'   maskminnew(i) = res
'   maskmaxnew(i) = 0
'Next i

'clears screen and makes mouse pointer into an hourglass
Main.GraphBox.Cls
Screen.MousePointer = 11

'finds yvar
yvar = Main.YVarBox.ListIndex + 1

dY = numdays
For i = 1 To numrecs
   If d(i, hrfld) = 0 Then
      dY = dY - 1
      If dY = dy1 Then
         'save recnum to calculate date below
         m = i
         'define the first base point
         x1 = dY / numdays * res: y1 = 0: z1 = res
         xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
         ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
         
         'draw data lines
         For hr = 0 To 23
            If d(i + hr, yvar) = ndflag Then
               y2 = 0
            Else
               y2 = d(i + hr, yvar) / ymax * res
            End If
            x2 = dY / numdays * res
            z2 = res - (hr / 23 * res)
            xs2 = (x2 - (z2 * cost) + (res * cost)) / (cost + 1)
            ys2 = (y2 - (z2 * sint) + (res * sint)) / (sint + 1)
            Main.GraphBox.Line (xs1, ys1)-(xs2, ys2)
            'Call D3Line(xs1, ys1, xs2, ys2, x1)
            xs1 = xs2
            ys1 = ys2
         Next hr
      
         'define the second base point in each daily load trace
         x2 = dY / numdays * res
         y2 = 0
         z2 = 0
         xs2 = (x2 - (z2 * cost) + (res * cost)) / (cost + 1)
         ys2 = (y2 - (z2 * sint) + (res * sint)) / (sint + 1)
         Main.GraphBox.Line (xs1, ys1)-(xs2, ys2)
         'Call D3Line(xs1, ys1, xs2, ys2, x1)
         
         'connect base points
         xs1 = xs2
         ys1 = ys2
         x2 = dY / numdays * res
         y2 = 0
         z2 = res
         xs2 = (x2 - (z2 * cost) + (res * cost)) / (cost + 1)
         ys2 = (y2 - (z2 * sint) + (res * sint)) / (sint + 1)
         Main.GraphBox.Line (xs1, ys1)-(xs2, ys2)
         'Call D3Line(xs1, ys1, xs2, ys2, x1)

         'draw vertical hour lines
         For hr = 0 To 23
            'define base point
            x1 = dY / numdays * res: y1 = 0: z1 = res * (1 - hr / 23)
            xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
            ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
            
            If d(i + hr, yvar) = ndflag Then
               y2 = 0
            Else
               y2 = d(i + hr, yvar) / ymax * res
            End If
            x2 = dY / numdays * res
            z2 = res - (hr / 23 * res)
            xs2 = (x2 - (z2 * cost) + (res * cost)) / (cost + 1)
            ys2 = (y2 - (z2 * sint) + (res * sint)) / (sint + 1)
            Main.GraphBox.Line (xs1, ys1)-(xs2, ys2)
            'Call D3Line(xs1, ys1, xs2, ys2, x1)
         Next hr

      End If
   End If
Next i

'draws x axis
x1 = 0: y1 = 0: z1 = 0: x2 = res: y2 = 0: z2 = 0
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
xs2 = (x2 - (z2 * cost) + (res * cost)) / (cost + 1)
ys2 = (y2 - (z2 * sint) + (res * sint)) / (sint + 1)
Main.GraphBox.Line (xs1, ys1)-(xs2, ys2)
'Call D3Line(xs1, ys1, xs2, ys2, x1)
'draws y axis
x1 = 0: y1 = 0: z1 = 0: x2 = 0: y2 = res: z2 = 0
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
xs2 = (x2 - (z2 * cost) + (res * cost)) / (cost + 1)
ys2 = (y2 - (z2 * sint) + (res * sint)) / (sint + 1)
Main.GraphBox.Line (xs1, ys1)-(xs2, ys2)
'Call D3Line(xs1, ys1, xs2, ys2, x1)
'draws z axis
x1 = 0: y1 = 0: z1 = 0: x2 = 0: y2 = 0: z2 = res
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
xs2 = (x2 - (z2 * cost) + (res * cost)) / (cost + 1)
ys2 = (y2 - (z2 * sint) + (res * sint)) / (sint + 1)
Main.GraphBox.Line (xs1, ys1)-(xs2, ys2)
'Call D3Line(xs1, ys1, xs2, ys2, x1)

'draws front x axis
x1 = 0: y1 = 0: z1 = res: x2 = res: y2 = 0: z2 = res
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
xs2 = (x2 - (z2 * cost) + (res * cost)) / (cost + 1)
ys2 = (y2 - (z2 * sint) + (res * sint)) / (sint + 1)
'Main.GraphBox.Line (xs1, ys1)-(xs2, ys2)
Call D3Line(xs1, ys1, xs2, ys2, x1)

'draws 23 hour label
x1 = (dy1 / numdays * res) + 15: y1 = 0: z1 = 0
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
Main.GraphBox.CurrentX = xs1: Main.GraphBox.CurrentY = ys1
Main.GraphBox.Print "23"
'draws 18 hour label
x1 = (dy1 / numdays * res) + 15: y1 = 0: z1 = res / 4
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
Main.GraphBox.CurrentX = xs1: Main.GraphBox.CurrentY = ys1
Main.GraphBox.Print "18"
x1 = (dy1 / numdays * res): y1 = 0: z1 = res * (1 - (18 / 23))
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
x2 = (dy1 / numdays * res) + 15
xs2 = (x2 - (z1 * cost) + (res * cost)) / (cost + 1)
Main.GraphBox.Line (xs1, ys1)-(xs2, ys1)
'draws 12 hour label
x1 = (dy1 / numdays * res) + 15: y1 = 0: z1 = res / 2
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
Main.GraphBox.CurrentX = xs1: Main.GraphBox.CurrentY = ys1
Main.GraphBox.Print "12"
x1 = (dy1 / numdays * res): y1 = 0: z1 = res * (1 - (12 / 23))
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
x2 = (dy1 / numdays * res) + 15
xs2 = (x2 - (z1 * cost) + (res * cost)) / (cost + 1)
Main.GraphBox.Line (xs1, ys1)-(xs2, ys1)
'draws 6 hour label
x1 = (dy1 / numdays * res) + 15: y1 = 0: z1 = res * 0.75
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
Main.GraphBox.CurrentX = xs1: Main.GraphBox.CurrentY = ys1
Main.GraphBox.Print "6"
x1 = dy1 / numdays * res: y1 = 0: z1 = res * (1 - (6 / 23))
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
x2 = (dy1 / numdays * res) + 15
xs2 = (x2 - (z1 * cost) + (res * cost)) / (cost + 1)
Main.GraphBox.Line (xs1, ys1)-(xs2, ys1)
'draw ymax line
x1 = 0: y1 = res: z1 = 0
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
x1 = 15: y1 = res: z1 = 0
xs2 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys2 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
Main.GraphBox.Line (xs1, ys1)-(xs2, ys2)
'draws variable label
x1 = 2: y1 = res + 20: z1 = 0
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
Main.GraphBox.CurrentX = xs1: Main.GraphBox.CurrentY = ys1
Main.GraphBox.Print Main.YVarBox.Text + " = " + Format(ymaxr)

'draws date label
x1 = dy1 / numdays * res: y1 = -10: z1 = res
xs1 = (x1 - (z1 * cost) + (res * cost)) / (cost + 1)
ys1 = (y1 - (z1 * sint) + (res * sint)) / (sint + 1)
Main.GraphBox.CurrentX = xs1: Main.GraphBox.CurrentY = ys1
timeint = 3
Call CalendarDate(m, ckdate$)
timeint = 4
'ckdate$ = Left$(ckdate$, Len(ckdate$) - 1)
Main.GraphBox.Print ckdate$

End Sub

Sub Trans(x(), xr As Integer, xc As Integer, xt(), xtr As Integer, xtc As Integer)

xtr = xc
xtc = xr

For R = 1 To xr
   For c = 1 To xc
      xt(c, R) = x(R, c)
   Next c
Next R


End Sub

Sub TxtToSer(txt$, sdate, errflag)
Dim cons As Integer

txt$ = RTrim$(LTrim$(txt$))

If timeint = 1 And Len(txt$) <> 2 And Len(txt$) <> 1 And Len(txt$) <> 0 Then
    errflag = True
    GoTo ttses:
ElseIf timeint = 2 And Len(txt$) <> 4 And Len(txt$) <> 1 And Len(txt$) <> 0 Then
    errflag = True
    GoTo ttses:
ElseIf timeint = 3 And Len(txt$) <> 6 And Len(txt$) <> 1 And Len(txt$) <> 0 Then
    errflag = True
    GoTo ttses:
ElseIf timeint = 4 And Len(txt$) <> 8 And Len(txt$) <> 1 And Len(txt$) <> 0 Then
    errflag = True
    GoTo ttses:
End If

'sets constant for proper date format
Select Case timeint
   Case 1
      cons = 10 ^ 6
   Case 2
      cons = 10 ^ 4
   Case 3
      cons = 10 ^ 2
   Case 4
      cons = 1
End Select

sdate = Val(txt$) * cons

ttses:
End Sub

Sub WYZCalc(ws(), zs(), ps, xrs, wyzs)
Dim i, j, s1

wyzs = 0
For j = 1 To xrs
   s1 = 0
   For i = 1 To xrs
       s1 = s1 + zs(i, 1) * ps ^ Abs(i - j)
   Next i
   wyzs = wyzs + ws(j, 1) * s1
Next j


End Sub


Public Sub UpdatePerDone(i, lasti, inc)

'logic of subroutine
'for i = 1 to lasti
   'update perdone when i is a multiple of inc
'next i

'update perdone
If Int(i / inc) * inc = i Then
   PerDone.Picture1.Line (0, 0)-(i / lasti * PerDone.Picture1.Width, PerDone.Picture1.Height), BLUE, BF
End If

End Sub




Public Sub PrintHeaderLine(gmodnum, modtype$)

'Main.ValueBar.Print " "; UCase$(filename$) + "   " + vn$(yfld) + " vs. " + vn$(xfld(1));
Main.ValueBar.Print " "; UCase$(filename$);
If Main.G1G2But.Value = False Then
  Main.ValueBar.Print " Single Group " + modtype$ + " model   ";
ElseIf pass = 1 Then
  Main.ValueBar.Print " " + g1name$ + " " + modtype$ + " model   ";
ElseIf pass = 2 Then
  Main.ValueBar.Print " " + g2name$ + " " + modtype$ + " model   ";
End If

End Sub

Public Sub PrintHeaderLineLog(gmodnum, modtype$)

Print #3, " "; UCase$(filename$) + "   " + vn$(yfld) + " vs. " + vn$(xfld(1));
If gmodnum = 1 Then
  Print #3, " Single Group " + modtype$ + " model"
ElseIf gmodnum = 2 And pass = 1 Then
  Print #3, " " + g1name$ + " " + modtype$ + " model"
ElseIf gmodnum = 2 And pass = 2 Then
  Print #3, " " + g2name$ + " " + modtype$ + " model"
End If

End Sub

Public Sub SaveFileParams()
Dim i, j
'saves original data and global file parameters

'from open data cmd
tnp2$ = tnp$
tfn2$ = tfn$
numrecs2 = numrecs
numflds2 = numflds
numfldsd2 = numfldsd
filename2$ = filename$
ndflag2 = ndflag

ReDim d2(numrecs2, numflds2), grpd2(numrecs2)
For i = 1 To numrecs2
   For j = 1 To numflds2
      d2(i, j) = d(i, j)
   Next j
   grpd2(i) = grpd(i)
Next i

'from open var subroutine
ReDim vn2(numflds2)
For i = 1 To numflds2
   vn2$(i) = vn$(i)
Next i

yrfld2 = yrfld
mofld2 = mofld
dyfld2 = dyfld
hrfld2 = hrfld
timeint2 = timeint

Main.MergeCmd.Enabled = True

End Sub



Public Sub CalendarDate2(recnum As Integer, ckdate$)
Dim t$

t$ = "#0"

Select Case timeint2
    Case 1
        ckdate$ = Format$(d2(recnum, yrfld2), t$)
    Case 2
        ckdate$ = Format$(d2(recnum, mofld2), t$) + "/" + Format$(d2(recnum, yrfld2), t$)
    Case 3
        ckdate$ = Format$(d2(recnum, mofld2), t$) + "/" + Format$(d2(recnum, dyfld2), t$) + "/" + Format$(d2(recnum, yrfld2), t$)
    Case 4
        ckdate$ = Format$(d2(recnum, mofld2), t$) + "/" + Format$(d2(recnum, dyfld2), t$) + "/" + Format$(d2(recnum, yrfld2), t$) + "/" + Format$(d2(recnum, hrfld2), t$)
End Select

'nht
If timeint = 4 And mnfld > 0 Then
   ckdate$ = ckdate$ + ":" + Format$(d(recnum, mnfld), t$)
End If

End Sub



Public Static Function DateOnly(datafile, recnum)
Dim t$
On Error GoTo dateonlyce
'returns only the date without an hour indicator

t$ = "#0"

If datafile = "d" Then
   Select Case timeint
       Case 1
           DateOnly = Format$(d(recnum, yrfld), t$)
       Case 2
           DateOnly = Format$(d(recnum, mofld), t$) + "/" + Format$(d(recnum, yrfld), t$)
       Case 3
           DateOnly = Format$(d(recnum, mofld), t$) + "/" + Format$(d(recnum, dyfld), t$) + "/" + Format$(d(recnum, yrfld), t$)
       Case 4
           DateOnly = Format$(d(recnum, mofld), t$) + "/" + Format$(d(recnum, dyfld), t$) + "/" + Format$(d(recnum, yrfld), t$)
   End Select
Else 'data file is d2()
   Select Case timeint2
       Case 1
           DateOnly = Format$(d2(recnum, yrfld2), t$)
       Case 2
           DateOnly = Format$(d2(recnum, mofld2), t$) + "/" + Format$(d2(recnum, yrfld2), t$)
       Case 3
           DateOnly = Format$(d2(recnum, mofld2), t$) + "/" + Format$(d2(recnum, dyfld2), t$) + "/" + Format$(d2(recnum, yrfld2), t$)
       Case 4
           DateOnly = Format$(d2(recnum, mofld2), t$) + "/" + Format$(d2(recnum, dyfld2), t$) + "/" + Format$(d2(recnum, yrfld2), t$)
   End Select
End If

dateonlyce:
Screen.MousePointer = 0
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume dateonlyes
End If
dateonlyes:

End Function

Public Function SerialDate(txt, errflag)
Dim i, dt, yr, mo, dY, hr
On Error GoTo serialdatece

txt = Trim(txt)
If timeint = 4 Then 'hourly data
   For i = Len(txt) To 1 Step -1
      If Mid$(txt, i, 1) = "/" Then Exit For
   Next i
   dt = Left$(txt, i - 1)
   'dt = txt
   yr = Year(dt)
   mo = Month(dt)
   dY = Day(dt)
   hr = Right$(txt, Len(txt) - i)
ElseIf timeint = 3 Then 'daily data
   yr = Year(txt)
   mo = Month(txt)
   dY = Day(txt)
   hr = 0
ElseIf timeint = 2 Then 'monthly data
   yr = Year(txt)
   mo = Month(txt)
   dY = 0
   hr = 0
ElseIf timeint = 1 Then 'yearly data
   yr = txt
   mo = 0
   dY = 0
   hr = 0
End If
'If yr < 100 Then yr = yr + 1900
If yr >= 50 And yr <= 99 Then yr = yr + 1900
If yr >= 0 And yr <= 49 Then yr = yr + 2000
SerialDate = yr * 10 ^ 6 + mo * 10 ^ 4 + dY * 10 ^ 2 + hr

serialdatece:
If Err Then
   errflag = True
   SerialDate = 0
   Resume serialdatees
End If
serialdatees:

End Function

Public Sub TimeScaleChange(tstype, period, periodfld, tsoutfile, numrecsnew, regnum, totregs, increment)
Dim i, j, curperiod
ReDim dtmp(numrecs / period + 2, numflds)
ReDim numdata(numflds), ndf(numflds)

numrecsnew = 0
For i = 1 To numrecs
   
   If d(i, periodfld) <> curperiod Then 'new dtmp record
      For j = 1 To numflds
         If (ndf(j) = True And makend = True) Or dtmp(numrecsnew, j) = "" Then 'set previous record to ndflag
            dtmp(numrecsnew, j) = ndflag
         End If
         'init ndf marker and counter (n)
         ndf(j) = False
         numdata(j) = 0
      Next j
      
      'increment record counter and reset current day
      numrecsnew = numrecsnew + 1
      curperiod = d(i, periodfld)
      
      'fill timestamp fields
      If tstype = "todaily" Then
         dtmp(numrecsnew, hrfld) = ndflag
         dtmp(numrecsnew, dyfld) = d(i, dyfld)
         dtmp(numrecsnew, mofld) = d(i, mofld)
         dtmp(numrecsnew, yrfld) = d(i, yrfld)
      ElseIf tstype = "tomonthly" Then
         dtmp(numrecsnew, hrfld) = ndflag
         dtmp(numrecsnew, dyfld) = ndflag
         dtmp(numrecsnew, mofld) = d(i, mofld)
         dtmp(numrecsnew, yrfld) = d(i, yrfld)
      ElseIf tstype = "toyearly" Then
         dtmp(numrecsnew, hrfld) = ndflag
         dtmp(numrecsnew, dyfld) = ndflag
         dtmp(numrecsnew, mofld) = ndflag
         dtmp(numrecsnew, yrfld) = d(i, yrfld)
      End If
   End If

   'calc new dtmp()
   For j = 1 To numflds
      If j <> hrfld And j <> dyfld And j <> mofld And j <> yrfld Then
         If d(i, j) <> ndflag Then
            If totvars(j) = 1 Then 'calc total
               dtmp(numrecsnew, j) = dtmp(numrecsnew, j) + d(i, j)
            Else 'calc average
               numdata(j) = numdata(j) + 1
               dtmp(numrecsnew, j) = ((numdata(j) - 1) * dtmp(numrecsnew, j) + d(i, j)) / numdata(j)
            End If
         Else 'at least one component of dtmp(numrecsnew,j) was nodata
            ndf(j) = True
         End If
      End If
   Next j

   'update perdone
   regnum = regnum + 1
   Call UpdatePerDone(regnum, totregs, increment)
Next i

For j = 1 To numflds
   If (ndf(j) = True And makend = True) Or dtmp(numrecsnew, j) = "" Then 'set last record to ndflag
      dtmp(numrecsnew, j) = ndflag
   End If
Next j

'transfer new data to d() and to file
Open tsoutfile For Output As #1
numrecs = numrecsnew
ReDim d(numrecs, numflds)
For i = 1 To numrecs
   For j = 1 To numflds
      d(i, j) = dtmp(i, j)
      Print #1, dtmp(i, j);
   Next j
   Print #1,
   
   'update perdone
   regnum = regnum + 1
   Call UpdatePerDone(regnum, totregs, increment)
Next i
ReDim dtmp(1, 1)
Close

End Sub


Public Sub TwoP(x(), y(), n As Integer, b1, seb1, b2, seb2, rmse, cvrmse, r2)
Dim betar As Integer, betac As Integer, numxvars

'dimension arrays
numxvars = 1
'ReDim beta(numxvars + 1, 1), sec(numxvars + 1), sigma(numxvars + 1)

'call regression engine
Call Reg(x(), n, numxvars + 1, y(), n, 1, beta(), betar, betac, sec())
'Call Inf(x(), n, numxvars + 1, y(), n, 1, beta(), betar, betac, sec(), rmse, cvrmse, r2, adjr2, sigma(), roe)
Call Inf(x(), n, numxvars + 1, y(), n, 1, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), roe)

'display results
b1 = beta(1, 1)
seb1 = sigma(1)
b2 = beta(2, 1)
seb2 = sigma(2)

End Sub

Public Sub ThreePMVR(x(), y(), n As Integer, numxvars, Index, xcp, sexcp, ycp, seycp, slope, seslope, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)
Dim i, j, k, xmin, xmax, numints, inc, itration, IndVar, n1, n2
Dim betar As Integer, betac As Integer
Dim regnum, totregs 'for perdone
'sets global variable modtype

'turns on VB error handling
On Error GoTo psmce

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo psmce:
End If

'shows perdone
Screen.MousePointer = 11
PerDone.Caption = "Processing Data..."
PerDone.Show 0
DoEvents

'If index = 0 Then 3PCooling else 3PHeating
modtype = "3P"

'dimensions arrays
'numxvars = 1
ReDim A(n, numxvars + 1)
'ReDim beta(numxvars + 1, 1)
'ReDim sec(numxvars + 1)
'ReDim sigma(numxvars + 1)


'finds x maxs and mins
For i = 1 To n
   If i = 1 Then
      xmin = x(1, 2)
      xmax = x(1, 2)
   End If
   If x(i, 2) < xmin Then xmin = x(i, 2)
   If x(i, 2) > xmax Then xmax = x(i, 2)
Next i

'calcs initial size of search grid
'numints = 10
numints = 100
inc = (xmax - xmin) / numints

'finds total regressions to be performed for perdone
totregs = 2 * (numints - 1) + 1 'for fine search

'finds cp using course then fine grid searches
For itration = 1 To 2
   If itration > 1 Then
      xmin = bestcp - inc
      inc = 2 * inc / numints
   End If

   'tries various changepoints
   For i = 1 To numints - 1
      
      'inits counters
      n1 = 0: n2 = 0

      'fills A()
      cp = xmin + i * inc
      For j = 1 To n
         If x(j, 2) <= cp Then
            n1 = n1 + 1
            If Index = 0 Then '3pc
               IndVar = 0
            Else '3ph
               IndVar = 1
            End If
         Else
            n2 = n2 + 1
            If Index = 0 Then '3pc
               IndVar = 1
            Else '3ph
               IndVar = 0
            End If
         End If
         A(j, 1) = 1
         A(j, 2) = IndVar * (x(j, 2) - cp)
         
         'loads additional independent variables for mvr into A()
         For k = 1 To numxvars - 1
            A(j, 2 + k) = x(j, 2 + k)
         Next k

      Next j

      'calls regression engine
      Call Reg(A(), n, numxvars + 1, y(), n, 1, beta(), betar, betac, sec())
      Call Inf(A(), n, numxvars + 1, y(), n, 1, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), roe)
     
      'records cp with minimum rmse
      If i = 1 Then
         rmsemin = rmse
         bestcp = cp
      ElseIf rmse < rmsemin Then
         rmsemin = rmse
         bestcp = cp
      End If
    
      'updates perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, 1)
   
   Next i
Next itration

'uses best cp to refill A() then rerun regression
n1 = 0: n2 = 0
cp = bestcp
For j = 1 To n
   If x(j, 2) <= cp Then
      n1 = n1 + 1
      If Index = 0 Then
         IndVar = 0
      Else
         IndVar = 1
      End If
   Else
      n2 = n2 + 1
      If Index = 0 Then
         IndVar = 1
      Else
         IndVar = 0
      End If
   End If
   A(j, 1) = 1
   A(j, 2) = IndVar * (x(j, 2) - cp)
   
    For k = 1 To numxvars - 1
       A(j, 2 + k) = x(j, 2 + k)
    Next k
Next j
Call Reg(A(), n, numxvars + 1, y(), n, 1, beta(), betar, betac, sec())
Call Inf(A(), n, numxvars + 1, y(), n, 1, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), roe)

'sets changepoints and slope
'if index = 0 then 3PCooling else 3PHeating
xcp = bestcp
sexcp = inc
ycp = beta(1, 1)
seycp = sigma(1)
slope = beta(2, 1)
seslope = sigma(2)
For k = 2 To numxvars
   'this sets y = 3p + x2*invcoefs(2) + x3*ivcoefs(3) ...
   ivcoefs(k) = beta(k + 1, 1)
   seivcoefs(k) = sigma(k + 1)
Next k


'calcs prediction interval
't = 1.96 + 2.7 / (n - 3)
'predint = t * rmse * (1 + 1 / n) ^ 0.5

'vb error handling
psmce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
  MsgBox """" + Error(Err) + """", , "Error"
  Resume psmes
End If
psmes:

End Sub

Public Sub FourPMVR(x(), y(), n As Integer, numxvars, xcp, sexcp, ycp, seycp, ls, sels, rs, sers, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)

Dim i, j, k, xmin, xmax, inc, numints
Dim itration, xmin4p, n1, n2, cp
Dim IndVar, rmsemin, bestcp, t, predint
Dim betar As Integer, betac As Integer
Dim totregs, regnum 'for perdone

'turns on VB error handling
On Error GoTo cpce

'show perdone
Screen.MousePointer = 11
PerDone.Caption = "Processing Data..."
PerDone.Show 0
DoEvents

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo cpce:
End If

'set modtype (global variable)
modtype = "4P"

'dimension arrays
'numxvars = 1
ReDim A(n, numxvars + 2) ', beta(numxvars + 2, 1), sec(numxvars + 2), sigma(numxvars + 2)

'finds x maxs and mins
For i = 1 To n
   If i = 1 Then
      xmin = x(1, 2)
      xmax = x(1, 2)
   End If
   If x(i, 2) < xmin Then xmin = x(i, 2)
   If x(i, 2) > xmax Then xmax = x(i, 2)
Next i

'calcs initial size of search grid
'numints = 10
numints = 100
inc = (xmax - xmin) / numints

'find total regressions to be performed for perdone
totregs = 2 * (numints - 1) + 1 'for fine search

'finds cp by using course then fine grid searches
For itration = 1 To 2
   
   'set xmin4p and inc depending on course or fine grid search
   If itration = 1 Then
      xmin4p = xmin
   Else '(itration = 2 for fine grid...
      xmin4p = bestcp - inc
      inc = 2 * inc / numints
   End If
   
   'tries various changepoints
   For i = 1 To numints - 1
      
      'inits counters
      n1 = 0: n2 = 0

      'fills A()
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
         
         'loads additional independent variables for mvr into A()
         For k = 1 To numxvars - 1
            A(j, 3 + k) = x(j, 2 + k)
         Next k
      Next j

      'calls regression engine
      Call Reg(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec())
      Call Inf(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), roe)
      
      'records cp with minimum rmse
      If i = 1 Then
         rmsemin = rmse
         bestcp = cp
      ElseIf rmse < rmsemin Then
         rmsemin = rmse
         bestcp = cp
      End If
      
      'updates perdone
      regnum = regnum + 1
      Call UpdatePerDone(regnum, totregs, 1)
   
   Next i
Next itration

'uses best cp to fill A, then reruns regression with best cp
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
   
   For k = 1 To numxvars - 1
      A(j, 3 + k) = x(j, 2 + k)
   Next k
Next j
Call Reg(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec())
Call Inf(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), roe)

'calculates changepoints and slopes from beta() for yhat = ycp + ls(t-xcp)-  +  rs(t-xcp)+
xcp = bestcp
sexcp = inc
ycp = beta(1, 1) + beta(2, 1) * bestcp
If Abs(beta(2, 1)) > 0 And Abs(xcp) > 0 Then
   seycp = (sigma(1) ^ 2 + (beta(2, 1) * xcp) ^ 2 * ((sigma(2) / beta(2, 1)) ^ 2 + (sexcp / xcp) ^ 2)) ^ 0.5
Else
   seycp = sexcp
End If
ls = beta(2, 1)
sels = sigma(2)
rs = beta(2, 1) + beta(3, 1)
sers = (sigma(2) ^ 2 + sigma(3) ^ 2) ^ 0.5

For k = 2 To numxvars
   ivcoefs(k) = beta(k + 2, 1)
   seivcoefs(k) = sigma(k + 2)
Next k

'calcs prediction interval
't = 1.96 + 2.7 / (n - 3)
'predint = t * rmse * (1 + 1 / n) ^ 0.5

'vb error handling
cpce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   MsgBox """" + Error(Err) + """"
   Resume cpes
End If
cpes:

End Sub

Public Sub FivePMVR(x(), y(), n As Integer, numxvars, xcp1, sexcp1, xcp2, sexcp2, ycp, seycp, ls, sels, rs, sers, ivcoefs(), seivcoefs(), rmse, cvrmse, r2)
Dim i, j, k, m, xmin, xmax, numints
Dim gridpass, numleft, numright, indvar1, indvar2, rmsebest, cp1best, cp2best
Dim betar As Integer, betac As Integer

Dim totregs, regnum 'for perdone

'turns on VB error handling
On Error GoTo p5ce

'prints error if n = 0
If n = 0 Then
   MsgBox "No data is available to model.", , "Error"
   GoTo p5ce:
End If

'show perdone
Screen.MousePointer = 11
PerDone.Caption = "Processing Data..."
PerDone.Show 0
DoEvents
 
'sets modtype
modtype = "5P"

'dims arrays
ReDim A(n, numxvars + 2) ', beta(numxvars + 2, 1), sec(numxvars + 2), sigma(numxvars + 2)

'sets number of search intervals
numints = 9
ReDim cpl(numints), cpr(numints)

'find total regressions to be performed for perdone
totregs = 16   'for fine search
For i = 1 To numints - 2   'for course search
   totregs = totregs + i
Next i

'finds x maxs and mins
For i = 1 To n
   If i = 1 Then
      xmin = x(1, 2)
      xmax = x(1, 2)
   End If
   If x(i, 2) < xmin Then xmin = x(i, 2)
   If x(i, 2) > xmax Then xmax = x(i, 2)
Next i

'finds best changepoints using course then fine grid search
For gridpass = 1 To 2

   'finds change points
   If gridpass = 1 Then 'course (first) grid search
      
      'calcs initial search grid
      inc = (xmax - xmin) / numints
      numleft = numints - 1
      numright = numints - 1
      For i = 1 To numints - 1
         cpl(i) = xmin + i * inc
         cpr(i) = xmin + i * inc
      Next i
      
   Else  'fine (fine) grid search
     
      'searches on either side of xcp1 and xcp2
      numleft = 4
      numright = 4
      cpl(1) = cp1best - 0.666 * inc
      cpl(2) = cp1best - 0.333 * inc
      cpl(3) = cp1best + 0.333 * inc
      cpl(4) = cp1best + 0.666 * inc
      cpr(1) = cp2best - 0.666 * inc
      cpr(2) = cp2best - 0.333 * inc
      cpr(3) = cp2best + 0.333 * inc
      cpr(4) = cp2best + 0.666 * inc
   End If
   
   'trys all combinations of change points
   For i = 1 To numleft
      For j = 1 To numright
                  
         'sets change points
         xcp1 = cpl(i)
         xcp2 = cpr(j)
         
         'performs regression
         If xcp1 < xcp2 Then
            numits = numits + 1
            For k = 1 To n
               If x(k, 2) <= xcp1 Then
                  indvar1 = 1
               Else
                  indvar1 = 0
               End If
                  
               If x(k, 2) <= xcp2 Then
                  indvar2 = 0
               Else
                  indvar2 = 1
               End If
               A(k, 1) = 1
               A(k, 2) = indvar1 * (x(k, 2) - xcp1)
               A(k, 3) = indvar2 * (x(k, 2) - xcp2)
               
               'loads additional independent variables for mvr into A()
               For m = 1 To numxvars - 1
                  A(k, 3 + m) = x(k, 2 + m)
               Next m
            Next k
            Call Reg(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec())
            Call Inf(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), P)

            'selects changepoints which give minimum RMSE
            If (gridpass = 1 And numits = 1) Or rmse < rmsebest Then
               rmsebest = rmse
               cp1best = xcp1
               cp2best = xcp2
            End If
            
            'update perdone
            regnum = regnum + 1
            Call UpdatePerDone(regnum, totregs, 1)

         End If
      Next j
   Next i
Next gridpass

'sets final change points equal to best-fit changepoints
xcp1 = cp1best
xcp2 = cp2best
sexcp1 = inc * 0.333
sexcp2 = inc * 0.333

'fills A() in accordance with xcp1 and xcp2
For k = 1 To n
   If x(k, 2) <= xcp1 Then
      indvar1 = 1
   Else
      indvar1 = 0
   End If
   If x(k, 2) <= xcp2 Then
      indvar2 = 0
   Else
      indvar2 = 1
   End If
   A(k, 1) = 1
   A(k, 2) = indvar1 * (x(k, 2) - xcp1)
   A(k, 3) = indvar2 * (x(k, 2) - xcp2)
   
   'loads additional independent variables for mvr into A()
   For m = 1 To numxvars - 1
      A(k, 3 + m) = x(k, 2 + m)
   Next m
Next k

'performs regression with best changepoints
Call Reg(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec())
Call Inf(A(), n, numxvars + 2, y(), n, 1, beta(), betar, betac, sec(), ymean, xmean(), rmse, cvrmse, r2, adjr2, sigma(), P)

'sets coefficients and standard error of coefficents
ycp = beta(1, 1)
seycp = sigma(1)
ls = beta(2, 1)
sels = sigma(2)
rs = beta(3, 1)
sers = sigma(3)
For k = 2 To numxvars
   ivcoefs(k) = beta(k + 2, 1)
   seivcoefs(k) = sigma(k + 2)
Next k

p5ce:
Screen.MousePointer = 0
Unload PerDone
If Err Then
   MsgBox """" + Error(Err) + """", , "Error"
   Resume p5es
End If
p5es:

End Sub

Public Sub FillXYGrp(pass, numxvars, x(), y(), n)
Dim j, k, count

'redimensions X(), Y(), and IndexNum()
ReDim IndexNum(numrecs) As Integer

If Main.G1G2But.Value = True And pass = 1 Then

    count = 0
    For j = 1 To numrecs
       'goes to next record if no data
       If grpd(j) = 2 Or d(j, yfld) = ndflag Then GoTo mlrnr1:
       For k = 1 To numxvars
          If d(j, xfld(k)) = ndflag Then GoTo mlrnr1:
       Next k
       
       'fills x() and y()
       count = count + 1
       x(count, 1) = 1
       For k = 1 To numxvars
          x(count, k + 1) = d(j, xfld(k))
       Next k
       y(count, 1) = d(j, yfld)
       IndexNum(count) = j
mlrnr1: Next j
    n = count

ElseIf Main.G1G2But.Value = True And pass = 2 Then

    count = 0
    For j = 1 To numrecs
       'goes to next record if nodata
       If grpd(j) = 1 Or d(j, yfld) = ndflag Then GoTo mlrnr2:
       For k = 1 To numxvars
          If d(j, xfld(k)) = ndflag Then GoTo mlrnr2:
       Next k
       
       'fills x() and y()
       count = count + 1
       x(count, 1) = 1
       For k = 1 To numxvars
          x(count, k + 1) = d(j, xfld(k))
       Next k
       y(count, 1) = d(j, yfld)
       IndexNum(count) = j
mlrnr2: Next j
    n = count

ElseIf Main.G1G2But.Value = False Then 'all-days......
    
    count = 0
    For j = 1 To numrecs
       'goes to next record if nodata
       If d(j, yfld) = ndflag Then GoTo mlrnr3:
       For k = 1 To numxvars
          If d(j, xfld(k)) = ndflag Then GoTo mlrnr3:
       Next k
       
       'fills x() and y()
       count = count + 1
       x(count, 1) = 1
       For k = 1 To numxvars
          x(count, k + 1) = d(j, xfld(k))
       Next k
       y(count, 1) = d(j, yfld)
       IndexNum(count) = j
mlrnr3: Next j
    n = count

End If

End Sub
