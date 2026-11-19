Attribute VB_Name = "GLOBAL"
'variable declarations common to entire program
Global Msg$

'when opening data file
Global tfn$, tnp$, fit$, filename$
Global LastChange As Integer
Global d(), numflds As Integer, numrecs ' As Integer
Global numfldsd As Integer
Global fopen
Global opencode

'when defining variable names
Global vn$(), g1name$, g2name$
Global yrfld As Integer, mofld As Integer
Global dyfld As Integer, hrfld As Integer
Global mnfld As Integer 'nht
Global timeint As Integer
Global grpd(), grpxy(), grpdefined, grptype

'when selecting dates
Global ckdate$, sdate
Global startd, endd, SDD(20), EDD(20)
Global numbdate As Integer, cons As Integer

'when setting variables
Global x(), y()
Global n As Integer, xfld(-1 To 7) As Integer, yfld As Integer
Global ymin, ymax, yrange
Global xmin, xmax, xrange
Global yminr, ymaxr, yranger, yroot, ypower, yfmt, yintindex
Global xminr, xmaxr, xranger, xroot, xpower, xfmt, xintindex

'when manually setting scales
Global numxlabels As Integer, numylabels As Integer

'for zooming timeseries
Global x1, y1, zoom As Integer, zoommode As Integer
Global i1 As Integer, i2 As Integer
Global lastx, lasty

'when using single scales for time series 2 graphs
Global ss, yfldo, ynameo$, yminss, ymaxss

'for histograms
Global histcancel As Integer, hnumints As Integer
Global hxmaxset, hxminset, hsetminmax

'when displaying data cursor
Global gtype$, newxygraph
Global currec As Integer, IndexNum() As Integer

'when animating
Global numpnts As Integer, delay As Integer, pause As Integer

'for hourly profiles and 3d
Global res, maskmin(), maskmax(), maskminnew(), maskmaxnew()
Global xclick, yclick, sint, cost, numdays, dy1, thetad

'when determining a model
Global xmean(7), ymean, ndflag, modtype$
Global beta(8, 1), sec(8), sigma(8), resx(), yhat()
Global xr As Integer, xc As Integer
Global yr As Integer, yc As Integer
Global betar As Integer, betac As Integer
Global pass As Integer
Global autocor, dw
'Global ymedian

'for binfit
Global binfitcancel As Integer, numints As Integer
Global xmaxset, xminset, setminmax

'for mlr
Global cancelmlr As Integer

'when modifying variables
Global cnst(), f$()

'when adding groups
Global agcancelflag As Integer
Global g1start, g1end
Global gmodnum

'when calculating savings
Global PerMeasErr, premodtype$, rmsepre, npre 'for all models
Global ymeanpre, ymedianpre 'for mean and median models
Global betapre(8) 'for 2P and mlr models
Global xcppre, ycppre, slopepre, indexpre 'for 3p models
Global lspre, rspre 'and xcppre, ypcpre for 4p models
Global xcp1pre, xcp2pre ' and ycppre, lspre, rspre 'for 5p models
Global numxvarspre, ivcoefspre() 'for mvr and cp+mvr models
Global ybarbf(), xminbf, xmaxbf 'for binfit
Global modelstats, modeleqn 'diplays baseline stats and eqn
Global calcsavedisplay 'for displaying timeseries with actual and predicted use

'for simulation
Global floorarea, fi, ua, tr, numoccs, laton
Global cavflow, vavflow, foafixed, thcfixed, tcc
Global cavon, econon, hdreseton, ophdreseton
Global numsims

'change timescales
Global dtmp(), totvars()
Global timeinthr, timeintdy, timeintmo, timeintyr
Global numrecshr, numrecsdy, numrecsmo, numrecsyr
Global tsoptscancel, makend, tsoptsshown

'opening multiple data files
Global numfiles
Global tnp2$, tfn2$, numrecs2  As Integer, numflds2  As Integer, numfldsd2  As Integer, filename2$, ndflag2
Global vn2$(), d2(), grpd2()
Global yrfld2, mofld2, dyfld2, hrfld2, timeint2
Global getvnsfrommerge

'for editor
Global filesaved, fscancel$, newfile, formcap
Global opath$, ofilename$, focancel$
Global findcancel, findtxt$, txtchanged, lastposition, txtposition
Global nsno, nscancel, nsyes

'for eseditor
Global fileselected

'calc lea
Global AdditionalVar, calcleadisplay

'Booleans
Global Const MODAL = 1

' KeyCode (KeyDown, KeyUp)
Global Const KEY_LEFT = &H25
Global Const KEY_RIGHT = &H27
Global Const KEY_PRINT = &H2A

' BackColor, ForeColor, FillColor (standard RGB colors: form, controls)
Global Const BLACK = &H0&
Global Const RED = &HFF&
Global Const GREEN = &HFF00&
Global Const YELLOW = &HFFFF&
Global Const BLUE = &HFF0000
Global Const MAGENTA = &HFF00FF
Global Const CYAN = &HFFFF00
Global Const WHITE = &HFFFFFF

Global Const CF_BITMAP = 2
Global Const CF_DIB = 8
Global Const CF_METAFILE = 3
Global Const CF_PALETTE = 9

