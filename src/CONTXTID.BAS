Attribute VB_Name = "ContextIDs"
Option Explicit
'=====================================================================
'=====================================================================
'
'This source code contains the following routines:
'  o SetAppHelp() 'Called in the main Form_Load event to register your
'                 'program with WINHELP.EXE
'  o QuitHelp()    'Deregisters your program with WINHELP.EXE. Should
'                  'be called in your main Form_Unload event
'  o ShowHelpTopic(Topicnum) 'Brings up context sensitive help based on
'                  'any of the following CONTEXT IDs
'  o ShowContents  'Displays the startup topic
'  o HelpWindowSize(x,y,dx,dy) ' Position help window in a screen
'                              ' independent manner
'  o SearchHelp()  'Brings up the windows help KEYWORD SEARCH dialog box
'***********************************************************************
'
'=====================================================================
'List of Context IDs for <EEHELP>
'=====================================================================
Global Const Hlp_Description = 10    'Main Help Window
Global Const Hlp_Getting_Startedx = 30    'Main Help Window
Global Const Hlp_Menu_Bar = 60    'Main Help Window
Global Const Hlp_Graph_Window = 70    'Main Help Window
Global Const Hlp_Variable_List = 80    'Main Help Window
Global Const Hlp_Graph_Buttons = 90    'Main Help Window
Global Const Hlp_Model_Buttons = 100    'Main Help Window
Global Const Hlp_Grouping_Buttons = 110    'Main Help Window
Global Const Hlp_Printing = 150    'Main Help Window
Global Const Hlp_Copying_Information = 170    'Main Help Window
Global Const Hlp_Text_Editor = 180    'Main Help Window
Global Const Hlp_Work_Window = 580    'Main Help Window
Global Const Hlp_Value_Window = 590    'Main Help Window
Global Const Hlp_File_Menu = 660    'Main Help Window
Global Const Hlp_Input_Data1 = 670    'Main Help Window
Global Const Hlp_Saving_Data1 = 710    'Main Help Window
Global Const Hlp_Edit_Menu = 730    'Main Help Window
Global Const Hlp_Data_Menu = 750    'Main Help Window
Global Const Hlp_Naming_Variables = 760    'Main Help Window
Global Const Hlp_Setting_NoxData = 770    'Main Help Window
Global Const Hlp_Creating_SubxSets = 780    'Main Help Window
Global Const Hlp_Displaying_The = 790    'Main Help Window
Global Const Hlp_Modifying_Data = 800    'Main Help Window
Global Const Hlp_Creating_New = 810    'Main Help Window
Global Const Hlp_Deleting_a = 820    'Main Help Window
Global Const Hlp_Grouping_Data = 830    'Main Help Window
Global Const Hlp_Changing_Time = 840    'Main Help Window
Global Const Hlp_Averaging_Intermittent = 850    'Main Help Window
Global Const Hlp_Graphical_Analysis = 860    'Main Help Window
Global Const Hlp_TimexSeries_Graphs = 870    'Main Help Window
Global Const Hlp_XY_Graphs = 880    'Main Help Window
Global Const Hlp_Animated_XY = 890    'Main Help Window
Global Const Hlp_Histograms = 900    'Main Help Window
Global Const Hlp_Value_Identification = 910    'Main Help Window
Global Const Hlp_Grid_Lines = 920    'Main Help Window
Global Const Hlp_Modeling = 930    'Main Help Window
Global Const Hlp_Mean_and = 940    'Main Help Window
Global Const Hlp_2P_and = 950    'Main Help Window
Global Const Hlp_3P_and = 960    'Main Help Window
Global Const Hlp_4P_and = 970    'Main Help Window
Global Const Hlp_5P_and = 980    'Main Help Window
Global Const Hlp_Bin_Fit = 990    'Main Help Window
Global Const Hlp_Running_a = 1000    'Main Help Window
Global Const Hlp_Modeling_Groups = 1010    'Main Help Window
Global Const Hlp_Prediction_Intervals = 1020    'Main Help Window
Global Const Hlp_Residual_Analysis = 1030    'Main Help Window
Global Const Hlp_Model_Parameters = 1040    'Main Help Window
Global Const Hlp_Measuring_Retrofit = 1070    'Main Help Window
Global Const Hlp_Savings_Measurement = 1080    'Main Help Window
Global Const Hlp_Acknowledgments = 1090    'Main Help Window
Global Const Hlp_Opening_Data1 = 1110    'Main Help Window
Global Const Hlp_Lean_Energy = 1120    'Main Help Window
Global Const Hlp_Analysis_Menu = 1150    'Main Help Window
Global Const Hlp_Energy_Explorer = 1160    'Main Help Window
Global Const Hlp_Energy_Explorer1 = 1170    'Main Help Window
Global Const Hlp_Exiting_Energy = 1180    'Main Help Window
'=====================================================================
'
'
'  Help engine section.

' Commands to pass WinHelp()
Global Const HELP_CONTEXT = &H1 '  Display topic in ulTopic
Global Const HELP_QUIT = &H2    '  Terminate help
Global Const HELP_FINDER = &HB  '  Display Contents tab
Global Const HELP_INDEX = &H3   '  Display index
Global Const HELP_HELPONHELP = &H4      '  Display help on using help
Global Const HELP_SETINDEX = &H5        '  Set the current Index for multi index help
Global Const HELP_KEY = &H101           '  Display topic for keyword in offabData
Global Const HELP_MULTIKEY = &H201
Global Const HELP_CONTENTS = &H3     ' Display Help for a particular topic
Global Const HELP_SETCONTENTS = &H5  ' Display Help contents topic
Global Const HELP_CONTEXTPOPUP = &H8 ' Display Help topic in popup window
Global Const HELP_FORCEFILE = &H9    ' Ensure correct Help file is displayed
Global Const HELP_COMMAND = &H102    ' Execute Help macro
Global Const HELP_PARTIALKEY = &H105 ' Display topic found in keyword list
Global Const HELP_SETWINPOS = &H203  ' Display and position Help window

#If Win32 Then
    Type HELPWININFO
      wStructSize As Long
      X As Long
      Y As Long
      dX As Long
      dY As Long
      wMax As Long
      rgChMember As String * 2
    End Type
    Declare Function WinHelp Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
    Declare Function WinHelpByInfo Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, dwData As HELPWININFO) As Long
    Declare Function WinHelpByStr Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData$) As Long
    Declare Function WinHelpByNum Lib "User32.dll" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, ByVal dwData&) As Long
    Dim m_hWndMainWindow as Long ' hWnd to tell WINHELP the helpfile owner

#Else
    Type HELPWININFO
        wStructSize As Integer
        X As Integer
        Y As Integer
        dX As Integer
        dY As Integer
        wMax As Integer
        rgChMember As String * 2
    End Type
    Declare Function WinHelp Lib "User" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData As Any) As Integer
    Declare Function WinHelpByInfo Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, dwData As HELPWININFO) As Integer
    Declare Function WinHelpByStr Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData$) As Integer
    Declare Function WinHelpByNum Lib "User" Alias "WinHelp" (ByVal hWnd As Integer, ByVal lpHelpFile As String, ByVal wCommand As Integer, ByVal dwData&) As Integer
    Dim m_hWndMainWindow as Integer ' hWnd to tell WINHELP the helpfile owner

#End If
Dim MainWindowInfo as HELPWININFO
Sub SetAppHelp (ByVal hWndMainWindow)
'=====================================================================
'To use these subroutines to access WINHELP, you need to add
'at least this one subroutine call to your code
'     o  In the Form_Load event of your main Form enter:
'        Call SetAppHelp(Me.hWnd) 'To setup helpfile variables
'         (If you are not interested in keyword searching or context
'         sensitive help, this is the only call you need to make!)
'=====================================================================
    m_hWndMainWindow = hWndMainWindow
    If Right$(Trim$(App.Path),1) = "\" then
        App.HelpFile = App.Path + "EEHELP.HLP"
    else
        App.HelpFile = App.Path + "\EEHELP.HLP"
    end if
#If Win32 Then
    MainWindowInfo.wStructSize = 26
#Else
    MainWindowInfo.wStructSize = 14
#End If 
    MainWindowInfo.X=256
    MainWindowInfo.Y=256
    MainWindowInfo.dX=512
    MainWindowInfo.dY=512
    MainWindowInfo.rgChMember=Chr$(0)+Chr$(0)
End Sub
Sub QuitHelp ()
    Dim Result as Variant
    Result = WinHelp(m_hWndMainWindow, App.HelpFile, HELP_QUIT, Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0))
End Sub
Sub ShowHelpTopic (ByVal ContextID As Long)
'=====================================================================
'  FOR CONTEXT SENSITIVE HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic(<any Hlpxxx entry above>)
'=====================================================================
'  TO ADD FORM LEVEL CONTEXT SENSITIVE HELP...
'=====================================================================
'     o  For FORM level context sensetive help, you should set each 
'        Me.HelpContext=<any Hlp_xxx entry above>
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXT, Clng(ContextID))

End Sub
Sub ShowHelpTopic2 (ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 2 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic2(<any Hlpxxx entry above>)
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile &">HlpWnd02", HELP_CONTEXT, Clng(ContextID))

End Sub
Sub ShowHelpTopic3 (ByVal ContextID As Long)
'=====================================================================
'  DISPLAY CONTEXT SENSITIVE HELP IN WINDOW 3 ...
'=====================================================================
'     o   For 'Help button' controls, you can call:
'         Call ShowHelpTopic3(<any Hlpxxx entry above>)
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile &">HlpWnd03", HELP_CONTEXT, Clng(ContextID))

End Sub
Sub ShowGlossary ()
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXT, Clng(64000))

End Sub
Sub ShowPopupHelp (ByVal ContextID As Long)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTEXTPOPUP, Clng(ContextID))

End Sub
Sub DoHelpMacro (ByVal MacroString As String)
'=====================================================================
'  FOR POPUP HELP IN RESPONSE TO A COMMAND BUTTON ...
'=====================================================================
    Dim Result as Variant

    Result = WinHelpByStr(m_hWndMainWindow, APP.HelpFile, HELP_COMMAND, ByVal(Macrostring))

End Sub
Sub ShowHelpContents ()
'=====================================================================
'  DISPLAY HELP STARTUP TOPIC IN RESPONSE TO A COMMAND BUTTON or MENU ...
'=====================================================================
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_CONTENTS, Clng(0))

End Sub
Sub ShowContentsTab ()
'=====================================================================
'  DISPLAY Contents tab (*.CNT)
'=====================================================================
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_FINDER, Clng(0))

End Sub
Sub ShowHelpOnHelp ()
'=====================================================================
'  DISPLAY HELP for WINHELP.EXE  ...
'=====================================================================
'
    Dim Result as Variant

    Result = WinHelpByNum(m_hWndMainWindow, APP.HelpFile, HELP_HELPONHELP, Clng(0))

End Sub

Sub SearchHelp ()
'=====================================================================
'  TO ADD KEYWORD SEARCH CAPABILITY...
'=====================================================================
'     o   In your Help|Search menu selection, simply enter:
'         Call SearchHelp() 'To invoke helpfile keyword search dialog
'
    Dim Result as Variant

    Result = WinHelp(m_hWndMainWindow, APP.HelpFile, HELP_PARTIALKEY, ByVal "" )

End Sub

Sub SearchHelpKeyWord (Argument as String)
'=====================================================================
'  TO ADD KEYWORD SEARCH CAPABILITY...
'=====================================================================
'     o   In your Help|Search menu selection, simply enter:
'         Call SearchHelp() 'To invoke helpfile keyword search dialog
'
    Dim Result as Variant

    Result = WinHelp(m_hWndMainWindow, APP.HelpFile, HELP_PARTIALKEY, ByVal Trim$(Argument))

End Sub
Sub HelpWindowSize (x As Integer, y As Integer, wx As Integer, wy As Integer)
'=====================================================================
'  TO SET THE SIZE AND POSITION OF THE MAIN HELP WINDOW...
'=====================================================================
'     o   Call HelpWindowSize(x, y, dx, dy), where:
'             x = 1-1024 (position from left edge of screen)
'             y = 1-1024 (position from top of screen)
'             dx= 1-1024 (width)
'             dy= 1-1024 (height)
'
    Dim Result as Variant
    MainWindowInfo.x = x
    MainWindowInfo.y = y
    MainWindowInfo.dx = wx
    MainWindowInfo.dy = wy
    Result = WinHelpByInfo(m_hWndMainWindow, App.HelpFile, HELP_SETWINPOS, MainWindowInfo)
End Sub
