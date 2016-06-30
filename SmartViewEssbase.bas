Attribute VB_Name = "SmartViewEssbase"
Public bpCancel As Boolean
Public bpMessages As Boolean
Public bpError As Boolean
Sub vbaSettings()
   bpCancel = False                 'used when the logon is cancelled
   bpError = False                  'used to cancel the connnection procedure and retrievals when an error occurs
   bpMessages = False                'make true to show all messages
End Sub
Sub vbaRetrieve(stSheet As String)
    Dim vtGrid As Variant
    Dim vtDimNames As Variant
    Dim vtPOVNames As Variant
    Dim stApplication As String
    Dim stApplicationFriendly As String
    Dim bIsConnected As Integer
    Dim stServer As String
    Dim stProvider As String
    Dim intConnection As Integer
    Dim stRange As Range, stRangeName As String
    
    If bpMessages Then application.ScreenUpdating = False
    
    Sheets(stSheet).Select
    Sheets(stSheet).Activate
    
    stApplication = Range("setApplication")                                     'Application to be refreshed
    stServer = Range("setServer")                                               'Server name
    stApplicationFriendly = stServer & "_" & stApplication & "_" & stApplication 'Frienly name
    stProvider = Range("setProvider")                                           'Provider URL
    stRangeName = Range("setDataRangeRetrieve")
    Set stRange = Range(stRangeName)
    
    'Set default settings
    Call vbaSettings
    
    'Delete any metadata available on the worksheet
    Call vbaHypDeleteMetaData
    
    'Connect to Smartview
    Call vbaConnect(stSheet, stApplication, stApplicationFriendly, stServer, stProvider)
    
    If SmartViewLogon.varCancel.Value Or bpCancel Then Exit Sub
    
    'If connected then set the options
    Call vbaOption
    
    x = HypRetrieveRange(stSheet, stRange, stApplicationFriendly)               'Use Range as retrieval
    'X = HypRetrieveRange(stSheet, Null, stApplicationFriendly)                 'This line is used to retrieve complete sheet
    
    ActiveSheet.Outline.ShowLevels rowlevels:=1, columnlevels:=1
    ActiveSheet.Outline.ShowLevels rowlevels:=0, columnlevels:=0
    
    'Delete any metadata available on the worksheet
    Call vbaHypDeleteMetaData
    
    If x = 0 Then
        If bpMessages Then MsgBox ("Retrieve sheet " & stSheet & " successful with application " & stApplication & ".")
    Else
        If bpMessages Then MsgBox ("Error: " & x & " Retrieve sheet " & stSheet & " failed with application " & stApplication & ". (retrieve)")
        bpError = True
    End If
    
    If bpMessages Then application.ScreenUpdating = True
 
 End Sub


Sub vbaConnect(stSheet As String, stApplication As String, stApplicationFriendly As String, stServer As String, stProvider As String)
    Dim intConnection As Integer
    Dim lngActive As Long
    
    intConnection = False
    intConnection = HypConnectionExists(stApplicationFriendly)
    lngActive = 1
    
    If Not intConnection Then
        SmartViewLogon.Show
        If SmartViewLogon.varCancel.Value Then
            Exit Sub
        Else
            Z = HypCreateConnection(stSheet, SmartViewLogon.varUsername.Value, SmartViewLogon.varPassword.Value, HYP_ESSBASE, stProvider, stServer, stApplication, stApplication, stApplicationFriendly, stApplicationFriendly)
            If Z = 0 Then
                'connect with the application
                Y = HypConnect(stSheet, SmartViewLogon.varUsername.Value, SmartViewLogon.varPassword.Value, stApplicationFriendly)
                If Not Y = 0 Then End
                'set connection as active
                lngActive = HypSetActiveConnection(stApplicationFriendly)
                If Y = 0 And lngActive = 0 Then
                    'connection is made with the application
                Else
                    If bpMessages Then MsgBox ("Error: " & Y & " Connection failed.")
                    bpCancel = True
                    bpError = True
                    Exit Sub
                End If
            Else
                If bpMessages Then MsgBox ("Error: " & Z & " Create connection failed.")
                bpCancel = True
                bpError = True
                Exit Sub
            End If
        End If
    Else
        lngActive = HypSetActiveConnection(stApplicationFriendly)
        If lngActive = 0 Then
            'sheet is connectetd with the application
        Else
            SmartViewLogon.Show
            If SmartViewLogon.varCancel.Value Then
                Exit Sub
            Else
                'connect with the application
                Y = HypConnect(stSheet, SmartViewLogon.varUsername.Value, SmartViewLogon.varPassword.Value, stApplicationFriendly)
                'set connection as active
                lngActive = HypSetActiveConnection(stApplicationFriendly)
                If Y = 0 And lngActive = 0 Then
                    'connection is made with the application
                    If bpMessages Then MsgBox ("Connection is made with the application: " & stApplicationFriendly)
                Else
                    If bpMessages Then MsgBox ("Error: " & Y & " Connection failed.")
                    bpError = True
                    
                    application.ScreenUpdating = True
                    MsgBox "Error, could not establish connection. Make sure you are on the Staples network and contact your BusSysHyperion team if problem remains", vbCritical, "Connection Error!"
                    End
                End If
            End If
        End If
    End If
   
    'Username and password are cleared with below settings
    SmartViewLogon.varUsername.Value = ""
    SmartViewLogon.varPassword.Value = ""
   
 End Sub
Sub vbaOption()
    
'Set POV to Excel addin style
sts = HypShowPov(False)

'Set Essbase options

sts = HypSetOption(1, 0, "")                    'Set zoom next level
sts = HypSetOption(2, True, "")                 'Included selection
sts = HypSetOption(5, 0, "")                    'No Identation
sts = HypSetOption(6, False, "")                'No suppression of missing
sts = HypSetOption(7, False, "")                'No suppression of zeros
sts = HypSetOption(8, False, "")                'No suppression of underscore
sts = HypSetOption(9, False, "")                'No suppression of noaccess
sts = HypSetOption(10, False, "")               'No suppression of repeatedmembers
sts = HypSetOption(11, False, "")               'No suppression of invalid
sts = HypSetOption(12, 1, "")                   'Ancestor at bottom
sts = HypSetOption(13, "#NumericZero", "")      'Missing is numeric zero
sts = HypSetOption(14, "#NumericZero", "")      'Noaccess is numeric zero
sts = HypSetOption(16, 0, "")                   'Show Name only
sts = HypSetOption(21, True, "")                'Preserve formulas
sts = HypSetOption(30, True, "")                'Use Excel formatting
'sts = HypSetOption(31, True, "")                'Retain Numeric formatting
sts = HypSetOption(36, False, "")               'Do not adjust cell width
sts = HypSetOption(101, False, "")              'Do not use doubleclick for ad-hoc
sts = HypSetOption(102, False, "")               'Enables undo
sts = HypSetOption(107, True, "")               'Reduce Excel file size
sts = HypSetOption(111, 0, "")                  '9 undo actions


'1    HSV_ZOOMIN = Number                       'Set zoom level
'              0 = next level
'              1 = all levels
'              2 = same level
'              3 = sibling level
'              4 = same level
'              5 = same generation
'              6 = formulas
'2    HSV_INCLUDE_SELECTION = Boolean           'Selects the Include Selections check box
'3    HSV_WITHIN_SELECTEDGROUP = Boolean        'Selects the Within Selected Group check box
'4    HSV_REMOVE_UNSELECTEDGROUP = Boolean
'5    HSV_INDENTATION = Number                  'Selects the Remove Unselected Groups check box
'                   0 = No Indentation
'                   1 = Indent sub items
'                   2 = Indent totals
'6    HSV_SUPPRESSROWS_MISSING = Boolean        'Suppresses rows that contain no data or are missing data
'7    HSV_SUPPRESSROWS_ZEROS = Boolean          'Suppresses rows that contain only zeroes
'8    HSV_SUPPRESSROWS_UNDERSCORE = Boolean     'Suppresses rows that contain underscore characters in member names
'9    HSV_SUPPRESSROWS_NOACCESS = Boolean       'Suppress rows that contain data that the user does not have the security access to view
'10   HSV_SUPPRESSROWS_REPEATEDMEMBERS = Boolean 'Suppresses rows that contain repeated member names, regardless of grid orientation.
'11   HSV_SUPPRESSROWS_INVALID = Boolean        'Suppresses rows that contain only invalid values
'12   HSV_ANCESTOR_POSITION = Number            'Specifies an ancestor position in hierarchies:
'                         0 = Top
'                         1 = Bottom
'13   HSV_MISSING_LABEL = Text                  'Displays #Missing, #Numeric Zero, or the text of your choice in data cells that contain missing data.
'14   HSV_NOACCESS_LABEL = Text                 'Displays #NoAccess, #Numeric Zero, or the text of your choice in data cells that the user does not have permission to view.
'15   HSV_CELL_STATUS
'16   HSV_MEMBER_DISPLAY = Number               'Specifies how to display member names in cells:
'                      0 = Name only
'                      1 = Name and Description
'                      2 = Description only
'17   HSV_INVALID_LABEL
'18   HSV_SUBMITZERO
'19   HSV_19                                    'unused reserved for future use
'20   HSV_20                                    'unused reserved for future use
'21   HSV_PRESERVE_FORMULA_COMMENT
'22   HSV_22                                    'unused reserved for future use
'23   HSV_FORMULA_FILL
'30   HSV_EXCEL_FORMATTING = Boolean            'Selects the Excel formatting check box
'31   HSV_RETAIN_NUMERIC_FORMATTING = Boolean   'When the user drills down in dimensions, uses the scale specified in HSV_ SCALE and/or number of decimal places from HSV_DECIMALPLACES for data.
'32   HSV_THOUSAND_SEPARATOR = Boolean          'Uses a comma or other thousands separator in numerical data. Do not use # or $ as the thousands separator in Excel International Options
'33   HSV_NAVIGATE_WITHOUTDATA = Boolean        'Enables the speeding up of operations such as Pivot, Zoom, Keep Only, and Remove Only by preventing the calculation of source data while you are navigating. When you are ready to retrieve data, disable Navigate without Data.
'34   HSV_ENABLE_FORMATSTRING
'35   HSV_ENHANCED_COMMENT_HANDLING
'36   HSV_ADJUSTCOLUMNWIDTH = Boolean           'Adjust column widths to cell content automatically
'37   HSV_DECIMALPLACES
'38   HSV_SCALE
'39   HSV_MOVEFORMATS_ON_ADHOC
'40   HSV_DISPLAY_INVALIDDATA
'41   HSV_SUPPRESSCOLUMNS_MISSING
'42   HSV_SUPPRESSCOLUMNS_ZEROS
'43   HSV_SUPPRESSCOLUMNS_NOACCESS
'44   HSV_SUPPRESS_MISSINGBLOCKS
'101  HSV_DOUBLECLICK_FOR_ADHOC = Boolean       'Specifies that double-clicking retrieves the default grid in a blank worksheet and thereafter zooms in or out on the cell contents.
'102  HSV_UNDO_ENABLE = Boolean                 'Enables and disables Undo.
'103  HSV_103                                   'unused reserved for future use
'104  HSV_LOGMESSAGE_DISPLAY
'105  HSV_ROUTE_LOGMESSAGE_TO_FILE
'106  HSV_CLEAR_LOG_ON_NEXTLAUNCH
'107  HSV_REDUCE_EXCEL_FILESIZE = Boolean       'Should always be enabled except in the following cases, when it should not be used
'108  HSV_ENABLE_RIBBON_CONTEXT
'109  HSV_DISPLAY_HOMEPANEL_ONSTARTUP
'110  HSV_SHOW_COMMENTDIALOG_ON_REFRESH
'111  HSV_NUMBER_OF_UNDO_ACTION = Number        'The number of Undo and Redo actions permitted on an operation (0 through 100).
'112  HSV_NUMBER_OF_MRU_ITEMS
'113  HSV_ROUTE_LOGMESSAGE_FILE_LOCATION
'114  HSV_DISABLE_SMARTVIEW_IN_OUTLOOK
'115  HSV_DISPLAY_SMARTVIEW_SHORTCUT_MENU_ONLY
'116  HSV_DISPLAY_DRILL_THROUGH_REPORT_TOOLTIP
'117  HSV_SHOW_PROGRESSINFORMATION
'118  HSV_PROGRESSINFO_TIMEDELAY
End Sub
Sub vbaHypDeleteMetaData()
Dim Ret As Long
Dim Workbook As Workbook
Dim Sheet As Worksheet

Set Workbook = ActiveWorkbook
Set Sheet = ActiveSheet

Ret = HypDeleteMetaData(oSheet, False, True) 'Mode 1 Delete all Smart View metadata only from the provided worksheet storage
'Ret = HypDeleteMetaData(oWorkbook, True, False) 'Mode 2 Delete all Smart View metadata only from the provided workbook storage
'Ret = HypDeleteMetaData(oWorkbook, True, True) 'Mode 3 Delete all Smart View metadata from the provided workbook storage and from all the worksheets’ storage

If bpMessages Then MsgBox ("Smartview metadata deleted: " & Ret)

End Sub


