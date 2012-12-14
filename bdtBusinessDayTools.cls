VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "bdtBusinessDayTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'****************************************************************************************
'* Name                 :   bdtBusinessDayTools Object Class Module                     *
'* Author               :   Joseph Solomon, CPA                                         *
'* Purpose              :   Class object holding the methods and properties required to *
'*                          work with Metlife specific Business Dates.                  *
'* Main Call            :   AdjustBusinessDate                                          *
'* Rules                :   1)  Holiday's must be defined in the IsHoliday Function and *
'*                              the array has to be adjusted for every holiday added or *
'*                              subtracted                                              *
'*                                                                                      *
'*  Usage example                                                                       *
'*  ~~~~~~~~~~~~~~~~~~                                                                  *
'*  Dim uobjHolidayScheule As bdtBusinessDayTools                                       *
'*  Dim dtmBizDay As Date                                                               *
'*  Set uobjHolidayScheule = New bdtBusinessDayTools                                    *
'*  dtmBizDay = uobjHolidayScheule.AdjustBusinessDate(CDate("12/30/2012"), bdtNext, 1)  *
'*                                                                                      *
'* Revision History-                                                                    *
'* Rev#     Revision Date   Description                                                 *
'* ~~~~     ~~~~~~~~~~~~~   ~~~~~~~~~~~                                                 *
'*  1       2012/12/06      Initial Release                                             *
'*  2       2012/12/07      Implemented "Global Reference Tables" database for a central*
'*                          location to hold the holidays.  Populated the Holiday       *
'*                          database table through 12/31/2015.  Implemented ADO         *
'*                          connection string for the "Global Reference Tables" database*
'*  3       2012/12/10      Added dtmPriorBusinessDay public method.                    *
'*  4       2012/12/13      Replaced static date table with dynamic holiday date        *
'*                          determination functions. Added blnIsHoliday function, and   *
'*                          the related date determination functions.                   *
'*  5       2012/12/14      Added error handling and finished documenting processes.    *
'****************************************************************************************

Option Compare Database
Option Explicit

Enum bdtMonth
    bdt01January = 1
    bdt02February
    bdt03March
    bdt04April
    bdt05May
    bdt06June
    bdt07July
    bdt08August
    bdt09September
    bdt10October
    bdt11November
    bdt12December
End Enum

Enum bdtDayofWeek
    bdt1Sunday = 1
    bdt2Monday
    bdt3Tuesday
    bdt4Wednesday
    bdt5Thursday
    bdt6Friday
    bdt7Saturday
End Enum

Enum bdtDirection
    bdtPrevious = -1
    bdtnext = 1
End Enum

'****************************************************************************************
'*  Methods-                                                                            *
'*  Name                    Last Revised    Description                                 *
'*  ~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~~~    ~~~~~~~~~~~~~~~~~~~~~~~~~                   *
'*  AdjustBusinessDate      2012/12/13      Function is the main call for the class     *
'*                                              object. Returns adjusted business day   *
'*  blnIsBusinessDay        2012/12/13      Tests whether the Date parameter passed is  *
'*                                              considered a business day.              *
'*  blnIsHoliday            2012/12/13      Runs a series of tests to determine if      *
'*                                              target date falls on calculated date.   *
'*  blnIsLastWkdayMonth     2012/12/13      Test date to determine if it falls on the   *
'*                                              last weekday specified as a parameter,  *
'*                                              of a month specified as a parameter.    *
'*  blnIsNumWkdayInMonth    2012/12/13      Test date to determine if it falls on a     *
'*                                              certain week of a certain month both    *
'*                                              specified as parameters on a weekday    *
'*                                              also specified  as a parameter.         *
'****************************************************************************************

Public Function AdjustBusinessDate( _
    ByVal pvTargetDate As Date, _
    ByVal pvDirection As bdtDirection, _
    Optional ByVal pvNumDays As Long = 1) As Date
    'pvTargetDate                       Original date being converted
    'pvDirection                        Move business date forward or backward
    'Optional pvNumDays                 Number of business days to add or subtract
    Dim dtmNewDate As Date              'Adjusted Business Date
    Dim lngBsnDayCount As Long          'Business day counter
        
    On Error GoTo AdjustBusinessDate_Err
        
    'Raise error if number of days to adjust by is set to zero
    If pvNumDays = 0 Then Err.Raise 22000
    
    dtmNewDate = pvTargetDate
    
    For lngBsnDayCount = 1 To pvNumDays
        
        'Adjust by 1 day
        dtmNewDate = DateAdd( _
            Interval:="d", _
            Number:=pvDirection, _
            Date:=dtmNewDate)
        
        'Adjust Date until business day is found.
        Do Until blnIsBusinessDay(dtmNewDate)
            dtmNewDate = DateAdd( _
                Interval:="d", _
                Number:=pvDirection, _
                Date:=dtmNewDate)
        Loop
    Next lngBsnDayCount
    AdjustBusinessDate = dtmNewDate
    
AdjustBusinessDate_Exit:
    Exit Function
    
AdjustBusinessDate_Err:
    Dim strErrorMessage As String
    Select Case Err.Number
        Case Is = 22000
            strErrorMessage = "Using 0 as the number of days to adjust will result " & _
                "in unexpected results."
        Case Else
            strErrorMessage = ""
    End Select
    Select Case ErrHandler(Err.Number, Err.Description, strErrorMessage, _
        "AdjustBusinessDate", "bdtBusinessDayTools")
        Case Is = vbIgnore: Resume Next
        Case Is = vbRetry: Resume
        Case Is = vbAbort: Exit Function
    End Select
    GoTo AdjustBusinessDate_Exit
End Function

Public Function blnIsBusinessDay(ByVal pvdtmTestDate As Date) As Boolean
    'pvdtmTestDate                      Date being tested
    On Error GoTo blnIsBusinessDay_Err
    
    'Assumed true until a non-business day test is met.
    blnIsBusinessDay = True
    
    'Test for Saturday
    If Weekday(pvdtmTestDate) = bdt7Saturday Then
        blnIsBusinessDay = False
        GoTo blnIsBusinessDay_Exit
    End If
    
    'Test for Sunday
    If Weekday(pvdtmTestDate) = bdt1Sunday Then
        blnIsBusinessDay = False
        GoTo blnIsBusinessDay_Exit
    End If
    
    'Test for holidays
    If blnIsHoliday(pvdtmTestDate) Then
        blnIsBusinessDay = False
        GoTo blnIsBusinessDay_Exit
    End If
    
blnIsBusinessDay_Exit:
    Exit Function
    
blnIsBusinessDay_Err:
    Select Case ErrHandler(Err.Number, Err.Description, , "blnIsBusinessDay", _
        "bdtBusinessDayTools")
        Case Is = vbIgnore: Resume Next
        Case Is = vbRetry: Resume
        Case Is = vbAbort: Exit Function
    End Select
    GoTo blnIsBusinessDay_Exit
End Function

Private Function blnIsHoliday(ByVal pvdtmTargetDate As Date) As Boolean
    'pvdtmTargetDate                    'Date being tested
    Dim blnaHoliday(1 To 10) As Boolean 'Results of a particular date test
    Dim lngX As Long                    'Counter to walk date test results
    
    'Note:
    'For each holiday test added or subtracted, increase or decrease the upper bound in
        'the blnaHoliday declaration at the beginning of this function.
    
    On Error GoTo blnIsHoliday_Err
    
    'Assume False unless a date test returns true
    blnIsHoliday = False
    
    'Test dates against holiday definitions
    '*New Years; Month=1, Day=1
    blnaHoliday(1) = (Month(pvdtmTargetDate) = 1 And Day(pvdtmTargetDate) = 1)
    '*MLK Jr. Day; 3rd Monday in January
    blnaHoliday(2) = blnIsNumWkdayInMonth(pvdtmTargetDate, 3, bdt2Monday, bdt01January)
    '*Presidents Day; 3rd Monday in February
    blnaHoliday(3) = blnIsNumWkdayInMonth(pvdtmTargetDate, 3, bdt2Monday, bdt02February)
    '*Memorial Day; last Monday in May
    blnaHoliday(4) = blnIsLastWkdayMonth(pvdtmTargetDate, bdt2Monday, bdt05May)
    '*Independence Day; Month=7, Day=4
    blnaHoliday(5) = (Month(pvdtmTargetDate) = 7 And Day(pvdtmTargetDate) = 4)
    '*Labor Day; 1st Monday in September
    blnaHoliday(6) = blnIsNumWkdayInMonth(pvdtmTargetDate, 1, bdt2Monday, bdt09September)
    '*Thanksgiving; 4th Thursday in November
    blnaHoliday(7) = blnIsNumWkdayInMonth(pvdtmTargetDate, 4, bdt5Thursday, bdt11November)
    '*Black Friday; 4th Friday in November
    blnaHoliday(8) = blnIsNumWkdayInMonth(pvdtmTargetDate, 4, bdt6Friday, bdt11November)
    '*Christmas Day; Month=12, Day=25
    blnaHoliday(9) = (Month(pvdtmTargetDate) = 12 And Day(pvdtmTargetDate) = 25)
    '*New Years Eve; ; Month=12, Day=31
    blnaHoliday(10) = (Month(pvdtmTargetDate) = 12 And Day(pvdtmTargetDate) = 31)
    
    'Return function true if any of the above returned true
    For lngX = LBound(blnaHoliday) To UBound(blnaHoliday)
        If blnaHoliday(lngX) = True Then
            blnIsHoliday = True
            Exit Function
        End If
    Next lngX
    
blnIsHoliday_Exit:
    Exit Function
    
blnIsHoliday_Err:
    Dim strErrorMesage As String
    
    Select Case Err.Number
        Case Is = 9
            strErrorMesage = "You may have tried to add a holiday without changing " & _
                "the size of the array.  Try adjusting the 'Dim blnaHoliday(1 To " & _
                UBound(blnaHoliday) & _
                ") As Boolean' statement at the beginning of the blnIsHoliday function."
        Case Else
            strErrorMesage = ""
    End Select
    
    Select Case ErrHandler(Err.Number, Err.Description, strErrorMesage, "blnIsHoliday", _
        "bdtBusinessDayTools")
        Case Is = vbIgnore: Resume Next
        Case Is = vbAbort: Exit Function
        Case Is = vbRetry: Resume
    End Select
    
End Function

Private Function blnIsLastWkdayMonth( _
    ByVal pvTargetDate As Date, _
    ByVal pvDayofWeek As bdtDayofWeek, _
    ByVal pvMonth As bdtMonth) As Boolean
    'pvTargetDate                           'Date being tested.
    'pvDayofWeek                            'The day of week the date being tested would
                                                'fall on for the function to return true
    'pvMonth                                'The month the date being tested would fall in
                                                ' for the function to return true.
    Dim dtmBOMonth As Date                  'First day of month being tested
    Dim dtmEOMonth As Date                  'Last day of month being tested
    Dim lngWeekdaysInMonth As Long          'Number of times the passed weekday
                                                '(pvDayofWeek) falls in the month.
    
    On Error GoTo blnIsLastWkdayMonth_Err
    
    'Assume false until all conditions are met.
    blnIsLastWkdayMonth = False
    
    'Skip if date being tested falls outside of the month.
    If Month(pvTargetDate) = pvMonth Then
        
        'Find beginning and end dates for the passed date to be tested.
        dtmBOMonth = DateSerial(Year(pvTargetDate), Month(pvTargetDate), 1)
        dtmEOMonth = DateSerial(Year(pvTargetDate), Month(pvTargetDate) + 1, 0)
        
        'Calculate the amount of passed weekdays in the month of passed date being tested.
        lngWeekdaysInMonth = DateDiff("ww", dtmBOMonth, dtmEOMonth, vbMonday, vbUseSystem)
        
        If DateDiff("ww", dtmBOMonth, pvTargetDate, vbMonday, vbUseSystem) = _
            lngWeekdaysInMonth And Weekday(pvTargetDate, vbMonday) = 1 Then
            'The difference in weeks between the first day of the month and the date being
                'tested equals the amount of weerkdays in the month, and the date being
                'tested also falls on the passed weekday.
            blnIsLastWkdayMonth = True
        End If
    End If
    
blnIsLastWkdayMonth_Exit:
    Exit Function
    
blnIsLastWkdayMonth_Err:
    Select Case ErrHandler(Err.Number, Err.Description, , "blnIsLastWkdayMonth", _
        "bdtBusinessDayTools")
        Case Is = vbIgnore: Resume Next
        Case Is = vbAbort: Exit Function
        Case Is = vbRetry: Resume
    End Select
End Function

Private Function blnIsNumWkdayInMonth( _
    ByVal pvTargetDate As Date, _
    ByVal pvWeekNum As Long, _
    ByVal pvDayofWeek As bdtDayofWeek, _
    ByVal pvMonth As bdtMonth) As Boolean
    'pvTargetDate                           'Date being tested.
    'pvWeekNum                              'The iteration of the weekday within the
                                                'month the date falls on.
    'pvDayofWeek                            'The day of week the date being tested would
                                                'fall on for the function to return true
    'pvMonth                                'The month the date being tested would fall in
                                                ' for the function to return true.
    Dim dtmBOMonth As Date                  'First day of the month of the date being
                                                'tested.
                                                
    On Error GoTo blnIsNumWkdayInMonth_Err
    
    'Assume false until all conditions are met.
    blnIsNumWkdayInMonth = False
    
    'Determine beginning of month of the date being tested.
    dtmBOMonth = DateSerial(Year(pvTargetDate), pvMonth, 1)
    
    'Fourth Friday in November
    If (DateDiff("ww", dtmBOMonth, pvTargetDate, pvDayofWeek, vbUseSystem) = pvWeekNum) _
        And (Weekday(pvTargetDate, vbSunday) = pvDayofWeek) Then
        'The difference in iterations of the weekday passed equals the iteration of the
            'weekday within the month the date falls on and the date also falls on the
            'passed weekday.
        blnIsNumWkdayInMonth = True
    End If
    
blnIsNumWkdayInMonth_Exit:
    Exit Function
    
blnIsNumWkdayInMonth_Err:
    Select Case ErrHandler(Err.Number, Err.Description, , "blnIsNumWkdayInMonth", _
        "bdtBusinessDayTools")
        Case Is = vbIgnore: Resume Next
        Case Is = vbAbort: Exit Function
        Case Is = vbRetry: Resume
    End Select
End Function

Private Function ErrHandler( _
ByVal pvlngErrNum As Long, _
ByVal pvstrErrDesc As String, _
Optional ByVal pvstrErrCustom As String, _
Optional ByVal pvstrErrProcedure As String, _
Optional ByVal pvstrErrModule As String) As Long
Dim strErrorMessage As String
On Error GoTo ErrHandler_Err
'****************************************************************************************
'*  Name            :       ErrHandler                                                  *
'*  Author          :       Joseph Solomon, CPA                                         *
'*  Purpose         :       This procedure is a general use Error Handling function. The*
'*                              function will log the Error to tblError.  If tblError   *
'*                              does not exist, this function will create it.           *
'*  Return Value    :       ErrHandler will return a value derived from the VBA MsgBox  *
'*                              Function's vbAbortRetryIgnore enumeration based on the  *
'*                              User's Selection.  If this function Errors out, it will *
'*                              not log the Error, but will prompt the user to Abort,   *
'*                              Retry, or Ignore.                                       *
'*  Last Update     :       2012/12/13                                                  *
'*                                                                                      *
'*  Parameters/Variables:   Description:                                                *
'*  ~~~~~~~~~~~~~~~~~~~~~   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~               *
'*  pvlngErrNum             Error number passed into the Error handler.                 *
'*  pvstrErrDesc            VBA's Error Description                                     *
'*  pvstrErrCustom          Optional custom Error Description                           *
'*  pvstrErrProcedure       Optional name of procedure calling the Error.               *
'*  pvstrErrModule          Optional name of module containing the procedure.           *
'*  strErrorMessage         Prompt that shows in the Error Message Box.                 *
'*                                                                                      *
'*  Usage Example:                                                                      *
'*  ~~~~~~~~~~~~~~~                                                                     *
'*  On Error GoTo MyProcedure_Err                                                       *
'*  <...code...>                                                                        *
'*  MyProcedure_Err:                                                                    *
'*                                                                                      *
'*  Dim strErrorMessage as String                                                       *
'*  Select Case Err.Number                                                              *
'*      Case is = 11 'Division by zero                                                  *
'*          strErrorMessage = "The Application tried to divide by zero.                 *
'*      Case Is = 75, 76 'Path/File Access Error                                        *
'*          strErrorMessage = "There was a problem accessing the network.  Check " & _  *
'*              "your connection and try again."                                        *
'*  End Select                                                                          *
'*                                                                                      *
'*  Select Case ErrHandler(Err.Number, Err.Description, strErrorMessage, "DataImport", _*
'*      "modImport")                                                                    *
'*      Case Is = vbIgnore: Resume Next                                                 *
'*      Case Is = vbAbort: Exit Sub                                                     *
'*      Case Is = vbRetry: Resume                                                       *
'*  End Select                                                                          *
'****************************************************************************************
    
    'Create Error Message
    If Not IsEmpty(pvstrErrCustom) Then 'Optional Error Message parameter was passed
         strErrorMessage = _
            pvstrErrCustom & vbCrLf & vbCrLf & _
            "Error # : " & pvlngErrNum & vbCrLf & _
            "Error Description : " & pvstrErrDesc & vbCrLf & vbCrLf & _
            "Please make a selection:"
    Else 'No Optional Error Message Passed.  Use default.
        strErrorMessage = _
            "Error # : " & pvlngErrNum & vbCrLf & _
            "Error Description : " & pvstrErrDesc & vbCrLf & vbCrLf & _
            "Please make a selection:"
    End If

    'Create Error table if it does not exist
    DoCmd.SetWarnings False
    If DCount("[Name]", "MSysObjects", "Type = 1 AND [Name] = 'tblErrorLog'") < 1 Then
        DoCmd.RunSQL _
            "CREATE TABLE tblErrorLog " & _
            "(ErrorID AUTOINCREMENT PRIMARY KEY, ErrorNumber INTEGER, " & _
            "ErrorDescription LONGTEXT, ProcedureName CHAR, ModuleName CHAR, " & _
            "ErrorTimeStamp DATETIME, User CHAR, ComputerName CHAR);"
    End If
    
    'Log the Error
    DoCmd.RunSQL _
        "INSERT INTO tblErrorLog " & _
        "(ErrorNumber, ErrorDescription, ProcedureName, ModuleName, " & _
        "ErrorTimeStamp, User, ComputerName) " & _
        "VALUES (" & _
        pvlngErrNum & ", '" & pvstrErrDesc & "', '" & pvstrErrProcedure & "', '" & _
        pvstrErrModule & "', #" & Now & "#, '" & Environ("USERNAME") & "', '" & _
        Environ("COMPUTERNAME") & "');"
    DoCmd.SetWarnings True
    
    ErrHandler = MsgBox( _
        Prompt:=strErrorMessage, _
        Buttons:=vbAbortRetryIgnore Or vbExclamation, _
        Title:="Error " & pvlngErrNum)
        
ErrHandler_Exit:
    Exit Function
    
ErrHandler_Err:
    On Error Resume Next
    Select Case MsgBox( _
            Prompt:="There was an Error handling the Error...Ironic, huh?" & vbCrLf & _
                "This Error will not be logged." & _
                vbCrLf & vbCrLf & _
                "Error # : " & pvlngErrNum & vbCrLf & _
                "Error Description : " & pvstrErrDesc & vbCrLf & vbCrLf & _
                "Please make a selection:", _
            Buttons:=vbAbortRetryIgnore Or vbExclamation, _
            Title:="Error Handling Error")
        Case Is = vbIgnore: Resume Next
        Case Is = vbRetry: Resume
        Case Is = vbAbort: Exit Function
    End Select
    GoTo ErrHandler_Exit

End Function

