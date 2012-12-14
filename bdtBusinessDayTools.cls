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
'*                          determination functions. Added blnHolidayTest function, and *
'*                          the related date determination functions.                   *
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
    bdtNext = 1
End Enum

'****************************************************************************************
'*  Methods-                                                                            *
'*  Name                    Last Revised    Description                                 *
'*  ~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~~~    ~~~~~~~~~~~~~~~~~~~~~~~~~                   *
'*  AdjustBusinessDate      2012/12/13      Function is the main call for the class     *
'*                                              object.                                 *
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
    
    dtmNewDate = pvTargetDate
    For lngBsnDayCount = 1 To pvNumDays
        'Adjust by 1 day
        dtmNewDate = DateAdd( _
            Interval:="d", _
            Number:=pvDirection * lngBsnDayCount, _
            Date:=dtmNewDate)
        
        'Adjust Date until business day is found.
        Do Until blnIsBusinessDay(dtmNewDate)
            dtmNewDate = DateAdd( _
            Interval:="d", _
            Number:=pvDirection * lngBsnDayCount, _
            Date:=dtmNewDate)
        Loop
    Next lngBsnDayCount
    AdjustBusinessDate = dtmNewDate
    
AdjustBusinessDate_Exit:
    Exit Function
    
AdjustBusinessDate_Err:
    Select Case ErrHandler(Err.Number, Err.Description, , "AdjustBusinessDate", _
        "bdtBusinessDayTools")
        Case Is = vbIgnore: Resume Next
        Case Is = vbRetry: Resume
        Case Is = vbAbort: Exit Function
    End Select
    GoTo AdjustBusinessDate_Exit
End Function


Private Function blnIsBusinessDay(ByVal pvdtmTestDate As Date) As Boolean
    'pvdtmTestDate       Date being tested
    On Error GoTo blnIsBusinessDay_Err
    
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
    
    'Passed all tests, considered a business day
    blnIsBusinessDay = True
    
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

Private Function blnIsHoliday(dtmTargetDate As Date) As Boolean
    Dim varHoliday(1 To 10) As Variant
    Dim lngX As Long
    blnIsHoliday = False
    
    'Test dates against holiday definitions
    '*New Years; Month=1, Day=1
    varHoliday(1) = (Month(dtmTargetDate) = 1 And Day(dtmTargetDate) = 1)
    '*MLK Jr. Day; 3rd Monday in January
    varHoliday(2) = blnIsNumWkdayInMonth(dtmTargetDate, 3, bdt2Monday, bdt01January)
    '*Presidents Day; 3rd Monday in February
    varHoliday(3) = blnIsNumWkdayInMonth(dtmTargetDate, 3, bdt2Monday, bdt02February)
    '*Memorial Day; last Monday in May
    varHoliday(4) = blnIsLastWkdayMonth(dtmTargetDate, bdt2Monday, bdt05May)
    '*Independence Day; Month=7, Day=4
    varHoliday(5) = (Month(dtmTargetDate) = 7 And Day(dtmTargetDate) = 4)
    '*Labor Day; 1st Monday in September
    varHoliday(6) = blnIsNumWkdayInMonth(dtmTargetDate, 1, bdt2Monday, bdt09September)
    '*Thanksgiving; 4th Thursday in November
    varHoliday(7) = blnIsNumWkdayInMonth(dtmTargetDate, 4, bdt5Thursday, bdt11November)
    '*Black Friday; 4th Friday in November
    varHoliday(8) = blnIsNumWkdayInMonth(dtmTargetDate, 4, bdt6Friday, bdt11November)
    '*Christmas Day; Month=12, Day=25
    varHoliday(9) = (Month(dtmTargetDate) = 12 And Day(dtmTargetDate) = 25)
    '*New Years Eve; ; Month=12, Day=31
    varHoliday(10) = (Month(dtmTargetDate) = 12 And Day(dtmTargetDate) = 31)
    
    'See if any returned true
    For lngX = LBound(varHoliday) To UBound(varHoliday)
        If varHoliday(lngX) = True Then
            blnIsHoliday = True
            Exit Function
        End If
    Next lngX
    End Function

Private Function blnIsNumWkdayInMonth( _
    ByVal pvTargetDate As Date, _
    ByVal pvWeekNum As Long, _
    ByVal pvDayofWeek As bdtDayofWeek, _
    ByVal pvMonth As bdtMonth) As Boolean
    Dim dtmBOMonth As Date
    
    dtmBOMonth = DateSerial(Year(pvTargetDate), pvMonth, 1)
    
    'Fourth Friday in November
    If (DateDiff("ww", dtmBOMonth, pvTargetDate, pvDayofWeek, vbUseSystem) = 4) And _
        (Weekday(pvTargetDate, pvDayofWeek) = pvWeekNum) Then
        blnIsNumWkdayInMonth = True
    Else
        blnIsNumWkdayInMonth = False
    End If
End Function

Private Function blnIsLastWkdayMonth( _
    ByVal pvTargetDate As Date, _
    ByVal pvDayofWeek As bdtDayofWeek, _
    ByVal pvMonth As bdtMonth) As Boolean
    Dim dtmBOMonth As Date
    Dim dtmEOMonth As Date
    Dim lngWeekdaysInMonth   As Long
    
    If Month(pvTargetDate) = pvMonth Then
        dtmBOMonth = DateSerial(Year(pvTargetDate), Month(pvTargetDate), 1)
        dtmEOMonth = DateSerial(Year(pvTargetDate), Month(pvTargetDate) + 1, 0)
        lngWeekdaysInMonth = DateDiff("ww", dtmBOMonth, dtmEOMonth, vbMonday, vbUseSystem)
        If DateDiff("ww", dtmBOMonth, pvTargetDate, vbMonday, vbUseSystem) = _
            lngWeekdaysInMonth And Weekday(pvTargetDate, vbMonday) = 1 Then
            blnIsLastWkdayMonth = True
        Else
            blnIsLastWkdayMonth = False
        End If
    End If
End Function
