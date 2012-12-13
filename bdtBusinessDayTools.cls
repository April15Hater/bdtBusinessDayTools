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
'* Main Call            :   DataImport                                                  *
'* Rules                :   1.  Database must reference Microsoft ActiveX Data Objects  *
'*                              2.8 Library (msado15.dll), and Microsoft Scripting      *
'*                              Runtime (scrrun.dll)                                    *
'*                          2.  User must have read-access to the "\\NS-TAMPVFS01\      *
'*                              Shared\Securities Financial Management\Derivatives\     *
'*                              Global Reference\" network location                     *
'*                                                                                      *
'*  Usage example                                                                       *
'*  ~~~~~~~~~~~~~~~~~~                                                                  *
'*  Dim uobjHolidayScheule As bdtBusinessDayTools                                       *
'*  Dim dtmBizDay As Date                                                               *
'*  Set uobjHolidayScheule = New bdtBusinessDayTools                                    *
'*  dtmBizDay = uobjHolidayScheule.pdtmNextBusinessDay(CDate("12/30/2012"))             *
'*                                                                                      *
'* Revision History-                                                                    *
'* Rev#     Revision Date   Description                                                 *
'* ~~~~     ~~~~~~~~~~~~~   ~~~~~~~~~~~                                                 *
'*  1       2012/12/06      Initial Release                                             *
'*  2       2012/12/07      Implemented "Global Reference Tables" database for a central*
'*                          location to hold the holidays.  Populated the Holiday       *
'*                          database table through 12/31/2015.  Implemented ADO         *
'*                          connection string for the "Global Reference Tables" database*
'*  3       2012/12/10      Added pdtmPriorBusinessDay public method.                   *
'****************************************************************************************

Option Compare Database
Option Explicit
Private mobjHolidaySchedule As Dictionary       'Object holding the holiday dates

'****************************************************************************************
'*  Properties-                                                                         *
'*  Name                    Last Revised    Description                                 *
'*  ~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~~~    ~~~~~~~~~~~~~~~~~~~~~~~~~                   *
'*  gobjHolidaySchedule     2012/12/06      Dictionary object holding holiday dates.    *
'*                                          The BuildHolidaySchedule Method populates   *
'*                                          the dictionary.                             *
'****************************************************************************************
Public Property Get gobjHolidaySchedule() As Dictionary
    If mobjHolidaySchedule Is Nothing Then BuildHolidaySchedule
    Set gobjHolidaySchedule = mobjHolidaySchedule
End Property

'****************************************************************************************
'*  Methods-                                                                            *
'*  Name                    Last Revised    Description                                 *
'*  ~~~~~~~~~~~~~~~~~~~~~~~ ~~~~~~~~~~~~    ~~~~~~~~~~~~~~~~~~~~~~~~~                   *
'*  pdtmNextBusinessDay     2012/12/06      Function returns the business day following *
'*                                          the passed Date parameter as a Date data    *
'*                                          type.                                       *
'*  pdtmPriorBusinessDay    2012/12/10      Function returns the business day prior to  *
'*                                          the passed Date parameter as a Date data    *
'*                                          type.                                       *
'*  BuildHolidaySchedule    2012/12/07      Stores the holiday dates into a dictionary  *
'*                                          object.                                     *
'*  blnIsBusinessDay        2012/12/06      Tests whether the Date parameter passed is  *
'*                                          considered a business day.                  *
'****************************************************************************************
Public Function pdtmNextBusinessDay(ByVal pvdtmDateValue As Date) As Date
    On Error GoTo pdtmNextBusinessDay_Err
    
    'Assign following day
    pvdtmDateValue = DateAdd(Interval:="d", Number:=1, Date:=pvdtmDateValue)
    
    'Test all following days until business day is found.
    Do Until blnIsBusinessDay(pvdtmDateValue)
        pvdtmDateValue = DateAdd(Interval:="d", Number:=1, Date:=pvdtmDateValue)
    Loop
    pdtmNextBusinessDay = pvdtmDateValue
    
pdtmNextBusinessDay_Exit:
    Exit Function
    
pdtmNextBusinessDay_Err:
    Select Case ErrHandler(Err.Number, Err.Description, "BuildHolidaySchedule")
        Case Is = vbIgnore: Resume Next
        Case Is = vbRetry: Resume
        Case Is = vbAbort: Exit Function
    End Select
    GoTo pdtmNextBusinessDay_Exit
End Function

Public Function pdtmPriorBusinessDay(ByVal pvdtmDateValue As Date) As Date
    On Error GoTo pdtmPriorBusinessDay_Err
    
    'Assign Prior day
    pvdtmDateValue = DateAdd(Interval:="d", Number:=-1, Date:=pvdtmDateValue)
    
    'Test all prior days until business day is found.
    Do Until blnIsBusinessDay(pvdtmDateValue)
        pvdtmDateValue = DateAdd(Interval:="d", Number:=-1, Date:=pvdtmDateValue)
    Loop
    pdtmPriorBusinessDay = pvdtmDateValue
    
pdtmPriorBusinessDay_Exit:
    Exit Function
    
pdtmPriorBusinessDay_Err:
    Select Case ErrHandler(Err.Number, Err.Description, "BuildHolidaySchedule")
        Case Is = vbIgnore: Resume
        Case Is = vbRetry: Resume
        Case Is = vbAbort: Exit Function
    End Select
    GoTo pdtmPriorBusinessDay_Exit
End Function

Private Sub BuildHolidaySchedule()
    On Error GoTo BuildHolidaySchedule_Err
    Dim cn As ADODB.Connection
    Dim strConnectionString As String
    Dim rs As ADODB.Recordset
    
    Set cn = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    Set mobjHolidaySchedule = New Dictionary
    mobjHolidaySchedule.CompareMode = BinaryCompare
    
    'Open recordset, populate dictionary object with all holiday dates.
    cn.Open _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & _
        "\\NS-TAMPVFS01\Shared\Securities Financial Management\Derivatives" & _
        "\Global Reference\Global Reference Tables.mdb; " & _
        "Jet OLEDB:System Database=" & _
        "\\NS-TAMPVFS01\Shared\Securities Financial Management\Derivatives" & _
        "\Global Reference\Security\ShareWorkgroup.mdw;"
    With rs
        .ActiveConnection = cn
        .Source = "tblHolidaySchedule"
        .CursorLocation = adUseServer
        .CursorType = adOpenForwardOnly
        .LockType = adLockReadOnly
        .Open
        Do Until .EOF
            mobjHolidaySchedule.Add CLng(!HolidayDate), !HolidayDate
            .MoveNext
        Loop
    End With
    
BuildHolidaySchedule_Exit:
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
    Exit Sub
    
BuildHolidaySchedule_Err:
    Select Case ErrHandler(Err.Number, Err.Description, "BuildHolidaySchedule")
        Case Is = vbIgnore: Resume Next
        Case Is = vbRetry: Resume
        Case Is = vbAbort: Exit Sub
    End Select
    GoTo BuildHolidaySchedule_Exit
End Sub

Private Function blnIsBusinessDay(ByVal pvdtmTestDate As Date) As Boolean
    On Error GoTo blnIsBusinessDay_Err
    'pvdtmTestDate       Date being tested

    'Test for weekend
    If Weekday(pvdtmTestDate) = vbSunday Or _
        Weekday(pvdtmTestDate) = vbSaturday Then
        blnIsBusinessDay = False
        GoTo blnIsBusinessDay_Exit
    End If
    
    'Get the holiday schedule if it does not yet exist.
    If mobjHolidaySchedule Is Nothing Then
        BuildHolidaySchedule
    End If
    
    'Test if the parameter date falls on a holiday.
    If gobjHolidaySchedule.Exists(pvdtmTestDate) Then
        blnIsBusinessDay = False
        GoTo blnIsBusinessDay_Exit
    End If
    
    'Passed all tests, considered a business day
    blnIsBusinessDay = True
    
blnIsBusinessDay_Exit:
    Exit Function
    
blnIsBusinessDay_Err:
    Select Case ErrHandler(Err.Number, Err.Description, "blnIsBusinessDay")
        Case Is = vbIgnore: Resume Next
        Case Is = vbRetry: Resume
        Case Is = vbAbort: Exit Function
    End Select
    GoTo blnIsBusinessDay_Exit
End Function

Private Function ErrHandler(ErrNum As Long, ErrDesc As String, Optional Proc As String) As Long
    Dim strErrorMessage As String
    Select Case ErrNum
        Case Is = 3705 'Recordset already open
            strErrorMessage = "Database connection already open.  If the problem " & _
            "continues, restart the database."
        Case Else 'Unknown Error

    End Select
    ErrHandler = MsgBox( _
        prompt:=strErrorMessage & vbCrLf & vbCrLf & _
            "Error #: " & ErrNum & vbCrLf & _
            "ErrDesc : " & ErrDesc & vbCrLf & vbCrLf & _
            "Please make a selection:", _
        Title:="Error " & ErrNum, _
            Buttons:=vbAbortRetryIgnore)
    Exit Function

End Function

