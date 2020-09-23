Attribute VB_Name = "Module_Reminders"
Public Function GetFutureTime(sTime As Date, nValue, pValue As String) As String
'Example: "11:56:45" + "5 minutes" = "12:01:45"
'         "11:23:34" + "11 hours"  = "22:23:34"
'sTime - Start time
'nValue - Number of (minutes or hours to be added to this time)
'       - Max value for nValue when adding "Minutes" is 60
'       - Max value for nValue when adding "Hours" is 24
'pValue - This is either "Minutes" or "Hours"
Dim sMinute, sHour As Integer
    sMinute = Minute(sTime) 'Get minutes
    sHour = Hour(sTime) 'Get hours
    'Assume no increase in day will occur
    'An increase in day will occur in following example:
    '"23:45:00" + "20 minutes"
    DateUpdateNeeded = False 'Here we assume this will not happen
    Select Case pValue
        Case "Minutes" 'We are adding minutes
            sMinute = sMinute + nValue
            If sMinute > 59 Then 'Hour change occured
                sHour = sHour + 1 'Increase hour by one
                sMinute = sMinute - 60 'Fix minutes
                If sHour > 23 Then 'Day change occured
                    'An increase of a day by one occured
                    'We will set DateUpdateNeeded to True for that mater
                    DateUpdateNeeded = True
                    sHour = sHour - 24 'Fix hour
                End If
            End If
        Case "Hours" 'If we are adding hours
            sHour = sHour + nValue
            If sHour > 23 Then 'Day change occured
                'An increase of a day by one occured
                'We will set DateUpdateNeeded to True for that mater
                DateUpdateNeeded = True
                sHour = sHour - 24 'Fix hour
            End If
    End Select
    'Return new time
    GetFutureTime = TimeSerial(sHour, sMinute, 0)
End Function

Public Function GetFutureDate(sDate As Date, nValue, pValue As String) As String
'Example: "11/11/2001" + "1 day  " = "11/12/2001"
'         "11/11/2001" + "1 month" = "12/11/2001"
'sDate - Start date
'nValue - Number of (days or weeks or months or years to be added to this date)
'       - Max value for nValue when adding "Days" is 31
'       - Max value for nValue when adding "Weeks" is 4
'       - Max value for nValue when adding "Months" is 12
'       - Max value for nValue when adding "Years" is 5
'pValue - This is either "Days" or "Weeks" or "Months" or "Years"
Dim sDay, sMonth, sYear, cmDays, nmDays As Integer
    sDay = Day(sDate) 'Get day
    sMonth = Month(sDate) 'Get month
    sYear = Year(sDate) 'Get year
    cmDays = DaysInMonth(sMonth, sYear) 'Number of days in a current month
    If sMonth < 12 Then nmDays = DaysInMonth(sMonth + 1, sYear) Else nmDays = DaysInMonth(1, sYear) 'Number of days in a next consecutive month
    If pValue = "Weeks" Then 'If we are adding weeks we are essentialy adding # of Day * 7
        'Here we assume that nValue has a maximum of 4 when adding weeks and therefore 4 * 7 = 28 which is less that the assumed maximum for days which is 31
        pValue = "Days"
        nValue = nValue * 7
    End If
    Select Case pValue
        Case "Days" 'Adding days (or weeks)
            sDay = sDay + nValue
            If sDay > cmDays + nmDays Then 'New day is 2 months from now
                sDay = sDay - cmDays - nmDays 'Fix day
                sMonth = sMonth + 2 'Increase months
                If sMonth > 12 Then 'Year change occured
                    sMonth = sMonth - 12 'Fix month
                    sYear = sYear + 1 'Add a year
                End If
            ElseIf sDay > cmDays Then 'New day is in 1 month from now
                sDay = sDay - cmDays 'Fix days
                sMonth = sMonth + 1 'Add a month
                If sMonth > 12 Then 'Year change occured
                    sMonth = sMonth - 12 'Fix month
                    sYear = sYear + 1 'Add a year
                End If
            End If
        Case "Months" 'Adding months
            sMonth = sMonth + nValue
            If sMonth > 12 Then 'Year change occured
                sMonth = sMonth - 12 'Fix months
                sYear = sYear + 1 'Add a year
            End If
        Case "Years" 'Adding years
            sYear = sYear + nValue
    End Select
    'Return new date
    GetFutureDate = DateSerial(sYear, sMonth, sDay)
End Function

Function DaysInMonth(iMonth, iYear) As Integer
    'Determines the number of days in a month
    Select Case (iMonth)
        Case 2 'February
            If (iYear Mod 4 = 0) And (iYear Mod 100 <> 0) Or (iYear Mod 400 = 0) Then DaysInMonth = 29 Else DaysInMonth = 28
        Case 4, 6, 9, 11 'April, June, September, November
            DaysInMonth = 30
        Case Else 'January, March, May, July, August, October, December
            DaysInMonth = 31
    End Select
End Function
