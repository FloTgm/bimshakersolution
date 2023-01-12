Imports System.Globalization

Public Module CalendarUtils
    Public Function GetWeekOfyear(myDate As Date) As Integer
        Dim myCal = New System.Globalization.GregorianCalendar()
        Dim weekOfYear As Integer = myCal.GetWeekOfYear(myDate, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday)
        If (myDate.Month = 1 AndAlso weekOfYear >= 52) Then weekOfYear = 0
        If (myDate.Month = 12 AndAlso weekOfYear = 1) Then weekOfYear = 53
        Return (weekOfYear)
    End Function

    Public Function GetActualWeekOfYear() As Integer
        Dim actualDate = Now.Date
        Return GetWeekOfyear(actualDate)
    End Function

    Public Function GetWeekOfyearAsString(myDate As Date) As String
        Dim weekOfYear = GetWeekOfyear(myDate)
        Return (If(weekOfYear < 10, "0" + weekOfYear.ToString, weekOfYear.ToString))
    End Function

    Public Function GetActualWeekOfyearAsString() As String
        Dim weekOfYear = GetActualWeekOfYear()
        Return (If(weekOfYear < 10, "0" + weekOfYear.ToString, weekOfYear.ToString))
    End Function

End Module
