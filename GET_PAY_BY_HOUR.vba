'The following function is used to calculate total wages of an employee who is paid by hour.
'This can be easily accomplished by the GET_PAY_BY_HOUR() function;
'The function takes an employee's times in and times out and the hourly rate that he is getting paid to calculate wages owed.

Function GET_PAY_BY_HOUR(time_in As Range, time_out As Range, rate As Currency) As Currency


GET_PAY_BY_HOUR = FormatCurrency((Application.WorksheetFunction.Sum(time_out) - Application.WorksheetFunction.Sum(time_in)) * rate * 24)

End Function


