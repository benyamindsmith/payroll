'The following is a Function used for calculating expected pay of an employee
'given the number of hours worked and the employee's hourly rate

Function pay(hours As Date, rate As Double) As Currency


Dim numhrs As Double

numhrs = hours * 24

pay = numhrs * rate

End Function

