Function payroll(time_in As Range, time_out As Range, rate As Double) As Currency

Dim cell_1 As Date
Dim cell_2 As Date


For Each cell_1 In time_in
Next cell_1

For Each cell_2 In time_out
Next cell_2

hrs As Range

hrs = cell_2 - cell_1

payroll = hrs * rate

End Function

