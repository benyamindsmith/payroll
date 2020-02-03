'Still in progress. This function does not work

Function Replicate(Rng As Range, Times As Integer, Location As Range) As Range

For i = 1 To Times

Rng.Copy Range(Location(i))

Next i

End Function
