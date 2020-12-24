
Dim arr4()
For I = 0 To 3
   For J = 0 To 3
'        ReDim Preserve arr4(I, J)
        ReDim arr4(I, J)
        arr4(I, J) = I * J

   Next J
Next I

 
 
For I = LBound(arr4) To UBound(arr4)
    For J = LBound(arr4, 2) To UBound(arr4, 2)

    Next

Next
