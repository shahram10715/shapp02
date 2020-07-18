'Excel vb macro

Private Sub CommandButton1_Click()


'this part calculates number of load cases

Dim nlc As Integer, i As Integer, j As Integer, lcase As String
i = 1
nlc = 1
lcase = Sheet2.Cells(4, 3).Value

Do While Sheet2.Cells(i + 4, 3) <> ""
    If Sheet2.Cells(i + 4, 3) <> lcase Then
        nlc = nlc + 1
        lcase = Sheet2.Cells(i + 4, 3).Value
    End If
    i = i + 1
Loop


'this part creates arrays for slope calculation
'first array

Dim arr1() As String
ReDim arr1(nlc)
i = 1       'sheet2 counter
j = 1       'arr1 counter
arr1(1) = Sheet2.Cells(4, 3).Value

Do While Sheet2.Cells(i + 4, 3) <> ""
    If Sheet2.Cells(i + 4, 3) <> arr1(j) Then
        j = j + 1
        arr1(j) = Sheet2.Cells(i + 4, 3).Value
    End If
    i = i + 1
Loop

'end of first array
 
 
 
 
 
 
 
 
 
 
 
 
 
Dim i2 As Integer
i2 = 1
Do While Cells(i2 + 3, 1) <> ""
Firstpoint = Cells(i2 + 3, 2)
secondpoint = Cells(i2 + 3, 3)
Dim length As Double
length = Cells(i2 + 3, 6)

'second array

Dim arr2() As Double
ReDim arr2(nlc, 2)
i = 1       'sheet2 counter
j = 1       'arr2 counter

Do While Sheet2.Cells(i + 3, 2) <> ""
    If Sheet2.Cells(i + 3, 2) = Firstpoint Then
        arr2(j, 1) = Sheet2.Cells(i + 3, 7)
        j = j + 1
    End If
    i = i + 1
Loop

i = 1
j = 1

Do While Sheet2.Cells(i + 3, 2) <> ""
    If Sheet2.Cells(i + 3, 2) = secondpoint Then
        arr2(j, 2) = Sheet2.Cells(i + 3, 7)
        j = j + 1
    End If
    i = i + 1
Loop

'end of second array





'third array
Dim arr3() As Double
ReDim arr3(nlc)

For i = 1 To nlc
    arr3(i) = Abs((arr2(i, 1) - arr2(i, 2)) / length)
Next i

'end of third array





'calculate max of array
Dim max As Double, maxi As Integer
max = arr3(1)
maxi = 1
For i = 2 To nlc
    If arr3(i) > max Then
        max = arr3(i)
        maxi = i
    End If
Next i
'end of max of array
    
Cells(i2 + 3, 8).Value = arr1(maxi)
Cells(i2 + 3, 9).Value = max
i2 = i2 + 1

Loop
    

End Sub




'For i = 1 To nlc
'    Cells(10 + i, 10) = arr2(i, 1)
'Next i

