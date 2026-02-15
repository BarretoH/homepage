Attribute VB_Name = "Module1"
'Modified from LMS.xls
'Oct 2014
'HB


Function LMS(myRangeY As Range, myRangeX As Range) As Variant
'runs the LMS fit routine
Application.Volatile (True)
On Error GoTo ErrorHandler:

Dim I As Integer
Dim J As Integer

'min screen flicker
Application.ScreenUpdating = False

'determine shape of original data and put down labels if available
Dim NumberRowsTotal As Long
Dim n As Long
Dim LabelBuffer As Integer

NumberRowsTotal = myRangeX.Rows.Count
n = NumberRowsTotal

Dim p As Long
Dim NumberSubsets As Long
p = 3 ' number of parameters estimated plus y
NumberSubsets = Application.WorksheetFunction.Fact(n) / (Application.WorksheetFunction.Fact(2) * Application.WorksheetFunction.Fact(n - 2))

'get LMS
'based on Rousseeuw [1984]
'create subset drawn to obs number mapping
Dim SubsetNumberArray() As Integer
ReDim SubsetNumberArray(1 To NumberSubsets, 1 To 2)
Dim First As Integer
First = 2 'the second in the Double starts at 2
Dim Counter As Integer
Counter = 1
For I = 1 To n - 1
    For J = First To n
        SubsetNumberArray(Counter, 1) = I
        SubsetNumberArray(Counter, 2) = J
        Counter = Counter + 1
    Next
    First = First + 1
Next

'load data
Dim DataArray() As Double
ReDim DataArray(1 To n, 1 To p)
For I = 1 To n
    DataArray(I, 1) = 1
    DataArray(I, 2) = myRangeX.Cells(I, 1)
    DataArray(I, 3) = myRangeY.Cells(I, 1)
Next

'sort the data by first x
qsort DataArray, 1, n, p, 2

Dim TwoObsArray(1 To 2, 1 To 3) As Double
Dim YMinusSlopeXArray() As Double
ReDim YMinusSlopeXArray(1 To n, 1 To 2) As Double '1st col is value, 2nd col is obs number
Dim ShortestHalfArray() As Double
ReDim ShortestHalfArray(Int(n / 2) + 1 To n, 1 To 2)
Dim MinShortestHalfValue As Double
Dim MinShortestHalfLocation As Integer
Dim ResidualsSquared() As Double
ReDim ResidualsSquared(1 To n)
Dim EstimatesArray() As Double
ReDim EstimatesArray(1 To NumberSubsets, 1 To 3)
Dim MinMedSR As Double
Dim MinMedSRLocation As Integer

For I = 1 To NumberSubsets
    'load two obs array
    For J = 1 To 2
        For k = 1 To 3
            TwoObsArray(J, k) = DataArray(SubsetNumberArray(I, J), k)
        Next
    Next
    'fit line through two points
    Dim Slope As Double
    Dim Intercept As Double
If TwoObsArray(2, 2) - TwoObsArray(1, 2) <> 0 Then 'checks to make sure we don't divide by zero July 19, 2006
    Slope = (TwoObsArray(2, 3) - TwoObsArray(1, 3)) / (TwoObsArray(2, 2) - TwoObsArray(1, 2))
    Intercept = TwoObsArray(1, 3) - Slope * TwoObsArray(1, 2)
    'do the shortest halfs to get the optimal intercept
        'create y - slope*x
        For k = 1 To n
            YMinusSlopeXArray(k, 1) = DataArray(k, 3) - Slope * DataArray(k, 2)
            YMinusSlopeXArray(k, 2) = k
        Next
        'sort
        qsort YMinusSlopeXArray, 1, n, 1, 1

        'find shortest half
        MinShortestHalfValue = YMinusSlopeXArray(Int(n / 2) + 1, 1) - YMinusSlopeXArray(1, 1) 'initialize with first half
        MinShortestHalfLocation = Int(n / 2) + 1
        For k = Int(n / 2) + 1 To n
            ShortestHalfArray(k, 1) = k
            ShortestHalfArray(k, 2) = YMinusSlopeXArray(k, 1) - YMinusSlopeXArray(k - Int(n / 2), 1)
            If MinShortestHalfValue > ShortestHalfArray(k, 2) Then
                MinShortestHalfValue = ShortestHalfArray(k, 2)
                MinShortestHalfLocation = k
            End If
        Next
        'find the intercept
        Dim InterceptOptimal As Double
        InterceptOptimal = MinShortestHalfValue / 2 + YMinusSlopeXArray(MinShortestHalfLocation - Int(n / 2), 1)
    'calculate the SSR at slope and optimal intercept
    Dim MedSR As Double
    MedSR = 0 'initialize
        'load residuals squared
        For k = 1 To n
            ResidualsSquared(k) = (DataArray(k, 3) - (InterceptOptimal + Slope * DataArray(k, 2))) ^ 2
        Next k
        MedSR = MedSR + Application.WorksheetFunction.Median(ResidualsSquared)
        
    'load estimates array
    EstimatesArray(I, 1) = InterceptOptimal
    EstimatesArray(I, 2) = Slope
    EstimatesArray(I, 3) = MedSR
    
    'track the min MedSR
    If I = 1 Then
        MinMedSR = MedSR 'initialize
        MinMedSRLocation = 1 'initialize
    Else
        If MinMedSR > MedSR Then
            MinMedSR = MedSR
            MinMedSRLocation = I
        End If
    End If
End If 'check divide by zero subset
Next I

'LMS = EstimatesArray(MinMedSRLocation, 2)

Dim FinalResults() As Variant
ReDim FinalResults(1 To p, 1 To 1)

FinalResults(1, 1) = EstimatesArray(MinMedSRLocation, 1)
FinalResults(2, 1) = EstimatesArray(MinMedSRLocation, 2)
FinalResults(3, 1) = MinMedSR

LMS = FinalResults


Application.ScreenUpdating = True
Exit Function
ErrorHandler:
LMS = "Can't do it. Sad."
End Function


'Sorting algorithm

' Thanks to David Maharry
' Aug 2000
' quick sort of array x which is x(l:r)
' l stands for left and r stands for right, but it's the same as top to bottom
' sorts array x[l..r,n] using the mth column for comparison
' n is the number of columns in the array x
' m is column that will be used for comparison (that is, m is the sorted column)

' Barreto made it sort decimal arrays by declaring several variables as Double
' If you make the temp variable in the swap sub an integer, you can see what the
' code is doing.

' To use, call qsort with the FIVE parameter values as described
' 1) The array to be sorted. This must be DIMmed as a Double.
' 2) l, stands for Left, or the first row number in the array; usually a 1
' 3) r, stands for Right, or the last row number in the array
' 4) n is the number of columns in the array
' 5) m is the column that is sorted from low to high
' Example
'Dim myArray(1 To 4, 1 To 3) As Double
'qsort myArray, 1, 4, 3, 2
' This sorts the values in myArray, which is a 4x3 matrix, using the second column

'Changed declarations to Long Mar-2002

Sub qsort(X() As Double, l As Long, r As Long, n As Long, m As Long)
    Dim p As Double
    Dim k As Long
    Dim pv As Double
    p = pivotValue(X, l, r, m)
    If (p <> -1) Then
        pv = X(p, m)
        Call MyPartition(X, l, r, pv, k, n, m)
        Call qsort(X, l, k - 1, n, m)
        Call qsort(X, k, r, n, m)
    End If
End Sub

Function pivotValue(X() As Double, l As Long, r As Long, m As Long)
    Dim Found As Boolean
    Dim I As Long
    Dim p As Long
    Dim middle As Long
    Dim x1 As Long, x2 As Long, x3 As Long
    Dim check As Integer, check2 As Integer, ptemp As Integer
    ' found is a check to see that there are at least two unequal
    ' entries in the list
    ' ptemp stores p in case median of three method fails
    ' Incorporates median of three method
    ' Reference is Sedgewick
    ' Help from Dave Maharry
    ' FH
    Found = False
    I = l
    
    While Found = False And I < r
        If X(I, m) <> X(I + 1, m) Then
            Found = True
            p = I + 1
        Else
            I = I + 1
        End If
    Wend
    ptemp = I
    If I < r Then
        If X(I + 1, m) > X(I, m) Then ptemp = I + 1
    End If
    ' ptemp is needed in case we have three equal pivot values
    ' in the median of three method
    '
    If Found = True Then

        If r - l > 1 Then
            x1 = X(l, m)
            x2 = X(r, m)
            middle = (l + r) / 2
            x3 = X(middle, m)
            Call CheckforSmaller(x1, x2, x3, l, r, middle, check, check2, p)
            If check = 0 Then
                If x1 > x2 Then
                    If x1 > x3 Then
                        If x2 > x3 Then
                            p = r
                        Else
                            p = middle
                        End If
                    Else
                        p = l
                    End If
                ElseIf x2 > x3 Then
                    If x1 > x3 Then
                        p = l
                    Else
                        p = middle
                    End If
                Else: p = r
                End If
            End If
            ' If check = 1 we assigned p already in the Check for smaller
            ' routine
         Else
            p = ptemp
            'If x(p, M) < x(p - 1, M) Then
            '   p = p - 1
            'End If
        End If
    Else
        p = -1
    End If
 If check2 = 1 Then p = ptemp
 ' check2 comes from CheckforSmaller and tells us
 ' that all three x1, x2, and x3 are equal
 'If all three are equal just use the original p

    pivotValue = p
End Function

Sub CheckforSmaller(x1 As Long, x2 As Long, x3 As Long, l As Long, r As Long, middle As Long, check As Integer, check2 As Integer, p As Long)
check2 = 0
check = 1
' This checks to see which element is smaller
'June 2002
' FH

If x1 = x2 Then
    If x1 > x3 Then
        p = l  ' pick x1 because it is larger than x3
    Else
        p = middle ' pick x3 because it may be bigger than x1 and x2
        If x1 = x3 Then check2 = 1
        ' This says that all three elements are equal, so
        ' we need to discard the result
        '
    End If
ElseIf x1 = x3 Then
    If x1 > x2 Then
        p = l  ' pick x1 because it is bigger than x2
    Else
        p = r  ' pick x2 because we know it is bigger than x1 and x3
    End If

ElseIf x2 = x3 Then
    If x2 > x1 Then
        p = r  ' pick x2
    Else
        p = l  ' pick x1 because we know that x1 <> x2
    End If
Else
    check = 0
End If
        
End Sub
Sub MyPartition(X() As Double, l As Long, r As Long, pv As Double, k As Long, n As Long, m As Long)
    Dim lo As Long
    Dim hi As Long
    lo = l
    hi = r
    While lo <= hi
        If X(lo, m) < pv Then
            lo = lo + 1
        ElseIf X(hi, m) >= pv Then
            hi = hi - 1
        Else
            Call swap(X, hi, lo, n)
            hi = hi - 1
            lo = lo + 1
        End If
    Wend
    
    k = lo

End Sub

Sub swap(X() As Double, I As Long, J As Long, n As Long)
    Dim k As Long
    Dim temp As Double
    For k = 1 To n
        temp = X(I, k)
        X(I, k) = X(J, k)
        X(J, k) = temp
    Next
End Sub

