Public Class karupKingoCoefficients
    Const Period = 12
    Const Precision = 6

    Public Sub New()
        Dim SeriesLendth As Integer = Period + 1
        Dim fraction As Double = 1 / Period

        Dim StartValue As Double = 0                              'beginning KKC value, everything populates from StartValue
        For i As Integer = 0 To KKC.GetLength(0) - 1
            If i = 0 Then
                KKC(0, 0) = StartValue                            'starts here
            Else
                KKC(i, 0) = fraction * i                          'first column values
            End If

            Dim A1 As Double = KKC(i, 0)
            KKC(i, 1) = (-A1 + 2 * A1 ^ 2 - A1 ^ 3) / 2          'Column 2
            KKC(i, 2) = (2 - 5 * A1 ^ 2 + 3 * A1 ^ 3) / 2        'Column 3
            KKC(i, 3) = (A1 + 4 * A1 ^ 2 - 3 * A1 ^ 3) / 2       'Column 4
            KKC(i, 4) = (-A1 ^ 2 + A1 ^ 3) / 2                   'Column 5
        Next
    End Sub

    Public KKC(12, 4) As Double
    Public Property KK() As Array
        Get
            Return KKC
        End Get
        Set(ByVal value As Array)
            KKC = value
        End Set
    End Property

    Private Function TableRow(time As Integer) As Integer
        If time = 0 Then
            Return 0
        Else
            Return time Mod 12
        End If

    End Function



    Function kkCoefficient(t, calctype) As Double
        Dim Result As Double = Nothing
        Dim A1 As Double = TableRow(t) / 12


        Select Case calctype
            Case 0
                Result = A1
            Case 1
                Result = (-A1 + 2 * A1 ^ 2 - A1 ^ 3) / 2
            Case 2
                Result = (2 - 5 * A1 ^ 2 + 3 * A1 ^ 3) / 2
            Case 3
                Result = (A1 + 4 * A1 ^ 2 - 3 * A1 ^ 3) / 2
            Case 4
                Result = (-A1 ^ 2 + A1 ^ 3) / 2

        End Select
        Return Result
    End Function


End Class
