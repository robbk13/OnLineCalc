Public Class MortalityProbability
    Public Property t As Integer
    Public Property Age As Integer
    Public Property Probability As Decimal
    Public Property Gender As constantGender

    Public Property ProbabilityofSurvivaltoAge_t As Decimal
    Public Property ProbabilityofSurvivaltoAge_t_exact As Decimal

    Public Property ResetProbabilityOfSurvivalToAge_t_ExactAge As Decimal

    Public Property UserName As String = "Tom"
End Class

Public Class MortalityProbabilitySeries
    Public Property SourceTableName As constantMortalityTableType = Nothing
    Public Property InitialAge As Integer = Nothing
    Public Property Gender As constantGender = Nothing

    Public Property PartialYear As Decimal = 0.5
    Private _InternalList As New List(Of MortalityProbability)
    Public ReadOnly Property MortalitySeries() As List(Of MortalityProbability)
        Get
            Dim DB As New MortalityDataContext

            Dim parameterGender As String = [Enum].GetName(GetType(constantGender), Me.Gender)
            Dim parameterSourceTableName = [Enum].GetName(GetType(constantMortalityTableType), Me.SourceTableName)

            Dim MTable = From M In DB.MortabilitySearchables Where M.SourceName = parameterSourceTableName And M.Gender = parameterGender And M.Age >= Me.InitialAge
                         Order By M.Age

            For Each Prob In MTable
                Dim P As New MortalityProbability With {.T = Prob.Age, .Age = Prob.Age, .Gender = Me.Gender, .Probability = Prob.Probability}
                _InternalList.Add(P)
            Next

            Fill_ProbabilityofSurvivaltoAge_t()
            Fill_ProbabilityofSurvivaltoAge_t_Exact()
            Fill_ResetProbabilityOfSurvivalToAge_t_ExactAge(t)

            Return _InternalList
        End Get

    End Property

    Private Sub Fill_ProbabilityofSurvivaltoAge_t()
        Dim V As Decimal = 1
        Dim IsFirst As Boolean = True
        Dim CurRow As MortalityProbability = Nothing
        Dim PrevRow As MortalityProbability = Nothing
        For Each P In _InternalList
            If IsFirst Then
                P.ProbabilityofSurvivaltoAge_t = 1S
                IsFirst = False
            Else
                Dim tempValue As Decimal = PrevRow.ProbabilityofSurvivaltoAge_t * (1S - PrevRow.Probability)
                P.ProbabilityofSurvivaltoAge_t = Math.Round(tempValue, 6)
            End If
            PrevRow = P
        Next
    End Sub
    Private Sub Fill_ProbabilityofSurvivaltoAge_t_Exact()
        For Each P In _InternalList
            Dim Pointer As Integer = _InternalList.IndexOf(P) + 1
            If Pointer < _InternalList.Count Then
                Dim NextRow As MortalityProbability = _InternalList(Pointer)
                If NextRow.ProbabilityofSurvivaltoAge_t = 0 And Me.PartialYear = 0 Then
                    P.ProbabilityofSurvivaltoAge_t_exact = Math.Round(P.ProbabilityofSurvivaltoAge_t ^ (1 - Me.PartialYear), 6)
                Else
                    P.ProbabilityofSurvivaltoAge_t_exact = Math.Round(NextRow.ProbabilityofSurvivaltoAge_t ^ Me.PartialYear * P.ProbabilityofSurvivaltoAge_t ^ (1 - Me.PartialYear), 6)
                End If
            End If
        Next
    End Sub
    Public Function Fill_ResetProbabilityOfSurvivalToAge_t_ExactAge(Optional t As Integer = 0) As Decimal
        Dim ResetProb(540) As Decimal   '540 is 12 months * 45 years
        Dim Demoninator As Decimal = Math.Round(_InternalList(0).ProbabilityofSurvivaltoAge_t_exact, 6)
        For Each p In _InternalList
            p.ResetProbabilityOfSurvivalToAge_t_ExactAge = Math.Round(p.ProbabilityofSurvivaltoAge_t_exact / Demoninator, 6)
            ResetProb(t) = p.ResetProbabilityOfSurvivalToAge_t_ExactAge
            t = t + 1
        Next
        Return ResetProb(540)

    End Function
End Class
