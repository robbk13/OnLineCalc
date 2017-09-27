Module Calc

    Public FMVt() As Single
    Dim Home1 As New myhome '.RepairsPrevMaint()
    Dim ResetProbabilityOfSurvivalExactAge() As Decimal
    Public ProbLastDeathBTween() As Decimal
    Dim AnnualPeriod(ProgramParameters_FMV.ProjectedLife_t) As Integer
    'Dim OptionActionRate(ProgramParameters_FMV.ProjectedLife_t) As Decimal
    Dim ProbNotTenderedMonthly(ProgramParameters_FMV.ProjectedLife_t) As Decimal
    Public RepairsPrevMaint() As Decimal
    Public t As Integer = 0
    Public x As Integer = 0
    Dim q As Integer
    Dim tTotal As Integer = 0  'total of t

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sub Main()

        Console.WriteLine("Program Starts")

        'initializeRun(2)




        'the values return by the 

        'Create Mortality table OBJECT MSeries matches column AQ-AR and AZ-BA on the Mort table
        'which use 73 male, 70 female for SS 
        Dim MSeries As New MortalityProbabilitySeries

        'set value that define the series
        'note SS requires an initial age parameter since that table begins with age 1
        With MSeries
            .InitialAge = 73
            .Gender = constantGender.M
            .SourceTableName = constantMortalityTableType.SS
            'MortalitySeries property is a list of the values retrieved from the database for the previous parameters
            'OutputData(.MortalitySeries)
            ProbNoTender(.MortalitySeries)

        End With

        'Dim list1() as Double = Moving_SummaryScalar(507, 507, 507, 507, 507)

        'Same series object can fetch a different set of data when you change one or more paramters
        'in this case switch to 70 year old female
        With MSeries
            .Gender = constantGender.F
            .InitialAge = 70
            'OutputData(.MortalitySeries)
            ProbNoTender(.MortalitySeries)
        End With

        'use parameter to fetch data from CDC table

        With MSeries
            .SourceTableName = constantMortalityTableType.CDC
            'age and gender carried over from previous setting
            'OutputData(.MortalitySeries)
            ProbNoTender(.MortalitySeries)

        End With


        Dim vMortalityTable As List(Of MortalityProbability) = Nothing

        Dim ReversArray(ProgramParameters_FMV.ProjectedLife_t) As Single '= Nothing

        Dim P = Nothing

    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Sub OutputData(vMortalityTable As List(Of MortalityProbability))


        For Each P In vMortalityTable
            'Console.WriteLine(String.Format("Age: {0} Q: {1:n6} {2} {3} {4}", P.Age, P.Probability, P.ProbabilityofSurvivaltoAge_t, P.ProbabilityofSurvivaltoAge_t_exact, P.ResetProbabilityOfSurvivalToAge_t_ExactAge))
            'Console.readline		
        Next
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'FMV

    Public Function GetFMVTable() As Single()
        'Sub OutputFMVTable()
        Dim FairMarketValue As New List(Of FMV)           'creates a list of FMV items as defined in Public Class FMV
        Dim PreviousHomeValue As Decimal = ProgramParameters_FMV.FMV_at_t0    ' gets that initial value from Parameters
        Dim Monthly_Factor_HPA = (1 + ProgramParameters_FMV.Annual_HPA) ^ (1 / 12)
        Dim FMV_at_t0 As Decimal = ProgramParameters_FMV.FMV_at_t0
        Dim t As Integer = 0
        Dim Deviation_Remaining_HPA As Decimal = ProgramParameters_FMV.Deviation_Remaining_HPA
        Dim First_Month_Deviation_Recovery_HPA = ProgramParameters_FMV.First_Month_Deviation_Recovery_HPA
        Dim FMVt(ProgramParameters_FMV.ProjectedLife_t) As Single         'Average FMV from t to t+1
        Dim CurrentFMV As Single

        For time As Integer = t To ProgramParameters_FMV.ProjectedLife_t
            'loop creates one row for each year, create a new row and assigns year and value properties
            Dim TableRow As New FMV With {.t = time, .Year = Int(t / 12)}

            If time = 0 Then
                With TableRow
                    .Simple_Proj_t = FMV_at_t0
                    .Biased_Proj_t = .Simple_Proj_t / (1 + ProgramParameters_FMV.Temporary_Deviation_HPA)
                    .Weight_Simple_Proj_t = Calc_Weight_Simple_Proj_t(.Simple_Proj_t, t)
                    .FMV_t = .Simple_Proj_t * .Weight_Simple_Proj_t + (1 - .Weight_Simple_Proj_t) * .Biased_Proj_t
                    .AvgFMV_t_t1 = (.FMV_t + .Simple_Proj_t * Monthly_Factor_HPA) / 2
                    PreviousHomeValue = .FMV_t
                    CurrentFMV = .AvgFMV_t_t1
                    ' t = t + 1
                End With
                FMVt(t) = CurrentFMV
            Else
                With TableRow
                    .Simple_Proj_t = PreviousHomeValue * Monthly_Factor_HPA
                    .Biased_Proj_t = PreviousHomeValue * Monthly_Factor_HPA
                    .Weight_Simple_Proj_t = Calc_Weight_Simple_Proj_t(.Simple_Proj_t, t)
                    .FMV_t = .Simple_Proj_t * .Weight_Simple_Proj_t + (1 - .Weight_Simple_Proj_t) * .Biased_Proj_t
                    .AvgFMV_t_t1 = (.FMV_t + .Simple_Proj_t * Monthly_Factor_HPA) / 2
                    PreviousHomeValue = .Simple_Proj_t
                    CurrentFMV = .AvgFMV_t_t1
                    't = t + 1
                End With
                FMVt(t) = CurrentFMV
            End If

            FairMarketValue.Add(TableRow)

            t = t + 1
        Next

        Return FMVt

        For Each trow As FMV In FairMarketValue 'iterates the collection row by row to output the table
            With trow
                ' Console.Write("T: " & .t)
                ' Console.Write(" Year: " & .Year)
                ' Console.Write(String.Format(" Simple_Proj_t: {0:c}", .Simple_Proj_t))
                ' Console.Write(String.Format(" Biased_Proj_t: {0:c}", .Biased_Proj_t))
                ' Console.Write(String.Format(" Weight_Simple_Proj_t: {0:p}", .Weight_Simple_Proj_t))
                ' Console.Write(String.Format(" FMV_t: {0:c}", .FMV_t))
                ' Console.Write(String.Format(" AvgFMV_t_t1: {0:c}", .AvgFMV_t_t1))
                ' Console.WriteLine()
            End With
        Next
        'Console.WriteLine("Press ENTER to Exit")
        'Console.ReadLine()
        'End Sub
    End Function
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''   
    Public Function Calc_Weight_Simple_Proj_t(vSimple_Proj_t As Decimal, ByVal t As Integer)
        If t = 0 Then
            Return 1
        Else
            Return 1 / (1 + (t ^ ((Math.Log(1 - ProgramParameters_FMV.First_Month_Deviation_Recovery_HPA) + Math.Log(1 - ProgramParameters_FMV.Deviation_Remaining_HPA) - Math.Log(ProgramParameters_FMV.First_Month_Deviation_Recovery_HPA * ProgramParameters_FMV.Deviation_Remaining_HPA)) / Math.Log(ProgramParameters_FMV.Month_When_Deviation_Remaining_Ends_HPA))) / (1 / ProgramParameters_FMV.First_Month_Deviation_Recovery_HPA) - 1)
        End If
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'MORTALITY AND PROB-NO TENDER

    Public Function ProbNoTender(vMortalityTable As List(Of MortalityProbability)) As WebSiteCalculatorOutputs ' Probabilities for Cancellation and Death - Assuming no Tenders
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        FMVt = GetFMVTable()          'To Set FMV value in the rest of the program
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	   
        Dim P = Nothing

        Dim ProbReset(540) As Double

        Dim CancellationTotal As Double = 0
        Dim DeathTotal As Double = 0
        Dim CandDTotal As Double = 0       'Total Cancellation probablilites and Death Probablilities should = 100%

        t = 0
        For Each P In vMortalityTable
            ProbReset(t) = P.ResetProbabilityOfSurvivalToAge_t_ExactAge
            t = t + 1
        Next

        'MORTALITY
        Dim ProbSurvivalTo_t(ProgramParameters_FMV.ProjectedLife_t + 20) As Double          'S    Probability of Survival to t   - from Mort tab
        Dim PServival As Decimal = 0
        Dim ProbLastDeath(ProgramParameters_FMV.ProjectedLife_t + 20) As Double             'q    Probability of Last Death Between t and t+1 Given Survival to t
        Dim PLDeath As Decimal = 0
        Dim ProbLastDeathBTween(ProgramParameters_FMV.ProjectedLife_t + 20) As Double       'f   Probability of Death Between t and t+1




        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'PROB-NO TENDER
        Dim OptionRate(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                    'A    Option Action Rate
        Dim DeathRate(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                     'D    Death Rate 
        Dim ProbOfCancellation(ProgramParameters_FMV.ProjectedLife_t + 6) As Double        'qC   Probability of Cancellation Between t and t+1 Given Present at t   
        Dim ProbDeathBTPres(ProgramParameters_FMV.ProjectedLife_t + 6) As Double           'qD   Probability of Death Between t and t+1 Given Present at t
        Dim ProbOfCancellationt(ProgramParameters_FMV.ProjectedLife_t + 6) As Double       'fC   Probability of Cancellation Between t and t+1                      
        Dim ProbDeathBT(ProgramParameters_FMV.ProjectedLife_t + 6) As Double               'fD   Probability of Death Between t and t+1
        Dim ProbBeingPresent(ProgramParameters_FMV.ProjectedLife_t + 6) As Double          'S    Probability of Being Present at t                                     

        Dim KKCo1 As karupKingoCoefficients = New karupKingoCoefficients

        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t   'Probability of Survival to t     'S - from Mort tab
            If t < 12 Then
                ProbSurvivalTo_t(t) = KKCo1.KKC(t Mod 12, 1) * (ProbReset(Int(t / 12) + 2) - 3 * ProbReset(Int(t / 12) + 1) + 3 * ProbReset(Int(t / 12))) + KKCo1.KKC(t Mod 12, 2) * ProbReset(Int(t / 12)) + KKCo1.KKC(t Mod 12, 3) * ProbReset(Int(t / 12) + 1) + KKCo1.KKC(t Mod 12, 4) * ProbReset(Int(t / 12) + 2)
            ElseIf t > 11 And t < 305 Then
                ProbSurvivalTo_t(t) = KKCo1.KKC(t Mod 12, 1) * ProbReset(Int(t / 12) - 1) + KKCo1.KKC(t Mod 12, 2) * ProbReset(Int(t / 12)) + KKCo1.KKC(t Mod 12, 3) * ProbReset(Int(t / 12) + 1) + KKCo1.KKC(t Mod 12, 4) * ProbReset(Int(t / 12) + 2)
            ElseIf t > 304 Then  'WTF????  fixed fields???
                ProbSurvivalTo_t(t) = ProbSurvivalTo_t(t - 1) / (1 + ((((ProbSurvivalTo_t(303) / ProbSurvivalTo_t(304)) - 1) ^ ((t - 267) / 36)) / (((ProbSurvivalTo_t(267) / ProbSurvivalTo_t(268)) - 1) ^ ((t - 303) / 36))))
            End If

            Console.WriteLine(ProbSurvivalTo_t(t))

            PServival = PServival + ProbSurvivalTo_t(t)                                     'S    Probability of Survival to t   - from Mort tab'to test to se if total matches the Mort column M
        Next

        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t

            If t < 6 Then
                Select Case t
                    Case 0
                        OptionRate(t) = Cancellations.CancelMonth0                           'A    Option Action Rate
                    Case 1
                        OptionRate(t) = Cancellations.CancelMonth1
                    Case 2
                        OptionRate(t) = Cancellations.CancelMonth2
                    Case 3
                        OptionRate(t) = Cancellations.CancelMonth3
                    Case 4
                        OptionRate(t) = Cancellations.CancelMonth4
                    Case 5
                        OptionRate(t) = Cancellations.CancelMonth5
                    Case Else
                        Exit Select
                End Select
            Else
                OptionRate(t) = 0
            End If

            If ProbSurvivalTo_t(t) = 0 Then         ''probablility of early death for Life Estate Calc
                ProbLastDeath(t) = 1
            Else
                ProbLastDeath(t) = 1 - ProbSurvivalTo_t(t + 1) / ProbSurvivalTo_t(t)             'q    Probability of Death Between t and t+1 Given Survival to t                                                                          'q
            End If
            PLDeath = PLDeath + ProbLastDeath(t)

            DeathRate(t) = ProbLastDeath(t)                                                      'D    Death Rate 

            '                                                                                     qC   Probability of Cancellation Between t and t+1 Given Present at t  
            ProbOfCancellation(t) = (1 - (1 - ProbLastDeath(t)) * (1 - OptionRate(t))) * OptionRate(t) / (ProbLastDeath(t) + OptionRate(t))

            '                                                                                     qD   Probability of Death Between t and t+1 Given Present at t
            ProbDeathBTPres(t) = (1 - (1 - DeathRate(t)) * (1 - OptionRate(t))) * DeathRate(t) / ((DeathRate(t) + OptionRate(t)))

            PerformLoopedOperation_A(ProbLastDeath, ProbOfCancellation, ProbBeingPresent, t)
            '                                                                                      fC   Probability of Cancellation Between t and t+1 
            ProbOfCancellationt(t) = ProbOfCancellation(t) * ProbBeingPresent(t)
            ProbLastDeathBTween(t) = ProbLastDeath(t) * ProbBeingPresent(t)                       'f   Probability of Death Between t and t+1

            ProbDeathBT(t) = ProbBeingPresent(t) * ProbDeathBTPres(t)                             'fD   Probability of Death Between t and t+1

            CancellationTotal = CancellationTotal + ProbOfCancellationt(t)
            DeathTotal = DeathTotal + ProbLastDeathBTween(t)
        Next

        CandDTotal = CancellationTotal + DeathTotal   'should show 100% probability


        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'OPTIONS - Cancellations and Life Estate

        Dim TenderRateAnnual(ProgramParameters_FMV.ProjectedLife_t / 12 + 7) As Double                          'Tender Rate Annual
        Dim ProbNotTenderedAnnual(ProgramParameters_FMV.ProjectedLife_t / 12 + 7) As Double                     'Probability Not Tendered Annual
        Dim OptionActionRate(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                                   'A       
        Dim CancelMonth As String = Nothing
        Dim TenderYr As String = Nothing

        For t As Integer = 0 To 49
            Select Case t
                Case 0
                    TenderRateAnnual(t) = LETenders.TenderYr0
                    ProbNotTenderedAnnual(t) = 1
                Case 1
                    TenderRateAnnual(t) = LETenders.TenderYr1
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 2
                    TenderRateAnnual(t) = LETenders.TenderYr2
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 3
                    TenderRateAnnual(t) = LETenders.TenderYr3
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 4
                    TenderRateAnnual(t) = LETenders.TenderYr4
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 5
                    TenderRateAnnual(t) = LETenders.TenderYr5
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 6
                    TenderRateAnnual(t) = LETenders.TenderYr6
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 7
                    TenderRateAnnual(t) = LETenders.TenderYr7
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 8
                    TenderRateAnnual(t) = LETenders.TenderYr8
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 9
                    TenderRateAnnual(t) = LETenders.TenderYr9
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 10
                    TenderRateAnnual(t) = LETenders.TenderYr10
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 11
                    TenderRateAnnual(t) = LETenders.TenderYr11
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 12
                    TenderRateAnnual(t) = LETenders.TenderYr12
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 13
                    TenderRateAnnual(t) = LETenders.TenderYr13
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 14
                    TenderRateAnnual(t) = LETenders.TenderYr14
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 15
                    TenderRateAnnual(t) = LETenders.TenderYr15
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 16
                    TenderRateAnnual(t) = LETenders.TenderYr16
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 17
                    TenderRateAnnual(t) = LETenders.TenderYr17
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 18
                    TenderRateAnnual(t) = LETenders.TenderYr18
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
                Case 19
                    TenderRateAnnual(t) = LETenders.TenderYr19
                    ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
            End Select

            If t > 19 Then
                TenderRateAnnual(t) = TenderRateAnnual(t - 1) * LETenders.ReductionFactorLE
                ProbNotTenderedAnnual(t) = ProbNotTenderedAnnual(t - 1) * (1 - TenderRateAnnual(t - 1))
            End If
        Next

        Dim KKCo As karupKingoCoefficients = New karupKingoCoefficients

        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t

            If t < 12 Then
                '                                                                                                   ProbNotTenderedMonthly
                ProbNotTenderedMonthly(t) = KKCo.KKC((t Mod 12), 1) * (ProbNotTenderedAnnual(Int(t / 12) + 2) - 3 * ProbNotTenderedAnnual(Int(t / 12) + 1) + 3 * ProbNotTenderedAnnual(Int(t / 12))) + KKCo.KKC((t Mod 12), 2) * ProbNotTenderedAnnual(Int(t / 12)) + KKCo.KKC(t Mod 12, 3) * ProbNotTenderedAnnual(Int(t / 12) + 1) + KKCo.KKC(t Mod 12, 4) * ProbNotTenderedAnnual(Int(t / 12) + 2)

                If t < 6 Then
                    AnnualPeriod(t) = 0  ' for 1 though 5
                    Select Case t
                        Case 0
                            OptionActionRate(t) = Cancellations.CancelMonth0
                        Case 1
                            OptionActionRate(t) = Cancellations.CancelMonth1
                        Case 2
                            OptionActionRate(t) = Cancellations.CancelMonth2
                        Case 3
                            OptionActionRate(t) = Cancellations.CancelMonth3
                        Case 4
                            OptionActionRate(t) = Cancellations.CancelMonth4
                        Case 5
                            OptionActionRate(t) = Cancellations.CancelMonth5
                    End Select
                Else
                    OptionActionRate(t) = 1 - ProbNotTenderedMonthly(t - 5) / ProbNotTenderedMonthly(t - 6)
                    AnnualPeriod(t) = Int((t + 6) / 12) ' Annual Period
                End If
            Else
                ProbNotTenderedMonthly(t) = KKCo.KKC(t Mod 12, 1) * ProbNotTenderedAnnual(Int(t / 12) - 1) + KKCo.KKC(t Mod 12, 2) * ProbNotTenderedAnnual(Int(t / 12)) + KKCo.KKC(t Mod 12, 3) * ProbNotTenderedAnnual(Int(t / 12) + 1) + KKCo.KKC(t Mod 12, 4) * ProbNotTenderedAnnual(Int(t / 12) + 2)
                OptionActionRate(t) = 1 - ProbNotTenderedMonthly(t - 5) / ProbNotTenderedMonthly(t - 6)
                AnnualPeriod(t) = Int((t + 6) / 12) ' Annual Period
            End If
        Next

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'PROB-FULL

        'Probabilities for Cancellation, Life Estate Purchase, and Death

        'Dim OptionActionRate(ProgramParameters_FMV.ProjectedLife_t) As Double                                                                            'A 
        'Dim ProbLastDeath(ProgramParameters_FMV.ProjectedLife_t) As Double   

        'Totals												TOTALS                                                                            'D
        Dim OptActRate1 As Double = 0                    'A     Option Action Rate
        Dim DeathRate1 As Double = 0                     'D     Death Rate 
        Dim CancelProb1 As Double = 0                    'qC    Probability of Cancellation Between t and t+1 Given Present at t
        Dim ProbTenderBTween1 As Double = 0              'qT    Probability of Tender Between t and t+1 Given Present at t
        Dim ProbDeathBTween1 As Double = 0               'qD    Probability of Death Between t and t+1 Given Present at t
        Dim Cancellations1 As Double = 0                 'fC    Probability of Cancellation Between t and t+1
        Dim NonTenderStatusDeath As Double = 0           'fDnts Probability of Non Tender Status Death Between t and t+1
        Dim ProbPresentNonTender1 As Double = 0          'S     Probability of Being Present in Non Tender Status at t
        Dim LifeEstateBuy As Double = 0                  'fP    Probability of Life Estate Buy Between t and t+1
        Dim TenderStatusDeath As Double = 0              'fDts  Probability of Tender Status Death Between t and t+1

        Dim TotalProbFull As Double = 0               '		 Total of fC + fP + fDts + fDnts

        ''''''''''''''''''''''''''''''''''''
        Dim OptActRateTotal As Double = 0                   'total of Option Action Rate %               
        Dim ProbLastDeathBTwenTotal As Double = 0           'totals
        Dim ProbOfCancellationBTween As Double = 0          'totals

        Dim CancelProb(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                     'qC - Probability of Cancellation Between t and t+1 Given Present at t
        Dim TenderProb(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                     'qT - Probability of Tender Between t and t+1 Given Present at t
        Dim ProbDeathBTween(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                'qD - Probability of Death Between t and t+1 Given Present at t
        Dim CancelProbBTween(ProgramParameters_FMV.ProjectedLife_t + 6) As Double               'fC - Probability of Cancellation Between t and t+1
        Dim ProbNonTenderDeathBTween(ProgramParameters_FMV.ProjectedLife_t + 6) As Double       'fDnts Probability of Non Tender Status Death Between t and t+1
        Dim ProbNonTenderPresent(ProgramParameters_FMV.ProjectedLife_t + 6) As Double           'S

        Dim ProbLEBuyBTween(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                'fP - Probability of Life Estate Buy Between t and t+1   
        Dim ProbTenderDeathBTween(ProgramParameters_FMV.ProjectedLife_t + 6) As Double          'fDts - Probability of Tender Status Death Between t and t+1

        Dim ProbPresentTenderT1(ProgramParameters_FMV.ProjectedLife_t + 6) As Double        'T1
        Dim ProbPresentTenderT2(ProgramParameters_FMV.ProjectedLife_t + 6) As Double        'T2
        Dim ProbPresentTenderT3(ProgramParameters_FMV.ProjectedLife_t + 6) As Double        'T3
        Dim ProbPresentTenderT4(ProgramParameters_FMV.ProjectedLife_t + 6) As Double        'T4
        Dim ProbPresentTenderT5(ProgramParameters_FMV.ProjectedLife_t + 6) As Double        'T5
        Dim ProbPresentTenderT6(ProgramParameters_FMV.ProjectedLife_t + 6) As Double        'T6

        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t
            'Console.writeline(t)
            'Console.writeline(ProbLastDeath(t))
            'Console.readline

            OptActRateTotal = OptActRateTotal + OptionActionRate(t)
            ProbLastDeathBTwenTotal = ProbLastDeathBTwenTotal + ProbLastDeathBTween(t)

            ProbDeathBTween(t) = (1 - (1 - ProbLastDeath(t)) * (1 - OptionActionRate(t))) * ProbLastDeath(t) / (ProbLastDeath(t) + OptionActionRate(t))                        'qD   16.82153  'Probability of Death Between t and t+1 Given Present at t                         

            If t = 0 Then                                                                   'Probability of Being Present in Non Tender Status at t                     'S    120.40304
                ProbNonTenderPresent(t) = 1
            ElseIf t < 7 Then
                ProbNonTenderPresent(t) = ProbNonTenderPresent(t - 1) * (1 - CancelProb(t - 1) - TenderProb(t - 1) - ProbDeathBTween(t - 1))
            Else
                ProbNonTenderPresent(t) = ProbNonTenderPresent(t - 1) * (1 - CancelProb(t - 1) - TenderProb(t - 1) - ProbDeathBTween(t - 1)) + LETenders.LEProbOfReneging * (1 - ProbLastDeathBTween(t - 1)) * ProbPresentTenderT6(t - 1)
            End If

            If t < 6 Then                                                                'Probability of Cancellation Between t and t+1 Given Present at t          
                CancelProb(t) = (1 - (1 - ProbLastDeath(t)) * (1 - OptionActionRate(t))) * OptionActionRate(t) / (ProbLastDeath(t) + OptionActionRate(t))            'qC    0.01598  Probability of Cancellation Between t and t+1 Given Present at t
                TenderProb(t) = 0                                                        ' Probability of Tender Between t and t+1 Given Present at t                 qT    1.19339 Probability of Tender Between t and t+1 Given Present at t
                CancelProbBTween(t) = ProbBeingPresent(t) * CancelProb(t)                'Probability of Cancellation Between t and t+1                              'fC    0.01577
            Else
                CancelProb(t) = 0                                                                                                                                     'qC
                TenderProb(t) = (1 - (1 - ProbLastDeath(t)) * (1 - OptionActionRate(t))) * OptionActionRate(t) / (ProbLastDeath(t) + OptionActionRate(t))             'qT
                CancelProbBTween(t) = 0                                                                                                                               'fC
            End If

            ProbNonTenderDeathBTween(t) = ProbDeathBTween(t) * ProbNonTenderPresent(t)                  'Probability of Non Tender Status Death Between t and t+1    fDnts

            If t = 0 Then
                ProbPresentTenderT1(t) = 0
                ProbPresentTenderT2(t) = 0
                ProbPresentTenderT3(t) = 0
                ProbPresentTenderT4(t) = 0
                ProbPresentTenderT5(t) = 0
                ProbPresentTenderT6(t) = 0

                ProbLEBuyBTween(t) = 0
                ProbTenderDeathBTween(t) = 0
            Else
                ProbPresentTenderT1(t) = TenderProb(t - 1) * ProbNonTenderPresent(t - 1)                 'T1 
                ProbPresentTenderT2(t) = ProbPresentTenderT1(t - 1) * (1 - ProbLastDeath(t - 1))         'T2
                ProbPresentTenderT3(t) = ProbPresentTenderT2(t - 1) * (1 - ProbLastDeath(t - 1))         'T3
                ProbPresentTenderT4(t) = ProbPresentTenderT3(t - 1) * (1 - ProbLastDeath(t - 1))         'T4
                ProbPresentTenderT5(t) = ProbPresentTenderT4(t - 1) * (1 - ProbLastDeath(t - 1))         'T5
                ProbPresentTenderT6(t) = ProbPresentTenderT5(t - 1) * (1 - ProbLastDeath(t - 1))         'T6

                ProbLEBuyBTween(t) = ProbPresentTenderT6(t) * (1 - ProbLastDeath(t)) * (1 - LETenders.LEProbOfReneging)                                              'fP
                ProbTenderDeathBTween(t) = ProbLastDeath(t) * (ProbPresentTenderT1(t) + ProbPresentTenderT2(t) + ProbPresentTenderT3(t) + ProbPresentTenderT4(t) + ProbPresentTenderT5(t) + ProbPresentTenderT6(t))
            End If

            Cancellations1 = Cancellations1 + CancelProbBTween(t)                                         'x
            LifeEstateBuy = LifeEstateBuy + ProbLEBuyBTween(t)                                            'x
            TenderStatusDeath = TenderStatusDeath + ProbTenderDeathBTween(t)                              'x
            NonTenderStatusDeath = NonTenderStatusDeath + ProbNonTenderDeathBTween(t)                     'x
            TotalProbFull = Cancellations1 + LifeEstateBuy + TenderStatusDeath + NonTenderStatusDeath
            OptActRate1 = OptActRate1 + OptionActionRate(t)
            DeathRate1 = DeathRate1 + ProbLastDeath(t)                                                       'D
            CancelProb1 = CancelProb1 + CancelProb(t)                                                        'qC    0.01598
            ProbTenderBTween1 = ProbTenderBTween1 + TenderProb(t)                                            'qT    1.19339
            ProbDeathBTween1 = ProbDeathBTween1 + ProbDeathBTween(t)                                         'qD   16.82153
            ProbPresentNonTender1 = ProbPresentNonTender1 + ProbNonTenderPresent(t)                          'S   120.40304

            If t = 499 Then
                Dim HUB As Decimal = 0
            End If

        Next

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        t = 0           'reset t to 0
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'Reversionary Costs - Fix-up prior to sale
        Dim ReversionaryCosts As Double
        'Dim RETCT As Double = Nothing

        Dim EstSqFt As Integer = HomeStatistics.HomeSize
        Dim HomeLotSize As Integer = HomeStatistics.LotSize

        Dim LSCT As Single, CLCT As Single, HVCT As Single, IPCT As Single, FCCT As Single, WCCT As Single, DMCT As Single, PVCT As Single

        CLCT = EstSqFt * ReversionCosts.CleaningPerSqFt * (PricingDefaultMultiplier.CleaningT + PricingDefaultMultiplier.CleaningL + PricingDefaultMultiplier.CleaningA + PricingDefaultMultiplier.CleaningG - 3)  'Cleaning
        HVCT = EstSqFt * ReversionCosts.HVACPerSqFt * (PricingDefaultMultiplier.HVACT + PricingDefaultMultiplier.HVACL + PricingDefaultMultiplier.HVACA + PricingDefaultMultiplier.HVACG - 3)  'HVAC
        IPCT = EstSqFt * ReversionCosts.InteriorPaintPerSqFt * (PricingDefaultMultiplier.InteriorPaintT + PricingDefaultMultiplier.InteriorPaintL + PricingDefaultMultiplier.InteriorPaintA + PricingDefaultMultiplier.InteriorPaintG - 3)  'InteriorPaint
        FCCT = EstSqFt * ReversionCosts.FloorCoveringsPerSqFt * (PricingDefaultMultiplier.FloorCoveringsT + PricingDefaultMultiplier.FloorCoveringsL + PricingDefaultMultiplier.FloorCoveringsA + PricingDefaultMultiplier.FloorCoveringsG - 3)  'FloorCoverings
        WCCT = EstSqFt * ReversionCosts.WindowCoveringsPerSqFt * (PricingDefaultMultiplier.WindowCoveringsT + PricingDefaultMultiplier.WindowCoveringsL + PricingDefaultMultiplier.WindowCoveringsA + PricingDefaultMultiplier.WindowCoveringsG - 3)  'WindowCoverings
        LSCT = (HomeLotSize - EstSqFt) * ReversionCosts.LandscapeMaintPerSqFt * (PricingDefaultMultiplier.LandscapingT + PricingDefaultMultiplier.LandscapingL + PricingDefaultMultiplier.LandscapingA + PricingDefaultMultiplier.LandscapingG - 3)  'landscaping
        DMCT = (HomeLotSize - EstSqFt) * ReversionCosts.DamagesPerSqFt * (PricingDefaultMultiplier.DamagesT + PricingDefaultMultiplier.DamagesL + PricingDefaultMultiplier.DamagesA + PricingDefaultMultiplier.DamagesG - 3)  'Damages
        PVCT = (HomeLotSize - EstSqFt) * ReversionCosts.PavingPerSqFt * (PricingDefaultMultiplier.PavingT + PricingDefaultMultiplier.PavingL + PricingDefaultMultiplier.PavingA + PricingDefaultMultiplier.PavingG - 3)  'Paving

        'POST-REVERSIONARY COSTS  - Costs incurred while waiting for sale to occur and close

        Dim LMCT As Double, ISCT As Double, UTCT As Double, RETCT As Double = Nothing, HOCT As Double
        'Dim ReversionaryCosts As Double

        Dim RSPM As Integer = ReversionCosts.ResaleInMo  ' the number of months until home can be resold

        ISCT = EstSqFt * ReversionCosts.InsurPerSqFt * RSPM / 12  'Home Insurance until sale
        RETCT = ReversionCosts.RETax * HomeStatistics.FMVt0 * RSPM / 12  'Real Estate Tax Accrual until sale
        UTCT = ReversionCosts.Utilities * RSPM / 12               'Utilities until saleLMCT = (HomeLotSize - EstSqFt) * ReversionCosts.LandscapeMaintPerSqFt * RSPM / 12   'Landscape maintenance while waiting for sale
        If LMCT < ReversionCosts.LandscapeMaintMin Then           'Landscaping Minimum - Not less than the minimum
            LMCT = ReversionCosts.LandscapeMaintMin
        End If

        If HomeStatistics.HomeType = "PUD" Or HomeStatistics.HomeType = "Condo" Then  'pay HO Dues and zero out outside costs covered by the asscociation CC&R's
            HOCT = ReversionCosts.HODues
            LSCT = 0
            PVCT = 0
            LMCT = 0
            ISCT = 0
        Else
            HOCT = 0
        End If

        ReversionaryCosts = CLCT + HVCT + IPCT + FCCT + WCCT + LSCT + LMCT + DMCT + PVCT + ISCT + HOCT + UTCT

        Dim ReversArray(ProgramParameters_FMV.ProjectedLife_t + +20) As Double '= Nothing                                         ReversArray()   
        Dim ResaleArray(ProgramParameters_FMV.ProjectedLife_t + +20) As Double '= Nothing                                         ResaleArray()
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Dim ReversArray5(ProgramParameters_FMV.ProjectedLife_t + +20) As Double
        Dim ReversArray4(ProgramParameters_FMV.ProjectedLife_t + +20) As Double
        Dim ReversArray3(ProgramParameters_FMV.ProjectedLife_t + +20) As Double
        Dim ReversArray2(ProgramParameters_FMV.ProjectedLife_t + +20) As Double
        Dim ReversArray1(ProgramParameters_FMV.ProjectedLife_t + +20) As Double
        Dim ReversArrayUponSale(ProgramParameters_FMV.ProjectedLife_t + +20) As Double
        ReversArray(0) = 0

        If RM.RMTI = 1 Then

            For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t
                ReversArray5(t) = RM.Revers5 * FMVt(t)
                ReversArray4(t) = RM.Revers4 * FMVt(t)
                ReversArray3(t) = RM.Revers3 * FMVt(t)
                ReversArray2(t) = RM.Revers2 * FMVt(t)
                ReversArray1(t) = RM.Revers1 * FMVt(t)
                ReversArrayUponSale(t) = RM.ReversUponSale * FMVt(t)
                ReversArray(t) = ReversArray(t) + ReversArray5(t) + ReversArray4(t) + ReversArray3(t) + ReversArray2(t) + ReversArray1(t) + ReversArrayUponSale(t)

                ResaleArray(t) = 0
            Next

        Else
            ''''''''''''''''''''''''''''''''''
            For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t
                ReversArray(t) = ReversionaryCosts * (1 + (HomeStatistics.CPI)) ^ (t / 12)
                ResaleArray(t) = Resale.LegalClosing + (Resale.CommissPct + Resale.MarketingPct) * FMVt(t) + (Resale.TitlePctPerThous + Resale.TransTaxPctPerThous + Resale.EscrClosingPctPerThous) * FMVt(t) / 1000

                If t = 499 Then
                    Dim HUB As Double = 0
                End If
            Next
        End If
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'MAINTENANCE ARRAY
        Dim MyHome As New myhome 'create a new instance of home inherits default values
        'Dim YourHome As New myhome 'creates another object with the same values and default

        MyHome.HomeType = HomeTypes.DetachedSingle '-- assigns a home type that overrides
        If MyHome.HomeType = HomeTypes.Duplex Then
            MyHome.PrevMaintM *= 2
        End If

        'Create a list of repair and maintenance values
        Dim MaintExp(ProgramParameters_FMV.ProjectedLife_t + 6)
        Dim PMaint(ProgramParameters_FMV.ProjectedLife_t + 6)
        Dim AppraisalC(ProgramParameters_FMV.ProjectedLife_t + 6)
        PMaint(0) = 0
        AppraisalC(0) = 0

        Dim AppraisalLast As Double = 0

        Dim MaintTotal(ProgramParameters_FMV.ProjectedLife_t + 1) As Double
        Dim MaintTotallast As Double = 0
        MaintExp(0) = 0

        For t As Integer = 1 To ProgramParameters_FMV.ProjectedLife_t + 1                   'ProjectedLife_t = 500
            If t Mod 15 = 0 Then
                PMaint(t) = MyHome.PrevMaintM * (1 + (MyHome.CPI)) ^ (t / 12)
                AppraisalC(t) = MyHome.AppraisalCostsM * (1 + (MyHome.CPI)) ^ (t / 12)
                AppraisalLast = AppraisalC(t)
            Else
                PMaint(t) = 0
                AppraisalC(t) = 0
            End If
            If t Mod RepairMaint.AppraisalPerYrs = 0 Then
                AppraisalC(t) = AppraisalLast * (1 + MyHome.CPI) ^ (t / 12)
            End If
            ''''''''''''''''''''''
            If RM.RMTI = 1 Then

                If t < RM.MaintBeginMo Then
                    MaintExp(t) = 0
                End If

                If t Mod RM.MaintInterval <> 0 Then
                    MaintExp(t) = 0
                End If

                If t = RM.MaintInterval Then
                    MaintExp(t) = RM.MaintInitial
                End If

                If t > RM.MaintBeginMo Then

                    If t Mod RM.MaintInterval = 0 Then
                        MaintExp(t) = RM.MaintInitial * RM.MaintIncreaseFactor ^ ((t - RM.MaintInterval) / RM.MaintInterval)
                    End If

                End If

            Else
                ''''''''''''''''''''''''''''''''''

                'MaintExp
                MaintExp(t) = (PMaint(t) + AppraisalC(t)) * MyHome.REFactor ' Creates Maintenance Array, REFactor = 1 to compute Maint, or 0 to Zero Out maintenance for subsequent transaction calculations
            End If
            '''''''''''''''''''''''''''''''
            MaintTotallast = MaintTotallast + MaintExp(t)
            MaintTotal(t) = MaintTotallast

            If t = 499 Then
                Dim HUB As Double = 0
            End If

        Next

        'Repairs Array
        Dim RepairsExp(ProgramParameters_FMV.ProjectedLife_t + 6)
        Dim RepairsTotal(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim RepairsTotallast As Double = 0
        RepairsExp(0) = 0

        Dim AddMonths As Integer
        AddMonths = DateDiff("m", DatesAll.DateRevalLast, DatesAll.DateRevalNew)  'Add months from Inception to current date to revalue all Reserves

        Dim EP(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim EL(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim RF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim PL(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim ST(ProgramParameters_FMV.ProjectedLife_t + 6) As Double

        For t As Integer = 1 To ProgramParameters_FMV.ProjectedLife_t
            If (t - RM.EPBal + AddMonths) Mod RM.EPBal = 0 Then
                EP(t) = RM.EPCT * (1 + (HomeStatistics.CPI)) ^ (t / 12)
            Else
                EP(t) = 0
            End If

            If (t - RM.ELBal + AddMonths) Mod RM.ELLife = 0 Then
                EL(t) = RM.ELCT * (1 + (HomeStatistics.CPI)) ^ (t / 12)
            Else
                EL(t) = 0
            End If

            If (t - RM.RFBal + AddMonths) Mod RM.RFBal = 0 Then
                RF(t) = RM.RFCT * (1 + (HomeStatistics.CPI)) ^ (t / 12)
            Else
                RF(t) = 0
            End If

            If (t - RM.PLBal + AddMonths) Mod RM.PLLife = 0 Then
                PL(t) = RM.PLCT * (1 + (HomeStatistics.CPI)) ^ (t / 12)
            Else
                PL(t) = 0
            End If

            If (t - RM.STBal + AddMonths) Mod RM.STLife = 0 Then
                ST(t) = RM.STCT * (1 + (HomeStatistics.CPI)) ^ (t / 12)
            Else
                ST(t) = 0
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If RM.RMTI = 1 Then

                If t < RM.RepairBeginMo Then
                    RepairsExp(t) = 0
                End If

                If t = RM.RepairBeginMo Then
                    RepairsExp(t) = RM.RepairInitial
                End If

                If t > RM.RepairBeginMo Then
                    If t Mod RM.RepairInterval = 1 Then
                        RepairsExp(t) = RM.RepairInitial * RM.RepairIncreaseFactor ^ ((t - RM.RepairBeginMo) / RM.RepairInterval)
                    End If

                    If t Mod RM.RepairInterval <> 1 Then
                        RepairsExp(t) = 0
                    End If

                End If

            Else
                ''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                RepairsExp(t) = 0 '(EP(t) + EL(t) + RF(t) + PL(t) + ST(t)) * MyHome.REFactor                               'RepairsExp
            End If

            RepairsTotallast = RepairsTotallast + RepairsExp(t)
            RepairsTotal(t) = RepairsTotallast

            If t = 499 Then
                Dim HUB As Double = 0
            End If

        Next

        'Taxes and Insurance Array

        Dim TaxesExp(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim TaxesTotal(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim TaxesTotalLast As Double = 0
        Dim InsurExp(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim InsurTotal(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim InsurTotalLast As Double = 0

        Dim TaxMo As Double                                                                                 'taxes increased by Prop13 growth

        'TAXES
        If TaxInsur.TaxInterval > 6 Then
            TaxMo = (TaxInsur.TaxAsPctOfFMV * FMVt(t) / 2) * TaxInsur.TaxAnnualIncrease                 'Second half Payment
        Else
            TaxMo = (TaxInsur.TaxAsPctOfFMV * FMVt(t) / 2)                                              'First half Payment
        End If



        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t

            If t Mod 12 = TaxInsur.TaxBegMo Then
                TaxesExp(t) = TaxMo
            ElseIf t Mod 12 = TaxInsur.TaxBegMo + TaxInsur.TaxInterval Then
                TaxesExp(t) = TaxMo * TaxInsur.TaxAnnualIncrease
            Else
                TaxesExp(t) = 0
            End If

            TaxesExp(t) = TaxesExp(t) * TaxInsur.TaxAnnualIncrease ^ Int(t / 12)                       'TaxesExp

            TaxesTotalLast = TaxesTotalLast + TaxesExp(t)
            TaxesTotal(t) = TaxesTotalLast

            'INSUR
            If t Mod 12 = TaxInsur.InsurBegMo Then
                InsurExp(t) = TaxInsur.InsurPctOfGSP * FMVt(t)                                          ' InsurExp
            Else
                InsurExp(t) = 0
            End If

            InsurTotalLast = InsurTotalLast + InsurExp(t)
            InsurTotal(t) = InsurTotalLast

        Next

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'HUB

        Dim tTime As Double = 0
        Dim RescDeath As Double = 0
        Dim CCancel As Double = 0
        Dim LEPurch As Double = 0
        Dim ReversDeath As Double = 0
        Dim tTotal As Double = 0
        Dim TurnOverInYears As Double = 0

        Dim AtT(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                                    'AtT - Option Action Rate (Cancellation or Tender)
        Dim qt(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                                     'qt - Probability of Dying Between t and t+1 Given Survival to t
        Dim ftR(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                                    'ftR - Probability of Rescissionary Death Between t and t+1
        Dim ftC(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                                    'ftC - Probability of Cancellation Between t and t+1
        Dim ftP(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                                    'ftP - Probability of Life Estate Purchase Between t and t+1
        Dim ftD(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                                    'ftD - Probability of Reversionary Death Between t and t+1
        Dim Stt(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                                    'Stt - Probability of Being Present at t
        Dim GSPt(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                                   'GSPt - Gross Sales Proceeds from t to t+1
        Dim EtM(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                                    'EtM - Maintenance Expense
        Dim EtR(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                                    'EtR - Repair Expense
        Dim EtT(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                                    'EtT - Taxes Expense
        Dim EtI(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                                    'EtI - Insurance Expense
        Dim ftiD(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                                   'ftiD - Probability of Death Between t and t+1 if No Tender
        Dim Sti(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                                    'Sti - Probability of Being Present at t if No Tender


        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t             'ProjectedLife_t = 500
            AtT(t) = OptionActionRate(t)
            qt(t) = ProbLastDeath(t)
            If t < 6 Then
                ftR(t) = ProbNonTenderDeathBTween(t) + ProbTenderDeathBTween(t)
            Else
                ftR(t) = 0
            End If
            ftC(t) = CancelProbBTween(t)
            ftP(t) = ProbLEBuyBTween(t)
            If t < 6 Then
                ftD(t) = 0
            Else
                ftD(t) = ProbNonTenderDeathBTween(t) + ProbTenderDeathBTween(t)
            End If
            Stt(t) = ProbNonTenderPresent(t) + ProbPresentTenderT1(t) + ProbPresentTenderT2(t) + ProbPresentTenderT3(t) + ProbPresentTenderT4(t) + ProbPresentTenderT5(t) + ProbPresentTenderT6(t)
            GSPt(t) = FMVt(t)
            EtM(t) = MaintExp(t)
            EtR(t) = RepairsExp(t)
            EtT(t) = TaxesExp(t)
            EtI(t) = InsurExp(t)
            ftiD(t) = ProbDeathBT(t)
            Sti(t) = ProbBeingPresent(t)

            tTime = tTime + t
            RescDeath = RescDeath + ftR(t)
            CCancel = CCancel + ftC(t)
            LEPurch = LEPurch + ftP(t)
            ReversDeath = ReversDeath + ftD(t)
        Next

        tTotal = RescDeath + CCancel + LEPurch + ReversDeath
        TurnOverInYears = ((tTime * RescDeath * CCancel * LEPurch * ReversDeath) + 6 * (LEPurch + ReversDeath) + 0.5) / 12

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'ENGINE

        Dim MoneyFactor(ProgramParameters_FMV.ProjectedLife_t + 13) As Double                    'vt  - Time Value of Money Factor

        Dim PRemaingPresentGP(ProgramParameters_FMV.ProjectedLife_t + 6) As Double          'pt  - Probability of Remaining Present to t+1 Given Present at t
        Dim PRescDeathGP(ProgramParameters_FMV.ProjectedLife_t + 6) As Double               'qtR - Probability of Rescissionary Death Between t and t+1 Given Present at t
        Dim PCancelGP(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                  'qtC - Probability of Cancellation Between t and t+1 Given Present at t
        Dim PLEPurchGP(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                 'qtP - Probability of Life Estate Purchase Between t and t+1 Given Present at t
        Dim PReversGP(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                  'qtD - Probability of Reversionary Death Between t and t+1 Given Present at t
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'See HUB																			See HUB
        Dim PPresent(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                   'St  - Probability of Being Present at t                     
        Dim PRescDeathBT(ProgramParameters_FMV.ProjectedLife_t + 6) As Double               'ftR - Probability of Rescissionary Death Between t and t+1
        Dim PCancelBT(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                  'ftC - Probability of Cancellation Between t and t+1
        Dim PLEPurchBT(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                 'ftP - Probability of Life Estate Purchase Between t and t+1
        Dim PReversBT(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                  'ftD - Probability of Reversionary Death Between t and t+1
        Dim PBeingPresentNT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double            'St' - Probability of Being Present at t if No Tender
        Dim PDeathBTNT(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                 'ft'D - Probability of Death Between t and t+1 if No Tender

        PBeingPresentNT(507) = 0
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim PRemaingPresentGPNT(ProgramParameters_FMV.ProjectedLife_t + 6) As Double        'pt' - Probability of Remaining Present to t+1 Given Present at t if No Tender
        Dim PReversGPNT(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                'qt' - Probability of Reversionary Death Between t and t+1 Given Present at t if No Tender

        Dim PRDBT As Double = 0                                                             'ftR total Rescisionary Deaths
        Dim PCBT As Double = 0                                                              'ftC total Cancellations
        Dim PLEP As Double = 0                                                              'ftP total Life Estate Purchases
        Dim PRBT As Double = 0                                                              'ftD total Reversionsary Deaths
        Dim TotMort As Double = 0                                                           'Total of all options probabilitis should all = 1
        Dim PBP As Double = 0                                                               'St  Probability of Being Present at t

        Dim MaintEx(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                    'EtM - Maintenance Expense at t
        Dim RepairEx(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                   'EtR - Repair Expense at t
        'Dim TaxesEx(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                    'EtT - Taxes Expense at t
        'Dim InsurEx(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                    'EtI - Insurance Expense at t

        Dim FMValue(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                        'FMV
        Dim ReversEx(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                   'Reversionary Expenses			
        Dim ResaleEx(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                   'Resale Expenses
        Dim NetSalesProceeds(ProgramParameters_FMV.ProjectedLife_t + 6) As Double           'Net Sale Proceeds

        Dim MoEstateMgtFees(ProgramParameters_FMV.ProjectedLife_t + 6) As Double            'FtE - Monthly Estate Management Fees Paid to Manager at t
        Dim MoPmtMgtFees(ProgramParameters_FMV.ProjectedLife_t + 6) As Double               'FtP - Monthly Payment Management Fees Paid to Manager at t
        Dim MoEstateMgtFeesCum(ProgramParameters_FMV.ProjectedLife_t + 6) As Double         'FtET - Cumulative Monthly Estate Management Fees Paid to Manager at t
        Dim MoPmtMgtFeesCum(ProgramParameters_FMV.ProjectedLife_t + 6) As Double            'FtPT - Cumulative Monthly Payment Management Fees Paid to Manager at t

        'BALLOON MONTHLY PAYMENT ALCI NOTE CALC
        Dim MoPmtIndicator(ProgramParameters_FMV.ProjectedLife_t + 6) As Double             'PtMi - Monthly Payment Indicator
        Dim MoPmtCumCoeff(ProgramParameters_FMV.ProjectedLife_t + 6) As Double              'PtMt - Monthly Payment Cumulative Coefficient
        Dim BalloonPmtIndicator(ProgramParameters_FMV.ProjectedLife_t + 6) As Double        'PtBi - Balloon Payment Indicator
        Dim BalloonPmtCoeff(ProgramParameters_FMV.ProjectedLife_t + 6) As Double            'PtBt - Balloon Payment Cumulative Coefficient

        Dim MgtFeesCancellation(ProgramParameters_FMV.ProjectedLife_t + 6) As Double        'FtC - Manager Fee on Cancellation Fee From t to t+1
        Dim MgrResaleFee(ProgramParameters_FMV.ProjectedLife_t + 6) As Double               'FtS - Sales Fee Paid to Manager Upon Reversion to Sale  for Transactions From t to t+1

        Dim GrossGainReversPurch(ProgramParameters_FMV.ProjectedLife_t + 6) As Double       'GtP - Gross Gain Upon Reversion to Sale Due to Purchase for Transactions From t to t+1
        Dim GrossGainReversDeath(ProgramParameters_FMV.ProjectedLife_t + 6) As Double       'GtD - Gross Gain Upon Reversion to Sale Due to Death for Transactions From t to t+1
        Dim FeeMgrReversPurch(ProgramParameters_FMV.ProjectedLife_t + 6) As Double          'FtGP - Fee Paid to Manager Upon Reversion to Sale Due to Purchase for Transactions From t to t+1
        Dim FeeMgrReversDeath(ProgramParameters_FMV.ProjectedLife_t + 6) As Double          'FtGD - Fee Paid to Manager Upon Reversion to Sale Due to Death for Transactions From t to t+1

        Dim MaintExCum(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                 'EtMT - Cumulative Maintenance Expense at t
        Dim RepairExCum(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                'EtRT - Cumulative Repair Expense at t
        Dim TaxesExCum(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                 'EtTT - Cumulative Taxes Expense at t
        Dim InsurExCum(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                 'EtIT - Cumulative Insurance Expense at t
        'Dim ReversCum(ProgramParameters_FMV.ProjectedLife_t + 6) As Double 

        MoEstateMgtFeesCum(0) = 0
        MoPmtMgtFeesCum(0) = 0

        MoPmtIndicator(0) = 0
        MoPmtCumCoeff(0) = 0
        BalloonPmtIndicator(0) = 0
        BalloonPmtCoeff(0) = 0

        MaintExCum(0) = 0
        RepairExCum(0) = 0
        TaxesExCum(0) = 0
        InsurExCum(0) = 0
        'ReversCum(0) = 0

        Dim BEofGSP(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                    'GSPtBE - Bequest Equity Portion of Gross Sales Proceeds for Transactions From t to t+1
        Dim CEofGSP(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                    'GSPtCE - Conserved Equity Portion of Gross Sales Proceeds for Transactions From t to t+1
        Dim MGRofGSP(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                   'GSPtMS - Manager's Share of Gross Sales Proceeds for Transactions From t to t+1
        Dim INVESTORofGSP(ProgramParameters_FMV.ProjectedLife_t + 6) As Double              'GSPtIS - Investor's Share of Gross Sales Proceeds for Transactions From t to t+1

        'EXPENDITURES
        Dim CancellationCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double         'C - Cancellation Fee From t to t+1
        Dim MaintCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                    'StEtM - Maintenance Expense Cash Flow at t
        Dim RepairsCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                  'StEtR - Repair Expense  Cash Flow at t
        Dim TaxesCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                    'StEtT - Taxes Expense  Cash Flow at t
        Dim InsurCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                    'StEtI - Insurance Expense  Cash Flow at t
        Dim ReversPurchCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double              '∑ft-6+kPEtSk - Reversionary Expenses Due to Purchase Cash Flow for t to t+1
        Dim ReversDeathCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double              'Includes Revers and Resale above
        Dim ResalePurchCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double              '∑ft-6+kDEtSk - Reversionary Expenses Due to Death Cash Flow for t to t+1
        Dim ResaleDeathCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double              'Includes Revers and Resale above

        Dim LEValue(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                    'LEPt - Life Estate from t to t+1
        LEValue(0) = 0                                                                      'FIX THE .. CREATE LEVALUE FUNCTION AND PULL ARRAY FROM FUCNTION

        'MANAGER FEES
        Dim MoEstMgtFeesCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double             'StFtE - Monthly Estate Management Fees Paid to Manager Cash Flow for t to t+1
        Dim MoPmtsMgtFeesCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double            'S'tFtP - Monthly Payment Management Fees Paid to Manager Cash Flow for t to t+1
        Dim CancellationFeesCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double         'ftCFtC - Manager Fee on Cancellation Fee Cash Flow From t to t+1
        Dim SalesFeesPurchReversCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double     'ft-6PFtS - Sales Fee Paid to Manager Due to Purchase Upon Reversion to Sale  Cash Flow From t to t+1
        Dim SalesFeesDeathReversCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double     'ft-6DFtS - Sales Fee Paid to Manager Due to Death Upon Reversion to Sale  Cash Flow From t to t+1
        Dim MgrFeesPurchReversCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double       'ft-6PFtGP - Fee Paid to Manager Upon Reversion to Sale Due to Purchase Cash Flow From t to t+1
        Dim MgrFeesDeathReversCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double       'ft-6DFtGD - Fee Paid to Manager Upon Reversion to Sale Due to Death Cash Flow From t to t+1


        'PAYMENTS AND LIFE ESTATE VALUE
        Dim At(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                         'At
        Dim Bsub1(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                      'Bsub1
        Dim Bsub2(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                      'Bsub2
        Dim Bt(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                         'Bt
        Dim Ct(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                         'Ct
        Dim Dsub1(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                      'Dsub1
        Dim Dsub2(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                      'Dsub2
        Dim Dt(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                         'Dt
        Dim Esub1(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                      'Esub1
        Dim Esub2(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                      'Esub2
        Dim Et(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                         'Et

        Dim tN(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                         '▼tN
        Dim tI(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                         '▼tI
        Dim tM(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                         '▼tM
        Dim tB(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                         '▼tB

        Dim sub1(ProgramParameters_FMV.ProjectedLife_t + 15) As Double                       'sub1
        Dim sub2(ProgramParameters_FMV.ProjectedLife_t + 15) As Double                       'sub2
        Dim sub3(ProgramParameters_FMV.ProjectedLife_t + 15) As Double                       'sub3
        Dim sub4(ProgramParameters_FMV.ProjectedLife_t + 15) As Double                       'sub4
        Dim sub5(ProgramParameters_FMV.ProjectedLife_t + 15) As Double                       'sub5
        Dim sub6(ProgramParameters_FMV.ProjectedLife_t + 15) As Double                       'sub6

        Dim sub7(ProgramParameters_FMV.ProjectedLife_t + 15) As Double                      'sub7
        Dim sub8(ProgramParameters_FMV.ProjectedLife_t + 15) As Double                      'sub8

        Dim APVLEP(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                     'APVt(LEP)
        Dim LEP(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                        'LEPt

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''		
        Dim MoneyFactor0 As Double = 0
        MoneyFactor0 = 1 / (1 + ManagerFees.AnnualIRR) ^ (1 / 12)

        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t + 13

            If t = 0 Then                                                                           'vt
                MoneyFactor(t) = MoneyFactor0 ^ 0.5
            Else
                MoneyFactor(t) = MoneyFactor(t - 1) * MoneyFactor0
            End If

        Next

        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t + 6

            'PROBABILITES''''''''''''''''''''''''''''''''''''''''''''''''''''''etso

            'St  - Probability of Being Present at t
            PPresent(t) = Stt(t)

            'ftR - Probability of Rescissionary Death Between t and t+1
            PRescDeathBT(t) = ftR(t)

            'ftC - Probability of Cancellation Between t and t+1
            PCancelBT(t) = ftC(t)

            'ftP - Probability of Life Estate Purchase Between t and t+1
            PLEPurchBT(t) = ftP(t)

            'ftD - Probability of Reversionary Death Between t and t+1
            PReversBT(t) = ftD(t)

            'St' - Probability of Being Present at t if No Tender
            PBeingPresentNT(t) = Sti(t)

            'ft'D - Probability of Death Between t and t+1 if No Tender
            PDeathBTNT(t) = ftiD(t)

            'qt' - Probability of Reversionary Death Between t and t+1 Given Present at t if No Tender
            If PBeingPresentNT(t) = 0 Then
                PReversGPNT(t) = 0
            Else
                PReversGPNT(t) = PDeathBTNT(t) / PBeingPresentNT(t)
            End If
        Next

        For t = ProgramParameters_FMV.ProjectedLife_t + 6 To 0 Step -1
            'pt' - Probability of Remaining Present to t+1 Given Present at t if No Tender
            If PBeingPresentNT(t) = 0 Then
                PRemaingPresentGPNT(t) = 0
            Else
                PRemaingPresentGPNT(t) = PBeingPresentNT(t + 1) / PBeingPresentNT(t)
            End If
        Next

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        'PAYMENTS TO MANAGER

        'Dim MgrResaleFee(ProgramParameters_FMV.ProjectedLife_t + 6) As Double

        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t
            If t = 0 Then
                MoEstateMgtFees(t) = 0
                MoEstateMgtFeesCum(t) = 0

                MoPmtMgtFees(t) = 0
                MoPmtMgtFeesCum(t) = 0

                MoPmtIndicator(t) = 0
                MoPmtCumCoeff(t) = 0

                BalloonPmtIndicator(t) = 0
                BalloonPmtCoeff(t) = 0
            Else
                MoEstateMgtFees(t) = ManagerFees.MoPmtMgt                                               'FtE - Monthly Estate Management Fees Paid to Manager at t - For ALCI Note
                MoEstateMgtFeesCum(t) = MoEstateMgtFeesCum(t - 1) + MoEstateMgtFees(t)                  'FtET - Cumulative Monthly Estate Management Fees Paid to Manager at t

                If t > ContractTerms.NumMoPmts Then
                    MoPmtIndicator(t) = 0
                Else
                    MoPmtIndicator(t) = 1
                    'NEEDS TO LOOKUP ContractTerms.NumMoPmts  FROM IRS TABLE - Reg72 - HUB I8
                    'FtPT - Cumulative Monthly Payment Management Fees Paid to Manager at t
                End If                                                                                  '
                MoPmtMgtFees(t) = ManagerFees.MoEStateMgt * MoPmtIndicator(t)                           'PtMi - Monthly Payment Indicator
                MoPmtMgtFeesCum(t) = MoPmtMgtFeesCum(t - 1) + MoPmtMgtFees(t)                           'FtPT - Monthly Payment Management Fees Paid to Manager at t

                MoPmtCumCoeff(t) = MoPmtCumCoeff(t - 1) + MoPmtIndicator(t)                             'PtMt - Monthly Payment Cumulative Coefficient

                If t = ContractTerms.NumMoPmts + 1 Then
                    BalloonPmtIndicator(t) = 1                                                          'PtBi - Balloon Payment Indicator
                Else
                    BalloonPmtIndicator(t) = 0
                End If

                BalloonPmtCoeff(t) = BalloonPmtCoeff(t - 1) + BalloonPmtIndicator(t)                    'PtBt - Balloon Payment Cumulative Coefficient

            End If

            MgtFeesCancellation(t) = ManagerFees.CancellationFee * ManagerFees.CancellationPct          'FtC - Manager Fee on Cancellation Fee From t to t+1
            MgrResaleFee(t) = FMVt(t) * ManagerFees.PostReversGSP                                       'FtS - Sales Fee Paid to Manager Upon Reversion to Sale  for Transactions From t to t+1
        Next


        'NET SALES PROCEEDS

        Dim ReversPurchCF5(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim ReversPurchCF4(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim ReversPurchCF3(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim ReversPurchCF2(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim ReversPurchCF1(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim ReversPurchCFUponSale(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim ResalePurchCFUponSale(ProgramParameters_FMV.ProjectedLife_t + 6) As Double

        Dim ReversDeathCF5(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim ReversDeathCF4(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim ReversDeathCF3(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim ReversDeathCF2(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim ReversDeathCF1(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim ReversDeathCFUponSale(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim ResaleDeathCFUponSale(ProgramParameters_FMV.ProjectedLife_t + 6) As Double

        Dim APVRever5(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim APVRever4(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim APVRever3(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim APVRever2(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim APVRever1(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        Dim APVReverTrans(ProgramParameters_FMV.ProjectedLife_t + 6) As Double

        'Dim FeeMgrReversPurch(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        'Dim FeeMgrReversDeath(ProgramParameters_FMV.ProjectedLife_t + 6) As Double

        'Dim SalesFeesPurchReversCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        'Dim SalesFeesDeathReversCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        'Dim MgrFeesPurchReversCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double
        'Dim MgrFeesDeathReversCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double

        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t

            FMValue(t) = FMVt(t)
            ReversEx(t) = ReversArray(t)
            ResaleEx(t) = ResaleArray(t)
            NetSalesProceeds(t) = FMValue(t) - ReversEx(t) - ResaleEx(t)

            'EXPENSES'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'MaintExp(t) 
            MaintExCum(t) = MaintTotal(t)                                                           'EtMT - Cumulative Maintenance Expense at t

            'RepairExp(t)
            RepairExCum(t) = RepairsTotal(t)                                                        'EtRT - Cumulative Repair Expense at t

            If t > 0 Then
                TaxesExCum(t) = TaxesExCum(t - 1) + TaxesExp(t)                                    'EtTT - Cumulative Taxes Expense at t

                InsurExCum(t) = InsurExCum(t - 1) + InsurExp(t)                                   'EtIT - Cumulative Insurance Expense at t
            End If

            'HUB proof totals''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            PRDBT = PRDBT + PRescDeathBT(t)            '  .01630   Y
            PCBT = PCBT + PCancelBT(t)                 '  .01577   Y
            PLEP = PLEP + PLEPurchBT(t)                '  .199629  very close
            PRBT = PRBT + PReversBT(t)                 '  .76830   Close

            PBP = PBP + ProbNonTenderPresent(t) + (ProbPresentTenderT1(t) + ProbPresentTenderT2(t) + ProbPresentTenderT3(t) + ProbPresentTenderT4(t) + ProbPresentTenderT5(t) + ProbPresentTenderT6(t))
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'CASH PAYOUTS ON DEATH, TERMINATION EXPIRATION, CANCELLATION, RESCISSION

            Dim SumGSPCE(ProgramParameters_FMV.ProjectedLife_t) As Double
            SumGSPCE(5) = 0
            For z = 6 To ProgramParameters_FMV.ProjectedLife_t
                SumGSPCE(z) = (1 + ManagerFees.MonthlyRepChrgToEstate) ^ (z + 0.5 - (z - 6)) * RepairEx(z - 6)
                SumGSPCE(z) = SumGSPCE(z) - SumGSPCE(z - 1)
            Next

            BEofGSP(t) = Math.Min(GSPt(t), ContractTerms.BequestEquity)                                     'GSPtBE - Bequest Equity Portion of Gross Sales Proceeds for Transactions From t to t+1
            tTotal = tTotal + 1                                                                         'total of t

            If t < 6 Then                                                                                   'GSPtCE - Conserved Equity Portion of Gross Sales Proceeds for Transactions From t to t+1					
                CEofGSP(t) = Math.Max(0, ContractTerms.ConservedEquityPct * (GSPt(t) - BEofGSP(t) - ReversEx(t) - ResaleEx(t) - MgrResaleFee(t)))
            Else                                                                                            '
                CEofGSP(t) = Math.Max(0, ContractTerms.ConservedEquityPct * (GSPt(t) - BEofGSP(t) - ReversEx(t) - ResaleEx(t) - MgrResaleFee(t)) - SumGSPCE(t))
            End If

            'GSPtMS - Manager's Share of Gross Sales Proceeds for Transactions From t to t+1
            MGRofGSP(t) = Math.Max(0, (FMVt(t) - (BEofGSP(t) + CEofGSP(t) + ReversEx(t) + ResaleEx(t) + MgrResaleFee(t))) * ManagerFees.EquityShare - ManagerFees.AdvEquityAmt)

            'GSPtIS - Investor's Share of Gross Sales Proceeds for Transactions From t to t+1
            INVESTORofGSP(t) = FMVt(t) - BEofGSP(t) - CEofGSP(t) - MGRofGSP(t)

            CancellationCF(t) = ManagerFees.CancellationFee
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'EXPENSES * PROBABILITIES

            MaintCF(t) = MaintExp(t) * PPresent(t)                                                          'StEtM - Maintenance Expense Cash Flow at t
            RepairsCF(t) = RepairsExp(t) * PPresent(t)                                                      'StEtR - Repair Expense  Cash Flow at t
            TaxesCF(t) = TaxesExp(t) * PPresent(t)                                                          'StEtT - Taxes Expense  Cash Flow at t
            InsurCF(t) = InsurExp(t) * PPresent(t)                                                          'StEtI - Insurance Expense  Cash Flow at t

            If RM.RMTI = 1 Then

                If t > 0 Then
                    ReversPurchCF5(t) = ReversArray5(t + 6) * PLEPurchBT(t - 1)
                End If
                If t > 1 Then
                    ReversPurchCF4(t) = ReversArray4(t + 6) * PLEPurchBT(t - 2)
                End If
                If t > 2 Then
                    ReversPurchCF3(t) = ReversArray3(t + 6) * PLEPurchBT(t - 3)
                End If
                If t > 3 Then
                    ReversPurchCF2(t) = ReversArray2(t + 6) * PLEPurchBT(t - 4)
                End If
                If t > 4 Then
                    ReversPurchCF1(t) = ReversArray1(t + 6) * PLEPurchBT(t - 5)
                End If
                If t > 5 Then
                    ReversPurchCFUponSale(t) = ReversArrayUponSale(t) * PLEPurchBT(t - 6)
                End If

                ReversPurchCF(t) = ReversPurchCF5(t) + ReversPurchCF4(t) + ReversPurchCF3(t) + ReversPurchCF2(t) + ReversPurchCF1(t) + ReversPurchCFUponSale(t)
                ResalePurchCFUponSale(t) = ResaleEx(t) * PLEPurchBT(t)

                If t > 0 Then
                    ReversDeathCF5(t) = ReversArray5(t + 6) * PReversBT(t - 1)
                End If
                If t > 1 Then
                    ReversDeathCF4(t) = ReversArray4(t + 6) * PReversBT(t - 2)
                End If
                If t > 2 Then
                    ReversDeathCF3(t) = ReversArray3(t + 6) * PReversBT(t - 3)
                End If
                If t > 3 Then
                    ReversDeathCF2(t) = ReversArray2(t + 6) * PReversBT(t - 4)
                End If
                If t > 4 Then
                    ReversDeathCF1(t) = ReversArray1(t + 6) * PReversBT(t - 5)
                End If
                If t > 5 Then
                    ReversDeathCFUponSale(t) = ReversArrayUponSale(t) * PReversBT(t - 6)
                End If

                ReversDeathCF(t) = ReversDeathCF5(t) + ReversDeathCF4(t) + ReversDeathCF3(t) + ReversDeathCF2(t) + ReversDeathCF1(t) + ReversDeathCFUponSale(t)
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Else

                ReversDeathCF(t) = ReversEx(t) * PReversBT(t)
                ResaleDeathCF(t) = ResaleEx(t) * PReversBT(t)

            End If

            MoEstMgtFeesCF(t) = MoEstateMgtFees(t) * PPresent(t)                                            'StFtE - Monthly Estate Management Fees Paid to Manager Cash Flow for t to t+1
            MoPmtsMgtFeesCF(t) = MoPmtMgtFees(t) * PBeingPresentNT(t)                                       'S'tFtP - Monthly Payment Management Fees Paid to Manager Cash Flow for t to t+1
            CancellationFeesCF(t) = MgtFeesCancellation(t) * PCancelBT(t)                                   'ftCFtC - Manager Fee on Cancellation Fee Cash Flow From t to t+1

            'FtGP - Fee Paid to Manager Upon Reversion to Sale Due to Purchase for Transactions From t to t+1
            FeeMgrReversPurch(t) = ManagerFees.PostReversNP * GrossGainReversPurch(t)

            'FtGD - Fee Paid to Manager Upon Reversion to Sale Due to Death for Transactions From t to t+1
            FeeMgrReversDeath(t) = ManagerFees.PostReversNP * GrossGainReversDeath(t)

            If t < 6 Then
                SalesFeesPurchReversCF(t) = 0                                                               'ft-6PFtS - Sales Fee Paid to Manager Due to Purchase Upon Reversion to Sale  Cash Flow From t to t+1
                SalesFeesDeathReversCF(t) = 0                                                               'ft-6DFtS - Sales Fee Paid to Manager Due to Death Upon Reversion to Sale  Cash Flow From t to t+1
                MgrFeesPurchReversCF(t) = 0                                                                 'ft-6PFtGP - Fee Paid to Manager Upon Reversion to Sale Due to Purchase Cash Flow From t to t+1
                MgrFeesDeathReversCF(t) = 0                                                                 'ft-6DFtGD - Fee Paid to Manager Upon Reversion to Sale Due to Death Cash Flow From t to t+1
            ElseIf t > 5 Then
                SalesFeesPurchReversCF(t) = MgrResaleFee(t + 6) * PLEPurchBT(t)                               'ft-6PFtS - Sales Fee Paid to Manager Due to Purchase Upon Reversion to Sale  Cash Flow From t to t+1
                SalesFeesDeathReversCF(t) = MgrResaleFee(t + 6) * PReversBT(t)                                'ft-6DFtS - Sales Fee Paid to Manager Due to Death Upon Reversion to Sale  Cash Flow From t to t+1
                MgrFeesPurchReversCF(t) = FeeMgrReversPurch(t + 6) * ProbLEBuyBTween(t)       'FIX!                 'ft-6PFtGP - Fee Paid to Manager Upon Reversion to Sale Due to Purchase Cash Flow From t to t+1
                MgrFeesDeathReversCF(t) = FeeMgrReversDeath(t + 6) * PReversBT(t)        'fIX                  'ft-6DFtGD - Fee Paid to Manager Upon Reversion to Sale Due to Death Cash Flow From t to t+1
            End If
        Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Dim x as integer

        For x = ProgramParameters_FMV.ProjectedLife_t + 6 To 1 Step -1
            If x < 502 Then
                APVRever5(x) = (MoneyFactor0 ^ 0.05) * (ReversPurchCF5(x - 1) + ReversDeathCF5(x - 1)) + MoneyFactor0 * APVRever5(x + 1)
            End If

            If x < 503 Then
                APVRever4(x) = (MoneyFactor0 ^ 0.05) * (ReversPurchCF4(x - 1) + ReversDeathCF4(x - 1)) + MoneyFactor0 * APVRever4(x + 1)
            End If

            If x < 504 Then
                APVRever3(x) = (MoneyFactor0 ^ 0.05) * (ReversPurchCF3(x - 1) + ReversDeathCF3(x - 1)) + MoneyFactor0 * APVRever3(x + 1)
            End If

            If x < 505 Then
                APVRever2(x) = (MoneyFactor0 ^ 0.05) * (ReversPurchCF2(x - 1) + ReversDeathCF2(x - 1)) + MoneyFactor0 * APVRever2(x + 1)
            End If

            If x < 506 Then
                APVRever1(x) = (MoneyFactor0 ^ 0.05) * (ReversPurchCF1(x - 1) + ReversDeathCF1(x - 1)) + MoneyFactor0 * APVRever1(x + 1)
            End If

            If x < 506 Then
                APVReverTrans(x) = (MoneyFactor0 ^ 0.05) * (ReversPurchCFUponSale(x - 1) + ReversDeathCFUponSale(x - 1)) + MoneyFactor0 * APVReverTrans(x + 1)
            End If

        Next

        APVRever5(0) = MoneyFactor0 * APVRever5(1)

        APVRever4(1) = MoneyFactor0 * APVRever4(2)
        APVRever4(0) = MoneyFactor0 * APVRever4(1)

        APVRever3(2) = MoneyFactor0 * APVRever3(3)
        APVRever3(1) = MoneyFactor0 * APVRever3(2)
        APVRever3(0) = MoneyFactor0 * APVRever3(1)

        APVRever2(3) = MoneyFactor0 * APVRever2(4)
        APVRever2(2) = MoneyFactor0 * APVRever2(3)
        APVRever2(1) = MoneyFactor0 * APVRever2(2)
        APVRever2(0) = MoneyFactor0 * APVRever2(1)

        APVRever1(4) = MoneyFactor0 * APVRever1(5)
        APVRever1(3) = MoneyFactor0 * APVRever1(4)
        APVRever1(2) = MoneyFactor0 * APVRever1(3)
        APVRever1(1) = MoneyFactor0 * APVRever1(2)
        APVRever1(0) = MoneyFactor0 * APVRever1(1)

        APVReverTrans(5) = MoneyFactor0 * APVReverTrans(6)
        APVReverTrans(4) = MoneyFactor0 * APVReverTrans(5)
        APVReverTrans(3) = MoneyFactor0 * APVReverTrans(4)
        APVReverTrans(2) = MoneyFactor0 * APVReverTrans(3)
        APVReverTrans(1) = MoneyFactor0 * APVReverTrans(2)
        APVReverTrans(0) = MoneyFactor0 * APVReverTrans(1)
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'CASH FLOW to NAV
        Dim MgrFeeReversPurch(ProgramParameters_FMV.ProjectedLife_t + 7) As Double              'ft-6PFtGP - Fee Paid to Manager Upon Reversion to Sale Due to Purchase Cash Flow From t to t+1
        Dim MgrFeeReversDeath(ProgramParameters_FMV.ProjectedLife_t + 7) As Double              'ft-6DFtGD - Fee Paid to Manager Upon Reversion to Sale Due to Death Cash Flow From t to t+1
        Dim LEPurchBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                      'ftPLEPt - Life Estate Purchase Cash Flow Between t and t+1
        Dim GSPPurchBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                     'ft-6PGSPt - Gross Sales Proceeds Due to Purchase Cash Flow Between t and t+1
        Dim BEGSPPurchBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                   'ft-6PGSPtBE - Bequest Equity Portion of Gross Sales Proceeds Due to Purchase Cash Flow Between t and t+1
        Dim CEGSPPurchBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                   'ft-6PGSPtCE - Conserved Equity Portion of Gross Sales Proceeds Due to Purchase Cash Flow Between t and t+1
        Dim MgrGSPPurchBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                  'ft-6PGSPtMS - Managers Share of Gross Sales Proceeds Due to Purchase Cash Flow Between t and t+1
        Dim InvGSPPurchBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                  'ft-6PGSPtIS - Investors Share of Gross Sales Proceeds Due to Purchase Cash Flow Between t and t+1
        Dim GSPDeathBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                     'ft-6DGSPt - Gross Sales Proceeds Due to Death Cash Flow Between t and t+1
        Dim BEGSPDeathBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                   'ft-6DGSPtBE - Bequest Equity Portion of Gross Sales Proceeds Due to Death Cash Flow Between t and t+1
        Dim CEGSPDeathBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                   'ft-6DGSPtCE - Conserved Equity Portion of Gross Sales Proceeds Due to Death Cash Flow Between t and t+1
        Dim MgrGSPDeathBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                  'ft-6DGSPtMS - Managers Share of Gross Sales Proceeds Due to Death Cash Flow Between t and t+1
        Dim InvGSPDeathBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                  'ft-6DGSPtIS - Investors Share of Gross Sales Proceeds Due to Death Cash Flow Between t and t+1
        Dim PmtCashFlow(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                    'St'Pt - Payment Cash Flow at t
        Dim RescPmtDeathCashFlowBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double         'ftRRtR - Rescinded Payments Due to Death Cash Flow Between t and t+1
        Dim ReturnPmtCancelCashFlow(ProgramParameters_FMV.ProjectedLife_t + 7) As Double        'ftCRtC - Returned Payments Due to Cancellation Cash Flow From t to t+1
        Dim CancelFeeCashFlowBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double            'ftCC - Cancellation Fee Cash Flow Between t and t+1
        Dim CashOutFlow_t(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                  'CFtOUT(B) - Cash OufFlows at t
        Dim CashOutFlowBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                  'CFtOUT(M) - Cash Outflows Between t and t+1
        Dim CashInFlowBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                   'CFtIN - Cash Inflows Between t and t+1
        Dim NAV(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                            'NAVt - Net Asset Value at time t

        Dim PmtAmt(ProgramParameters_FMV.ProjectedLife_t + 7) As Double
        Dim RescAmtDeath(ProgramParameters_FMV.ProjectedLife_t + 7) As Double
        Dim ReturnedAmtCancellation(ProgramParameters_FMV.ProjectedLife_t + 7) As Double

        '''''''''''''''''''''''''''''''''''''''''
        NAV(ProgramParameters_FMV.ProjectedLife_t + 1) = 0

        'APV CALCULATIONS
        Dim APVMaint(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                       'APVt(EM) - APVt(EM)
        Dim APVRepair(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                      'APVt(ER) - APV of Repair Expense at t
        Dim APVTaxes(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                       'APVt(ET) - APV of Taxes Expense at t
        Dim APVInsur(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                       'APVt(EI) - APV of Insurance Expense at t
        Dim APVRevers(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                      'APVt(ES0) - APV of Reversionary Expenses Concurrent with Transactions from t to t+1
        Dim APVEstMgtFee(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                   'APVt(FE) - APV of Monthly Estate Management Fees Paid to Manager at t
        Dim APVPmtMgtFee(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                   'APVt(FP) - APV of Monthly Payment Management Fees Paid to Manager at t
        Dim APVMgrFeeCancel(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                'APVt(FC) - APV of Manager Fee on Cancellation Fee From t to t+1
        Dim APVMgrSalesFeeRevers(ProgramParameters_FMV.ProjectedLife_t + 7) As Double           'APVt(FS) - APV of Sales Fee Paid to Manager Upon Reversion to Sale  for Transactions From t to t+1
        Dim APVMgrSalesFeeReversPurch(ProgramParameters_FMV.ProjectedLife_t + 7) As Double      'APVt(FGP) - APV of Fee Paid to Manager Upon Reversion to Sale Due to Purchase for Transactions From t to t+1
        Dim APVMgrSalesFeeReversDeath(ProgramParameters_FMV.ProjectedLife_t + 7) As Double      'APVt(FGD) - APV of Fee Paid to Manager Upon Reversion to Sale Due to Death for Transactions From t to t+1
        Dim APVGSP(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                         'APVt(GSP) - APV of Gross Sales Proceeds for Transactions From t and t+1
        Dim APVBEGSP(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                       'APVt(GSPBE) - APV of Bequest Equity Portion of Gross Sales Proceeds for Transactions From t to t+1
        Dim APVCEGSP(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                       'APVt(GSPCE) - APV of Conserved Equity Portion of Gross Sales Proceeds for Transactions From t to t+1
        Dim APVMgrGSP(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                      'APVt(GSPMS) - APV of Manager's Share of Gross Sales Proceeds for Transactions From t to t+1
        Dim APVInvGSP(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                      'APVt(GSPIS) - APV of Investor's Share of Gross Sales Proceeds for Transactions From t to t+1
        Dim APVPmt(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                         'APVt(P) - APV of Payment Amount at t
        Dim APVRescPmtDeath(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                'APVt(RR) - APV of Rescinded Payments Due to Death Amount From t to t+1
        Dim APVReturnPmtCancel(ProgramParameters_FMV.ProjectedLife_t + 7) As Double             'APVt(RC) - APV of Returned Payments Due to Cancellation Amount From t to t+1
        Dim APVCancelFee(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                   'APVt(C) - APV of Cancellation Fee From t to t+1
        Dim APVCashOutFlows(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                'APVt(CFOUT(B)) - APV of Cash OufFlows at t
        Dim APVCashOutFlowsBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double              'APVt(CFOUT(M)) - APV of Cash Outflows Between t and t+1
        Dim APVCashInFlowsBT(ProgramParameters_FMV.ProjectedLife_t + 7) As Double               'APVt(CFIN) - APV of Cash Inflows Between t and t+1
        Dim APVLEPt(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                        'APVt(LEP) - APV of Life Estate from t to t+1
        Dim APVNAV(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                         'NAVt - Net Asset Value at time t

        APVMaint(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVRepair(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVTaxes(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVInsur(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVInsur(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVRevers(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVEstMgtFee(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVPmtMgtFee(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVMgrFeeCancel(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVMgrSalesFeeRevers(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVMgrSalesFeeReversPurch(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVMgrSalesFeeReversDeath(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVGSP(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVBEGSP(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVCEGSP(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVMgrGSP(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVInvGSP(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVPmt(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVRescPmtDeath(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVReturnPmtCancel(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVCashOutFlows(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVCashOutFlows(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVCashOutFlowsBT(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVCashInFlowsBT(ProgramParameters_FMV.ProjectedLife_t + 7) = 0
        APVNAV(ProgramParameters_FMV.ProjectedLife_t + 7) = 0

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For t = ProgramParameters_FMV.ProjectedLife_t + 2 To 0 Step -1
            If t < 6 Then
                APVRevers(t) = MoneyFactor0 * APVRevers(t + 1)
                APVMgrSalesFeeRevers(t) = MoneyFactor0 * APVMgrSalesFeeRevers(t + 1)
                APVMgrSalesFeeReversPurch(t) = MoneyFactor0 * APVMgrSalesFeeReversPurch(t + 1)
                APVMgrSalesFeeReversDeath(t) = MoneyFactor0 * APVMgrSalesFeeReversDeath(t + 1)
                APVGSP(t) = MoneyFactor0 * APVGSP(t + 1)
                APVBEGSP(t) = MoneyFactor0 * APVBEGSP(t + 1)

                APVCEGSP(t) = MoneyFactor0 * APVCEGSP(t + 1)
                APVMgrGSP(t) = MoneyFactor0 * APVMgrGSP(t + 1)
                APVInvGSP(t) = MoneyFactor0 * APVInvGSP(t + 1)
            Else
                APVRevers(t) = (MoneyFactor0 ^ 0.5) * (ReversPurchCF(t) + ResalePurchCF(t) + ReversDeathCF(t) + ResaleDeathCF(t)) + MoneyFactor0 * APVRevers(t + 1)
                APVMgrSalesFeeRevers(t) = (MoneyFactor0 ^ 0.5) * (SalesFeesPurchReversCF(t) + SalesFeesDeathReversCF(t)) + MoneyFactor0 * APVMgrSalesFeeRevers(t + 1)
                APVMgrSalesFeeReversPurch(t) = (MoneyFactor0 ^ 0.5) * MgrFeeReversPurch(t) + MoneyFactor0 * APVMgrSalesFeeReversPurch(t + 1)
                APVMgrSalesFeeReversDeath(t) = (MoneyFactor0 ^ 0.5) * MgrFeeReversDeath(t) + MoneyFactor0 * APVMgrSalesFeeReversDeath(t + 1)
                APVGSP(t) = (MoneyFactor0 ^ 0.5) * (GSPPurchBT(t) + GSPDeathBT(t)) + MoneyFactor0 * APVGSP(t + 1)
                APVBEGSP(t) = (MoneyFactor0 ^ 0.5) * (BEGSPPurchBT(t) + BEGSPDeathBT(t)) + MoneyFactor0 * APVBEGSP(t + 1)
                APVCEGSP(t) = (MoneyFactor0 ^ 0.5) * (CEGSPPurchBT(t) + CEGSPDeathBT(t)) + MoneyFactor0 * APVCEGSP(t + 1)
                APVMgrGSP(t) = (MoneyFactor0 ^ 0.5) * (MgrGSPPurchBT(t) + MgrGSPDeathBT(t)) + MoneyFactor0 * APVMgrGSP(t + 1)
                APVInvGSP(t) = (MoneyFactor0 ^ 0.5) * (InvGSPPurchBT(t) + InvGSPDeathBT(t)) + MoneyFactor0 * APVInvGSP(t + 1)
            End If
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t
            If t < 6 Then
                MgrFeeReversPurch(t) = 0
                MgrFeeReversDeath(t) = 0
                'LEPurchBT(t) = 0
                GSPPurchBT(t) = 0
                BEGSPPurchBT(t) = 0
                CEGSPPurchBT(t) = 0
                MgrGSPPurchBT(t) = 0
                InvGSPPurchBT(t) = 0
                GSPDeathBT(t) = 0
                BEGSPDeathBT(t) = 0
                CEGSPDeathBT(t) = 0
                MgrGSPDeathBT(t) = 0
                InvGSPDeathBT(t) = 0
            Else
                MgrFeeReversPurch(t) = PLEPurchBT(t - 6) * FeeMgrReversPurch(t)
                MgrFeeReversDeath(t) = PReversBT(t - 6) * FeeMgrReversDeath(t)
                GSPPurchBT(t) = PLEPurchBT(t - 6) * FMVt(t)
                BEGSPPurchBT(t) = PLEPurchBT(t - 6) * BEofGSP(t)
                CEGSPPurchBT(t) = PLEPurchBT(t - 6) * CEofGSP(t)
                MgrGSPPurchBT(t) = PLEPurchBT(t - 6) * MGRofGSP(t)
                InvGSPPurchBT(t) = PLEPurchBT(t - 6) * INVESTORofGSP(t)
                GSPDeathBT(t) = PReversBT(t - 6) * FMVt(t)
                BEGSPDeathBT(t) = PReversBT(t - 6) * BEofGSP(t)
                CEGSPDeathBT(t) = PReversBT(t - 6) * CEofGSP(t)
                MgrGSPDeathBT(t) = PReversBT(t - 6) * MGRofGSP(t)
                InvGSPDeathBT(t) = PReversBT(t - 6) * INVESTORofGSP(t)
            End If
            LEPurchBT(t) = LEP(t) * PLEPurchBT(t)
            PmtCashFlow(t) = PmtAmt(t) * PBeingPresentNT(t)                                 'St'Pt      Payment Cash Flow at t
            RescPmtDeathCashFlowBT(t) = RescAmtDeath(t) * PRescDeathBT(t)                   'ftRRtR     Rescinded Payments Due to Death Cash Flow Between t and t+1
            ReturnPmtCancelCashFlow(t) = ReturnedAmtCancellation(t) * PCancelBT(t)          'ftCRtC     Returned Payments Due to Cancellation Cash Flow From t to t+1
            CancelFeeCashFlowBT(t) = CancellationCF(t) * PCancelBT(t)                   'ftCC       Cancellation Fee Cash Flow Between t and t+1
            If t = 0 Then                                                                   'CFtOUT(B)  Cash OufFlows at t
                CashOutFlow_t(t) = LEPurchBT(t) + MoPmtsMgtFeesCF(t) + MoEstMgtFeesCF(t) + RepairsCF(t) + MaintCF(t) + TaxesCF(t) + InsurCF(t) + Origination.CommissExp + ManagerFees.AdvEquityAmt + ManagerFees.InceptionExp
            Else
                CashOutFlow_t(t) = CashOutFlow_t(t) = LEPurchBT(t) + MoPmtsMgtFeesCF(t) + MoEstMgtFeesCF(t) + RepairsCF(t) + MaintCF(t) + TaxesCF(t) + InsurCF(t)
            End If
            'CFtOUT(M)    Cash Outflows Between t and t+1
            CashOutFlowBT(t) = CancellationFeesCF(t) + SalesFeesPurchReversCF(t) + SalesFeesDeathReversCF(t) + MgrFeesPurchReversCF(t) + MgrFeesDeathReversCF(t) + LEPurchBT(t) + ReversPurchCF(t) + ResalePurchCF(t) + ReversDeathCF(t) + ResaleDeathCF(t)
            'CFtIN        Cash Inflows Between t and t+1
            CashInFlowBT(t) = InvGSPPurchBT(t) + InvGSPDeathBT(t) + RescPmtDeathCashFlowBT(t) + ReturnPmtCancelCashFlow(t) + CancelFeeCashFlowBT(t)
            'NAVt         Net Asset Value at time t
            NAV(t) = (MoneyFactor0 ^ 0.5) * (CashInFlowBT(t) - CashOutFlow_t(t)) - CashOutFlow_t(t) + NAV(t + 1) * MoneyFactor0
        Next

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''PUT BELOW PAYMENT CALC OR FUNCTION
        'FINAL PAYMENT SCHEDULE

        'From Primary Pricing Relation
        Dim PPR_12N As Double = Nothing
        Dim PPR_12I As Double = Nothing
        Dim PPR_12M As Double = Nothing
        Dim PPR_12B As Double = Nothing

        'From Investor Ambivalence Relation
        Dim IAR_N As Double = Nothing
        Dim IAR_I As Double = Nothing
        Dim IAR_M As Double = Nothing
        Dim IAR_B As Double = Nothing

        'Payment Coefficient
        Dim PC_I As Double = Nothing
        Dim PC_M As Double = Nothing
        Dim PC_B As Double = Nothing

        'Pricing Methods
        Dim ScalarInitial = 0
        Dim ScalarMonthly = 0
        Dim ScalarBalloon = 0
        Dim ScalarTotal = 0

        Dim InitialDollarInitial = Nothing
        Dim InitialDollarMonthly = Nothing
        Dim InitialDollarBalloon = Nothing
        Dim InitialDollarTotal = Nothing

        Dim InitialSolutionInitial = Nothing
        Dim InitialSolutionMonthly = Nothing
        Dim InitialSolutionBalloon = Nothing
        Dim InitialSolutionTotal = Nothing

        Dim BalloonSolutionInitial = Nothing
        Dim BalloonSolutionMonthly = Nothing
        Dim BalloonSolutionBalloon = Nothing
        Dim BalloonSolutionTotal = Nothing

        Dim FactorNormBallonPmt = Nothing

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''	
        'APV Calc

        APVPmt(ProgramParameters_FMV.ProjectedLife_t + 2) = 0

        For t = ProgramParameters_FMV.ProjectedLife_t + 2 To 0 Step -1
            APVMaint(t) = MaintCF(t) + MoneyFactor0 * APVMaint(t + 1)                   'APVt(EM)  APV of Maintenance Expense at t
            APVRepair(t) = RepairsCF(t) + MoneyFactor0 * APVRepair(t + 1)               'APVt(ER)  APV of Repair Expense at t
            APVTaxes(t) = TaxesCF(t) + MoneyFactor0 * APVTaxes(t + 1)                   'APVt(ET)  APV of Taxes Expense at t
            APVInsur(t) = InsurCF(t) + MoneyFactor0 * APVInsur(t + 1)                   'APVt(EI)  APV of Insurance Expense at t
            APVEstMgtFee(t) = MoEstMgtFeesCF(t) + MoneyFactor0 * APVEstMgtFee(t + 1)
            APVPmtMgtFee(t) = PBeingPresentNT(t) * MoPmtMgtFees(t) + MoneyFactor0 * APVPmtMgtFee(t + 1)
            APVMgrFeeCancel(t) = (MoneyFactor0 ^ 0.5) * CancellationFeesCF(t) + MoneyFactor0 * APVMgrFeeCancel(t + 1)
            If t < 6 Then
                APVRevers(t) = MoneyFactor0 * APVRevers(t + 1)
                APVMgrSalesFeeRevers(t) = MoneyFactor0 * APVMgrSalesFeeRevers(t + 1)
                APVMgrSalesFeeReversPurch(t) = MoneyFactor0 * APVMgrSalesFeeReversPurch(t + 1)
                APVMgrSalesFeeReversDeath(t) = MoneyFactor0 * APVMgrSalesFeeReversDeath(t + 1)
                APVGSP(t) = MoneyFactor0 * APVGSP(t + 1)
                APVBEGSP(t) = MoneyFactor0 * APVBEGSP(t + 1) '
                APVCEGSP(t) = MoneyFactor0 * APVCEGSP(t + 1)
                APVMgrGSP(t) = MoneyFactor0 * APVMgrGSP(t + 1)
                APVInvGSP(t) = MoneyFactor0 * APVInvGSP(t + 1)
            Else
                APVRevers(t) = (MoneyFactor0 ^ 0.5) * (ReversPurchCF(t) + ResalePurchCF(t) + ReversDeathCF(t) + ResaleDeathCF(t)) + MoneyFactor0 * APVRevers(t + 1)
                APVMgrSalesFeeRevers(t) = (MoneyFactor0 ^ 0.5) * (SalesFeesPurchReversCF(t) + SalesFeesDeathReversCF(t)) + MoneyFactor0 * APVMgrSalesFeeRevers(t + 1)
                APVMgrSalesFeeReversPurch(t) = (MoneyFactor0 ^ 0.5) * MgrFeeReversPurch(t) + MoneyFactor0 * APVMgrSalesFeeReversPurch(t + 1)
                APVMgrSalesFeeReversDeath(t) = (MoneyFactor0 ^ 0.5) * MgrFeeReversDeath(t) + MoneyFactor0 * APVMgrSalesFeeReversDeath(t + 1)
                APVGSP(t) = (MoneyFactor0 ^ 0.5) * (GSPPurchBT(t) + GSPDeathBT(t)) + MoneyFactor0 * APVGSP(t + 1)
                APVBEGSP(t) = (MoneyFactor0 ^ 0.5) * (BEGSPPurchBT(t) + BEGSPDeathBT(t)) + MoneyFactor0 * APVBEGSP(t + 1)
                APVCEGSP(t) = (MoneyFactor0 ^ 0.5) * (CEGSPPurchBT(t) + CEGSPDeathBT(t)) + MoneyFactor0 * APVCEGSP(t + 1)
                APVMgrGSP(t) = (MoneyFactor0 ^ 0.5) * (MgrGSPPurchBT(t) + MgrGSPDeathBT(t)) + MoneyFactor0 * APVMgrGSP(t + 1)
                APVInvGSP(t) = (MoneyFactor0 ^ 0.5) * (InvGSPPurchBT(t) + InvGSPDeathBT(t)) + MoneyFactor0 * APVInvGSP(t + 1)
            End If
            APVPmt(t) = PmtCashFlow(t) + MoneyFactor0 * APVPmt(t + 1)
            APVRescPmtDeath(t) = (MoneyFactor0 ^ 0.5) * RescPmtDeathCashFlowBT(t) + MoneyFactor0 * APVRescPmtDeath(t + 1)
            APVReturnPmtCancel(t) = (MoneyFactor0 ^ 0.5) * ReturnPmtCancelCashFlow(t) + MoneyFactor0 * APVReturnPmtCancel(t + 1)
            APVCancelFee(t) = (MoneyFactor0 ^ 0.5) * CancelFeeCashFlowBT(t) + MoneyFactor0 * APVCancelFee(t + 1)
            APVCashOutFlows(t) = CashOutFlow_t(t) + MoneyFactor0 * APVCashOutFlows(t + 1)
            APVCashOutFlowsBT(t) = (MoneyFactor0 ^ 0.5) * CashOutFlowBT(t) + MoneyFactor0 * APVCashOutFlowsBT(t + 1)
            APVCashInFlowsBT(t) = (MoneyFactor0 ^ 0.5) * CashInFlowBT(t) + MoneyFactor0 * APVCashInFlowsBT(t + 1)
            APVLEPt(t) = (MoneyFactor0 ^ 0.5) * LEPurchBT(t) + MoneyFactor0 * APVLEPt(t + 1)
            APVNAV(t) = APVInvGSP(t) + APVRescPmtDeath(t) + APVReturnPmtCancel(t) + APVCancelFee(t) - APVPmt(t) - APVLEPt(t) - APVMgrFeeCancel(t) - APVEstMgtFee(t) - APVPmtMgtFee(t)
            APVNAV(t) = APVNAV(t) - APVMgrSalesFeeRevers(t) - APVMgrSalesFeeReversPurch(t) - APVMgrSalesFeeReversDeath(t) - APVMaint(t) - APVRepair(t) - APVTaxes(t) - APVInsur(t)
            APVNAV(t) = APVNAV(t) - APVRevers(t) - Origination.CommissExp - ManagerFees.InceptionExp - ManagerFees.AdvEquityAmt
        Next

        'If t > 11 Then
        '    APVLEP(t) = tN(t) + tI(t) * ProgramParameters_FMV.FMV_at_t0 + tM(t) * ManagerFees.CancellationFee + tB(t)
        '    APVLEP(t) = APVLEP(t) * (Origination.CommissOnContract * (1 - ContractTerms.ConservedEquityPct) * (ProgramParameters_FMV.FMV_at_t0 - ContractTerms.BequestEquity) + Origination.CommissOnPotential * (ContractTerms.ConservedEquityPct * ProgramParameters_FMV.FMV_at_t0 + (1 - ContractTerms.ConservedEquityPct * ContractTerms.BequestEquity)))
        '    LEP(t) = (APVLEP(t) - MoneyFactor(0) * APVLEP(t + 1)) / (MoneyFactor(0) ^ 0.5 * MoPmtMgtFees(t))
        'End If
        'Next

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim PPR_12NSumProduct1(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim PPR_12NSumProduct2(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim PPR_12NSumProduct3(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim PPR_12NSumProduct4(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim PPR_12NSumProduct5(ProgramParameters_FMV.ProjectedLife_t + 10) As Double

        Dim PPR_12ISumProduct1(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim PPR_12ISumProduct2(ProgramParameters_FMV.ProjectedLife_t + 10) As Double

        Dim PPR_12MSumProduct1(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim PPR_12MSumProduct2(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim PPR_12MSumProduct3(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim PPR_12MSumProduct4(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim PPR_12MSumProduct5(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim PPR_12MSumProduct6(ProgramParameters_FMV.ProjectedLife_t + 10) As Double

        Dim PPR_12BSumProduct1(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim PPR_12BSumProduct2(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim PPR_12BSumProduct3(ProgramParameters_FMV.ProjectedLife_t + 10) As Double

        t = 0

        PPR_12NSumProduct1(0) = 0
        PPR_12NSumProduct2(0) = 0
        PPR_12NSumProduct3(0) = 0
        PPR_12NSumProduct4(0) = 0
        PPR_12NSumProduct5(0) = 0

        PPR_12ISumProduct1(0) = 0
        PPR_12ISumProduct2(0) = 0

        PPR_12MSumProduct1(0) = 0
        PPR_12MSumProduct2(0) = 0
        PPR_12MSumProduct3(0) = 0
        PPR_12MSumProduct4(0) = 0
        PPR_12MSumProduct5(0) = 0
        PPR_12MSumProduct6(0) = 0

        PPR_12BSumProduct1(0) = 0
        PPR_12BSumProduct2(0) = 0
        PPR_12BSumProduct3(0) = 0

        For t As Integer = 1 To ProgramParameters_FMV.ProjectedLife_t - 1
            PPR_12NSumProduct1(t) = PPR_12NSumProduct1(t - 1) + MoneyFactor(t + 6) * (ReversArray(t + 6) + ResaleArray(t + 6)) * (PLEPurchBT(t) + PReversBT(t))  '(ProbLEBuyBTween(t) + ProbNonTenderDeathBTween(t) + ProbTenderDeathBTween(t))
            PPR_12NSumProduct2(t) = PPR_12NSumProduct2(t - 1) + MoneyFactor(t + 6) * (PLEPurchBT(t) + PReversBT(t))                                              '(ReversArray(t + 7) + ResaleArray(t + 7))  '(ProbLEBuyBTween(t) + ProbNonTenderDeathBTween(t) + ProbTenderDeathBTween(t))
            PPR_12NSumProduct3(t) = PPR_12NSumProduct3(t - 1) + MoneyFactor(t + 6) * (PLEPurchBT(t) + PReversBT(t)) * (MaintExCum(t) + RepairExCum(t) + TaxesExCum(t) + InsurExCum(t) + MoEstateMgtFeesCum(t) + MoPmtMgtFeesCum(t))   '(ProbLEBuyBTween(t) + ProbNonTenderDeathBTween(t) + ProbTenderDeathBTween(t)) * (MaintExCum(t) + RepairExCum(t) + TaxesExCum(t) + InsurExCum(t) + MoEstateMgtFeesCum(t) + MoPmtMgtFeesCum(t))          
            PPR_12NSumProduct4(t) = PPR_12NSumProduct4(t - 1) + MoneyFactor(t + 6) * (PLEPurchBT(t) / PBeingPresentNT(t + 1)) * sub1(t)
            PPR_12NSumProduct5(t) = PPR_12NSumProduct5(t - 1) + (PLEPurchBT(t) / PBeingPresentNT(t + 1)) * sub2(t)

            If t < 7 Then
                PPR_12ISumProduct1(t) = PPR_12ISumProduct1(t - 1) + MoneyFactor(t) * (PRescDeathBT(t) + PCancelBT(t))
            End If
            PPR_12ISumProduct2(t) = PPR_12ISumProduct2(t - 1) + MoneyFactor(t + 7) * (PLEPurchBT(t) + PReversBT(t))
        Next

        For t As Integer = 1 To ContractTerms.NumMoPmts
            PPR_12MSumProduct1(t) = PPR_12MSumProduct1(t - 1) + MoneyFactor(t + 1) * PBeingPresentNT(t + 1)
            If t < 7 Then
                PPR_12MSumProduct2(t) = PPR_12MSumProduct2(t - 1) + MoneyFactor(t - 1) * (PRescDeathBT(t - 1) + PCancelBT(t - 1)) * (t - 1)
            End If
        Next

        For t As Integer = 1 To ContractTerms.NumMoPmts + 1
            PPR_12MSumProduct3(t) = PPR_12MSumProduct3(t - 1) + MoneyFactor(t + 6) * (PLEPurchBT(t - 1) + PReversBT(t - 1)) * (t - 1)
        Next

        For t As Integer = 0 + ContractTerms.NumMoPmts + 1 To 500 - ContractTerms.NumMoPmts + 1
            PPR_12MSumProduct4(t) = PPR_12MSumProduct4(t - 1) + MoneyFactor(t + 6) * (PLEPurchBT(t) + PReversBT(t))

            'PPR_12BSumProduct1(t) = PPR_12BSumProduct1(t-1) + MoneyFactor(t + 6) * (PLEPurchBT(t) + PReversBT(t))
        Next

        For t As Integer = 1 To ProgramParameters_FMV.ProjectedLife_t + 6
            PPR_12MSumProduct5(t) = PPR_12MSumProduct5(t - 1) + MoneyFactor(t + 6) * (PLEPurchBT(t) / PBeingPresentNT(t + 1)) * sub3(t)
            PPR_12MSumProduct6(t) = PPR_12MSumProduct6(t - 1) + (PLEPurchBT(t) / PBeingPresentNT(t + 1)) * sub4(t)

            PPR_12BSumProduct2(t) = PPR_12BSumProduct2(t - 1) + MoneyFactor(t + 6) * (PLEPurchBT(t) / PBeingPresentNT(t + 1)) * sub5(t)
            PPR_12BSumProduct3(t) = PPR_12BSumProduct3(t - 1) + (PLEPurchBT(t) / PBeingPresentNT(t + 1)) * sub6(t)

        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'PRICING

        'From Primary Pricing Relation
        t = 499

        PPR_12N = (1 / ((1 - ManagerFees.PostReversNP * MoneyFactor0 ^ 6) * MoneyFactor0 ^ 12)) * ((1 - ManagerFees.PostReversNP) * (APVInvGSP(0) - APVMgrSalesFeeRevers(0)) + APVCancelFee(0) - APVMgrFeeCancel(0) - APVEstMgtFee(0) - APVPmtMgtFee(0) - APVMaint(0) - APVRepair(0) - APVTaxes(0) - APVInsur(0) - APVRevers(0) + ManagerFees.PostReversNP * PPR_12NSumProduct1(t)) - (ManagerFees.AdvEquityAmt + ManagerFees.InceptionExp + Origination.CommissExp) * (1 - ManagerFees.PostReversNP * PPR_12NSumProduct2(t)) + ManagerFees.PostReversNP * PPR_12NSumProduct3(t) + ManagerFees.PostReversNP * PPR_12NSumProduct4(t) + (ManagerFees.PostReversNP / MoneyFactor0 ^ 0.5) * PPR_12NSumProduct5(t)

        PPR_12I = 1 / ((1 - ManagerFees.PostReversNP * MoneyFactor0 ^ 6) * MoneyFactor0 ^ 12) * (-1 + PPR_12ISumProduct1(t) + ManagerFees.PostReversNP * PPR_12ISumProduct2(t))

        PPR_12M = (1 / ((1 - ManagerFees.PostReversNP * MoneyFactor0 ^ 6) * MoneyFactor0 ^ 12) * (-1 / MoneyFactor0 ^ 0.5 * PPR_12MSumProduct1(ContractTerms.NumMoPmts)) + PPR_12MSumProduct2(6) + ManagerFees.PostReversNP * PPR_12MSumProduct3(t) + ManagerFees.PostReversNP * ContractTerms.NumMoPmts * PPR_12MSumProduct4(t) + ManagerFees.PostReversNP * PPR_12MSumProduct5(t) + ManagerFees.PostReversNP / MoneyFactor0 ^ 0.5 * PPR_12MSumProduct6(t))

        PPR_12B = 1 / ((1 - ManagerFees.PostReversNP * MoneyFactor0 ^ 6) * MoneyFactor0 ^ 12) * (-1 / MoneyFactor0 ^ 0.5 * MoneyFactor(ContractTerms.NumMoPmts + 1) * PBeingPresentNT(ContractTerms.NumMoPmts + 1) + ManagerFees.PostReversNP * PPR_12BSumProduct1(t) + ManagerFees.PostReversNP * PPR_12BSumProduct2(t) + (ManagerFees.PostReversNP / MoneyFactor0 ^ 0.5) * PPR_12BSumProduct3(t))
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'PAYMENTS AND LIFE ESTATE VALUE

        Dim btSum1 As Double = 0
        Dim btSum2 As Double = 0
        Dim btSum3 As Double = 0
        Dim ctSum1 As Double = 0
        Dim dtSum1 As Double = 0
        Dim dtSum2 As Double = 0
        Dim dtSum3 As Double = 0
        Dim etSum1 As Double = 0
        Dim etSum2 As Double = 0
        Dim etSum3 As Double = 0

        Dim btSumproduct1(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim btSumproduct2(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim btSumproduct3(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim ctSumproduct1(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim dtSumproduct1(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim dtSumproduct2(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim dtSumproduct3(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim etSumproduct1(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim etSumproduct2(ProgramParameters_FMV.ProjectedLife_t + 10) As Double
        Dim etSumproduct3(ProgramParameters_FMV.ProjectedLife_t + 10) As Double

        Dim CY As Double
        CY = 0

        For t As Integer = 0 To 11  'ProgramParameters_FMV.ProjectedLife_t + 1
            At(t) = 0
            Bsub1(t) = 0
            Bsub2(t) = 0
            Bt(t) = 0
            Ct(t) = 0
            Dsub1(t) = 0
            Dsub2(t) = 0
            Dt(t) = 0
            Esub1(t) = 0
            Esub2(t) = 0
            Et(t) = 0
            tN(t) = 0
            tI(t) = 0
            tM(t) = 0
            tB(t) = 0

            btSumproduct1(t) = 0
            btSumproduct2(t) = 0
            btSumproduct3(t) = 0
            ctSumproduct1(t) = 0
            dtSumproduct1(t) = 0
            dtSumproduct2(t) = 0
            dtSumproduct3(t) = 0
            etSumproduct1(t) = 0
            etSumproduct2(t) = 0
            etSumproduct3(t) = 0
        Next

        t = 11

        For k As Integer = t + 1 To ProgramParameters_FMV.ProjectedLife_t - 1
            At(k) = MoneyFactor0 * (1 - (ManagerFees.PostReversNP * MoneyFactor0 ^ 6 * PLEPurchBT(k)) / (PPresent(t + 1) * ((1 / LETenders.AntiSelection) - ManagerFees.PostReversNP * MoneyFactor0 ^ 6)))
        Next

        For j As Integer = 1 To ProgramParameters_FMV.ProjectedLife_t - 9
            If j < 12 Then
                CY = 0
            Else
                For q As Integer = j + 1 To j + 6
                    CY = CY + PBeingPresentNT(q) * MoPmtMgtFees(q)
                Next
                Bsub1(j) = CY
                CY = 0
            End If
        Next

        For j As Integer = 1 To ProgramParameters_FMV.ProjectedLife_t
            If j < 12 Then
                CY = 0
            Else
                For q As Integer = j + 7 To ProgramParameters_FMV.ProjectedLife_t
                    CY = CY + MoneyFactor(q) * PBeingPresentNT(q) * MoPmtMgtFees(q)
                Next
                Bsub2(j) = CY
                CY = 0
            End If
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        btSumproduct1(11) = 0
        btSumproduct2(11) = 0
        btSumproduct3(11) = 0
        ctSumproduct1(11) = 0
        dtSumproduct1(11) = 0

        For j = 12 To ProgramParameters_FMV.ProjectedLife_t
            For k = j To ProgramParameters_FMV.ProjectedLife_t - 2
                btSum1 = btSum1 + MoneyFactor(k + 7) * (PLEPurchBT(k + 1) + PReversBT(k + 1)) * (ManagerFees.InceptionExp + ManagerFees.AdvEquityAmt + Origination.CommissExp + MaintExCum(k + 1) + RepairExCum(k + 1) + TaxesExCum(k + 1) + InsurExCum(k + 1) + MoEstateMgtFeesCum(k + 1) + MoPmtMgtFeesCum(k + 1))
                btSum2 = btSum2 + MoneyFactor(k + 7) * (PLEPurchBT(k + 1) / PBeingPresentNT(k + 2)) * Bsub1(k + 1)
                btSum3 = btSum3 + (PLEPurchBT(k + 1) / PBeingPresentNT(k + 2)) * Bsub2(k + 1)
                ctSum1 = ctSum1 + MoneyFactor(k + 7) * (PLEPurchBT(k + 1) + PReversBT(k + 1))
                dtSum1 = dtSum1 + MoneyFactor(k + 7) * (PLEPurchBT(k + 1) + PDeathBTNT(k + 1)) * MoPmtCumCoeff(k + 1)
            Next
            btSumproduct1(j) = btSum1
            btSumproduct2(j) = btSum2
            btSumproduct3(j) = btSum3
            ctSumproduct1(j) = ctSum1
            dtSumproduct1(j) = dtSum1

            btSum1 = 0
            btSum2 = 0
            btSum3 = 0
            ctSum1 = 0
            dtSum1 = 0
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim Bt1(ProgramParameters_FMV.ProjectedLife_t) As Double
        Dim Bt2(ProgramParameters_FMV.ProjectedLife_t) As Double
        Dim Bt3(ProgramParameters_FMV.ProjectedLife_t) As Double
        Dim Bt4(ProgramParameters_FMV.ProjectedLife_t) As Double
        Dim Bt5(ProgramParameters_FMV.ProjectedLife_t) As Double
        Dim Bt6(ProgramParameters_FMV.ProjectedLife_t) As Double
        Dim Bt7(ProgramParameters_FMV.ProjectedLife_t) As Double
        Dim Bt8(ProgramParameters_FMV.ProjectedLife_t) As Double
        Dim Bt9(ProgramParameters_FMV.ProjectedLife_t) As Double
        Dim Bt10(ProgramParameters_FMV.ProjectedLife_t) As Double

        ProjectedLifeSeries(ReversArray5, ReversArray4, ReversArray3, ReversArray2, ReversArray1, ReversArrayUponSale, MoneyFactor, PPresent, PLEPurchBT, PBeingPresentNT, MoEstateMgtFeesCum, MoPmtMgtFeesCum, MgrResaleFee, MaintExCum, RepairExCum, TaxesExCum, InsurExCum, INVESTORofGSP, Bsub1, Bsub2, Bt, MoneyFactor0, APVRever5, APVRever4, APVRever3, APVRever2, APVRever1, APVReverTrans, APVMaint, APVRepair, APVTaxes, APVInsur, APVEstMgtFee, APVMgrSalesFeeRevers, APVInvGSP, btSumproduct1, btSumproduct2, btSumproduct3, Bt1, Bt2, Bt3, Bt4, Bt5, Bt6, Bt7, Bt8, Bt9, Bt10)

        For j As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t
            If j < 12 Then
                CY = 0
            Else
                For q As Integer = j To ProgramParameters_FMV.ProjectedLife_t
                    CY = CY + MoneyFactor(q + 7) * (PLEPurchBT(q + 1) + PReversBT(q + 1))
                Next
                ctSumproduct1(j) = CY
                CY = 0
            End If
        Next

        t = 11
        For t = t + 1 To ProgramParameters_FMV.ProjectedLife_t
            Ct(t) = ManagerFees.PostReversNP * (MoneyFactor0 ^ 0.5) * PLEPurchBT(t) / (PPresent(t + 1) * ((1 / LETenders.AntiSelection) - ManagerFees.PostReversNP * MoneyFactor0 ^ 6)) * ((MoneyFactor0 ^ 6) * PPresent(t + 1) - (1 / MoneyFactor(t)) * ctSumproduct1(t))
        Next

        'Dsub1(t) 
        For j As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t - 12
            If j < 12 Then
                CY = 0
            Else
                For q As Integer = j + 1 To j + 6
                    CY = CY + PBeingPresentNT(q) * MoPmtIndicator(q)
                Next
                Dsub1(j) = CY
                CY = 0
            End If
        Next

        'Dsub2(t) 
        For j As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t - 12
            If j < 12 Then
                CY = 0
            Else
                For q As Integer = j + 7 To ProgramParameters_FMV.ProjectedLife_t
                    CY = CY + PBeingPresentNT(q) * MoneyFactor(q) * MoPmtIndicator(q)
                Next
                Dsub2(j) = CY
                CY = 0
            End If
        Next

        '''''''''''''''''''''''''''''''''''''''''''
        Dim DT1 As Double
        Dim DT2 As Double
        Dim DT3 As Double
        DT1 = 0
        DT2 = 0
        DT3 = 0

        For j As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t - 12
            If j < 12 Then
                DT1 = 0
                DT2 = 0
                DT3 = 0
            Else
                For k As Integer = j To ProgramParameters_FMV.ProjectedLife_t - 12
                    DT1 = DT1 + MoneyFactor(k + 7) * (PLEPurchBT(k + 1) + PDeathBTNT(k + 1)) * MoPmtCumCoeff(k + 1)
                    DT2 = DT2 + MoneyFactor(k + 7) * (PLEPurchBT(k + 1) / PBeingPresentNT(k + 2)) * Dsub1(k + 1)
                    DT3 = DT3 + (PLEPurchBT(k + 1) / PBeingPresentNT(k + 2)) * Dsub2(k + 1)
                Next
            End If
            dtSumproduct1(j) = DT1
            dtSumproduct2(j) = DT2
            dtSumproduct3(j) = DT3

            DT1 = 0
            DT2 = 0
            DT3 = 0
        Next
        '''''''''''''''''''''''''''''''''''''''''''	

        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t - 1
            Dt(t) = ManagerFees.PostReversNP * (MoneyFactor0 ^ 0.5) * PLEPurchBT(t) / (PPresent(t + 1) * ((1 / LETenders.AntiSelection) - ManagerFees.PostReversNP * MoneyFactor0 ^ 6)) * ((MoneyFactor0 ^ 6) * PPresent(t + 1) * MoPmtCumCoeff(t) + (MoneyFactor0 ^ 6) * (PPresent(t + 1) / PBeingPresentNT(t + 1)) * Dsub1(t) + (MoneyFactor0 ^ 0.5) * (PPresent(t + 1) / (PBeingPresentNT(t + 1) * MoneyFactor(t + 1))) * Dsub2(t) - (1 / MoneyFactor(t)) * dtSumproduct1(t) - (1 - MoneyFactor(t)) * dtSumproduct2(t) - (1 / (MoneyFactor(t) * (MoneyFactor0 ^ 0.5))) * dtSumproduct3(t))
            If Dt(t) < 0 Then
                Dt(t) = 0
            End If
        Next

        'Esub1(t)
        For j As Integer = 1 To ProgramParameters_FMV.ProjectedLife_t
            If j < 12 Then
                CY = 0
            Else
                For q As Integer = j + 1 To j + 6
                    CY = CY + PBeingPresentNT(q) * BalloonPmtIndicator(q)
                Next
                Esub1(j) = CY
                CY = 0
            End If
        Next

        'Esub2(t) 
        For j As Integer = 1 To ProgramParameters_FMV.ProjectedLife_t
            If j < 12 Then
                CY = 0
            Else
                For q As Integer = j + 7 To ProgramParameters_FMV.ProjectedLife_t
                    CY = CY + MoneyFactor(q) * PBeingPresentNT(q) * BalloonPmtIndicator(q)
                Next
                Esub2(j) = CY
                CY = 0
            End If
        Next

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim ET1 As Double
        Dim ET2 As Double
        Dim ET3 As Double
        ET1 = 0
        ET2 = 0
        ET3 = 0

        For j As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t - 12
            If j < 12 Then
                ET1 = 0
                ET2 = 0
                ET3 = 0
            Else
                For k As Integer = 1 To ProgramParameters_FMV.ProjectedLife_t - 2
                    ET1 = ET1 + MoneyFactor(k + 7) * (PLEPurchBT(k + 1) + PReversBT(k + 1)) * BalloonPmtCoeff(k + 1)
                    ET2 = ET2 + MoneyFactor(k + 7) * (PLEPurchBT(k + 1) / PBeingPresentNT(k + 2)) * Esub1(k + 1)
                    ET3 = ET3 + (PLEPurchBT(k + 1) / PBeingPresentNT(k + 2)) * Esub2(k + 1)
                Next
            End If
            etSumproduct1(j) = ET1
            etSumproduct2(j) = ET2
            etSumproduct3(j) = ET3

            ET1 = 0
            ET2 = 0
            ET3 = 0
        Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For j As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t - 12
            If j < 12 Then
                Et(j) = 0
            Else
                Et(j) = ManagerFees.PostReversNP * (MoneyFactor0 ^ 0.5) * PLEPurchBT(j) / (PPresent(j + 1) * ((1 / LETenders.AntiSelection) - ManagerFees.PostReversNP * MoneyFactor0 ^ 6)) * ((MoneyFactor0 ^ 6) * PPresent(j + 1) * BalloonPmtCoeff(j) + (MoneyFactor0 ^ 6) * ((PPresent(j + 1) / PBeingPresentNT(j + 1)) * Esub1(j) + (MoneyFactor0 ^ 0.5) * (PPresent(j + 1) / (PBeingPresentNT(j + 1) * MoneyFactor(j + 1))) * Esub2(j) - (1 / MoneyFactor(j)) * etSumproduct1(j) - (1 / MoneyFactor(j)) * etSumproduct2(j) - (1 / (MoneyFactor(j) * (MoneyFactor0 ^ 0.5))) * etSumproduct3(j)))
            End If
        Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Pricing

        tN(ProgramParameters_FMV.ProjectedLife_t) = 0
        tI(ProgramParameters_FMV.ProjectedLife_t) = 0
        tM(ProgramParameters_FMV.ProjectedLife_t) = 0
        tB(ProgramParameters_FMV.ProjectedLife_t) = 0
        APVLEP(ProgramParameters_FMV.ProjectedLife_t) = 0
        LEP(ProgramParameters_FMV.ProjectedLife_t) = 0

        For t = ProgramParameters_FMV.ProjectedLife_t - 1 To 0 Step -1
            tN(t) = Bt(t) + At(t) * tN(t + 1)
            tI(t) = Ct(t) + At(t) * tI(t + 1)
            tM(t) = Dt(t) + At(t) * tM(t + 1)
            tB(t) = Et(t) + At(t) * tB(t + 1)
        Next

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'From Investor Ambivalence Relation
        IAR_N = tN(12)
        IAR_I = tI(12)
        IAR_M = tM(12)
        IAR_B = tB(12)

        'Payment Coefficient

        PC_I = (IAR_I - PPR_12I) / (PPR_12N - IAR_N)
        PC_M = (IAR_M - PPR_12M) / (PPR_12N - IAR_N)
        PC_B = (IAR_B - PPR_12B) / (PPR_12N - IAR_N)

        'Factor for Normal Ballon Payment

        For t As Integer = 0 + ContractTerms.NumMoPmts + 1 To 500 - ContractTerms.NumMoPmts
            FactorNormBallonPmt = FactorNormBallonPmt + (PBeingPresentNT(t) * MoneyFactor(t)) / (MoneyFactor(ContractTerms.NumMoPmts + 1) * PBeingPresentNT(ContractTerms.NumMoPmts + 1))
        Next


        'SCALAR METHOD
        ScalarMonthly = 1 / (PC_I * PaymentSolutionsMethod.ScalarMonthly + PC_M + PC_B * PaymentSolutionsMethod.ScalarNormBalloon * FactorNormBallonPmt)
        ScalarInitial = ScalarMonthly * PaymentSolutionsMethod.ScalarMonthly
        ScalarBalloon = PaymentSolutionsMethod.ScalarNormBalloon * FactorNormBallonPmt * ScalarMonthly
        ScalarTotal = ScalarInitial + ScalarMonthly * ContractTerms.NumMoPmts + ScalarBalloon

        'INITIAL DOLLAR METHOD
        InitialDollarInitial = PaymentSolutionsMethod.InitialDollarInitial
        InitialDollarMonthly = (1 - PC_I * InitialDollarInitial) / (PC_M + PC_B * PaymentSolutionsMethod.InitialDollarNormBalloon * FactorNormBallonPmt)
        InitialDollarBalloon = PaymentSolutionsMethod.InitialDollarNormBalloon * FactorNormBallonPmt * InitialDollarMonthly
        InitialDollarTotal = InitialDollarInitial + InitialDollarMonthly * ContractTerms.NumMoPmts + InitialDollarBalloon

        'INITIAL SOLUTION METHOD
        InitialSolutionInitial = (1 - PaymentSolutionsMethod.InitialSolutionMonthly * (PC_M + PC_B * PaymentSolutionsMethod.InitialSolutionNormBalloon * FactorNormBallonPmt)) / PC_I
        InitialSolutionMonthly = PaymentSolutionsMethod.InitialSolutionMonthly
        InitialSolutionBalloon = PaymentSolutionsMethod.InitialSolutionNormBalloon * FactorNormBallonPmt * InitialSolutionMonthly
        InitialSolutionTotal = InitialSolutionInitial + InitialSolutionMonthly * ContractTerms.NumMoPmts + InitialSolutionBalloon

        'BALLOON SOLUTIONS METHOD
        BalloonSolutionMonthly = PaymentSolutionsMethod.BalloonSolutionMonthly
        BalloonSolutionInitial = BalloonSolutionMonthly * PaymentSolutionsMethod.BalloonSolutionInitialMonthly
        BalloonSolutionBalloon = (1 - BalloonSolutionMonthly * (PC_I * PaymentSolutionsMethod.BalloonSolutionInitialMonthly + PC_M)) / PC_B
        BalloonSolutionTotal = BalloonSolutionInitial * PaymentSolutionsMethod.BalloonSolutionInitialMonthly + PaymentSolutionsMethod.BalloonSolutionMonthly * ContractTerms.NumMoPmts + BalloonSolutionBalloon

        Dim PmtAmtCum(ProgramParameters_FMV.ProjectedLife_t + 7) As Double                           'PtT  Cumulative Payment Amounts at t

        For t = ProgramParameters_FMV.ProjectedLife_t - 1 To 0 Step -1
            If t > 11 Then
                APVLEP(t) = tN(t) + tI(t) * ProgramParameters_FMV.FMV_at_t0 + tM(t) * ManagerFees.CancellationFee + tB(t)
                APVLEP(t) = APVLEP(t) * (Origination.CommissOnContract * (1 - ContractTerms.ConservedEquityPct) * (ProgramParameters_FMV.FMV_at_t0 - ContractTerms.BequestEquity) + Origination.CommissOnPotential * (ContractTerms.ConservedEquityPct * ProgramParameters_FMV.FMV_at_t0 + (1 - ContractTerms.ConservedEquityPct * ContractTerms.BequestEquity)))
                LEP(t) = (APVLEP(t) - MoneyFactor(0) * APVLEP(t + 1)) / (MoneyFactor(0) ^ 0.5 * MoPmtMgtFees(t))
            End If
        Next

        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t
            'assumes Scalar for now
            If t < ContractTerms.NumMoPmts Then
                PmtAmt(t) = ScalarMonthly
            ElseIf t = ContractTerms.NumMoPmts Then
                PmtAmt(t) = InitialDollarBalloon
            ElseIf t > ContractTerms.NumMoPmts Then
                PmtAmt(t) = 0
            End If

            If t = 0 Then
                PmtAmtCum(t) = ScalarMonthly
                RescAmtDeath(t) = ScalarMonthly
                ReturnedAmtCancellation(t) = ScalarMonthly
            Else
                PmtAmtCum(t) = PmtAmtCum(t - 1) + PmtAmt(t)
                RescAmtDeath(t) = RescAmtDeath(t - 1) + PmtAmt(t)
                ReturnedAmtCancellation(t) = ReturnedAmtCancellation(t - 1) + PmtAmt(t)
            End If
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'CashFlow to NAV

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t
            If t < 6 Then
                MgrFeeReversPurch(t) = 0
                MgrFeeReversDeath(t) = 0
                'LEPurchBT(t) = 0
                GSPPurchBT(t) = 0
                BEGSPPurchBT(t) = 0
                CEGSPPurchBT(t) = 0
                MgrGSPPurchBT(t) = 0
                InvGSPPurchBT(t) = 0
                GSPDeathBT(t) = 0
                BEGSPDeathBT(t) = 0
                CEGSPDeathBT(t) = 0
                MgrGSPDeathBT(t) = 0
                InvGSPDeathBT(t) = 0
            Else
                MgrFeeReversPurch(t) = PLEPurchBT(t - 6) * FeeMgrReversPurch(t)
                MgrFeeReversDeath(t) = PReversBT(t - 6) * FeeMgrReversDeath(t)
                GSPPurchBT(t) = PLEPurchBT(t - 6) * FMVt(t)
                BEGSPPurchBT(t) = PLEPurchBT(t - 6) * BEofGSP(t)
                CEGSPPurchBT(t) = PLEPurchBT(t - 6) * CEofGSP(t)
                MgrGSPPurchBT(t) = PLEPurchBT(t - 6) * MGRofGSP(t)
                InvGSPPurchBT(t) = PLEPurchBT(t - 6) * INVESTORofGSP(t)
                GSPDeathBT(t) = PReversBT(t - 6) * FMVt(t)
                BEGSPDeathBT(t) = PReversBT(t - 6) * BEofGSP(t)
                CEGSPDeathBT(t) = PReversBT(t - 6) * CEofGSP(t)
                MgrGSPDeathBT(t) = PReversBT(t - 6) * MGRofGSP(t)
                InvGSPDeathBT(t) = PReversBT(t - 6) * INVESTORofGSP(t)
            End If
            LEPurchBT(t) = LEP(t) * PLEPurchBT(t)
            PmtCashFlow(t) = PmtAmt(t) * PBeingPresentNT(t)                                 'St'Pt      Payment Cash Flow at t
            RescPmtDeathCashFlowBT(t) = RescAmtDeath(t) * PRescDeathBT(t)                   'ftRRtR     Rescinded Payments Due to Death Cash Flow Between t and t+1
            ReturnPmtCancelCashFlow(t) = ReturnedAmtCancellation(t) * PCancelBT(t)          'ftCRtC     Returned Payments Due to Cancellation Cash Flow From t to t+1
            CancelFeeCashFlowBT(t) = CancellationFeesCF(t) * PCancelBT(t)                   'ftCC       Cancellation Fee Cash Flow Between t and t+1
            If t = 0 Then                                                                   'CFtOUT(B)  Cash OufFlows at t
                CashOutFlow_t(t) = LEPurchBT(t) + MoPmtsMgtFeesCF(t) + MoEstMgtFeesCF(t) + RepairsCF(t) + MaintCF(t) + TaxesCF(t) + InsurCF(t) + Origination.CommissExp + ManagerFees.AdvEquityAmt + ManagerFees.InceptionExp
            Else
                CashOutFlow_t(t) = CashOutFlow_t(t) = LEPurchBT(t) + MoPmtsMgtFeesCF(t) + MoEstMgtFeesCF(t) + RepairsCF(t) + MaintCF(t) + TaxesCF(t) + InsurCF(t)
            End If
            'CFtOUT(M)    Cash Outflows Between t and t+1
            CashOutFlowBT(t) = CancellationFeesCF(t) + SalesFeesPurchReversCF(t) + SalesFeesDeathReversCF(t) + MgrFeesPurchReversCF(t) + MgrFeesDeathReversCF(t) + LEPurchBT(t) + ReversPurchCF(t) + ResalePurchCF(t) + ReversDeathCF(t) + ResaleDeathCF(t)
            'CFtIN        Cash Inflows Between t and t+1
            CashInFlowBT(t) = InvGSPPurchBT(t) + InvGSPDeathBT(t) + RescPmtDeathCashFlowBT(t) + ReturnPmtCancelCashFlow(t) + CancelFeeCashFlowBT(t)
            'NAVt         Net Asset Value at time t
            NAV(t) = (MoneyFactor0 ^ 0.5) * (CashInFlowBT(t) - CashOutFlow_t(t)) - CashOutFlow_t(t) + NAV(t + 1) * MoneyFactor0
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        For t = ProgramParameters_FMV.ProjectedLife_t - 1 To 0 Step -1
            If t > 11 Then
                APVLEP(t) = tN(t) + tI(t) * ScalarInitial + tM(t) * ScalarMonthly + tB(t) * ScalarBalloon
                LEP(t) = (APVLEP(t) - MoneyFactor(0) * APVLEP(t + 1)) / (MoneyFactor(0) ^ 0.5 * PLEPurchBT(t))
            End If

        Next
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim S1Sub As Double
        Dim S2Sub As Double
        Dim S3Sub As Double
        Dim S4Sub As Double
        Dim S5Sub As Double
        Dim S6Sub As Double
        Dim S7Sub As Double
        Dim S8Sub As Double

        S1Sub = 0
        S2Sub = 0
        S3Sub = 0
        S4Sub = 0
        S5Sub = 0
        S6Sub = 0
        S7Sub = 0
        S8Sub = 0

        For j As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t - 3
            For k As Integer = 1 To 6
                S1Sub = S1Sub + PBeingPresentNT(k + j) * MoPmtMgtFees(k + j)
                S3Sub = S3Sub + PBeingPresentNT(k + j) * MoPmtIndicator(k + j)
                S5Sub = S5Sub + PBeingPresentNT(k + j) * BalloonPmtIndicator(k + j)
                S7Sub = S7Sub + PBeingPresentNT(k + j) * PmtAmt(k + j)
            Next
            sub1(j) = S1Sub
            sub3(j) = S3Sub
            sub5(j) = S5Sub
            sub7(j) = S7Sub

            S1Sub = 0
            S3Sub = 0
            S5Sub = 0
            S7Sub = 0
        Next

        For j As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t
            For k As Integer = 7 + j To ProgramParameters_FMV.ProjectedLife_t - 2
                S2Sub = S2Sub + MoneyFactor(k) * PBeingPresentNT(k) * MoPmtMgtFees(k)
                S4Sub = S4Sub + MoneyFactor(k) * PBeingPresentNT(k) * MoPmtIndicator(k)
                S6Sub = S6Sub + MoneyFactor(k) * PBeingPresentNT(k) * BalloonPmtIndicator(k)
                S8Sub = S8Sub + MoneyFactor(k) * PBeingPresentNT(k) * PmtAmt(k)

            Next
            sub2(j) = S2Sub
            sub4(j) = S4Sub
            sub6(j) = S6Sub
            sub8(j) = S8Sub

            S2Sub = 0
            S4Sub = 0
            S6Sub = 0
            S8Sub = 0
        Next

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SumGG1 As Double = 0
        Dim SumGG2 As Double = 0
        Dim SumGG3 As Double = 0
        Dim SumGG4 As Double = 0

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'Dim SalesFeesPurchReversCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double                 'ftP - Probability of Life Estate Purchase Between t and t+1
        Dim SalesFeesPurchDeathCF(ProgramParameters_FMV.ProjectedLife_t + 6) As Double

        PerformLoopedOperation_B(MaintTotal, RepairsTotal, TaxesTotal, InsurTotal, MoneyFactor, PBeingPresentNT, ReversEx, ResaleEx, MoEstateMgtFeesCum, MoPmtMgtFeesCum, MgrResaleFee, GrossGainReversPurch, GrossGainReversDeath, FeeMgrReversPurch, FeeMgrReversDeath, INVESTORofGSP, sub1, sub2, sub7, sub8, LEP, MoneyFactor0, PmtAmtCum)

        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t
            If t < 6 Then
                SalesFeesPurchReversCF(t) = 0
                SalesFeesPurchDeathCF(t) = 0
            ElseIf t > 5 Then
                SalesFeesPurchReversCF(t) = FeeMgrReversPurch(t) * PLEPurchBT(t)
                SalesFeesPurchDeathCF(t) = FeeMgrReversDeath(t - 6) * PReversBT(t)
            End If
        Next
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'DIFFERENCE CHECKING
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        TotMort = TotMort + PRDBT + PCBT + PLEP + PRBT   'total should equal 1,  see HUB
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        'INPUTS - WEBSITE CALCULATOR

        'In Public Class ContractTerms

        'FMV = FMVt0
        'BequestEquity = BequestEquity
        'Conserved% = ConservedEquityPct

        'BDay1								#xx/xx/xxxx#  - You'll need to convert it to Age
        'Sex1								"M" or "F"
        'BDay2								#xx/xx/xxxx#  - You'll need to convert it to Age	
        'Sex2								"M" or "F"

        'OUTPUTS - WEBSITE CALCULATOR
        Dim TotCashWS As Double
        Dim RMValueWS As Double
        Dim TIValueWS As Double
        Dim TotInHomeValue As Double
        Dim ConservedEqWS As Double
        Dim WalkAwayYr10WS As Double
        Dim TermValueWS As Double


        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim RValues As New WebSiteCalculatorOutputs
        With RValues
            .CashWS = InitialSolutionInitial
            .BequestWS = BEofGSP(0)
            .TotCashWS = InitialSolutionInitial + BEofGSP(0)
            .RMValueWS = APVMaint(0) + APVRepair(0)
            .TIValueWS = APVTaxes(0) + APVInsur(0)
            .TotInHomeValue = RMValueWS + TIValueWS
            .ConservedEqWS = CEofGSP(0)
            .WalkAwayYr10WS = LEP(120)
            .TermValueWS = ConservedEqWS + WalkAwayYr10WS
            .AdjPmtsWS = ScalarMonthly
            .TotalValueWS = TotCashWS + TotInHomeValue + TermValueWS
        End With
        Return RValues
    End Function

    Private Sub ProjectedLifeSeries(ReversArray5() As Double, ReversArray4() As Double, ReversArray3() As Double, ReversArray2() As Double, ReversArray1() As Double, ReversArrayUponSale() As Double, MoneyFactor() As Double, PPresent() As Double, PLEPurchBT() As Double, PBeingPresentNT() As Double, MoEstateMgtFeesCum() As Double, MoPmtMgtFeesCum() As Double, MgrResaleFee() As Double, MaintExCum() As Double, RepairExCum() As Double, TaxesExCum() As Double, InsurExCum() As Double, INVESTORofGSP() As Double, Bsub1() As Double, Bsub2() As Double, Bt() As Double, MoneyFactor0 As Double, APVRever5() As Double, APVRever4() As Double, APVRever3() As Double, APVRever2() As Double, APVRever1() As Double, APVReverTrans() As Double, APVMaint() As Double, APVRepair() As Double, APVTaxes() As Double, APVInsur() As Double, APVEstMgtFee() As Double, APVMgrSalesFeeRevers() As Double, APVInvGSP() As Double, btSumproduct1() As Double, btSumproduct2() As Double, btSumproduct3() As Double, Bt1() As Double, Bt2() As Double, Bt3() As Double, Bt4() As Double, Bt5() As Double, Bt6() As Double, Bt7() As Double, Bt8() As Double, Bt9() As Double, Bt10() As Double)
        For t = 0 To ProgramParameters_FMV.ProjectedLife_t - 1
            If t < 12 Then
                Bt(t) = 0
            Else
                Bt1(t) = (MoneyFactor0 ^ 0.5) * PLEPurchBT(t) / (PPresent(t + 1) * ((1 / LETenders.AntiSelection) - ManagerFees.PostReversNP * MoneyFactor0 ^ 6))
                Bt2(t) = (MoneyFactor0 ^ 0.5) * APVMaint(t + 1) + (MoneyFactor0 ^ 0.5) * APVRepair(t + 1) + (MoneyFactor0 ^ 0.5) * APVTaxes(t + 1) + (MoneyFactor0 ^ 0.5) * APVInsur(t + 1) + (MoneyFactor0 ^ 0.5) * APVEstMgtFee(t + 1)
                Bt3(t) = -(MoneyFactor0 ^ 0.5) * (1 - ManagerFees.PostReversNP) * ((MoneyFactor0 ^ 6) * APVInvGSP(t + 7) - (MoneyFactor0 ^ 6) * APVMgrSalesFeeRevers(t + 7) - (MoneyFactor0 ^ 6) * APVReverTrans(t + 7) - (MoneyFactor0 ^ 5) * APVRever1(t + 6) - (MoneyFactor0 ^ 4) * APVRever2(t + 5) - (MoneyFactor0 ^ 3) * APVRever3(t + 4) - (MoneyFactor0 ^ 2) * APVRever4(t + 3) - (MoneyFactor0 ^ 1) * APVRever5(t + 2))
                Bt4(t) = +(MoneyFactor0 ^ 6) * PPresent(t + 1) * (1 - ManagerFees.PostReversNP) * INVESTORofGSP(t + 6) - (MoneyFactor0 ^ 6) * PPresent(t + 1) * (1 - ManagerFees.PostReversNP) * MgrResaleFee(t + 6) - (MoneyFactor0 ^ 6) * PPresent(t + 1) * (1 - ManagerFees.PostReversNP) * ReversArrayUponSale(t + 6) - (MoneyFactor0 ^ 5) * PPresent(t + 1) * (1 - MoneyFactor0 ^ 1 * ManagerFees.PostReversNP) * ReversArray1(t + 6) - (MoneyFactor0 ^ 4) * PPresent(t + 1) * (1 - MoneyFactor0 ^ 2 * ManagerFees.PostReversNP) * ReversArray2(t + 6) - (MoneyFactor0 ^ 3) * PPresent(t + 1) * (1 - MoneyFactor0 ^ 3 * ManagerFees.PostReversNP) * ReversArray3(t + 6) - (MoneyFactor0 ^ 2) * PPresent(t + 1) * (1 - MoneyFactor0 ^ 4 * ManagerFees.PostReversNP) * ReversArray4(t + 6) - (MoneyFactor0 ^ 1) * PPresent(t + 1) * (1 - MoneyFactor0 ^ 5 * ManagerFees.PostReversNP) * ReversArray5(t + 6)
                Bt5(t) = -ManagerFees.PostReversNP * (1 / MoneyFactor(t)) * btSumproduct1(t)
                Bt6(t) = +ManagerFees.PostReversNP * (MoneyFactor0 ^ 6) * PPresent(t + 1) * (ManagerFees.InceptionExp + ManagerFees.AdvEquityAmt + Origination.CommissExp + MaintExCum(t) + RepairExCum(t) + TaxesExCum(t) + InsurExCum(t) + MoEstateMgtFeesCum(t) + MoPmtMgtFeesCum(t))
                Bt7(t) = -ManagerFees.PostReversNP * (1 / MoneyFactor(t)) * btSumproduct2(t)
                Bt8(t) = -((ManagerFees.PostReversNP / MoneyFactor(t)) / (MoneyFactor0 ^ 0.5)) * btSumproduct3(t)
                Bt9(t) = +ManagerFees.PostReversNP * (MoneyFactor0 ^ 6) * (PPresent(t + 1) / PBeingPresentNT(t + 1)) * Bsub1(t)
                Bt10(t) = +ManagerFees.PostReversNP * (MoneyFactor0 ^ 0.05) * (PPresent(t + 1) / (PBeingPresentNT(t + 1) * MoneyFactor(t + 1)) * Bsub2(t))

                Bt(t) = Bt1(t) * (Bt2(t) + Bt3(t) + Bt4(t) + Bt5(t) + Bt6(t) + Bt7(t) + Bt8(t) + Bt9(t) + Bt10(t))
            End If
        Next
    End Sub

    Private Sub PerformLoopedOperation_B(MaintTotal() As Double, RepairsTotal() As Double, TaxesTotal() As Double, InsurTotal() As Double, MoneyFactor() As Double, PBeingPresentNT() As Double, ReversEx() As Double, ResaleEx() As Double, MoEstateMgtFeesCum() As Double, MoPmtMgtFeesCum() As Double, MgrResaleFee() As Double, GrossGainReversPurch() As Double, GrossGainReversDeath() As Double, FeeMgrReversPurch() As Double, FeeMgrReversDeath() As Double, INVESTORofGSP() As Double, sub1() As Double, sub2() As Double, sub7() As Double, sub8() As Double, LEP() As Double, MoneyFactor0 As Double, PmtAmtCum() As Double)
        For t As Integer = 0 To ProgramParameters_FMV.ProjectedLife_t '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            If t < 6 Then
                GrossGainReversPurch(t) = 0                                                         'GtP - Gross Gain Upon Reversion to Sale Due to Purchase for Transactions From t to t+1
                GrossGainReversDeath(t) = 0                                                         'GtD - Gross Gain Upon Reversion to Sale Due to Death for Transactions From t to t+1
                FeeMgrReversPurch(t) = 0                                                            'FtGP - Fee Paid to Manager Upon Reversion to Sale Due to Purchase for Transactions From t to t+1
                FeeMgrReversDeath(t) = 0                                                            'FtGD - Fee Paid to Manager Upon Reversion to Sale Due to Death for Transactions From t to t+1
            ElseIf t > 5 Then
                'GtP - Gross Gain Upon Reversion to Sale Due to Purchase for Transactions From t to t+1
                GrossGainReversPurch(t) = INVESTORofGSP(t) - (ManagerFees.AdvEquityAmt + ReversEx(t) + ResaleEx(t) + MgrResaleFee(t) + LEP(t - 6) + ManagerFees.InceptionExp + Origination.CommissExp + MaintTotal(t - 6) + RepairsTotal(t - 6) + TaxesTotal(t - 6) + InsurTotal(t - 6) + MoEstateMgtFeesCum(t - 6) + MoPmtMgtFeesCum(t - 6) + PmtAmtCum(t - 6) + (1 / PBeingPresentNT(t - 5)) * sub1(t - 6) + (1 / PBeingPresentNT(t - 5)) * sub7(t - 6) + ((MoneyFactor0 ^ 0.05) / (PBeingPresentNT(t - 5) * MoneyFactor(t + 1)) * sub2(t - 6) + ((MoneyFactor0 ^ 0.05) / (PBeingPresentNT(t - 5) * MoneyFactor(t + 1)) * sub8(t - 6))))

                'GtD - Gross Gain Upon Reversion to Sale Due to Death for Transactions From t to t+1               
                GrossGainReversDeath(t) = (INVESTORofGSP(t) - (ManagerFees.AdvEquityAmt + (ReversEx(t) + ResaleEx(t)) + MgrResaleFee(t) + ManagerFees.InceptionExp + Origination.CommissExp + MaintTotal(t - 6) + RepairsTotal(t - 6) + TaxesTotal(t - 6) + InsurTotal(t - 6) - PmtAmtCum(t - 6)))

                'FtGP - Fee Paid to Manager Upon Reversion to Sale Due to Purchase for Transactions From t to t+1
                FeeMgrReversPurch(t) = ManagerFees.PostReversNP * GrossGainReversPurch(t)

                'FtGD - Fee Paid to Manager Upon Reversion to Sale Due to Death for Transactions From t to t+1
                FeeMgrReversDeath(t) = ManagerFees.PostReversNP * GrossGainReversDeath(t)
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Next
    End Sub

    Private Sub PerformLoopedOperation_A(ProbLastDeath() As Double, ProbOfCancellation() As Double, ProbBeingPresent() As Double, t As Integer)
        If t = 0 Then
            ProbBeingPresent(t) = 1
        Else                                                                                  'S    Probability of Being Present at t  
            ProbBeingPresent(t) = ProbBeingPresent(t - 1) * (1 - ProbOfCancellation(t - 1) - ProbLastDeath(t - 1))
        End If
    End Sub

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Public Class MovingSummary_Parameters
        Public Property SumGroupLength As Integer = 5
        Public Property StartingRowNumber As Integer = 13
        'Public Shared Property SumGroupLength As Integer = 5
        'Public Shared Property StartingRowNumber As Integer = 13
    End Class

    Public Class PaymentSolutionsMethod
        Public Shared Property ScalarMonthly = 1
        Public Shared Property ScalarNormBalloon = 1
        Public Shared Property InitialDollarInitial = 0
        Public Shared Property InitialDollarNormBalloon = 1
        Public Shared Property InitialSolutionMonthly = 0
        Public Shared Property InitialSolutionNormBalloon = 1
        Public Shared Property BalloonSolutionInitialMonthly = 10
        Public Shared Property BalloonSolutionMonthly = 1000
    End Class

    Public Class ContractTerms
        'Public Shared Property InceptionDate = #1/1/2016#
        'Public Shared Property FMVt0 = 1000000                                                    'FMV -  Fair Market Value at t=0 
        Public Shared Property BequestEquity = 10000   '$$                                         'B - Bequest Equity
        Public Shared Property ConservedEquityPct = 0.1                                            'αCE - Conserved Equity
        Public Shared Property ConservedEquityAmt = 0 '10000
        Public Shared Property NumMoPmts = 165                                                     'n - Number of Monthly Payments for IRS ALCI Note - NEED THE LOOKUP TABLE
        Public Shared Property PmtSolMeth = "S"    'Payment Solutions Method - S = Scalar, ID = Initial Dollar, IS = Initial Solutions, BS = Balloon Solution

        Public Shared Property BDay1 = #1/1/1943#   'Primary
        Public Shared Property Sex1 = "F"
        Public Shared Property SmokerStatus1 = "N" 'N = NonSmoker, Y = Smoker, C = Cancer????

        Public Shared Property BDay2 = #1/1/1943#   'Primary
        Public Shared Property Sex2 = "M"
        Public Shared Property SmokerStatus2 = "N" 'N = NonSmoker, Y = Smoker, C = Cancer????
    End Class

    Public Class ManagerFees
        Public Shared Property AnnualIRR = 0.08

        Public Shared Property Inception = 0.01                                                        'αMI -Fee on Inception of FMV
        Public Shared Property InceptionExp = ManagerFees.Inception * ProgramParameters_FMV.FMV_at_t0                                                    'EI - Inception Expense  
        Public Shared Property MoEStateMgt = 10.0                                                    'P -  Monthly Estate Management Fee
        Public Shared Property MoPmtMgt = 80.0                                                       'E -  Monthly Payment Management Fee
        Public Shared Property EquityShare = 0.1                                                      'αMS -Equity Share 
        Public Shared Property AdvEquityShare = 0.03                                                   'αMA - Advance on Equity Share, proportion of initial FMV

        Public Shared Property AdvEquityAmt = ManagerFees.AdvEquityShare * (ProgramParameters_FMV.FMV_at_t0 - ContractTerms.BequestEquity) * (1 - ContractTerms.ConservedEquityPct)

        Public Shared Property PostReversGSP = 0.02                                                    'αMS - Fee on Post Reversion Gross Sales Proceeds    GSP = Gross Sales Price
        Public Shared Property PostReversNP = 0.05                                                     'αMG - Fee on Post Reversion Net Gains   NP = Net Profits
        Public Shared Property CancellationPct = 0.65                                                  'αMC - Fee on the Cancellation Fee
        Public Shared Property CancellationFeePct = 0.01
        Public Shared Property CancellationFee = ManagerFees.CancellationFeePct * (1 - ContractTerms.ConservedEquityPct) * (ProgramParameters_FMV.FMV_at_t0 - ContractTerms.BequestEquity)                                            'C -   Cancellation Fee


        Public Shared Property AnnualRepChrgToEstate = 0.005                                           'i - Annual Repair Charge to Estate
        Public Shared Property MonthlyRepChrgToEstate = (1 + ManagerFees.AnnualRepChrgToEstate) ^ (1 / 12) - 1

    End Class

    Public Class DatesAll
        Public Shared Property DateInception As Date = #1/1/2016#
        Public Shared Property DateRevalLast As Date = #1/1/2016#
        Public Shared Property DateRevalNew As Date = #4/1/2016#
    End Class

    Public Class ProgramParameters_FMV                                               'ProgramParameters_FMV
        Public Shared Property FMV_at_t0 As Decimal = 1000000
        Public Shared Property Annual_HPA As Decimal = 0.034
        Public Shared Property Monthly_Factor_HPA As Decimal = 1 / (1 + Annual_HPA) ^ (1 / 12)
        Public Shared Property Temporary_Deviation_HPA As Decimal = 0.0
        Public Shared Property First_Month_Deviation_Recovery_HPA As Decimal = 0.01
        Public Shared Property Month_When_Deviation_Remaining_Ends_HPA As Integer = 60
        Public Shared Property Deviation_Remaining_HPA As Decimal = 0.05
        Public Shared Property ProjectedLife_t As Integer = 500

        ''Public Class Tx  'term
        '	Public Shared Property t0 = 0
        '	Public Shared Property t6 = 6
        '	Public Shared Property t12 = 12

        Public Shared Property FMVt() As Single
    End Class

    Public Class FMV                                                                  'FMV
        Public Property t As Integer
        Public Property Year As Integer
        Public Property HomeValue As Decimal
        Public Property PreviousHomeValue As Decimal
        Public Property Simple_Proj_t As Decimal
        Public Property Biased_Proj_t As Decimal
        Public Property Weight_Simple_Proj_t As Decimal
        Public Property FMV_t As Decimal
        Public Property AvgFMV_t_t1 As Decimal
        Public Property Log As Object
        Public Shared Property FMVt() As Single
    End Class

    Public Class HomeStatistics                                                       'HomeStatistics
        Public Shared Property HomeSize As Integer = 1800
        Public Shared Property LotSize As Integer = 2500
        Public Shared Property NumberOfStories As Integer = 1
        Public Shared Property HomeAge As Integer = 40
        Public Shared Property Bedrooms As Integer = 3
        Public Shared Property Bathrooms As Integer = 2
        Public Shared Property Totalrooms As Integer = 5

        Public Shared Property HomeType As String = "Detached" 'Add Condo, Co-op, DetachedSingle, Duplex, PUD, ZeroLotLine
        Public Shared Property RoofType As String = "Composite" 'Add Asphalt, BuiltUp, CompositeShingle, Metal, SinglePlyBitumen, Slate, Tile, WoodShake
        Public Shared Property ExteriorType As String = "Stucco"  'Add AluminumVinyl, Brick, Stone, Stucco, Wood

        Public Shared Property FMVt0 As Single = 500000
        Public Shared Property CPI As Decimal = 0.03
    End Class

    Public Class Origination                                                          'Origination
        Public Shared Property CommissOnContract As Decimal = 0.02
        Public Shared Property CommissOnPotential As Decimal = 0.005
        Public Shared Property CommissExp As Decimal = Origination.CommissOnContract * (1 - ContractTerms.ConservedEquityPct) * (ProgramParameters_FMV.FMV_at_t0 - ContractTerms.BequestEquity) + Origination.CommissOnPotential * (ContractTerms.ConservedEquityPct * ProgramParameters_FMV.FMV_at_t0 + (1 - ContractTerms.ConservedEquityPct) * ContractTerms.BequestEquity)
    End Class


    Public Class ClosingCosts                                                         'ClosingCosts
        Public Property TitleAcq As Single = 5.0  'per thousand
        Public Property EscAcq As Single = 3.0    'per thousand

        Public Property MiscAcq As Single = 350
        Public Property TermiteAcq As Single = 300
        Public Property HomeInspAcq As Single = 250
        Public Property AppraisalAcq As Single = 400
    End Class

    Public Class ApprecHPA                                                           'HPA = Home Price Appreciation
        Public Property HPAYr1 As Decimal = 0.034
        Public Property HPAYr2 As Decimal = 0.034
        Public Property HPAYr3 As Decimal = 0.034
        Public Property HPAYr4 As Decimal = 0.034
        Public Property HPAYr5 As Decimal = 0.034
        Public Property HPAYr6 As Decimal = 0.034

        Public Property HPAAdjYr1 As Decimal = 0.001
        Public Property HPAAdjYr2 As Decimal = 0.001
        Public Property HPAAdjYr3 As Decimal = 0.001
        Public Property HPAAdjYr4 As Decimal = 0.001
        Public Property HPAAdjYr5 As Decimal = 0.001
        Public Property HPAAdjYr6 As Decimal = 0.001
    End Class

    Public Class TaxInsur                                                            'Taxes and Insurance

        'Taxes
        Public Shared Property TaxAsPctOfFMV As Decimal = 0.0075
        Public Shared Property TaxBegMo As Integer = 5
        Public Shared Property TaxInterval As Integer = 4
        Public Shared Property TaxInitialHalf As Decimal = 0.0075 / 2
        Public Shared Property TaxAnnualIncrease As Decimal = 1.02

        'Insurance
        Public Shared Property InsurPctOfGSP As Decimal = 0.0018                                     'percent of Gross Sales Price
        Public Shared Property InsurBegMo As Integer = 2                                             'Insur beginning month
    End Class

    Public Class RepairMaint                                                        'Repairs and Maintenance

        'Preventive Maintenance
        Public Shared Property PrevMaintAmt As Single = 200.0
        Public Shared Property AppraisalPerYrs As Integer = 3
        Public Shared Property AppraisalCosts As Single = 35.0
        Public Shared Property AppraisalUpdate As Single = 10

        'Repairs
        Public Shared Property ExtPaintPerSqFt As Decimal = 0.68                             'per square foot
        Public Shared Property RoofPerSqFt As Decimal = 1.52                                 'per square foot
        Public Shared Property ElectricalPerSqFt As Decimal = 0.26                           'per square foot
        Public Shared Property PlumbingPerSqFt As Decimal = 0.26                             'per square foot
        Public Shared Property StructuralPerSqFt As Decimal = 0.21                           'per square foot
    End Class

    Public Class ReversionCosts                                                     'Reversion
        Public Shared Property ResaleInMo As Integer = 6                                     'Resale to happen in X months
        Public Shared Property InsurPerSqFt As Decimal = 0.5                                 'per square foot
        Public Shared Property HODues As Single = 150.0                                      'per month
        Public Shared Property Utilities As Single = 80                                      'per month
        Public Shared Property RETax As Decimal = 0.0075                                     'annual new RE tax rate
        Public Shared Property LandscapeMaintPerSqFt As Decimal = 0.16                       'per square foot
        Public Shared Property LandscapeMaintMin As Single = 75.0                            'Landscape Minimum
        Public Shared Property CleaningPerSqFt As Decimal = 0.16                             'per square foot
        Public Shared Property HVACPerSqFt As Decimal = 0.16                                 'per square foot
        Public Shared Property InteriorPaintPerSqFt As Decimal = 1.05                        'per square foot
        Public Shared Property FloorCoveringsPerSqFt As Decimal = 0.91                       'per square foot
        Public Shared Property WindowCoveringsPerSqFt As Decimal = 0.11                      'per square foot
        Public Shared Property LandscapingPerSqFt As Decimal = 0.15                          'per square foot
        Public Shared Property DamagesPerSqFt As Decimal = 0.16                              'per square foot
        Public Shared Property PavingPerSqFt As Decimal = 0.16                               'per square foot
    End Class

    Public Class Resale                                                            'Resale
        Public Shared Property CommissPct As Decimal = 0.05                                  'Commission percent
        Public Shared Property MarketingPct As Decimal = 0.02                                'marketing percent
        Public Shared Property TitlePctPerThous As Single = 6.0                              'per month
        Public Shared Property TransTaxPctPerThous As Single = 6.0                           'per month
        Public Shared Property EscrClosingPctPerThous As Decimal = 3.0                       'per 
        Public Shared Property LegalClosing As Single = 250.0
    End Class

    Public Class PricingDefaults                                                   'PricingDefaults

        'Roof type                                                                     Roof type
        Public Shared Property WoodShakeLE As Decimal = 25                                   'Wood shake Life Expectancy
        Public Shared Property WoodShakePct As Decimal = 2                                   'Wood shake percent of Standard
        Public Shared Property BuiltUPLE As Decimal = 10                                     'Built Up -Tar and Gravel Life Expectancy
        Public Shared Property BuiltUpPct As Decimal = 1                                     'Built Up -Tar and Gravel percent of Standard
        Public Shared Property CompositeShingleLE As Decimal = 20                            'Composite Life Expectancy
        Public Shared Property CompositePct As Decimal = 1                                   'Composite percent of Standard
        Public Shared Property SinglePlyBitumenLE As Decimal = 30                            'Bitumen Life Expectancy 
        Public Shared Property SinglePlyBitumenPct As Decimal = 1.5                          'Bitumen percent of Standard
        Public Shared Property AsphaltLE As Decimal = 5                                      'Asphalt Life Expectancy
        Public Shared Property AsphaltPct As Decimal = 1                                     'Asphalt percent of Standard
        Public Shared Property TileLE As Decimal = 50                                        'Tile Life Expectancy
        Public Shared Property TilePct As Decimal = 0.25                                     'Tile percent of Standard
        Public Shared Property SlateLE As Decimal = 50                                       'Slate Life Expectancy
        Public Shared Property SlatePct As Decimal = 0.25                                    'Slate percent of Standard
        Public Shared Property MetalLE As Decimal = 50                                       'Metal Life Expectancy
        Public Shared Property MetalPct As Decimal = 0.25                                    'Metal percent of Standard

        'Exterior Finish                                                               Exterior Finish
        Public Shared Property StuccoLE As Decimal = 5                                       'Stucco Life Expectancy
        Public Shared Property StuccoPct As Decimal = 1                                      'Stucco percent of Standard
        Public Shared Property WoodLE As Decimal = 50                                        'Wood Life Expectancy
        Public Shared Property WoodPct As Decimal = 1                                        'Wood percent of Standard
        Public Shared Property BrickLE As Decimal = 25                                       'Brick Life Expectancy
        Public Shared Property BrickPct As Decimal = 0.5                                     'Brick percent of Standard
        Public Shared Property StoneLE As Decimal = 25                                       'Stone Life Expectancy
        Public Shared Property StonePct As Decimal = 0.5                                     'Stone percent of Standard
        Public Shared Property AlumVinylLE As Decimal = 100                                  'Aluminum/Vinyl Life Expectancy
        Public Shared Property AlumVinylPct As Decimal = 0.25                                'Aluminum/Vinyl of Standard
    End Class

    Public Class PricingDefaultMultiplier                                         'Resale PricingDefaultMultiplier

        'Preventive Maintenance                                                       Preventive Maintenance
        Public Shared Property PrevMaintAmtT As Decimal = 1                                 'Type
        Public Shared Property PrevMaintAmtL As Decimal = 1                                 'Locale
        Public Shared Property PrevMaintAmtA As Decimal = 1                                 'Age
        Public Shared Property PrevMaintAmtG As Decimal = 1                                 'Grade

        Public Shared Property AppraisalCostsT As Decimal = 1                               'Type
        Public Shared Property AppraisalCostsL As Decimal = 1                               'Locale
        Public Shared Property AppraisalCostsA As Decimal = 1                               'Age
        Public Shared Property AppraisalCostsG As Decimal = 1                               'Grade

        Public Shared Property AppraisalUpdateT As Decimal = 1                              'Type
        Public Shared Property AppraisalUpdateL As Decimal = 1                              'Locale
        Public Shared Property AppraisalUpdateA As Decimal = 1                              'Age
        Public Shared Property AppraisalUpdateG As Decimal = 1                              'Grade

        'Repairs                                                                       Repairs
        Public Shared Property ExtPaintT As Decimal = 1                                     'Type
        Public Shared Property ExtPaintL As Decimal = 1                                     'Locale
        Public Shared Property ExtPaintA As Decimal = 1                                     'Age
        Public Shared Property ExtPaintG As Decimal = 1                                     'Grade
        Public Shared Property ExtPaintRBal As Integer = 5                                  'Remaining Balance

        Public Shared Property RoofT As Decimal = 1                                         'Type
        Public Shared Property RoofL As Decimal = 1                                         'Locale
        Public Shared Property RoofA As Decimal = 1                                         'Age
        Public Shared Property RoofG As Decimal = 1                                         'Grade
        Public Shared Property RoofRBal As Integer = 15                                     'Remaining Balance

        Public Shared Property ElectricalT As Decimal = 1                                   'Type
        Public Shared Property ElectricalL As Decimal = 1                                   'Locale
        Public Shared Property ElectricalA As Decimal = 1                                   'Age
        Public Shared Property ElectricalG As Decimal = 1                                   'Grade
        Public Shared Property ElectricalLife As Integer = 50                               'Life
        Public Shared Property ElectricalRBal As Integer = 25                               'Remaining Balance

        Public Shared Property PlumbingT As Decimal = 1                                     'Type
        Public Shared Property PlumbingL As Decimal = 1                                     'Locale
        Public Shared Property PlumbingA As Decimal = 1                                     'Age
        Public Shared Property PlumbingG As Decimal = 1                                     'Grade
        Public Shared Property PlumbingLife As Integer = 50                                 'Life
        Public Shared Property PlumbingRBal As Integer = 25                                 'Remaining Balance

        Public Shared Property StructuralT As Decimal = 1                                   'Type
        Public Shared Property StructuralL As Decimal = 1                                   'Locale
        Public Shared Property StructuralA As Decimal = 1                                   'Age
        Public Shared Property StructuralG As Decimal = 1
        Public Shared Property StructuralLife As Integer = 50
        Public Shared Property StructuralRBal As Integer = 25 'Grade

        'Reversionary Costs                                                            Reversionary Costs
        Public Shared Property CleaningT As Decimal = 1                                     'Type
        Public Shared Property CleaningL As Decimal = 1                                     'Locale
        Public Shared Property CleaningA As Decimal = 1                                     'Age
        Public Shared Property CleaningG As Decimal = 1                                     'Grade

        Public Shared Property HVACT As Decimal = 1                                         'Type
        Public Shared Property HVACL As Decimal = 1                                         'Locale
        Public Shared Property HVACA As Decimal = 1                                         'Age
        Public Shared Property HVACG As Decimal = 1                                         'Grade

        Public Shared Property InteriorPaintT As Decimal = 1                                'Type
        Public Shared Property InteriorPaintL As Decimal = 1                                'Locale
        Public Shared Property InteriorPaintA As Decimal = 1                                'Age
        Public Shared Property InteriorPaintG As Decimal = 1                                'Grade

        Public Shared Property FloorCoveringsT As Decimal = 1                               'Type
        Public Shared Property FloorCoveringsL As Decimal = 1                               'Locale
        Public Shared Property FloorCoveringsA As Decimal = 1                               'Age
        Public Shared Property FloorCoveringsG As Decimal = 1                               'Grade

        Public Shared Property WindowCoveringsT As Decimal = 1                              'Type
        Public Shared Property WindowCoveringsL As Decimal = 1                              'Locale
        Public Shared Property WindowCoveringsA As Decimal = 1                              'Age
        Public Shared Property WindowCoveringsG As Decimal = 1                              'Grade

        Public Shared Property LandscapingT As Decimal = 1                                  'Type
        Public Shared Property LandscapingL As Decimal = 1                                  'Locale
        Public Shared Property LandscapingA As Decimal = 1                                  'Age
        Public Shared Property LandscapingG As Decimal = 1                                  'Grade

        Public Shared Property DamagesT As Decimal = 1                                      'Type
        Public Shared Property DamagesL As Decimal = 1                                      'Locale
        Public Shared Property DamagesA As Decimal = 1                                      'Age
        Public Shared Property DamagesG As Decimal = 1                                      'Grade

        Public Shared Property PavingT As Decimal = 1                                       'Type
        Public Shared Property PavingL As Decimal = 1                                       'Locale
        Public Shared Property PavingA As Decimal = 1                                       'Age
        Public Shared Property PavingG As Decimal = 1                                       'Grade

    End Class

    Public Class myhome
        Public Property HomeSize As Integer = 1800
        Public Property LotSize As Integer = 2500
        Public Property NumberOfStories As Integer = 1
        Public Property HomeAge As Integer = 40
        Public Property Bedrooms As Integer = 3
        Public Property Bathrooms As Integer = 2
        Public Property Totalrooms As Integer = 5
        Public Property HomeType As String = HomeTypes.DetachedSingle
        Public Property RoofType As String = RoofTypes.CompositeShingle
        Public Property ExteriorType As String = ExteriorTypes.Stucco
        Public Property FMVt0 As Single = 500000
        Public Property CPI As Decimal = 0.03
        Public Property REFactor As Integer = 1
        Public Property ResalePct As Decimal = Resale.CommissPct + Resale.MarketingPct
        Public Property PrevMaintM As Decimal = RepairMaint.PrevMaintAmt * (PricingDefaultMultiplier.PrevMaintAmtT + PricingDefaultMultiplier.PrevMaintAmtL + PricingDefaultMultiplier.PrevMaintAmtA + PricingDefaultMultiplier.PrevMaintAmtG - 3) 'M = the multiplier value
        Public Property AppraisalCostsM As Decimal = RepairMaint.AppraisalCosts * (PricingDefaultMultiplier.AppraisalCostsT + PricingDefaultMultiplier.AppraisalCostsL + PricingDefaultMultiplier.AppraisalCostsA + PricingDefaultMultiplier.AppraisalCostsG - 3)  'M = the multiplier value
        Public Property AppraisalUpdateM As Decimal = RepairMaint.AppraisalUpdate * (PricingDefaultMultiplier.AppraisalUpdateT + PricingDefaultMultiplier.AppraisalUpdateL + PricingDefaultMultiplier.AppraisalUpdateA + PricingDefaultMultiplier.AppraisalUpdateG - 3)  'M = the multiplier value
    End Class

    Public Class Cancellations  'Monthly
        Public Shared Property CancelMonth
        Public Shared Property CancelMonth0 As Decimal = 0.004
        Public Shared Property CancelMonth1 As Decimal = 0.003
        Public Shared Property CancelMonth2 As Decimal = 0.002
        Public Shared Property CancelMonth3 As Decimal = 0.001
        Public Shared Property CancelMonth4 As Decimal = 0.002
        Public Shared Property CancelMonth5 As Decimal = 0.004
        Public Shared Property CancelReductionFactor As Decimal = 1.0
    End Class

    Public Class LETenders  'Life Estate Tender Rates - Annual
        Public Shared Property TenderYr
        Public Shared Property TenderYr0 As Decimal = 0.02
        Public Shared Property TenderYr1 As Decimal = 0.021
        Public Shared Property TenderYr2 As Decimal = 0.022
        Public Shared Property TenderYr3 As Decimal = 0.023
        Public Shared Property TenderYr4 As Decimal = 0.024
        Public Shared Property TenderYr5 As Decimal = 0.025
        Public Shared Property TenderYr6 As Decimal = 0.026
        Public Shared Property TenderYr7 As Decimal = 0.027
        Public Shared Property TenderYr8 As Decimal = 0.028
        Public Shared Property TenderYr9 As Decimal = 0.029
        Public Shared Property TenderYr10 As Decimal = 0.03
        Public Shared Property TenderYr11 As Decimal = 0.03
        Public Shared Property TenderYr12 As Decimal = 0.03
        Public Shared Property TenderYr13 As Decimal = 0.03
        Public Shared Property TenderYr14 As Decimal = 0.03
        Public Shared Property TenderYr15 As Decimal = 0.03
        Public Shared Property TenderYr16 As Decimal = 0.03
        Public Shared Property TenderYr17 As Decimal = 0.03
        Public Shared Property TenderYr18 As Decimal = 0.03
        Public Shared Property TenderYr19 As Decimal = 0.03
        Public Shared Property ReductionFactorLE As Decimal = 1.0
        Public Shared Property LEProbOfReneging As Decimal = 0.15

        Public Shared Property AntiSelection As Decimal = 0.9

    End Class

    Public ReadOnly Property EPLife() As Decimal
        Get
            Select Case HomeStatistics.ExteriorType
                Case ExteriorTypes.Stucco
                    Return PricingDefaults.StuccoLE * 12
                Case ExteriorTypes.Wood
                    Return PricingDefaults.WoodLE * 12
                Case ExteriorTypes.Stone
                    Return PricingDefaults.StoneLE * 12
                Case ExteriorTypes.Brick
                    Return PricingDefaults.BrickLE * 12
                Case ExteriorTypes.AluminumVinyl
                    Return PricingDefaults.AlumVinylLE * 12
                Case Else
                    Return 0
            End Select
        End Get
    End Property

    Public ReadOnly Property RFLife() As Decimal
        Get
            Select Case HomeStatistics.RoofType
                Case RoofTypes.Asphalt
                    Return PricingDefaults.AsphaltLE * 12
                Case RoofTypes.BuiltUp
                    Return PricingDefaults.BuiltUPLE * 12
                Case RoofTypes.CompositeShingle
                    Return PricingDefaults.CompositeShingleLE * 12
                Case RoofTypes.Metal
                    Return PricingDefaults.MetalLE * 12
                Case RoofTypes.SinglePlyBitumen
                    Return PricingDefaults.SinglePlyBitumenLE * 12
                Case RoofTypes.Slate
                    Return PricingDefaults.SlateLE * 12
                Case RoofTypes.Tile
                    Return PricingDefaults.TileLE * 12
                Case RoofTypes.WoodShake
                    Return PricingDefaults.WoodShakeLE * 12
                Case Else
                    Return 0
            End Select
        End Get
    End Property

    Public Class RM
        Public Shared Property EPBal As Integer = PricingDefaultMultiplier.ExtPaintRBal * 12                                        'Exterior Paint remaining balance in months
        Public Shared Property EPArea As Decimal = 40 * HomeStatistics.NumberOfStories * Math.Sqrt(HomeStatistics.HomeSize)         'Exterior Paint area in SqFt
        Public Shared Property RFBal As Integer = PricingDefaultMultiplier.RoofRBal * 12                                            'Roof remaining balance in months
        Public Shared Property RFArea As Decimal = 1.2 * HomeStatistics.HomeSize                                                    'Home size multiplier
        Public Shared Property ELLife As Integer = PricingDefaultMultiplier.ElectricalLife * 12                                     'Electircal life in months
        Public Shared Property ELBal As Integer = PricingDefaultMultiplier.ElectricalRBal * 12                                      'Electrical remaining balance
        Public Shared Property PLLife As Integer = PricingDefaultMultiplier.PlumbingLife * 12                                       'Plumbing life in months
        Public Shared Property PLBal As Integer = PricingDefaultMultiplier.PlumbingRBal * 12                                        'Plumbing remaining balance
        Public Shared Property STLife As Integer = PricingDefaultMultiplier.StructuralLife * 12
        Public Shared Property STBal As Integer = PricingDefaultMultiplier.StructuralRBal * 12
        Public Shared Property EstSqft As Integer = HomeStatistics.HomeSize   'Estimate home sqFt
        Public Shared Property EPCT As Decimal = EPArea * RepairMaint.ExtPaintPerSqFt * (PricingDefaultMultiplier.ExtPaintT + PricingDefaultMultiplier.ExtPaintL + PricingDefaultMultiplier.ExtPaintA + PricingDefaultMultiplier.ExtPaintG - 3)

        Public Shared Property RFCT As Decimal = RFArea * RepairMaint.RoofPerSqFt * (PricingDefaultMultiplier.RoofT + PricingDefaultMultiplier.RoofL + PricingDefaultMultiplier.RoofA + PricingDefaultMultiplier.RoofG - 3)
        Public Shared Property ELCT As Decimal = EstSqft * RepairMaint.ElectricalPerSqFt * (PricingDefaultMultiplier.ElectricalT + PricingDefaultMultiplier.ElectricalL + PricingDefaultMultiplier.ElectricalA + PricingDefaultMultiplier.ElectricalG - 3)
        Public Shared Property PLCT As Decimal = EstSqft * RepairMaint.PlumbingPerSqFt * (PricingDefaultMultiplier.PlumbingT + PricingDefaultMultiplier.PlumbingL + PricingDefaultMultiplier.PlumbingA + PricingDefaultMultiplier.PlumbingG - 3)
        Public Shared Property STCT As Decimal = EstSqft * RepairMaint.StructuralPerSqFt * (PricingDefaultMultiplier.StructuralT + PricingDefaultMultiplier.StructuralL + PricingDefaultMultiplier.StructuralA + PricingDefaultMultiplier.StructuralG - 3)

        Public Shared Property RepairInitial As Double = 1000
        Public Shared Property RepairBeginMo As Integer = 16
        Public Shared Property RepairInterval As Integer = 15
        Public Shared Property RepairIncreaseFactor As Double = 1.033

        Public Shared Property MaintInitial As Double = 250
        Public Shared Property MaintBeginMo As Integer = 15
        Public Shared Property MaintInterval As Integer = 15
        Public Shared Property MaintIncreaseFactor As Double = 1.033

        Public Shared Property Revers5 As Double = 0.0075
        Public Shared Property Revers4 As Double = 0.0075
        Public Shared Property Revers3 As Double = 0.003
        Public Shared Property Revers2 As Double = 0.001
        Public Shared Property Revers1 As Double = 0.001
        Public Shared Property ReversUponSale As Double = 0.052

        Public Shared Property RMTI As Integer = 1                     '1 = use BRE model for R&M, 0 = use TAGoren model for RMTI
    End Class
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''


    Public ReadOnly Property EPCT() As Decimal
        Get
            Dim base As Decimal = RM.EPArea * RepairMaint.ExtPaintPerSqFt * (PricingDefaultMultiplier.ExtPaintT + PricingDefaultMultiplier.ExtPaintL + PricingDefaultMultiplier.ExtPaintA + PricingDefaultMultiplier.ExtPaintG - 3)
            Select Case HomeStatistics.ExteriorType
                Case ExteriorTypes.Stucco
                    Return base * PricingDefaults.StuccoPct
                Case ExteriorTypes.Wood
                    Return base * PricingDefaults.WoodPct

                Case ExteriorTypes.Stone
                    Return base * PricingDefaults.StonePct

                Case ExteriorTypes.Brick
                    Return base * PricingDefaults.BrickPct

                Case ExteriorTypes.AluminumVinyl
                    Return base * PricingDefaults.AlumVinylLE
                Case Else
                    Return 0
            End Select
            'End Get

            If HomeStatistics.HomeType = HomeTypes.ZeroLotLine Then

                EPCT *= 0.5
            End If

            If HomeStatistics.HomeType = HomeTypes.Condo Then
                RM.EPCT = 0
                RM.RFCT = 0
                RM.ELCT = 0
                RM.PLCT = 0
                RM.STCT = 0
            End If

        End Get
    End Property


    Sub initializeRun(SetNumber As Integer)
        Select Case SetNumber
            Case 1
                LETenders.TenderYr = 3
            Case 2
                LETenders.TenderYr = 500

        End Select

    End Sub

    ''''''''''''''''''''''''''''''''''''''''''''''''''''




End Module