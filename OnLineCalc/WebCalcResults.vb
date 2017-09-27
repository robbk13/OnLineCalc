Public Class WebSiteInputs
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

    Public Property FMV As Double
    Public Property BequestEquity As Double
    Public Property ConservedPercentage As Double

    Public Property DateofBirth1 As Date
    Public Property Gender1 As String

    Public Property DateofBirth2 As Date
    Public Property Gender2 As String

    Private _Age1 As Integer
    Public ReadOnly Property Age1() As Integer
        Get


            Return DateDiff(DateInterval.Year, Me.DateofBirth1, Today)
        End Get
    End Property
    Public ReadOnly Property Age2() As Integer
        Get


            Return DateDiff(DateInterval.Year, Me.DateofBirth2, Today)
        End Get
    End Property
End Class
