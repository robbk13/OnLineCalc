Imports OnLineCalc.Calc
Public Class Calculator
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Me.dob.Text = "9/13/47"
            Me.txtfmv.Text = "250000"
            Me.bequest.Text = "10000"
        End If
    End Sub

    Private Sub calcbtn_Click(sender As Object, e As EventArgs) Handles calcbtn.Click

        Dim DateofBirth As Date = CDate(Me.dob.Text)
        Dim InitialAge1 As Integer = DateDiff(DateInterval.Year, DateofBirth, Today)

        Dim GenderSelection As constantGender
        If Me.RadioButtonList1.SelectedValue = "Male" Then
            GenderSelection = constantGender.M
        Else
            GenderSelection = constantGender.F
        End If

        ProgramParameters_FMV.FMV_at_t0 = CInt(Me.txtfmv.Text.Replace(",", ""))
        ContractTerms.BequestEquity = CInt(Me.bequest.Text.Replace(",", ""))
        Dim RVals As WebSiteCalculatorOutputs

        Dim MSeries As New MortalityProbabilitySeries
        With MSeries
            t = 0
            .InitialAge = InitialAge1
            .Gender = GenderSelection
            .SourceTableName = constantMortalityTableType.SS
            RVals = ProbNoTender(.MortalitySeries)
        End With

        Me.CashWS.Text = String.Format("{0:c}", RVals.CashWS)
        Me.bx.Text = String.Format("{0:c}", RVals.BequestWS)
        Me.TotCashWS.Text = String.Format("{0:c}", RVals.TotCashWS)
        Me.TIValueWS.Text = String.Format("{0:c}", RVals.RMValueWS)
        Me.tax.Text = String.Format("{0:c}", RVals.TIValueWS)
        Me.tinhome.Text = String.Format("{0:c}", RVals.TotInHomeValue)
        Me.ce.Text = String.Format("{0:c}", RVals.ConservedEqWS)
        Me.walk.Text = String.Format("{0:c}", RVals.WalkAwayYr10WS)
        Me.term.Text = String.Format("{0:c}", RVals.TermValueWS)
        Me.adj.Text = String.Format("{0:c}", RVals.AdjPmtsWS)
        Me.Total.Text = String.Format("{0:c}", RVals.TotCashWS + RVals.TotInHomeValue)

        RVals = Nothing
        MSeries = Nothing
        ProgramParameters_FMV.FMV_at_t0 = Nothing
        ContractTerms.BequestEquity = Nothing


    End Sub
End Class