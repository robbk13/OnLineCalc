Imports OnLineCalc.support
Imports OnLineCalc.MortalityProbabilitySeries
Public Class _default
    Inherits System.Web.UI.Page



    'Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    '    If IsPostBack = False Then
    '        Gethelp()
    '    End If
    'End Sub

    'Protected Sub calc_Click(sender As Object, e As EventArgs) Handles calc.Click


    'End Sub


    'Public Sub Gethelp()

    '    Dim db As New helpDataContext

    '    Dim Help As String = (From H In db.Contents Where H.PageName = "calculator" And H.Heading1 = "conserve" And H.PageType = "popup" Select H.Block).FirstOrDefault
    '    Me.help_conserve_Literal.Text = Help

    '    Dim Bequest As String = (From H In db.Contents Where H.PageName = "calculator" And H.Heading1 = "bequest" And H.PageType = "popup" Select H.Block).FirstOrDefault
    '    Me.help_bequest_Literal.Text = Bequest
    'End Sub



    'Public Sub AssignUserInputsToDefaultValues()
    '    ProgramParameters_FMV.FMV_at_t0 = CDbl(Me.txtfmv.Text)

    'End Sub
End Class