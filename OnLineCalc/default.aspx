<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="default.aspx.vb" Inherits="OnLineCalc._default" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
    <script>
        $(function () {
            $("#calc").tooltip();
        });
    </script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">


    <div class="row">
        <div class="col-md-3">

            <div class="panel panel-default">
                <div class="panel-heading">On-Line Calculator</div>
                <div class="panel-body">
                    <div class="well well-sm">
                        <a href="Calculator.aspx">
                            <img id="calc" data-toggle="tooltip" data-placement="right" src="calc.jpg" title="Calculate Your Life Estate" /></a>
                    </div>
                </div>
                <div class="col-md-6">
                </div>
            </div>
        </div>
    </div>
</asp:Content>
