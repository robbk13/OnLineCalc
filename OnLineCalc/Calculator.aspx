<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="Calculator.aspx.vb" Inherits="OnLineCalc.Calculator" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">


    <div class="row">
        <div class="col-md-2">
        </div>
        <div class="col-md-8">
            <div class="panel panel-default">
                <div class="panel-heading"><span class=" fa fa-id-badge fa-2x" style="color: #0026ff"></span>&nbsp;Your Information</div>
                <div class="panel-body">

                    <div class="row">
                        <div class="col-md-6">
                            <div class="input-group">
                                <span class="input-group-addon">First name:</span>
                                <asp:TextBox ID="fname" CssClass="form-control" runat="server"></asp:TextBox>
                            </div>
                            <div class="input-group">
                                <span class="input-group-addon">Last name:</span>
                                <asp:TextBox ID="lname" CssClass="form-control" runat="server"></asp:TextBox>
                            </div>
                        </div>
                        <div class="col-md-6">
                            <div class="input-group">
                                <span class="input-group-addon">Email:</span>
                                <asp:TextBox ID="email" TextMode="email" CssClass="form-control" runat="server"></asp:TextBox>
                            </div>
                            <div class="input-group">
                                <span class="input-group-addon">Date of birth:</span>
                                <asp:TextBox ID="dob" CssClass="form-control" runat="server"></asp:TextBox>
                            </div>
                            <div class="input-group">
                                <span class="input-group-addon">Gender:</span>
                                <asp:RadioButtonList ID="RadioButtonList1" CssClass="form-control" runat="server" RepeatDirection="Horizontal">
                                    <asp:ListItem Selected="True">Male</asp:ListItem>
                                    <asp:ListItem>Female</asp:ListItem>
                                </asp:RadioButtonList>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
            <div class="panel panel-default">
                <div class="panel-heading"><span style="color: #ff6a00;" class="fa fa-calculator fa-2x"></span>&nbsp;Your Life Estate</div>
                <div class="panel-body">

                    <div class="row">
                        <div class="col-md-6">
                            <div class="input-group">
                                <span class="input-group-addon">Home Value:</span>
                                <asp:TextBox ID="txtfmv" CssClass="form-control" runat="server"></asp:TextBox>
                            </div>

                            <div class="input-group">
                                <span class="input-group-addon">Bequest Equity:</span>
                                <asp:TextBox ID="bequest" CssClass="form-control" runat="server"></asp:TextBox>
                                <span class=" input-group-btn">
                                    <button type="button" class="btn btn-default" data-toggle="modal" data-target="#help_bequest">
                                        <span class=" fa fa-info-circle" style="color: #0026ff"></span>
                                    </button>
                                </span>
                            </div>

                            <div class="input-group">
                                <span class="input-group-addon">Conserve Percentage:</span>
                                <asp:TextBox ID="conserve" CssClass="form-control" runat="server"></asp:TextBox>
                                <span class=" input-group-btn">
                                    <button type="button" class="btn btn-default" data-toggle="modal" data-target="#help_conserve">
                                        <span class=" fa fa-info-circle" style="color: #0026ff"></span>
                                    </button>
                                </span>
                            </div>


                        </div>
                        <div class="col-md-6">
                            <div class="input-group">
                                <span class=" input-group-btn">
                                    <asp:Button ID="calcbtn" runat="server" CssClass="btn btn-default" Text="Calculate" />
                                </span>
                            </div>
                        </div>
                    </div>

                </div>
            </div>


            <div class="row">
                <div class="col-md-6">
                    <div class="panel panel-default">
                        <div class="panel-heading"><span class=" fa fa-dollar fa-2x" style="color: #156f20"></span>&nbsp;Estate Benefits:</div>
                        <div class="panel-body">
                            <div class="input-group">
                                <span class="input-group-addon">LumpSum Amount</span>
                                <asp:TextBox ID="CashWS" runat="server" ReadOnly="true" class="form-control text-right" placeholder="$"></asp:TextBox>
                            </div>
                            <div class="input-group">
                                <span class="input-group-addon">Bequest</span>
                                <asp:TextBox ID="bx" runat="server" ReadOnly="true" class="form-control text-right" placeholder="$"></asp:TextBox>
                            </div>
                            <div class="input-group">
                                <span class="input-group-addon">Total(Lumpsum + Bequest)</span>
                                <asp:TextBox ID="TotCashWS" runat="server" ReadOnly="true" class="form-control text-right" placeholder="$"></asp:TextBox>
                            </div>
                            <div class="input-group">
                                <span class="input-group-addon">Repairs and Maintenance</span>
                                <asp:TextBox ID="TIValueWS" runat="server" ReadOnly="true" class="form-control text-right" placeholder="$"></asp:TextBox>
                            </div>
                            <div class="input-group">
                                <span class="input-group-addon">Taxes and Insurance</span>
                                <asp:TextBox ID="tax" runat="server" ReadOnly="true" class="form-control text-right" placeholder="$"></asp:TextBox>
                            </div>
                            <div class="input-group">
                                <span class="input-group-addon">Total In-Home</span>
                                <asp:TextBox ID="tinhome" runat="server" ReadOnly="true" class="form-control text-right" placeholder="$"></asp:TextBox>
                            </div>
                            <div class="input-group">
                                <span class="input-group-addon">Conserved Equity Yr 10 </span>
                                <asp:TextBox ID="ce" runat="server" ReadOnly="true" class="form-control text-right" placeholder="$"></asp:TextBox>
                            </div>
                            <div class="input-group">
                                <span class="input-group-addon">Walking Away Yr 10 </span>
                                <asp:TextBox ID="walk" runat="server" ReadOnly="true" class="form-control text-right" placeholder="$"></asp:TextBox>
                            </div>
                            <div class="input-group">
                                <span class="input-group-addon">Total Termination </span>
                                <asp:TextBox ID="term" runat="server" ReadOnly="true" class="form-control text-right" placeholder="$"></asp:TextBox>
                            </div>
                            <div class="input-group">
                                <span class="input-group-addon">Monthly Pmt </span>
                                <asp:TextBox ID="adj" runat="server" ReadOnly="true" class="form-control text-right" placeholder="$"></asp:TextBox>
                            </div>
                            <div class="input-group">
                                <span class="input-group-addon">Total</span>
                                <asp:TextBox ID="Total" runat="server" ReadOnly="true" class="form-control text-right" placeholder="$"></asp:TextBox>
                            </div>
                        </div>
                    </div>

                </div>
                <div class="col-md-6">
                </div>
            </div>
        </div>

    </div>
    <div class="modal fade" id="help_conserve">
        <div class="modal-dialog" role="document">
            <div class="modal-content">
                <div class="modal-header">
                    <h3 class="modal-title">Conserve Percentage</h3>
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                        <span aria-hidden="true">&times;</span>
                    </button>
                </div>
                <div class="modal-body">
                    <asp:Literal ID="help_conserve_Literal" runat="server"></asp:Literal>
                </div>
            </div>
        </div>
    </div>

    <div class="modal fade" id="help_bequest">
        <div class="modal-dialog" role="document">
            <div class="modal-content">
                <div class="modal-header">
                    <h3 class="modal-title">Conserve Percentage</h3>
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                        <span aria-hidden="true">&times;</span>
                    </button>
                </div>
                <div class="modal-body">
                    <asp:Literal ID="help_bequest_Literal" runat="server"></asp:Literal>
                </div>
            </div>
        </div>
    </div>

</asp:Content>
