<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="PreviewData.aspx.cs" Inherits="UI.ReadExcel.PreviewData" %>

<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <div class="panel panel-primary">
        <div class="panel-body">
            <div class="row">
                <div class="col-md-12">
                    <%--<asp:Button ID="btnDownload" runat="server" Text="Download" CssClass="btn btn-danger" OnClick="btnDownload_Click" />--%>
                    <asp:Button ID="btnBack" runat="server" Text="Back" CssClass="btn btn-danger" OnClick="btnBack_Click" />
                </div>
            </div>
        </div>
    </div>


    <div class="panel panel-primary">
        <div class="panel-heading">Data</div>
        <div class="panel-body">
            <asp:GridView ID="grdViewResult" runat="server" AllowPaging="true" PageSize="25" AutoGenerateColumns="true" GridLines="Both" CssClass="table table-striped table-bordered table-consendensed table-hover" OnPageIndexChanging="grdViewResult_PageIndexChanging">
                <PagerStyle CssClass="GridPager" />
                <HeaderStyle BackColor="Black" Font-Bold="true" ForeColor="White" />
            </asp:GridView>
        </div>
    </div>
</asp:Content>
