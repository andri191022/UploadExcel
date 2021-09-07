<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="UI.ReadExcel._Default" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <div class="panel panel-primary">
        <div class="panel-body">
            <div class="row">
                <div class="col-md-2">
                    <asp:Label ID="Label1" runat="server" Text="Excel File"></asp:Label>
                </div>
                <div class="col-md-5">
                    <asp:FileUpload ID="fileUploadExcel" runat="server" CssClass="form-control" Width="95%" accept=".xls,.xlsx" />
                </div>
                <div class="col-md-5">
                    <asp:Button ID="btnProcess" runat="server" Text="Process" CssClass="btn btn-danger" OnClick="btnProcess_Click" />
                </div>
            </div>
        </div>
    </div>

    <div class="panel panel-primary">
        <div class="panel-heading">Data</div>
        <div class="panel-body">
            <asp:GridView ID="grdViewResult" runat="server" AllowPaging="true" PageSize="25" AutoGenerateColumns="true" GridLines="Both" CssClass="table table-striped table-bordered table-consendensed table-hover" OnRowCommand="grdViewResult_RowCommand" OnPageIndexChanging="grdViewResult_PageIndexChanging">
                <Columns>
                    <asp:TemplateField>
                        <ItemTemplate>
                            <asp:Button ID="btnPreview" Text="Preview" runat="server" CssClass="btn btn-danger" CommandName="Preview" CommandArgument="<%# Container.DataItemIndex %>" />
                            <asp:Button ID="btnDownload" Text="Download" runat="server" CssClass="btn btn-danger" CommandName="Download" CommandArgument="<%# Container.DataItemIndex %>" />
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
                <PagerStyle CssClass="GridPager" />
                <HeaderStyle BackColor="Black" Font-Bold="true" ForeColor="White" />
            </asp:GridView>
        </div>
    </div>
</asp:Content>
