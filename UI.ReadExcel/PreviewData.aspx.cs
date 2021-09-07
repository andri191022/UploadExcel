using Dta.ReadExcel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace UI.ReadExcel
{
    public partial class PreviewData : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                DataTable dtResult = new DataTable();
                string errMessage = string.Empty;

                string tblName = Request.QueryString["uq"];
                dtResult = GetTableData.GetAllDataByTableName(tblName, out errMessage);
                grdViewResult.DataSource = dtResult;
                grdViewResult.DataBind();
            }

        }

        protected void btnDownload_Click(object sender, EventArgs e)
        {

        }

        protected void btnBack_Click(object sender, EventArgs e)
        {
            Response.Redirect("~/Default.aspx");
        }

        protected void grdViewResult_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            DataTable dtResult = new DataTable();
            string errMessage = string.Empty;

            string tblName = Request.QueryString["uq"];
            dtResult = GetTableData.GetAllDataByTableName(tblName, out errMessage);
            grdViewResult.DataSource = dtResult; 
            grdViewResult.PageIndex = e.NewPageIndex;
            grdViewResult.DataBind();
        }
    }
}