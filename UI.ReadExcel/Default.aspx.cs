using Dta.ReadExcel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Logic.ReadExcel;
using System.Text;
using System.Net;
using NPOI.SS.UserModel;

namespace UI.ReadExcel
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                string errMesage = string.Empty;
                DataTable dt = GetTable.GetAll(out errMesage);

                grdViewResult.DataSource = dt;
                grdViewResult.DataBind();
            }
        }

        protected void grdViewResult_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            int rowIndex = Convert.ToInt32(e.CommandArgument);
            GridViewRow row = grdViewResult.Rows[rowIndex];
            string tblName = row.Cells[1].Text;
            DataTable dtResult = new DataTable();
            string errMessage = string.Empty; MemoryStream stream = new MemoryStream();
            if (e.CommandName == "Download")
            {
                dtResult = GetTableData.GetAllDataByTableName(tblName, out errMessage);

                stream = GenerateExcelFile.GenerateXLS(dtResult, tblName);
                string FilePath = Server.MapPath("~/tempResult/") + tblName + ".xlsx";

                if (File.Exists(FilePath))
                {
                    File.Delete(FilePath);
                }

                GenerateExcelFile.SaveDataSetAsExcelX(dtResult, FilePath);

                Response.Clear();
                Response.Buffer = true;
                Response.Charset = "";
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                Response.AddHeader("content-disposition", "attachment;filename=" + tblName + ".xlsx");
                Response.TransmitFile(FilePath);
                Response.Flush();
                Response.End();   
            }
            else
            {
                Response.Redirect("~/PreviewData.aspx?uq=" + tblName);
            }

            //  ClientScript.RegisterStartupScript(this.GetType(), "alert", "alert('Name: " + name + "\\nCountry: " + country + "');", true);
        }

        protected void btnProcess_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(fileUploadExcel.FileName))
            {
                string ext = Path.GetExtension(fileUploadExcel.FileName).ToLower();
                string path = Path.GetFileName(fileUploadExcel.FileName);
                string fileName = Path.GetFileNameWithoutExtension(fileUploadExcel.FileName);

                string errMessage = string.Empty; string eMessage = string.Empty;
                path = path.Replace(" ", "");
                fileUploadExcel.SaveAs(Server.MapPath("~/temp/") + path);
                DataTable dt = ReadData.Read(Server.MapPath("~/temp/") + path);
                string cSQL = string.Empty; StringBuilder cSQLB = new StringBuilder();
                if (dt.Rows.Count > 0)
                {
                    cSQLB = GenerateScript.GenerateQueryCreateTable(dt.Columns, fileName);

                    RunQueryData.RunX(cSQLB.ToString(), out errMessage);

                    foreach (DataRow row in dt.Rows)
                    {
                        cSQL = GenerateScript.GenerateQueryByRow(row, dt.Columns, fileName);

                        RunQueryData.RunX(cSQL, out errMessage);
                    }
                }

               
                DataTable dta = GetTable.GetAll(out eMessage);

                grdViewResult.DataSource = dta;
                grdViewResult.DataBind();

            }
        }

        protected void grdViewResult_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {

            string errMesage = string.Empty;
            DataTable dt = GetTable.GetAll(out errMesage);

            grdViewResult.DataSource = dt;
            grdViewResult.PageIndex = e.NewPageIndex;
            grdViewResult.DataBind();

        }
    }
}