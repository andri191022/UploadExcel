using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Dta.ReadExcel
{
    public class GetTable
    {
        public static DataTable GetAll(out string errMessage)
        {
            DataTable dtReturn = new DataTable();
            string conSTR = ConfigurationManager.ConnectionStrings["connReadExcel"].ToString();

            string cSQL = @"SELECT name as TableName FROM sys.Tables where type='U'";          
            try
            {
                using (var conn = new SqlConnection(conSTR))
                {
                    using (var cmd = new SqlCommand(cSQL, conn))
                    {
                        using (var da = new SqlDataAdapter(cmd))
                        {
                            cmd.CommandType = CommandType.Text;
                            da.Fill(dtReturn);
                        }
                    }
                }


                errMessage = string.Empty;
                return dtReturn;
            }
            catch (Exception err)
            {
                errMessage = err.Message;
                return dtReturn;
            }
        }
    }
}
