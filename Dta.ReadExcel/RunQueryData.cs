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
   public class RunQueryData
    {
        public static void RunX(string cSQL, out string errMessage)
        {
            string conSTR = ConfigurationManager.ConnectionStrings["connReadExcel"].ToString();           
            try
            {
                using (var conn = new SqlConnection(conSTR))
                {
                    conn.Open();
                    var cmd = new SqlCommand(cSQL, conn);
                    cmd.ExecuteNonQuery();
                    conn.Close();
                }

                errMessage = string.Empty;
                
            }
            catch (Exception err)
            {
                errMessage = err.Message;      
               
            }
        }

    }
}
