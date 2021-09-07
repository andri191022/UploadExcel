using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Logic.ReadExcel
{
    public class GenerateScript
    {
        public static string GenerateQueryByRow(DataRow row, DataColumnCollection col, string tblName)
        {
            string strRetrun = string.Empty;
            if (row.ItemArray.Length != col.Count) { return strRetrun; }

            string colName = string.Empty; string rowVal = string.Empty;
            for (int i = 0; i < col.Count; i++)
            {
                string delimenter = i == col.Count - 1 ? "" : ",";
                colName += col[i].ColumnName + delimenter;
            }

            for (int i = 0; i < row.ItemArray.Length; i++)
            {
                string delimenter = i == row.ItemArray.Length - 1 ? "'" : "',";
                rowVal += "'" + row.ItemArray[i].ToString().Replace("'","''") + delimenter;
            }


            strRetrun = @"insert into " + tblName + "(" + colName + ") Values (" + rowVal + ")";

            return strRetrun;
        }


        public static StringBuilder GenerateQueryCreateTable(DataColumnCollection col, string tblName)
        {
            StringBuilder strRetrun = new StringBuilder();
            string cField = string.Empty;
            for (int i = 0; i < col.Count; i++)
            {
                string delimenter = i == col.Count - 1 ? "" : ",";
                cField += "[" + col[i].ColumnName + "] [nvarchar](255) NULL" + delimenter;
            }


            strRetrun.AppendLine("IF OBJECT_ID('[" + tblName + "]','u') IS NOT NULL " );
            strRetrun.AppendLine("DROP TABLE[dbo].[" + tblName + "];");
           // strRetrun.AppendLine("GO");

            strRetrun.AppendLine("Create Table " + tblName + "(");
            strRetrun.AppendLine(cField);
            strRetrun.Append(")  ON [PRIMARY] ");

            return strRetrun;
        }


    }
}
