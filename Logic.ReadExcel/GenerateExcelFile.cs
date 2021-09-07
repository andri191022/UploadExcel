using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Logic.ReadExcel
{
    public class GenerateExcelFile
    {
        public static MemoryStream GenerateXLS(DataTable dt, string fileName)
        {
            // Declare HSSFWorkbook object for create sheet  
            var workbook = new HSSFWorkbook();
            var sheet = workbook.CreateSheet("Sheet1");
            var row = sheet.CreateRow(0);
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                var cell = row.CreateCell(i);
                cell.SetCellValue(dt.Columns[i].ColumnName);
                //string delimenter = i == col.Count - 1 ? "" : ",";
                //colName += col[i].ColumnName + delimenter;
            }


            //// Convert datatable into json  
            //string JSON = JsonConvert.SerializeObject(dt);

            //// Convert json into SummaryClass class list  
            //var items = JsonConvert.DeserializeObject<List<SummaryClass>>(JSON);

            //// Set column name this column name use for fetch data from list  
            //var columns = new[] { "ID", "Name" };

            //// Set header name this header use for set name in excel first row  
            //var headers = new[] { "ID", "Name" };

            //var headerRow = sheet.CreateRow(0);

            ////Below loop is create header  
            //for (int i = 0; i < columns.Length; i++)
            //{
            //    var cell = headerRow.CreateCell(i);
            //    cell.SetCellValue(headers[i]);
            //}

            ////Below loop is fill content  
            //for (int i = 0; i < items.Count; i++)
            //{
            //    var rowIndex = i + 1;
            //    var row = sheet.CreateRow(rowIndex);

            //    for (int j = 0; j < columns.Length; j++)
            //    {
            //        var cell = row.CreateCell(j);
            //        var o = items[i];
            //        cell.SetCellValue(o.GetType().GetProperty(columns[j]).GetValue(o, null).ToString());
            //    }
            //}

            // Declare one MemoryStream variable for write file in stream  
            var stream = new MemoryStream();
            workbook.Write(stream);

            return stream;
        }


        public static void SaveDataSetAsExcelX(DataTable dataTable, string exceloutFilePath)
        {
            using (var fs = new FileStream(exceloutFilePath, FileMode.CreateNew, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("Sheet1");
                List<string> columns = new List<string>();
                IRow row = excelSheet.CreateRow(0);
                int columnIndex = 0;

                foreach (System.Data.DataColumn column in dataTable.Columns)
                {
                    columns.Add(column.ColumnName);
                    row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                    columnIndex++;
                }

                int rowIndex = 1;
                foreach (DataRow dsrow in dataTable.Rows)
                {
                    row = excelSheet.CreateRow(rowIndex);
                    int cellIndex = 0;
                    foreach (String col in columns)
                    {
                        row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                        cellIndex++;
                    }

                    rowIndex++;
                }
                // return workbook;
                workbook.Write(fs);
            }
        }

    }
}
