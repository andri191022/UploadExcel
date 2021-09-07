using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Logic.ReadExcel
{
    public class ReadData
    {
        public static DataTable Read(string Path)
        {
            //XSSFWorkbook wb;
            ISheet sh;
            //String Sheet_name;

            //using (var fs = new FileStream(Path, FileMode.Open, FileAccess.Read))
            //{
            //    wb = new XSSFWorkbook(fs);

            //    Sheet_name = wb.GetSheetAt(0).SheetName;  //get first sheet name
            //}


            DataTable DT = new DataTable();
            DT.Rows.Clear();
            DT.Columns.Clear();

            // get sheet
            //sh = (XSSFSheet)wb.GetSheet(Sheet_name);

            sh = GetFileStream(Path);

            int i = 0;
            while (sh.GetRow(i) != null)
            {
                // add neccessary columns
                if (DT.Columns.Count < sh.GetRow(i).Cells.Count)
                {
                    for (int j = 0; j < sh.GetRow(i).Cells.Count; j++)
                    {

                        string sb_trim = sh.GetRow(i).GetCell(j).ToString().Replace(" ","").Replace(".","");
                        
                        DT.Columns.Add(sb_trim, typeof(string));
                    }
                }

                // add row
                DT.Rows.Add();

                // write row value
                for (int j = 0; j < sh.GetRow(i).Cells.Count; j++)
                {
                    var cell = sh.GetRow(i).GetCell(j);

                    if (cell != null)
                    {
                        // TODO: you can add more cell types capatibility, e. g. formula
                        switch (cell.CellType)
                        {
                            case NPOI.SS.UserModel.CellType.Numeric:
                                DT.Rows[i][j] = sh.GetRow(i).GetCell(j).NumericCellValue;
                                //dataGridView1[j, i].Value = sh.GetRow(i).GetCell(j).NumericCellValue;

                                break;
                            case NPOI.SS.UserModel.CellType.String:
                                DT.Rows[i][j] = sh.GetRow(i).GetCell(j).StringCellValue;

                                break;
                        }
                    }
                }

                i++;
            }

            DT.Rows[0].Delete();
            DT.AcceptChanges();

            return DT;
        }


        private static ISheet GetFileStream(string fullFilePath)
        {
            var fileExtension = Path.GetExtension(fullFilePath);
            string sheetName;
            ISheet sheet = null;
            switch (fileExtension)
            {
                case ".xlsx":
                    using (var fs = new FileStream(fullFilePath, FileMode.Open, FileAccess.Read))
                    {
                        var wb = new XSSFWorkbook(fs);
                        sheetName = wb.GetSheetAt(0).SheetName;
                        sheet = (XSSFSheet)wb.GetSheet(sheetName);
                    }
                    break;
                case ".xls":
                    using (var fs = new FileStream(fullFilePath, FileMode.Open, FileAccess.Read))
                    {
                        var wb = new HSSFWorkbook(fs);
                        sheetName = wb.GetSheetAt(0).SheetName;
                        sheet = (HSSFSheet)wb.GetSheet(sheetName);
                    }
                    break;
            }
            return sheet;
        }

    }
}
