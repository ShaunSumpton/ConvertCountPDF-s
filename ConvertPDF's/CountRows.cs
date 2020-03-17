using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace ConvertPDF_s
{
    class CountRows
    {

        public static void GetCountRows(string dir)
        {

            Application application = new Application();
            

            var files = new DirectoryInfo(dir);
            var XLFiles = Program.GetFilesByExtensions(files, ".CSV", ".xls", ".xlsx");

            object missing = System.Reflection.Missing.Value;
            foreach (FileInfo XLFile in XLFiles)
            {

                string filename = XLFile.FullName; // Full FileName including extentions and directory
                string filen = XLFile.Name; // get just file name
                string dirf = XLFile.DirectoryName; // get just directory

                Workbook ExcelDoc = application.Workbooks.Open(filename);
                Worksheet ws; // create worksheet
                application.Visible = true;

                ws = (Worksheet)ExcelDoc.Worksheets[0]; // worksheet assigned to 1st sheet in workbook

                int lastRow = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row; // find the last row 



                using (StreamWriter sw = new StreamWriter(dirf + @"\RowCount.txt", true))
                {

                    sw.WriteLine(filen + "| Number of Rows: " + lastRow); // name of page 1
                    sw.WriteLine(" ");

                }

            }



        }
    }
}
