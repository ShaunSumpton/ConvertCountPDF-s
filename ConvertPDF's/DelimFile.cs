using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Excel;


namespace ConvertPDF_s
{
    class DelimFile
    {

        public static void GetDelimFile(string fpath)
        {

            var excelApp = new Application();
            var files = new DirectoryInfo(fpath);
            var CSVFiles = Program.GetFilesByExtensions(files, ".CSV");

            object missing = System.Reflection.Missing.Value;
            foreach (FileInfo CSVFile in CSVFiles)
            {
                if (CSVFile.FullName.Contains("EXP"))
                {
                    Object filename = (Object)CSVFile.FullName;
                    string jn = CSVFile.Name.ToString().Substring(0, 6);
                    string dirf = CSVFile.DirectoryName;


                    excelApp.Workbooks.OpenText(filename.ToString(), missing, 3, XlTextParsingType.xlDelimited, XlTextQualifier.xlTextQualifierNone, missing, missing, missing, true,
                     missing, missing, missing, missing, missing, missing, missing, missing, missing);
                    excelApp.Visible = true;
                    excelApp.ActiveWorkbook.SaveAs(dirf + @"\"+ jn + ".xlsx" );
                   
                }
                else
                {
                   // do nothing
                }

            }

            

            

        }
    }
}
