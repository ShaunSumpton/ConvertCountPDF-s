using System;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Collections.Generic;
using System.Linq;

namespace ConvertPDF_s
{
    static class Program
    {

       
        static void Main(string[] args)
        {




            // Create a new Microsoft Word application object
            Application word = new Application();

            // C# doesn't have optional arguments so we'll need a dummy value
            object oMissing = System.Reflection.Missing.Value;

            // Get list of Word files in specified directory
            
            Console.Write("Enter directory containing PDF's:");  // ask user for location of files
            string dir = Console.ReadLine();
            Console.WriteLine("");

            Console.WriteLine("1) Count Pages in PDF's");
            Console.WriteLine("2) Convert word to PDF's");
            Console.WriteLine("3) Split PDF's");
            Console.WriteLine("4) Import Delim File" + Environment.NewLine);
            // ask use to select an option
            Console.Write("Select Option e.g 1: ");
            int Option = Convert.ToInt32(Console.ReadLine().Trim());

            if (Option == 1)
            {
                CountPDFPages.GetPages(dir);
            }
            else if (Option == 3)
            {
                SplitPDFs.GetSplitPDFs(dir);

            }
            else if (Option == 2)
            {
                WordtoPDF.GetWordtoPDF(dir);
            }
            else

                DelimFile.GetDelimFile(dir);
                

           

        }

        public static IEnumerable<FileInfo> GetFilesByExtensions(this DirectoryInfo dirInfo, params string[] extensions)
        {
            var allowedExtensions = new HashSet<string>(extensions, StringComparer.OrdinalIgnoreCase);

            return dirInfo.EnumerateFiles()
                          .Where(f => allowedExtensions.Contains(f.Extension));
        }
            



    }


}
