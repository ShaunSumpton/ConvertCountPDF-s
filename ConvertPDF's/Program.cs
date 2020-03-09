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
            Console.WriteLine("3) Split PDF's" + Environment.NewLine);
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
            else
            {


                var files = new DirectoryInfo(dir);
                var wordFiles = GetFilesByExtensions(files, ".doc", ".docx"); // get all files in specified directory with .doc and .docx file extensions

                word.Visible = false;
                word.ScreenUpdating = false; // make sure work doesnt show

                foreach (FileInfo wordFile in wordFiles)
                {
                    // Cast as Object for word Open method
                    Object filename = (Object)wordFile.FullName;

                    // Use the dummy value as a placeholder for optional arguments
                    Document doc = word.Documents.Open(ref filename, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                    doc.Activate();

                    object outputFileName = wordFile.FullName; // fullname of word doc

                    if (wordFile.FullName.EndsWith(".doc"))
                    {
                        outputFileName = wordFile.FullName.Replace(".doc", ".pdf");
                    }
                    else
                    {
                        outputFileName = wordFile.FullName.Replace(".docx", ".pdf"); // changing extension
                    }


                    object fileFormat = WdSaveFormat.wdFormatPDF;


                    // Save document into PDF Format
                    doc.SaveAs(ref outputFileName,
                        ref fileFormat, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing); // saving 

                    // Close the Word document, but leave the Word application open.
                    // doc has to be cast to type _Document so that it will find the
                    // correct Close method.                
                    object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                    ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                    doc = null;


                }

            }
                    // word has to be cast to type _Application so that it will find
                    // the correct Quit method.
                    ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
                                    word = null;


           

        }

        public static IEnumerable<FileInfo> GetFilesByExtensions(this DirectoryInfo dirInfo, params string[] extensions)
        {
            var allowedExtensions = new HashSet<string>(extensions, StringComparer.OrdinalIgnoreCase);

            return dirInfo.EnumerateFiles()
                          .Where(f => allowedExtensions.Contains(f.Extension));
        }
            



    }


}
