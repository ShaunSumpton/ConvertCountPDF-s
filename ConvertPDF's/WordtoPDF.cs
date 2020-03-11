using System;
using Microsoft.Office.Interop.Word;
using System.IO;


namespace ConvertPDF_s
{
    class WordtoPDF
    {
        public static void GetWordtoPDF(String dir)
        {
            // Create a new Microsoft Word application object
            Application word = new Application();

            // C# doesn't have optional arguments so we'll need a dummy value
            object oMissing = System.Reflection.Missing.Value;

            var files = new DirectoryInfo(dir);
            var wordFiles = Program.GetFilesByExtensions(files, ".doc", ".docx"); // get all files in specified directory with .doc and .docx file extensions

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

                // word has to be cast to type _Application so that it will find
                // the correct Quit method.
                ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
                    word = null;

        }
       


     }
}
