using System;
using System.IO;
using iText.Kernel.Pdf;
using iText.Kernel.Utils;

namespace ConvertPDF_s
{
    class SplitPDFs
    {
        public static void GetSplitPDFs(string fpath)
        {

            var files = new DirectoryInfo(fpath); // get directory info from file path
            var pdfFiles = Program.GetFilesByExtensions(files, ".pdf"); // get all files with the extension of .PDF

            foreach (FileInfo pdfFile in pdfFiles)
            {

                string filename = pdfFile.FullName; // Full FileName including extentions and directory
                string filen = pdfFile.Name; // get just file name
                string dirf = pdfFile.DirectoryName; // get just directory

                


                using (PdfDocument pdfDoc = new PdfDocument(new PdfReader(filename)))
                {
                    var splitDocuments = new MyPdfSplitter(pdfDoc, dirf,filen).SplitByPageCount(1); // split PDF
                    foreach (var splitDocument in splitDocuments)
                    {
                        splitDocument.Close();
                    }
                }

             

            }

            Console.WriteLine("Operation Complete");
        }


        public class MyPdfSplitter : PdfSplitter // class to split PDF's
        {
            private readonly string _destFolder;
            private int _pageNumber;
            private readonly string _filename;
            public MyPdfSplitter(PdfDocument pdfDocument, string destFolder, string filen) : base(pdfDocument)
            {
                _destFolder = destFolder;
                _filename = filen;
            }

            protected override PdfWriter GetNextPdfWriter(PageRange documentPageRange)
            {
                _pageNumber++;
                return new PdfWriter(Path.Combine(_destFolder, $"p{_pageNumber} {_filename}.pdf")); // Page number and filename
            }
        }

    }
}
