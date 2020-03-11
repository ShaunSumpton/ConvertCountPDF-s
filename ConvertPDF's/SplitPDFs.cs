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

            var files = new DirectoryInfo(fpath);
            var pdfFiles = Program.GetFilesByExtensions(files, ".pdf");

            foreach (FileInfo pdfFile in pdfFiles)
            {

                string filename = pdfFile.FullName;
                string filen = pdfFile.Name;
                string dirf = pdfFile.DirectoryName;

                


                using (PdfDocument pdfDoc = new PdfDocument(new PdfReader(filename)))
                {
                    var splitDocuments = new MyPdfSplitter(pdfDoc, dirf,filen).SplitByPageCount(1);
                    foreach (var splitDocument in splitDocuments)
                    {
                        splitDocument.Close();
                    }
                }

             

            }

            Console.WriteLine("Operation Complete");
        }


        public class MyPdfSplitter : PdfSplitter
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
                return new PdfWriter(Path.Combine(_destFolder, $"p{_pageNumber} {_filename}.pdf"));
            }
        }

    }
}
