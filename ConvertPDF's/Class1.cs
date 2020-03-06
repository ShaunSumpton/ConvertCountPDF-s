using System;
using System.IO;
using iText.Kernel.Pdf;

namespace ConvertPDF_s
{
    class CountPDFPages
    {

        public static void GetPages(string fpath)
        {
            // Right side of equation is location of YOUR pdf file
           

            var files = new DirectoryInfo(fpath);
            var pdfFiles = Program.GetFilesByExtensions(files, ".pdf");

            foreach (FileInfo pdfFile in pdfFiles)
            {

                string filename = pdfFile.FullName;
                string filen = pdfFile.Name;
                PdfDocument pdfDoc = new PdfDocument(new PdfReader(filename));
                int numberOfPages = pdfDoc.GetNumberOfPages();
                Console.WriteLine(filen + "| Number of Pages: " + numberOfPages);
                

            }
        }




    }






    }

