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
                string dirf = pdfFile.DirectoryName;

                PdfDocument pdfDoc = new PdfDocument(new PdfReader(filename));
                int numberOfPages = pdfDoc.GetNumberOfPages();

                using (StreamWriter sw = new StreamWriter(dirf + @"\counts.txt", true))
                {

                    sw.WriteLine(filen + ", Number of Pages: " + numberOfPages); // name of page 1
                    sw.WriteLine(" ");

                }
                    //Console.WriteLine(filen + "| Number of Pages: " + numberOfPages);

            }
        }




    }






    }

