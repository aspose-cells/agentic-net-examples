using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsHtmlToPdfDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file
            string htmlPath = "input.html";

            // Path for the output PDF file
            string pdfPath = "output.pdf";

            // Load the HTML file into a Workbook instance
            // Aspose.Cells can directly load HTML documents
            Workbook workbook = new Workbook(htmlPath);

            // Create PDF save options (using the provided PdfSaveOptions constructor)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"HTML file '{htmlPath}' has been converted to PDF at '{pdfPath}'.");
        }
    }
}