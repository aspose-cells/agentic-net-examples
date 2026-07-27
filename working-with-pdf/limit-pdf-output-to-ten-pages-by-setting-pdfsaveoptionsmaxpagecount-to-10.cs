using System;
using Aspose.Cells;

namespace AsposeCellsPdfLimitDemo
{
    // Author: Aspose.Cells .NET example – limit PDF output to ten pages
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Initialize PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Limit the output to a maximum of 10 pages.
            // The documented property for controlling page count is PageCount.
            pdfOptions.PageCount = 10;

            // Save the workbook as PDF with the specified options
            workbook.Save("output.pdf", pdfOptions);

            Console.WriteLine("PDF saved with a maximum of 10 pages.");
        }
    }
}