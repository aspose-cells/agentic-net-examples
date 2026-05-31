using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfMinimalist
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // creates a blank workbook
            // Example: load an existing file
            // Workbook workbook = new Workbook("input.xlsx");

            // Set each worksheet to print in black and white.
            // This causes background colors to be ignored in the PDF output.
            foreach (Worksheet ws in workbook.Worksheets)
            {
                ws.PageSetup.BlackAndWhite = true;
            }

            // Configure PDF save options (optional additional settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure that pages containing only styles are not printed,
                // which also helps to keep the layout minimal.
                PrintingPageType = PrintingPageType.IgnoreStyle
            };

            // Save the workbook to PDF using the configured options.
            workbook.Save("MinimalistOutput.pdf", pdfOptions);

            Console.WriteLine("Workbook exported to PDF with background colors ignored.");
        }
    }
}