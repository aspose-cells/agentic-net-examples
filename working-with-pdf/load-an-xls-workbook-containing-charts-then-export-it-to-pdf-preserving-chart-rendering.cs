using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsChartToPdf
{
    class Program
    {
        static void Main()
        {
            // Load the existing Excel workbook that contains charts.
            // The constructor with a string parameter follows the provided load rule.
            Workbook workbook = new Workbook("input.xls");

            // Create PDF save options and enable chart cache refresh.
            // This ensures that charts are rendered correctly in the PDF.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                RefreshChartCache = true
            };

            // Save the workbook as a PDF file using the save rule that accepts
            // a filename and SaveOptions.
            workbook.Save("output.pdf", pdfOptions);

            Console.WriteLine("Workbook with charts successfully exported to PDF.");
        }
    }
}