using System;
using Aspose.Cells;

namespace AsposeCellsPdfFitColumns
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data (optional, just for demonstration)
            for (int col = 0; col < 30; col++)
            {
                sheet.Cells[0, col].PutValue($"Header {col + 1}");
                sheet.Cells[1, col].PutValue($"Data {col + 1}");
            }

            // Configure PDF save options to fit all columns on a single page per sheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // This property forces all column content of each sheet onto one PDF page
                AllColumnsInOnePagePerSheet = true
            };

            // Save the workbook as PDF using the configured options
            workbook.Save("AllColumnsOnePage.pdf", pdfOptions);
        }
    }
}
// Author: Aspose.Cells .NET example – fits all columns onto one PDF page per sheet.