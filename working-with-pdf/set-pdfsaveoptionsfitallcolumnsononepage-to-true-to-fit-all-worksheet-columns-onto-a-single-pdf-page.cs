using System;
using Aspose.Cells;

namespace AsposeCellsPdfDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            for (int col = 0; col < 30; col++)
            {
                sheet.Cells[0, col].PutValue($"Header {col + 1}");
                sheet.Cells[1, col].PutValue($"Data {col + 1}");
            }

            // Configure PDF save options to fit all columns on one page per sheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                AllColumnsInOnePagePerSheet = true
            };

            // Save the workbook as PDF with the configured options
            workbook.Save("output.pdf", pdfOptions);

            // Author note: This example demonstrates fitting all worksheet columns onto a single PDF page.
        }
    }
}