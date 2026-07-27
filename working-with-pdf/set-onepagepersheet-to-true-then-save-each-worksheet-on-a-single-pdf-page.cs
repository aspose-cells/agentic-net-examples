using System;
using Aspose.Cells;

namespace AsposeCellsOnePagePerSheetDemo
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data to two worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            for (int i = 0; i < 20; i++)
            {
                sheet1.Cells[i, 0].PutValue($"Sheet1 Row {i + 1}");
                sheet2.Cells[i, 0].PutValue($"Sheet2 Row {i + 1}");
            }

            // Configure PDF save options to force each sheet onto a single page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook as a PDF; each worksheet will occupy one page
            workbook.Save("WorkbookOnePagePerSheet.pdf", pdfOptions);
        }
    }
}