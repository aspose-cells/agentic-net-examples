using System;
using Aspose.Cells;

namespace AsposeCellsOnePagePerSheetDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add two worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";

            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");

            // Populate each sheet with some sample data
            for (int i = 0; i < 20; i++)
            {
                sheet1.Cells[i, 0].PutValue($"FirstSheet Row {i + 1}");
                sheet2.Cells[i, 0].PutValue($"SecondSheet Row {i + 1}");
            }

            // Create PDF save options and force each worksheet onto a separate PDF page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook as a PDF using the configured options
            workbook.Save("OnePagePerSheetOutput.pdf", pdfOptions);
        }
    }
}