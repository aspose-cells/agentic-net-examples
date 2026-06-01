using System;
using Aspose.Cells;

namespace AsposeCellsOnePagePerSheetDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook with two worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";

            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");

            // Populate some sample data in each sheet
            for (int i = 0; i < 10; i++)
            {
                sheet1.Cells[i, 0].PutValue($"FirstSheet Row {i + 1}");
                sheet2.Cells[i, 0].PutValue($"SecondSheet Row {i + 1}");
            }

            // Create PDF save options and enable OnePagePerSheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true // Each worksheet will be saved as a separate PDF page
            };

            // Save the workbook as PDF using the specified options
            workbook.Save("OnePagePerSheetOutput.pdf", pdfOptions);
        }
    }
}