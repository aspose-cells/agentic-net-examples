using System;
using Aspose.Cells;

namespace AsposeCellsOnePagePerSheetDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook with default first worksheet
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            for (int i = 0; i < 20; i++)
            {
                sheet1.Cells[i, 0].PutValue($"Row {i + 1} - Sheet1");
            }

            // Add a second worksheet and add sample data
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            for (int i = 0; i < 30; i++)
            {
                sheet2.Cells[i, 0].PutValue($"Row {i + 1} - Sheet2");
            }

            // Configure PDF save options to output each sheet on a separate page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true // Ensures each worksheet is rendered on its own page
            };

            // Save the workbook as a PDF file using the configured options
            workbook.Save("Workbook_OnePagePerSheet.pdf", pdfOptions);
        }
    }
}