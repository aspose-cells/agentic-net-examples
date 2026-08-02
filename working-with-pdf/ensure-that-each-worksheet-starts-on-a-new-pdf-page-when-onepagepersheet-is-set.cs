using System;
using Aspose.Cells;

namespace AsposeCellsPdfDemo
{
    // Author: Aspose.Cells .NET example – ensures each worksheet starts on a new PDF page
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet and add sample data
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "FirstSheet";
            for (int i = 0; i < 20; i++)
                sheet1.Cells[i, 0].PutValue($"Sheet1 Row {i + 1}");

            // Add a second worksheet and populate it
            Worksheet sheet2 = workbook.Worksheets.Add("SecondSheet");
            for (int i = 0; i < 20; i++)
                sheet2.Cells[i, 0].PutValue($"Sheet2 Row {i + 1}");

            // Configure PDF save options to start each worksheet on a new page
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true // ensures each sheet is rendered on its own page
            };

            // Save the workbook to PDF (lifecycle: save)
            workbook.Save("WorksheetsPerPage.pdf", pdfOptions);

            Console.WriteLine("PDF saved successfully with each worksheet on a separate page.");
        }
    }
}