using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPrintDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            for (int i = 2; i <= 20; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Product {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Configure page setup to fit both width and height on a single page
            sheet.PageSetup.SetFitToPages(1, 1);
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Create print options
            ImageOrPrintOptions printOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            // Save the workbook as PDF (printing not supported in .NET Core)
            string pdfPath = "FitToPageWorkbook.pdf";
            workbook.Save(pdfPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF to \"{pdfPath}\" successfully.");

            // Save the workbook for reference
            workbook.Save("FitToPageWorkbook.xlsx");
        }
    }
}