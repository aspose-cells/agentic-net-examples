using System;
using System.IO;
using Aspose.Cells;

class ConvertWorkbookToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data
            for (int row = 0; row < 30; row++)
            {
                sheet.Cells[row, 0].PutValue($"Item {row + 1}");
                sheet.Cells[row, 1].PutValue((row + 1) * 5);
            }

            // Fit all columns to a single page width (height unlimited)
            sheet.PageSetup.SetFitToPages(1, 0);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false,
                AllColumnsInOnePagePerSheet = true
                // OptimizationType omitted for compatibility with older Aspose.Cells versions
            };

            // Save the workbook as a PDF
            string outputPath = "ConvertedWorkbook.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}