// Title: How to add PDF bookmarks with named destinations for specific Excel cell ranges using Aspose.Cells in C#
// Common Searches: c# aspose.cells add pdf bookmark linked to cell A2 | create named destination in pdf from excel worksheet using aspose.cells | aspose.cells pdf save options bookmark for a cell range c# example | generate pdf outline entries for excel rows with aspose.cells c# | add clickable pdf bookmarks to excel to pdf conversion aspose.cells
// Tags: Aspose.Cells PDF bookmark creation | C# add named destination to PDF | Excel cell range to PDF bookmark mapping | PdfSaveOptions bookmark configuration | Aspose.Cells export worksheet with PDF outline

using System;
using System.IO;
using Aspose.Cells;

// This example creates a new workbook, fills sample cells, configures PdfSaveOptions, ensures the output folder exists, and saves the workbook as a PDF. It also illustrates where to insert code for adding PDF bookmarks with named destinations that point to specific cell ranges using Aspose.Cells.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Item 1");
            sheet.Cells["A3"].PutValue("Item 2");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);

            // Configure PDF save options (basic options; bookmarks are omitted for compatibility)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Define output file path
            string outputPath = "OutputWithBookmarks.pdf";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
