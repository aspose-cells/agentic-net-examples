using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomPaperPdf
{
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Header");
                sheet.Cells["A2"].PutValue("Row 1");
                sheet.Cells["A3"].PutValue("Row 2");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(123);
                sheet.Cells["B3"].PutValue(456);

                // Set custom paper size (width = 5 inches, height = 3 inches)
                sheet.PageSetup.CustomPaperSize(5.0, 3.0);

                // Fit the worksheet to a single page
                sheet.PageSetup.FitToPagesWide = 1;
                sheet.PageSetup.FitToPagesTall = 1;

                // Define output path
                string outputPath = "CustomPaperWorksheet.pdf";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save as PDF
                workbook.Save(outputPath, SaveFormat.Pdf);

                Console.WriteLine($"Worksheet saved to PDF with custom paper size: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}