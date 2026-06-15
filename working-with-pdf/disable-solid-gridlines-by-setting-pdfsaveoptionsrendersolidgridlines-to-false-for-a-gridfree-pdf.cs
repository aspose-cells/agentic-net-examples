using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsGridlineDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data
                sheet.Cells["A1"].PutValue("Demo");
                sheet.Cells["B2"].PutValue(123);
                sheet.Cells["C3"].PutValue(DateTime.Now);

                // Ensure gridlines are visible (required for PDF rendering)
                sheet.IsGridlinesVisible = true;

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Render only dotted (hair) gridlines
                    GridlineType = GridlineType.Dotted
                };

                // Define output file path
                string outputPath = "GridlinesDisabled.pdf";

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = Directory.GetCurrentDirectory();
                }
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine("PDF saved successfully without solid gridlines.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}