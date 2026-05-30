using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportPrintAreaHtmlDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Fill the worksheet with sample data
                for (int i = 0; i < 20; i++)
                {
                    for (int j = 0; j < 10; j++)
                    {
                        worksheet.Cells[i, j].PutValue($"Cell {i + 1},{j + 1}");
                    }
                }

                // Define the print area (e.g., B2:F10)
                worksheet.PageSetup.PrintArea = "B2:F10";

                // Set HTML save options
                HtmlSaveOptions options = new HtmlSaveOptions
                {
                    // Export only the defined print area
                    ExportPrintAreaOnly = true,
                    // Disable CSS generation; use only inline styles
                    DisableCss = true
                };

                // Output file path
                string outputPath = "PrintAreaOnly_NoCss.html";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML with the specified options
                workbook.Save(outputPath, options);
                Console.WriteLine($"HTML file saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportPrintAreaHtmlDemo.Run();
        }
    }
}