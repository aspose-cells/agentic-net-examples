using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportPrintAreaWithNoCss
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (optional, just for demonstration)
                for (int row = 0; row < 20; row++)
                {
                    for (int col = 0; col < 10; col++)
                    {
                        worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Define the print area that should be exported
                worksheet.PageSetup.PrintArea = "B2:F10";

                // Configure HTML save options:
                // - ExportPrintAreaOnly = true  => only the defined print area is saved
                // - DisableCss = true           => use only inline styles, no external CSS
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportPrintAreaOnly = true,
                    DisableCss = true
                };

                string outputPath = "PrintArea_NoCss.html";

                // Save the workbook as HTML using the configured options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any runtime errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportPrintAreaWithNoCss.Run();
        }
    }
}