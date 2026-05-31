using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportPrintAreaToHtml
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate the worksheet with sample data
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    worksheet.Cells[row, col].PutValue($"Cell {row + 1},{col + 1}");
                }
            }

            // Define the print area (e.g., B2:F10)
            worksheet.PageSetup.PrintArea = "B2:F10";

            // Set HTML save options:
            // - ExportPrintAreaOnly = true  => only the defined print area is exported
            // - ExportWorksheetProperties = false => omit worksheet properties in the HTML output
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportPrintAreaOnly = true,
                ExportWorksheetProperties = false
            };

            // Determine output file path
            string outputPath = "PrintAreaOnly.html";

            // Ensure the directory exists (in case a relative path is used)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as an HTML file using the configured options
            workbook.Save(outputPath, htmlOptions);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}