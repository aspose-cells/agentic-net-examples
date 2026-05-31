using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class ExportPrintAreaToHtml
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
                        worksheet.Cells[row, col].PutValue($"Cell {row + 1},{col + 1}");
                    }
                }

                // Define the print area that should be exported
                worksheet.PageSetup.PrintArea = "B2:F10";

                // Configure HTML save options to export only the defined print area
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    ExportPrintAreaOnly = true, // Export only the print area
                    ExportGridLines = true      // Include grid lines in the output (optional)
                };

                // Save the workbook as HTML; only the print area will be present in the file
                workbook.Save("PrintAreaOnly.html", htmlOptions);
                Console.WriteLine("HTML file 'PrintAreaOnly.html' created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportPrintAreaToHtml.Run();
        }
    }
}