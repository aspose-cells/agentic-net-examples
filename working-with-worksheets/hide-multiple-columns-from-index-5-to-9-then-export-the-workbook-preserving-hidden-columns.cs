using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class HideColumnsAndExport
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate sample data (10 rows x 12 columns)
                for (int row = 0; row < 10; row++)
                {
                    for (int col = 0; col < 12; col++)
                    {
                        cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Hide columns with zero‑based indexes 5 through 9 (5 columns total)
                int startColumn = 5;      // Column F
                int columnCount = 5;      // Columns F, G, H, I, J
                cells.HideColumns(startColumn, columnCount);

                // Export to HTML while preserving hidden columns (they remain hidden in the output)
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Ensure hidden columns are kept hidden rather than removed
                    HiddenColDisplayType = HtmlHiddenColDisplayType.Hidden
                };

                string outputPath = "HiddenColumnsPreserved.html";

                // Save the workbook with the specified options
                workbook.Save(outputPath, htmlOptions);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
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
            HideColumnsAndExport.Run();
        }
    }
}