using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DisableAutoStyleAndCleanup
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
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // NOTE: In newer Aspose.Cells versions the property to disable automatic style creation
                // may not be available. The default behavior is acceptable for this example.

                // Add sample data with explicit styles
                Worksheet sheet = workbook.Worksheets[0];
                for (int i = 0; i < 5; i++)
                {
                    Cell cell = sheet.Cells[i, 0];
                    cell.PutValue($"Item {i + 1}");

                    // Create a distinct style for each cell
                    Style style = workbook.CreateStyle();
                    style.Font.Name = "Arial";
                    style.Font.Size = 10 + i;
                    style.Font.IsBold = i % 2 == 0;
                    cell.SetStyle(style);
                }

                // Delete rows to leave some styles unused
                sheet.Cells.DeleteRows(3, 2);

                // Remove all unused styles
                workbook.RemoveUnusedStyles();

                // Prepare output path
                string outputPath = "CleanedWorkbook.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Run error: {ex.Message}");
            }
        }
    }
}