using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class DisableAutoStyleAndCleanup
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook instance
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data with different styles to generate a style pool
                for (int i = 0; i < 10; i++)
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

                // Delete some rows to leave unused styles in the pool
                sheet.Cells.DeleteRows(5, 5);

                // NOTE: In older Aspose.Cells versions the property
                // WorkbookSettings.EnableAutomaticStyleCreation does not exist.
                // The cleanup works without explicitly disabling it.

                // Remove all styles that are no longer referenced by any cell
                workbook.RemoveUnusedStyles();

                // Save the cleaned workbook
                string outputPath = "CleanedWorkbook.xlsx";

                // Ensure the directory exists before saving
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            DisableAutoStyleAndCleanup.Run();
        }
    }
}