using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SplitWorksheetVerticallyAtColumnFive
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Split the worksheet window (creates two panes)
                worksheet.Split();

                // Set the first visible column of the right pane to column index 5 (zero‑based)
                PaneCollection panes = worksheet.GetPanes();
                panes.FirstVisibleColumnOfRightPane = 5; // Splits vertically at column F (index 5)

                // Define output file path
                string outputPath = "SplitAtColumnFive.xlsx";

                // Ensure the directory exists before saving
                string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
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
            SplitWorksheetVerticallyAtColumnFive.Run();
        }
    }
}