using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SplitPanesDemo
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

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (optional, just for demonstration)
                for (int row = 0; row < 20; row++)
                {
                    for (int col = 0; col < 10; col++)
                    {
                        worksheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                    }
                }

                // Enable split (parameterless overload)
                worksheet.Split();

                // Retrieve the pane collection to configure split positions
                PaneCollection panes = worksheet.GetPanes();

                // Set the first visible column of the right pane (zero‑based index)
                panes.FirstVisibleColumnOfRightPane = 3;

                // Set the first visible row of the bottom pane (zero‑based index)
                panes.FirstVisibleRowOfBottomPane = 4;

                // Define output file path
                string outputPath = "SplitPanesDemo.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }
}