using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SplitPanesDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Add sample data to visualize the split
                for (int i = 0; i < 30; i++)
                {
                    for (int j = 0; j < 15; j++)
                    {
                        worksheet.Cells[i, j].Value = $"R{i + 1}C{j + 1}";
                    }
                }

                // Split the worksheet window
                worksheet.Split();

                // Configure split positions
                PaneCollection panes = worksheet.GetPanes();
                panes.FirstVisibleRowOfBottomPane = 10;   // rows 0‑9 stay in the top pane
                panes.FirstVisibleColumnOfRightPane = 5; // columns 0‑4 stay in the left pane

                // Define output file
                string outputPath = "SplitPanesDemo.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
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
            SplitPanesDemo.Run();
        }
    }
}