using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SetRowHeightsWithLoop
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Loop through rows 0 to 9 and set incremental heights
                // Example: start at 15 points and increase by 2 points for each subsequent row
                for (int rowIndex = 0; rowIndex < 10; rowIndex++)
                {
                    double height = 15.0 + (rowIndex * 2.0);
                    cells.SetRowHeight(rowIndex, height);
                }

                // Define output file path
                string outputPath = "RowHeightsLoopDemo.xlsx";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SetRowHeightsWithLoop.Run();
        }
    }
}