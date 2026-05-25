using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class MergedCellsCountDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Merge sample ranges
                cells.Merge(0, 0, 2, 2); // A1:B2
                cells.Merge(3, 3, 3, 2); // D4:E6
                cells.Merge(9, 6, 1, 1); // G10 (single cell)

                // Retrieve merged areas
                CellArea[] mergedAreas = cells.GetMergedAreas();

                // Calculate total cells covered by merged regions
                int totalMergedCellCount = 0;
                foreach (CellArea area in mergedAreas)
                {
                    int rows = area.EndRow - area.StartRow + 1;
                    int cols = area.EndColumn - area.StartColumn + 1;
                    totalMergedCellCount += rows * cols;
                }

                // Output results
                Console.WriteLine($"Number of merged areas: {mergedAreas.Length}");
                Console.WriteLine($"Total cells covered by merged regions: {totalMergedCellCount}");

                // Save workbook
                string outputPath = "MergedCellsCountDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            MergedCellsCountDemo.Run();
        }
    }
}