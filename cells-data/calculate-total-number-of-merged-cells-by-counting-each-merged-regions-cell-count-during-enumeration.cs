// Title: How to calculate the total number of merged cells in an Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to list all merged ranges in a worksheet and returns the sum of cells covered by those ranges. | Show how to call Cells.GetMergedAreas, compute each area’s row and column span, and aggregate the cell count with Aspose.Cells. | Create a method that prints both the number of merged regions and the combined merged‑cell total for a workbook.
// Common Searches: C# Aspose.Cells count merged cells in a worksheet | total cells covered by merged ranges using Aspose.Cells .NET | how to sum merged area sizes with Aspose.Cells API | retrieve merged areas and calculate merged cell count Aspose.Cells | Aspose.Cells GetMergedAreas example for merged cell total
// Tags: Aspose.Cells GetMergedAreas merged cell count | calculate merged region size C# | enumerate merged areas workbook .NET | sum cells of merged ranges Aspose | merged cells total calculation Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The example creates a workbook, merges several ranges, retrieves all merged areas via Cells.GetMergedAreas, iterates each CellArea to compute its row and column span, aggregates these values to obtain the total merged cell count, outputs both the number of merged regions and the combined cell total, and saves the workbook.
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

                // Merge a few sample ranges
                // A1:B2  (2 rows x 2 columns)
                cells.Merge(0, 0, 2, 2);
                // D4:E6  (3 rows x 2 columns)
                cells.Merge(3, 3, 3, 2);
                // G1:G3  (3 rows x 1 column)
                cells.Merge(0, 6, 3, 1);

                // Retrieve all merged areas
                CellArea[] mergedAreas = cells.GetMergedAreas();

                // Calculate total number of cells covered by merged regions
                int totalMergedCellCount = 0;
                foreach (CellArea area in mergedAreas)
                {
                    int rows = area.EndRow - area.StartRow + 1;
                    int cols = area.EndColumn - area.StartColumn + 1;
                    totalMergedCellCount += rows * cols;
                }

                // Output the result
                Console.WriteLine($"Number of merged areas: {mergedAreas.Length}");
                Console.WriteLine($"Total merged cells (by counting each region's cells): {totalMergedCellCount}");

                // Save the workbook (optional, just to demonstrate lifecycle)
                workbook.Save("MergedCellsCountDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            MergedCellsCountDemo.Run();
        }
    }
}
