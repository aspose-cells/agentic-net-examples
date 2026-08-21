// Title: Aspose.Cells for .NET – Compute total merged cells in a worksheet
// Description: This C# sample creates a workbook, merges three ranges, retrieves all merged areas via GetMergedAreas, calculates each area’s row and column span, multiplies them to get the cell count per region, sums the counts, prints the total, and saves the file.
// Keywords: Aspose.Cells merged cells count | GetMergedAreas .NET | CellArea row span | calculate merged region size | C# Excel merged cells | Aspose.Cells workbook merge enumeration | total merged cells worksheet
// Common Searches: How to count merged cells with Aspose.Cells .NET | Get total merged cell count in Excel using Aspose | Aspose.Cells enumerate merged areas | C# count cells in merged ranges | Sum merged cells Aspose.Cells example
// Developer Intent: Show how to determine the aggregate number of cells that belong to merged ranges in an Aspose.Cells worksheet.
// Use Cases: Validate that merged regions stay within size limits before publishing a workbook. | Create a summary report that lists each merged area's dimensions and the overall merged cell total. | Adjust layout logic or formula references based on the cumulative count of merged cells.
// AI Prompts: Write a reusable method that returns the total merged cell count for any Worksheet using Aspose.Cells. | Add robust error handling to compute merged cell totals when a workbook contains no merged areas or corrupted data. | Generate console‑output code that logs each merged region’s row/column span and the running total of merged cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# sample creates a workbook, merges three ranges, retrieves all merged areas via GetMergedAreas, calculates each area’s row and column span, multiplies them to get the cell count per region, sums the counts, prints the total, and saves the file.
    public class MergedCellsCountDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Merge some sample ranges
                // A1:B2  (2 rows x 2 columns)
                cells.Merge(0, 0, 2, 2);
                // D4:E6  (3 rows x 2 columns)
                cells.Merge(3, 3, 3, 2);
                // G1:G3  (3 rows x 1 column)
                cells.Merge(0, 6, 3, 1);

                // Retrieve all merged areas
                CellArea[] mergedAreas = cells.GetMergedAreas();

                // Calculate total number of merged cells by summing each area's cell count
                int totalMergedCellCount = 0;
                foreach (CellArea area in mergedAreas)
                {
                    int rows = area.EndRow - area.StartRow + 1;
                    int cols = area.EndColumn - area.StartColumn + 1;
                    totalMergedCellCount += rows * cols;
                }

                // Output the result
                Console.WriteLine($"Total merged cells count: {totalMergedCellCount}");

                // Save the workbook (lifecycle rule: save)
                workbook.Save("MergedCellsCountDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            MergedCellsCountDemo.Run();
        }
    }
}
