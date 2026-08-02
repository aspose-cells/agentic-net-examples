// Title: Aspose.Cells for .NET (C#) – Compute Total Merged Cells by Summing Each Merged Region
// Description: This C# sample creates a workbook, merges three ranges (A1:B2, D4:E6, G1:G3), retrieves all merged areas with Cells.GetMergedAreas(), determines the row and column span of each CellArea, multiplies them to get the cell count per region, aggregates the counts to obtain the overall merged‑cell total, prints the results, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | merged cells count | GetMergedAreas | CellArea | Excel merge ranges | calculate merged cells | worksheet automation | Excel file processing
// Common Searches: Aspose.Cells count merged cells C# | How to sum cells in merged areas using Aspose.Cells | Get total merged cell number in .NET workbook | Calculate merged region size with Aspose.Cells | C# code to enumerate merged ranges in Excel
// Developer Intent: Determine how many individual cells belong to merged blocks in an Excel worksheet.
// Use Cases: Enforce a maximum merged‑cell threshold before converting a sheet to PDF. | Produce a diagnostic report that lists each merged block’s dimensions and the cumulative merged‑cell count. | Adjust page‑layout or printing parameters based on the total merged cells to avoid unexpected page breaks.
// AI Prompts: Generate a reusable C# method that returns the total merged‑cell count for any Worksheet object. | Provide code that safely handles worksheets with no merged areas or a null result from GetMergedAreas. | Show how to log the start row, end row, start column, and end column of each merged region while computing the total.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# sample creates a workbook, merges three ranges (A1:B2, D4:E6, G1:G3), retrieves all merged areas with Cells.GetMergedAreas(), determines the row and column span of each CellArea, multiplies them to get the cell count per region, aggregates the counts to obtain the overall merged‑cell total, prints the results, and saves the file.
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

                // Calculate total number of merged cells by summing each area's cell count
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

                // Save the workbook (optional, just to visualize the merged cells)
                string outputPath = "MergedCellsCountDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
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
