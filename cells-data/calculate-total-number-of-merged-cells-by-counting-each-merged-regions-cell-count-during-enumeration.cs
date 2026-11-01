// Title: How to compute the total number of merged cells in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, iterates every worksheet, and returns the sum of cells covered by all merged ranges. | Create a reusable method in C# that accepts a Workbook object and calculates the total merged cell count by multiplying rows and columns of each CellArea.
// Common Searches: C# Aspose.Cells count all merged cells in a workbook | How to determine the number of merged cells across all worksheets using Aspose.Cells .NET | Compute cell count for each merged range with Aspose.Cells API | Calculate merged region cell total in Excel file using Aspose.Cells C# | Aspose.Cells enumerate merged cells and compute total cells
// Tags: Aspose.Cells merged region cell count | C# enumerate merged ranges | Workbook merged cell total calculation | CellArea size computation Aspose.Cells | Excel merged region aggregation .NET

using System;
using Aspose.Cells;

// Loads an Excel workbook, loops through each worksheet's MergedCells collection, computes the cell count for each merged region (rows × columns), aggregates the totals, and outputs the overall merged cell count.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        string filePath = "input.xlsx";
        Workbook workbook = new Workbook(filePath);

        // Variable to hold the total number of merged cells across all worksheets
        int totalMergedCellCount = 0;

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the collection of merged cells for the current worksheet
            var mergedCells = sheet.Cells.MergedCells;

            // Enumerate each merged region
            foreach (CellArea area in mergedCells)
            {
                // Calculate the number of cells in this merged region
                int rows = area.EndRow - area.StartRow + 1;
                int cols = area.EndColumn - area.StartColumn + 1;
                int cellCount = rows * cols;

                // Add to the running total
                totalMergedCellCount += cellCount;
            }
        }

        // Output the result
        Console.WriteLine($"Total number of merged cells: {totalMergedCellCount}");
    }
}
