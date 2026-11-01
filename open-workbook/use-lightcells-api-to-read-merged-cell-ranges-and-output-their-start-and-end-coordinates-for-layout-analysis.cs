// Title: Read merged cell ranges from an Excel worksheet using Aspose.Cells LightCells API in C# and output their start and end coordinates
// Common Searches: Aspose.Cells LightCells example to read merged cells in C# | How to enumerate merged ranges with LightCells API | C# code to list merged cell start and end rows and columns from an .xlsx file | Retrieve merged cell areas for layout analysis using Aspose.Cells | Extract merged cell boundaries from the first worksheet with LightCells
// Tags: LightCells merged area extraction | Aspose.Cells merged range coordinates | C# enumerate merged cells Excel | Excel merged cell boundaries Aspose | layout analysis with LightCells merged cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel file, accesses the first worksheet, uses Cells.GetMergedAreas() via the LightCells API to obtain all merged cell ranges, converts zero‑based indices to 1‑based Excel coordinates, and prints each range's start and end row/column positions.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be analyzed
        string filePath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(filePath);

            // Index of the worksheet to inspect (0 = first worksheet)
            int sheetIndex = 0;

            // Ensure the worksheet index is within range
            if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
            {
                Console.WriteLine($"Error: Worksheet index {sheetIndex} is out of range.");
                return;
            }

            // Retrieve the worksheet
            Worksheet sheet = workbook.Worksheets[sheetIndex];

            // Get all merged cell ranges in the specified worksheet (new API)
            CellArea[] mergedRanges = sheet.Cells.GetMergedAreas();

            // Output the start and end coordinates of each merged range
            foreach (CellArea area in mergedRanges)
            {
                // Convert zero‑based indices to 1‑based Excel coordinates for readability
                int startRow = area.StartRow + 1;
                int startColumn = area.StartColumn + 1;
                int endRow = area.EndRow + 1;
                int endColumn = area.EndColumn + 1;

                Console.WriteLine($"Merged Range: Start=({startRow},{startColumn}) End=({endRow},{endColumn})");
            }
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
