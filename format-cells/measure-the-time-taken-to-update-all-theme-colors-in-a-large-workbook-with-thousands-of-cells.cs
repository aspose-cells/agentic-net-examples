// Title: How to measure the execution time of updating theme colors for every cell in a large Excel workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to change the foreground theme color of all used cells in a workbook and reports the elapsed time with Stopwatch. | Show a performance‑focused example that iterates only the used range of each worksheet, applies a solid red style, and prints the total seconds taken. | Provide a snippet that loads a large .xlsx file, updates cell styles in bulk, measures the operation duration, and saves the modified workbook.
// Common Searches: Aspose.Cells benchmark updating cell theme colors in large workbook .NET | C# measure time for bulk style changes across all worksheets in Excel file | How long does it take to apply a theme color to every cell using Aspose.Cells | Performance testing of iterating used range cells with Aspose.Cells in .NET
// Tags: bulk theme color update Aspose.Cells .NET | measure cell style modification performance | stopwatch timing Excel formatting Aspose.Cells | iterate used range worksheets Aspose.Cells | benchmark foreground color change large workbook

using System;
using System.Diagnostics;
using Aspose.Cells;
using System.Drawing;

// The example loads a large Excel workbook, iterates through the used range of each worksheet, sets each cell's foreground color to red via a modified Style, measures the total operation time with a Stopwatch, outputs the elapsed seconds, and saves the updated workbook.
class ThemeColorUpdateTimer
{
    static void Main()
    {
        // Load the existing workbook (replace with actual file path)
        string inputPath = "LargeWorkbook.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Start timing the theme color update operation
        Stopwatch stopwatch = new Stopwatch();
        stopwatch.Start();

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the maximum used row and column to limit the iteration
            int maxRow = sheet.Cells.MaxDataRow;
            int maxCol = sheet.Cells.MaxDataColumn;

            // Loop through each cell within the used range
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = sheet.Cells[row, col];
                    // Retrieve the current style
                    Style style = cell.GetStyle();

                    // Example: Change the foreground theme color to a new RGB value
                    // (Here we simply set it to a solid red color)
                    style.ForegroundColor = Color.Red;
                    style.Pattern = BackgroundType.Solid; // Ensure the color is applied

                    // Apply the modified style back to the cell
                    cell.SetStyle(style);
                }
            }
        }

        // Stop timing
        stopwatch.Stop();

        // Output the elapsed time
        Console.WriteLine($"Time taken to update theme colors: {stopwatch.Elapsed.TotalSeconds} seconds");

        // Save the modified workbook (replace with desired output path)
        string outputPath = "LargeWorkbook_Updated.xlsx";
        workbook.Save(outputPath);
    }
}
