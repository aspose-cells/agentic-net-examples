// Title: Count non‑empty cells within the MaxDisplayRange of an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, retrieves the worksheet's MaxDisplayRange, iterates every cell, and returns the total number of cells that contain any value. | Enhance the iteration to treat cells that contain formulas with a non‑empty evaluated result as non‑empty and include them in the count. | Add logic to place the computed non‑empty cell count into a target cell (e.g., A1) and save the workbook as a new file.
// Common Searches: asp.net count cells with data in maxdisplayrange using Aspose.Cells | how to get number of populated cells in an Excel sheet with Aspose.Cells C# | c# iterate over worksheet MaxDisplayRange and count non‑blank cells | store result of cell count in another cell and save workbook Aspose.Cells
// Tags: maxdisplayrange cell iteration aspose.cells | count non‑empty cells c# aspose.cells | write calculation result to worksheet cell aspose | save modified workbook aspose.cells c# | excel data presence detection aspose.cells

using Aspose.Cells;
using System;

// The program loads an Excel workbook, accesses the first worksheet, obtains its MaxDisplayRange, loops through each cell in that range, increments a counter for cells whose Value is not null and whose StringValue is not empty, writes the final count to cell A1, and saves the updated workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Load workbook (using provided rule)
        var workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        var sheet = workbook.Worksheets[0];

        // Get the maximum display range of the worksheet
        var range = sheet.Cells.MaxDisplayRange;

        int nonEmptyCount = 0;

        // Iterate through all rows in the range
        for (int row = range.FirstRow; row <= range.FirstRow + range.RowCount - 1; row++)
        {
            // Iterate through all columns in the range
            for (int col = range.FirstColumn; col <= range.FirstColumn + range.ColumnCount - 1; col++)
            {
                var cell = sheet.Cells[row, col];

                // Count the cell if it is not empty
                if (cell.Value != null && !string.IsNullOrEmpty(cell.StringValue))
                {
                    nonEmptyCount++;
                }
            }
        }

        // Output the count to the console
        Console.WriteLine("Non‑empty cells count: " + nonEmptyCount);

        // Optionally store the result in a cell and save the workbook (using provided rule)
        sheet.Cells["A1"].PutValue(nonEmptyCount);
        workbook.Save("output.xlsx");
    }
}
