// Title: Read a cell value with zero‑based row/column indices using Aspose.Cells for .NET
// Description: Demonstrates how to insert a string into cell B2 by addressing it with zero‑based row 1 and column 2, retrieve the cell's StringValue, output the value and cell name to the console, and optionally save the workbook as Output.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | .NET | zero based cell access | read cell value | Console.WriteLine | Workbook.Save | cell.StringValue | C# spreadsheet API | index based cell reference | Aspose.Cells tutorial
// Common Searches: Aspose.Cells read cell by row and column index | C# zero based cell address Aspose.Cells | How to get cell value in Aspose.Cells .NET | Log spreadsheet cell content to console | Save workbook after reading cell Aspose.Cells
// Developer Intent: Get the value of a specific spreadsheet cell using zero‑based row and column numbers and display it in a console application.
// Use Cases: Quickly verify data entry by printing a cell's content during development. | Extract configuration settings stored in a known cell location before processing larger datasets. | Debug spreadsheet transformations by logging individual cell values and then persisting changes.
// AI Prompts: Show C# code to read a cell using zero‑based indices with Aspose.Cells and handle null or empty values. | Provide an example that accesses cell B2, prints its text, and saves the workbook. | Explain how to convert the retrieved StringValue to int, double, DateTime, and handle conversion errors with Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to insert a string into cell B2 by addressing it with zero‑based row 1 and column 2, retrieve the cell's StringValue, output the value and cell name to the console, and optionally save the workbook as Output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Put a sample value into a cell using zero‑based indices (row 1, column 2 => B2)
        cells[1, 2].PutValue("Sample Text");

        // Access the same cell using zero‑based row and column indices
        Cell cell = cells[1, 2];

        // Read the cell's value as a string
        string value = cell.StringValue;

        // Log the value to the console
        Console.WriteLine($"Value at row 1, column 2 (cell {cell.Name}): {value}");

        // Save the workbook (optional)
        workbook.Save("Output.xlsx");
    }
}
