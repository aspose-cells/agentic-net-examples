// Title: Read and log a cell's string value using zero‑based row and column indices with Aspose.Cells for .NET
// AI Prompts: Write C# code that retrieves the string value of a worksheet cell by specifying zero‑based row and column indices using Aspose.Cells. | Show how to output the retrieved cell value to the console and optionally save the workbook to an .xlsx file with Aspose.Cells. | Demonstrate inserting a value into a cell using zero‑based indices before reading it back in C# with Aspose.Cells.
// Common Searches: C# Aspose.Cells retrieve cell content using numeric row and column positions | Example of reading a worksheet cell with zero‑based coordinates in Aspose.Cells | How to output a cell's string value to console with Aspose.Cells .NET | Saving a workbook after cell manipulation using Aspose.Cells in C# | Insert and then read a cell value using index‑based access in Aspose.Cells
// Tags: zero-based cell access Aspose.Cells | read string from worksheet cell C# | console logging of cell value Aspose.Cells | save workbook as xlsx Aspose.Cells | insert cell value by index Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, writes "Sample Text" to the cell at zero‑based row 2, column 3, reads the cell's string value, prints it to the console, and saves the workbook as Output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Put a sample value into a cell (row 2, column 3) using zero‑based indices
        cells[2, 3].PutValue("Sample Text");

        // Access the same cell using zero‑based row and column indices
        int rowIndex = 2;      // zero‑based row index
        int columnIndex = 3;   // zero‑based column index
        Cell cell = cells[rowIndex, columnIndex];

        // Read the cell's value as a string
        string value = cell.StringValue;

        // Log the value to the console
        Console.WriteLine($"Value at row {rowIndex}, column {columnIndex}: {value}");

        // Save the workbook (optional)
        workbook.Save("Output.xlsx");
    }
}
