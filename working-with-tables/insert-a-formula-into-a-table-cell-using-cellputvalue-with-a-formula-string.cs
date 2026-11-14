// Title: Insert a SUM formula into a specific ListObject cell using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add a SUM formula to a cell inside an Excel table (ListObject) | how to set Cell.Formula for a ListObject column using Aspose.Cells .NET | programmatically write a formula to a table cell in an Excel workbook with Aspose.Cells | Aspose.Cells example for inserting formulas into ListObject rows in C# | set formula in a specific cell of a ListObject range using Aspose.Cells API
// Tags: Aspose.Cells ListObject cell formula assignment | C# insert SUM formula into Excel table cell | Cell.Formula property usage Aspose.Cells | programmatic Excel table formula with Aspose.Cells .NET | add formula to ListObject range C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a new workbook, adds a ListObject covering A1:B5, assigns the formula "=SUM(A2:A5)" to the cell at column B of the second table row using the Cell.Formula property, and saves the file as Output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define the range for the table (A1:B5)
            int firstRow = 0;      // Row index for A1
            int firstColumn = 0;   // Column index for A1
            int totalRows = 5;     // Number of rows in the table
            int totalColumns = 2;  // Number of columns in the table

            // Add a ListObject (table) to the defined range
            int tableIndex = sheet.ListObjects.Add(
                firstRow,
                firstColumn,
                firstRow + totalRows,
                firstColumn + totalColumns,
                true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable";

            // Insert a formula into a specific cell of the table
            // Example: put a SUM formula into the cell at column B, second row of the table
            int targetRow = table.StartRow + 1;        // second row of the table (0‑based index)
            int targetColumn = table.StartColumn + 1; // column B within the table
            Cell formulaCell = sheet.Cells[targetRow, targetColumn];
            formulaCell.Formula = "=SUM(A2:A5)";

            // Define output path
            string outputPath = "Output.xlsx";

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
