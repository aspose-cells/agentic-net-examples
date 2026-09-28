// Title: Enumerate initialized cells that lack a value and add a "Missing value" comment with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that iterates through all initialized cells in a worksheet, records the addresses of cells where the Value property is null, and inserts a comment "Missing value" into each of those cells. | Modify the example to export the list of empty‑cell addresses to a CSV file instead of writing them to the console, while still adding comments to the cells.
// Common Searches: aspocells c# enumerate initialized cells with null value | how to add comment to empty cells using Aspose.Cells .NET | list addresses of cells without data in an Excel worksheet Aspose.Cells | detect and flag cells that have no value in a workbook with Aspose.Cells
// Tags: iterate over initialized cells Aspose.Cells | detect null cell values Aspose.Cells C# | annotate empty cells Aspose.Cells | write empty cell list to CSV Aspose.Cells | save workbook with comments Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The sample loads an existing workbook, iterates over every initialized cell in the first worksheet, captures the addresses of cells whose Value property is null, adds a "Missing value" comment to each, outputs the addresses, and saves the workbook with the new comments.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // List to collect addresses of initialized cells that have no value
        List<string> cellsWithoutValue = new List<string>();

        // Enumerate all initialized cells in the worksheet
        foreach (Cell cell in worksheet.Cells)
        {
            // Check if the cell's Value property is null (i.e., no value assigned)
            if (cell.Value == null)
            {
                // Flag the cell by storing its address
                cellsWithoutValue.Add(cell.Name);

                // Optionally add a comment to the cell for visual identification
                int commentIndex = worksheet.Comments.Add(cell.Row, cell.Column);
                Comment comment = worksheet.Comments[commentIndex];
                comment.Note = "Missing value";
            }
        }

        // Output the flagged cell addresses for further analysis
        Console.WriteLine("Initialized cells lacking a value:");
        foreach (string address in cellsWithoutValue)
        {
            Console.WriteLine(address);
        }

        // Save the workbook (the comments will be persisted)
        workbook.Save("output.xlsx");
    }
}
