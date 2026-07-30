// Title: Copy a Worksheet Between Workbooks and Freeze the Header Row with Aspose.Cells for .NET (C#)
// Description: Shows how to load a source workbook, copy its first worksheet into a new workbook, apply FreezePanes to lock the top row as a header, and save the file as XLSX using Aspose.Cells for .NET.
// Keywords: Aspose.Cells copy worksheet C# | Aspose.Cells FreezePanes C# | copy sheet between workbooks .NET | freeze top row Aspose.Cells | C# Excel worksheet duplication | Aspose.Cells example | Excel FreezePanes C#
// Common Searches: Aspose.Cells copy worksheet to another workbook | How to freeze the first row after copying a sheet with Aspose.Cells | C# Aspose.Cells FreezePanes example | Copy sheet and preserve frozen panes Aspose.Cells | Aspose.Cells duplicate worksheet and set FreezePanes
// Developer Intent: Copy a worksheet from one workbook to another and lock its first row with FreezePanes.
// Use Cases: Create a template file, import a data sheet from an external workbook, and freeze the header for easy scrolling. | Merge multiple source workbooks into a consolidated report while keeping each sheet’s frozen header rows. | Automate worksheet duplication in generated Excel files, ensuring consistent UI by preserving frozen panes.
// AI Prompts: Generate C# code using Aspose.Cells to copy a worksheet from a source workbook to a new workbook and freeze the top row. | Explain the parameters of the FreezePanes method for locking only the header row after copying a sheet with Aspose.Cells. | Provide an example that copies several worksheets and applies FreezePanes to the first row of each using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to load a source workbook, copy its first worksheet into a new workbook, apply FreezePanes to lock the top row as a header, and save the file as XLSX using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the source workbook from a file
        Workbook sourceWorkbook = new Workbook("source.xlsx");
        // Get the worksheet you want to copy (first worksheet in this example)
        Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

        // Create a new (empty) destination workbook
        Workbook destinationWorkbook = new Workbook();
        // Remove the default sheet that comes with a new workbook
        destinationWorkbook.Worksheets.Clear();

        // Add a new worksheet to the destination workbook with the same name as the source
        Worksheet destinationSheet = destinationWorkbook.Worksheets.Add(sourceSheet.Name);
        // Copy the contents and formatting from the source worksheet
        destinationSheet.Copy(sourceSheet);

        // Freeze the header row (first row). 
        // Parameters: row index, column index, number of frozen rows, number of frozen columns.
        // To freeze only the first row: row = 1 (zero‑based index of the row below the freeze line),
        // column = 0, frozenRows = 1, frozenColumns = 0.
        destinationSheet.FreezePanes(1, 0, 1, 0);

        // Save the resulting workbook
        destinationWorkbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
