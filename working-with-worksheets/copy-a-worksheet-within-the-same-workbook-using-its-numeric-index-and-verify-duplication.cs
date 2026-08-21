// Title: Copy a Worksheet Using Its Numeric Index in Aspose.Cells (C#) and Verify Duplication
// Description: This C# example shows how to duplicate a worksheet inside the same workbook by referencing its zero‑based index with the AddCopy method, assign a new name to the clone, compare cell values to confirm an exact copy, and save the workbook as an Excel file.
// Keywords: Aspose.Cells | C# | .NET | AddCopy | duplicate sheet via index | verify worksheet data | rename sheet | Excel workbook manipulation | programmatic sheet copy | worksheet cloning
// Common Searches: Aspose.Cells duplicate sheet by index C# | AddCopy method example Aspose.Cells | how to copy a worksheet in Aspose.Cells | verify copied worksheet values Aspose | rename copied worksheet Aspose.Cells | save workbook after copying sheet .NET
// Developer Intent: The developer wants to programmatically clone an existing worksheet using its numeric position, give the clone a distinct name, and ensure the cloned sheet contains identical cell contents before persisting the workbook.
// Use Cases: Generate multiple report sections from a single template sheet while preserving the original layout. | Create a backup copy of a worksheet before applying data transformations, enabling easy rollback. | Automate workbook creation where the same sheet structure is reused with different data sets for each user. | Prepare Excel files that require identical worksheets for separate departments or regions.
// AI Prompts: Write C# code that copies a worksheet at index 0 using Aspose.Cells, renames the copy, and asserts that all cell values match the source. | Provide a robust method for cloning a worksheet by its numeric index with error handling and a report of any mismatched cells. | Generate a reusable function that returns the duplicated Worksheet object and prints a comparison summary of original versus copy.

using System;
using Aspose.Cells;

// This C# example shows how to duplicate a worksheet inside the same workbook by referencing its zero‑based index with the AddCopy method, assign a new name to the clone, compare cell values to confirm an exact copy, and save the workbook as an Excel file.
class Program
{
    static void Main()
    {
        // Create a new workbook (default contains one worksheet)
        Workbook workbook = new Workbook();

        // Access the first worksheet (index 0) and add sample data
        Worksheet originalSheet = workbook.Worksheets[0];
        originalSheet.Name = "Original";
        originalSheet.Cells["A1"].PutValue("Original Sheet");
        originalSheet.Cells["B2"].PutValue(12345);

        // Copy the worksheet using its numeric index (0)
        // AddCopy returns the index of the newly created copy
        int copiedIndex = workbook.Worksheets.AddCopy(0);
        Worksheet copiedSheet = workbook.Worksheets[copiedIndex];
        copiedSheet.Name = "CopyOfOriginal";

        // Verify that the copy contains the same data as the original
        Console.WriteLine("Original A1: " + originalSheet.Cells["A1"].StringValue);
        Console.WriteLine("Copied   A1: " + copiedSheet.Cells["A1"].StringValue);
        Console.WriteLine("Original B2: " + originalSheet.Cells["B2"].IntValue);
        Console.WriteLine("Copied   B2: " + copiedSheet.Cells["B2"].IntValue);

        // Save the workbook to demonstrate that both sheets exist
        workbook.Save("WorksheetCopyDemo.xlsx");
    }
}
